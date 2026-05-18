import { useMemo, useRef, useState } from 'react';
import { executeSingleRow, fetchOAuthToken } from './lib/requestRunner';
import { extractEndpoints, generateTestSuite, parseDefinition, resolveServerFromDefinition } from './lib/openApi';
import { exportTestSuiteToExcel, importTestSuiteFromExcel } from './lib/excelIO';

const INITIAL_OAUTH = {
  tokenUrl: '',
  clientId: '',
  clientSecret: '',
  scope: '',
};

function responseClass(code) {
  const status = Number(code);
  if (!status) {
    return '';
  }

  if (status >= 200 && status < 300) {
    return 'ok';
  }

  return 'fail';
}

function getFirstBasePath(path) {
  const segments = String(path || '')
    .split('/')
    .filter(Boolean);

  if (segments.length === 0) {
    return '/';
  }

  return `/${segments[0]}`;
}

function includesFilter(endpoint, filterValue) {
  if (!filterValue) {
    return true;
  }

  const haystack = `${endpoint.path} ${endpoint.methods.join(' ')}`.toLowerCase();
  return haystack.includes(filterValue);
}

function App() {
  const excelInputRef = useRef(null);
  const [definitionText, setDefinitionText] = useState('');
  const [definition, setDefinition] = useState(null);
  const [fileName, setFileName] = useState('No file selected');
  const [parseError, setParseError] = useState('');
  const [status, setStatus] = useState('Choose a Swagger/OpenAPI JSON file to begin.');

  const [config, setConfig] = useState({
    protocol: 'https',
    serverPath: '',
  });

  const [endpoints, setEndpoints] = useState([]);
  const [testSuite, setTestSuite] = useState([]);
  const [isRunning, setIsRunning] = useState(false);

  const [oauth, setOauth] = useState(INITIAL_OAUTH);
  const [authToken, setAuthToken] = useState('');

  const [endpointFilter, setEndpointFilter] = useState('');
  const [collapsedSections, setCollapsedSections] = useState({
    config: false,
    endpoints: false,
    tests: false,
    definitionJson: true,
  });
  const [collapsedTestGroups, setCollapsedTestGroups] = useState({});

  const endpointCount = endpoints.length;
  const testCount = testSuite.length;

  const groupedEndpoints = useMemo(() => {
    const normalizedFilter = endpointFilter.trim().toLowerCase();
    const groups = {};

    endpoints.forEach((item) => {
      if (!includesFilter(item, normalizedFilter)) {
        return;
      }

      const basePath = getFirstBasePath(item.path);
      if (!groups[basePath]) {
        groups[basePath] = [];
      }

      groups[basePath].push(item);
    });

    return Object.entries(groups)
      .map(([basePath, items]) => ({
        basePath,
        items,
      }))
      .sort((a, b) => a.basePath.localeCompare(b.basePath));
  }, [endpoints, endpointFilter]);

  const summary = useMemo(() => {
    const success = testSuite.filter((row) => {
      const code = Number(row.responseCode);
      return code >= 200 && code < 300;
    }).length;

    const failed = testSuite.filter((row) => {
      const code = Number(row.responseCode);
      return code && (code < 200 || code >= 300);
    }).length;

    return { success, failed };
  }, [testSuite]);

  const groupedTestSuite = useMemo(() => {
    const groups = {};

    testSuite.forEach((row) => {
      const basePath = getFirstBasePath(row.path);
      if (!groups[basePath]) {
        groups[basePath] = [];
      }

      groups[basePath].push(row);
    });

    return Object.entries(groups)
      .map(([basePath, rows]) => ({
        basePath,
        rows,
      }))
      .sort((a, b) => a.basePath.localeCompare(b.basePath));
  }, [testSuite]);

  async function onFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    const text = await file.text();
    setDefinitionText(text);
    setFileName(file.name);

    const parsed = parseDefinition(text);
    if (parsed.error) {
      setDefinition(null);
      setParseError(parsed.error);
      setStatus('Definition file has invalid JSON.');
      return;
    }

    setDefinition(parsed.parsed);
    setParseError('');

    const suggestedServer = resolveServerFromDefinition(parsed.parsed, config.serverPath);
    setConfig((prev) => ({
      ...prev,
      serverPath: suggestedServer,
    }));
    setStatus('Definition loaded. You can extract endpoints or generate the test suite.');
  }

  function clearExtractDefinition() {
    setEndpoints([]);
    setEndpointFilter('');
    setStatus('Endpoint extraction cleared.');
  }

  function clearTestSuite() {
    setTestSuite([]);
    setCollapsedTestGroups({});
    setStatus('Test suite cleared.');
  }

  function handleExtractEndpoints() {
    if (!definition) {
      setStatus('Load a valid definition before extracting endpoints.');
      return;
    }

    const extracted = extractEndpoints(definition);
    setEndpoints(extracted);
    setStatus(`Extracted ${extracted.length} endpoints.`);
  }

  function handleGenerateTestSuite() {
    if (!definition) {
      setStatus('Load a valid definition before generating tests.');
      return;
    }

    const rows = generateTestSuite(definition, config);
    setTestSuite(rows);
    setCollapsedTestGroups({});
    setStatus(`Generated ${rows.length} test rows.`);
  }

  function updateRow(rowId, key, value) {
    setTestSuite((prev) => prev.map((row) => (row.id === rowId ? { ...row, [key]: value } : row)));
  }

  function handleExportExcel() {
    if (testSuite.length === 0) {
      setStatus('No test rows available to export.');
      return;
    }

    exportTestSuiteToExcel(testSuite);
    setStatus(`Exported ${testSuite.length} rows to Excel.`);
  }

  function triggerImportExcel() {
    if (excelInputRef.current) {
      excelInputRef.current.click();
    }
  }

  async function handleImportExcel(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    try {
      const importedRows = await importTestSuiteFromExcel(file);
      if (importedRows.length === 0) {
        setStatus('Excel file is empty or does not contain supported columns.');
      } else {
        setTestSuite(importedRows);
        setStatus(`Imported ${importedRows.length} rows from ${file.name}.`);
      }
    } catch (error) {
      setStatus(`Excel import failed: ${error.message}`);
    } finally {
      event.target.value = '';
    }
  }

  function toggleSection(sectionName) {
    setCollapsedSections((prev) => ({
      ...prev,
      [sectionName]: !prev[sectionName],
    }));
  }

  function toggleTestGroup(basePath) {
    setCollapsedTestGroups((prev) => ({
      ...prev,
      [basePath]: !prev[basePath],
    }));
  }

  async function runSingle(rowId) {
    const row = testSuite.find((item) => item.id === rowId);
    if (!row) {
      return;
    }

    setStatus(`Running ${row.method} ${row.path}`);
    const updated = await executeSingleRow(row, authToken);
    setTestSuite((prev) => prev.map((item) => (item.id === rowId ? updated : item)));
    setStatus(`Completed ${row.method} ${row.path} with status ${updated.responseCode}.`);
  }

  async function runTestGroup(basePath) {
    const rows = groupedTestSuite.find((group) => group.basePath === basePath)?.rows || [];
    if (rows.length === 0) {
      setStatus(`No rows found for ${basePath}.`);
      return;
    }

    setIsRunning(true);
    setStatus(`Running ${rows.length} requests in ${basePath}...`);

    let nextRows = [...testSuite];
    for (let index = 0; index < rows.length; index += 1) {
      const current = rows[index];
      const updated = await executeSingleRow(current, authToken);
      nextRows = nextRows.map((item) => (item.id === current.id ? updated : item));
      setTestSuite([...nextRows]);
    }

    setIsRunning(false);
    setStatus(`Completed ${rows.length} requests in ${basePath}.`);
  }

  async function runAll() {
    if (testSuite.length === 0) {
      setStatus('Generate tests before running.');
      return;
    }

    setIsRunning(true);
    setStatus(`Running ${testSuite.length} requests...`);

    const nextRows = [...testSuite];
    for (let index = 0; index < nextRows.length; index += 1) {
      const row = nextRows[index];
      const updated = await executeSingleRow(row, authToken);
      nextRows[index] = updated;
      setTestSuite([...nextRows]);
    }

    setIsRunning(false);
    setStatus('All requests completed.');
  }

  async function handleGetToken() {
    if (!oauth.tokenUrl || !oauth.clientId || !oauth.clientSecret) {
      setStatus('Token URL, client ID, and client secret are required for OAuth2 token retrieval.');
      return;
    }

    try {
      setStatus('Requesting OAuth2 token...');
      const token = await fetchOAuthToken(oauth);
      setAuthToken(token);
      setStatus('OAuth2 token received. It will be used for requests unless Authorization header is set manually.');
    } catch (error) {
      setStatus(`Token request failed: ${error.message}`);
    }
  }

  return (
    <div className="app-shell">
      <div className="aurora" />
      <header className="hero">
        <h1>SwiftAPI - Simple, Fast and Dynamic API Testing solution</h1>
        <p>OpenAPI parser, test-suite generator, and API execution console merged into one tool.</p>
      </header>

      <section className="panel">
        <div className="panel-head">
          <h2>Configuration and Authentication</h2>
          <button
            className="ghost icon-toggle"
            onClick={() => toggleSection('config')}
            title={collapsedSections.config ? 'Expand section' : 'Minimize section'}
            aria-label={collapsedSections.config ? 'Expand section' : 'Minimize section'}
          >
            {collapsedSections.config ? '+' : '-'}
          </button>
        </div>

        {!collapsedSections.config ? (
          <div className="grid-two">
            <div>
              <h2>Definition Input</h2>
              <p className="muted">Selected file: {fileName}</p>
              <input type="file" accept=".json" onChange={onFileChange} />
              {parseError ? <p className="error">{parseError}</p> : null}

              <div className="inline-fields">
                <label>
                  Protocol
                  <select
                    value={config.protocol}
                    onChange={(e) => setConfig((prev) => ({ ...prev, protocol: e.target.value }))}
                  >
                    <option value="https">https</option>
                    <option value="http">http</option>
                  </select>
                </label>
                <label>
                  Server Path (without protocol)
                  <input
                    value={config.serverPath}
                    onChange={(e) => setConfig((prev) => ({ ...prev, serverPath: e.target.value }))}
                    placeholder="api.example.com/v1"
                  />
                </label>
              </div>

              <div className="action-row">
                <button onClick={handleExtractEndpoints}>Extract Endpoints</button>
                <button onClick={clearExtractDefinition} className="ghost">
                  Clear Extract
                </button>
                <button onClick={handleGenerateTestSuite}>Generate Test Suite</button>
                <button onClick={clearTestSuite} className="ghost">
                  Clear Tests
                </button>
              </div>
            </div>

            <div>
              <h2>OAuth2 Client Credentials</h2>
              <div className="inline-fields oauth">
                <label>
                  Token URL
                  <input
                    value={oauth.tokenUrl}
                    onChange={(e) => setOauth((prev) => ({ ...prev, tokenUrl: e.target.value }))}
                    placeholder="https://login.microsoftonline.com/.../token"
                  />
                </label>
                <label>
                  Client ID
                  <input
                    value={oauth.clientId}
                    onChange={(e) => setOauth((prev) => ({ ...prev, clientId: e.target.value }))}
                  />
                </label>
                <label>
                  Client Secret
                  <input
                    type="password"
                    value={oauth.clientSecret}
                    onChange={(e) => setOauth((prev) => ({ ...prev, clientSecret: e.target.value }))}
                  />
                </label>
                <label>
                  Scope
                  <input
                    value={oauth.scope}
                    onChange={(e) => setOauth((prev) => ({ ...prev, scope: e.target.value }))}
                    placeholder="api://app-id/.default"
                  />
                </label>
              </div>
              <div className="action-row">
                <button onClick={handleGetToken}>Get Token</button>
              </div>
              <p className="token-preview">Token loaded: {authToken ? `${authToken.slice(0, 24)}...` : 'none'}</p>
            </div>
          </div>
        ) : null}
      </section>

      <section className="panel">
        <div className="panel-head">
          <h2>Extracted Endpoints ({endpointCount})</h2>
          <button
            className="ghost icon-toggle"
            onClick={() => toggleSection('endpoints')}
            title={collapsedSections.endpoints ? 'Expand section' : 'Minimize section'}
            aria-label={collapsedSections.endpoints ? 'Expand section' : 'Minimize section'}
          >
            {collapsedSections.endpoints ? '+' : '-'}
          </button>
        </div>

        {!collapsedSections.endpoints ? (
          <>
            {endpointCount === 0 ? (
              <p className="muted">No endpoints extracted.</p>
            ) : (
              <>
                <div className="endpoint-toolbar">
                  <input
                    value={endpointFilter}
                    onChange={(e) => setEndpointFilter(e.target.value)}
                    placeholder="Filter by endpoint path or method"
                  />
                </div>

                {groupedEndpoints.length === 0 ? (
                  <p className="muted">No endpoints match the current filter.</p>
                ) : (
                  <div className="table-wrap endpoint-table-wrap">
                    <table className="endpoint-table endpoint-matrix-table">
                      <thead>
                        <tr>
                          <th>Method Name</th>
                          <th>GET</th>
                          <th>POST</th>
                          <th>DELETE</th>
                          <th>PUT</th>
                          <th>PATCH</th>
                        </tr>
                      </thead>
                      <tbody>
                        {groupedEndpoints.flatMap((group) => [
                          <tr key={`group:${group.basePath}`} className="group-row">
                            <td colSpan={6}>{group.basePath}</td>
                          </tr>,
                          ...group.items.map((item) => {
                            const hasMethod = (name) => item.methods.includes(name);
                            return (
                              <tr key={`${group.basePath}:${item.path}`}>
                                <td>{`"${item.path}"`}</td>
                                <td>{hasMethod('get') ? '\u2713' : ''}</td>
                                <td>{hasMethod('post') ? '\u2713' : ''}</td>
                                <td>{hasMethod('delete') ? '\u2713' : ''}</td>
                                <td>{hasMethod('put') ? '\u2713' : ''}</td>
                                <td>{hasMethod('patch') ? '\u2713' : ''}</td>
                              </tr>
                            );
                          }),
                        ])}
                      </tbody>
                    </table>
                  </div>
                )}
              </>
            )}
          </>
        ) : null}
      </section>

      <section className="panel">
        <div className="panel-head">
          <h2>Test Suite ({testCount})</h2>
          <button
            className="ghost icon-toggle"
            onClick={() => toggleSection('tests')}
            title={collapsedSections.tests ? 'Expand section' : 'Minimize section'}
            aria-label={collapsedSections.tests ? 'Expand section' : 'Minimize section'}
          >
            {collapsedSections.tests ? '+' : '-'}
          </button>
        </div>

        {!collapsedSections.tests ? (
          <>
            <div className="section-top">
              <div className="summary">
                <span className="ok">Success: {summary.success}</span>
                <span className="fail">Failed: {summary.failed}</span>
              </div>
              <button onClick={triggerImportExcel} className="ghost" disabled={isRunning}>
                Import Excel
              </button>
              <button onClick={handleExportExcel} className="ghost" disabled={isRunning || testSuite.length === 0}>
                Export Excel
              </button>
              <button onClick={runAll} disabled={isRunning || testSuite.length === 0}>
                {isRunning ? 'Running...' : 'Run All'}
              </button>
              <input
                ref={excelInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleImportExcel}
                style={{ display: 'none' }}
              />
            </div>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Method</th>
                    <th>Path</th>
                    <th>URL</th>
                    <th>Params</th>
                    <th>Param Values</th>
                    <th>Headers</th>
                    <th>Body</th>
                    <th>Status</th>
                    <th>Resp. Time (ms)</th>
                    <th>Response Body</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {groupedTestSuite.flatMap((group) => [
                    <tr key={`group:${group.basePath}`} className="group-row test-group-row">
                      <td colSpan={8}>{group.basePath}</td>
                      <td>{group.rows.length} APIs</td>
                      <td className="test-group-actions" colSpan={2}>
                        <button onClick={() => runTestGroup(group.basePath)} disabled={isRunning} className="ghost">
                          Run this group
                        </button>
                        <button
                          className="ghost icon-toggle"
                          onClick={() => toggleTestGroup(group.basePath)}
                          title={collapsedTestGroups[group.basePath] ? 'Expand group' : 'Minimize group'}
                          aria-label={collapsedTestGroups[group.basePath] ? 'Expand group' : 'Minimize group'}
                        >
                          {collapsedTestGroups[group.basePath] ? '+' : '-'}
                        </button>
                      </td>
                    </tr>,
                    ...(collapsedTestGroups[group.basePath]
                      ? []
                      : group.rows.map((row) => (
                          <tr key={row.id}>
                            <td>{row.method}</td>
                            <td>{row.path}</td>
                            <td>
                              <textarea
                                value={row.urlTemplate}
                                onChange={(e) => updateRow(row.id, 'urlTemplate', e.target.value)}
                              />
                            </td>
                            <td>
                              <textarea value={row.params} readOnly />
                            </td>
                            <td>
                              <textarea
                                value={row.paramValues}
                                onChange={(e) => updateRow(row.id, 'paramValues', e.target.value)}
                              />
                            </td>
                            <td>
                              <textarea
                                value={row.headerValues}
                                onChange={(e) => updateRow(row.id, 'headerValues', e.target.value)}
                              />
                            </td>
                            <td>
                              <textarea value={row.body} onChange={(e) => updateRow(row.id, 'body', e.target.value)} />
                            </td>
                            <td className={responseClass(row.responseCode)}>{row.responseCode}</td>
                            <td>{row.responseTimeMs}</td>
                            <td>
                              <textarea value={row.responseBody} readOnly />
                            </td>
                            <td>
                              <button onClick={() => runSingle(row.id)} disabled={isRunning}>
                                Run
                              </button>
                            </td>
                          </tr>
                        ))),
                  ])}
                </tbody>
              </table>
            </div>
          </>
        ) : null}
      </section>

      <footer className="status-bar">
        <strong>Status:</strong> {status}
      </footer>

      <section className="panel">
        <div className="panel-head">
          <h2>Loaded Definition (JSON)</h2>
          <button
            className="ghost icon-toggle"
            onClick={() => toggleSection('definitionJson')}
            title={collapsedSections.definitionJson ? 'Expand section' : 'Minimize section'}
            aria-label={collapsedSections.definitionJson ? 'Expand section' : 'Minimize section'}
          >
            {collapsedSections.definitionJson ? '+' : '-'}
          </button>
        </div>

        {!collapsedSections.definitionJson ? (
          <div className="json-preview">
            <pre>{definitionText || 'No definition loaded.'}</pre>
          </div>
        ) : null}
      </section>
    </div>
  );
}

export default App;
