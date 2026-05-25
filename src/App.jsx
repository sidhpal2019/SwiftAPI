import { useEffect, useMemo, useRef, useState } from 'react';
import { executeSingleRow, fetchOAuthToken } from './lib/requestRunner';
import { extractEndpoints, generateTestSuite, parseDefinition, resolveServerFromDefinition } from './lib/openApi';
import { exportTestSuiteToExcel, importTestSuiteFromExcel } from './lib/excelIO';

const WORKSPACE_STORAGE_KEY = 'swiftapi-workspace-library-v1';
const LIBRARY_FILE_TYPE = 'swiftapi-library-entry';
const LIBRARY_FILE_VERSION = 1;

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

function endpointMethodKey(path, method) {
  return `${path}::${method}`;
}

function buildEndpointSelectionMap(items) {
  const selection = {};

  items.forEach((item) => {
    item.methods.forEach((method) => {
      selection[endpointMethodKey(item.path, method)] = true;
    });
  });

  return selection;
}

function createId(prefix) {
  return `${prefix}:${Date.now().toString(36)}:${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeServerPath(value) {
  return String(value || '')
    .replace(/^https?:\/\//i, '')
    .replace(/\/+$/, '')
    .trim();
}

function getWorkspaceRootName(serverPath) {
  return normalizeServerPath(serverPath) || 'Unassigned Server';
}

function formatSnapshotTime(value) {
  if (!value) {
    return '';
  }

  try {
    return new Intl.DateTimeFormat(undefined, {
      dateStyle: 'medium',
      timeStyle: 'short',
    }).format(new Date(value));
  } catch {
    return String(value);
  }
}

function loadWorkspaceLibrary() {
  if (typeof window === 'undefined') {
    return [];
  }

  try {
    const raw = window.localStorage.getItem(WORKSPACE_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function sanitizeFileNamePart(value) {
  const cleaned = String(value || '')
    .replace(/[^a-z0-9.-]+/gi, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');

  return cleaned || 'library-entry';
}

function toSnapshotArray(items, kind) {
  if (!Array.isArray(items)) {
    return [];
  }

  return items
    .filter((item) => item && typeof item === 'object')
    .map((item) => ({
      ...item,
      id: createId(kind),
      savedAt: item.savedAt || new Date().toISOString(),
    }));
}

function normalizeWorkspaceEntry(rawEntry) {
  if (!rawEntry || typeof rawEntry !== 'object') {
    return null;
  }

  const serverPath = normalizeServerPath(rawEntry.serverPath || rawEntry.serverUrlPath || '');
  const serverPathKey = serverPath || '__unassigned__';

  return {
    id: createId('workspace'),
    serverPath,
    serverPathKey,
    definitions: toSnapshotArray(rawEntry.definitions, 'definition'),
    testSuites: toSnapshotArray(rawEntry.testSuites, 'suite'),
  };
}

function parseLibraryFilePayload(parsed, fallbackName) {
  if (!parsed) {
    return [];
  }

  if (Array.isArray(parsed)) {
    return parsed.map(normalizeWorkspaceEntry).filter(Boolean);
  }

  if (Array.isArray(parsed.workspaceLibrary)) {
    return parsed.workspaceLibrary.map(normalizeWorkspaceEntry).filter(Boolean);
  }

  if (parsed.type === LIBRARY_FILE_TYPE && parsed.entry) {
    const serverPath = normalizeServerPath(parsed.serverPath || parsed.entry.serverPath || '');
    const definitionSnapshot = parsed.entry.definitionSnapshot || null;
    const suiteSnapshot = parsed.entry.testSuiteSnapshot || null;

    return [
      normalizeWorkspaceEntry({
        serverPath,
        definitions: definitionSnapshot
          ? [
              {
                ...definitionSnapshot,
                name: definitionSnapshot.name || `${fallbackName} Definition`,
              },
            ]
          : [],
        testSuites: suiteSnapshot
          ? [
              {
                ...suiteSnapshot,
                name: suiteSnapshot.name || `${fallbackName} Test Suite`,
              },
            ]
          : [],
      }),
    ].filter(Boolean);
  }

  if (Array.isArray(parsed.definitions) || Array.isArray(parsed.testSuites)) {
    return [normalizeWorkspaceEntry(parsed)].filter(Boolean);
  }

  return [];
}

function groupByBasePath(items, getPath) {
  const groups = {};

  items.forEach((item) => {
    const basePath = getFirstBasePath(getPath(item));
    if (!groups[basePath]) {
      groups[basePath] = [];
    }

    groups[basePath].push(item);
  });

  return Object.entries(groups)
    .map(([basePath, rows]) => ({
      basePath,
      rows,
    }))
    .sort((a, b) => a.basePath.localeCompare(b.basePath));
}

function SnapshotHierarchy({ snapshot, groups, renderRow, emptyMessage, selected, onLoad, renderControls }) {
  const [collapsedGroups, setCollapsedGroups] = useState({});

  function toggleGroup(basePath) {
    setCollapsedGroups((prev) => ({
      ...prev,
      [basePath]: !prev[basePath],
    }));
  }

  return (
    <details className={`workspace-snapshot ${selected ? 'selected' : ''}`} open>
      <summary>
        <span className="workspace-snapshot-title">{snapshot.name}</span>
        <button
          type="button"
          className="snapshot-load"
          onClick={(event) => {
            event.preventDefault();
            event.stopPropagation();
            onLoad();
          }}
        >
          Load
        </button>
      </summary>

      <div className="workspace-snapshot-body">
        {renderControls ? <div className="snapshot-controls">{renderControls(snapshot)}</div> : null}
        {groups.length === 0 ? (
          <p className="muted tree-empty">{emptyMessage}</p>
        ) : (
          <ul className="hierarchy-list">
            {groups.map((group) => (
              <li key={`${snapshot.id}:${group.basePath}`} className="hierarchy-group">
                <div className="hierarchy-group-head">
                  <div className="hierarchy-group-title">{group.basePath}</div>
                  <div className="hierarchy-group-actions">
                    <span className="hierarchy-group-count">{group.rows.length} APIs</span>
                    <button
                      type="button"
                      className="hierarchy-toggle-link"
                      onClick={() => toggleGroup(group.basePath)}
                      aria-label={collapsedGroups[group.basePath] ? `Expand ${group.basePath}` : `Collapse ${group.basePath}`}
                      title={collapsedGroups[group.basePath] ? 'Expand group' : 'Minimize group'}
                    >
                      {collapsedGroups[group.basePath] ? '+' : '-'}
                    </button>
                  </div>
                </div>

                {!collapsedGroups[group.basePath] ? (
                  <ul className="hierarchy-rows">
                    {group.rows.map((row) => renderRow(row))}
                  </ul>
                ) : null}
              </li>
            ))}
          </ul>
        )}
      </div>
    </details>
  );
}

function App() {
  const excelInputRef = useRef(null);
  const libraryInputRef = useRef(null);
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
  const [endpointSelections, setEndpointSelections] = useState({});
  const [endpointSelectionEnabled, setEndpointSelectionEnabled] = useState(false);
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
  const [workspaceLibrary, setWorkspaceLibrary] = useState(loadWorkspaceLibrary);
  const [selectedDefinitionSnapshotId, setSelectedDefinitionSnapshotId] = useState('');
  const [selectedTestSuiteSnapshotId, setSelectedTestSuiteSnapshotId] = useState('');
  const [testSuiteNameDrafts, setTestSuiteNameDrafts] = useState({});

  const endpointCount = endpoints.length;
  const testCount = testSuite.length;
  const currentServerPath = normalizeServerPath(resolveServerFromDefinition(definition, config.serverPath));
  const currentWorkspaceRootName = getWorkspaceRootName(currentServerPath);
  const hasExtractedEndpoints = endpointCount > 0;

  const selectedEndpointMethodCount = useMemo(() => {
    return endpoints.reduce((count, item) => {
      return (
        count +
        item.methods.reduce((methodCount, method) => {
          const key = endpointMethodKey(item.path, method);
          return methodCount + (endpointSelections[key] ? 1 : 0);
        }, 0)
      );
    }, 0);
  }, [endpoints, endpointSelections]);

  const canGenerateTestSuite = hasExtractedEndpoints && endpointSelectionEnabled && selectedEndpointMethodCount > 0;

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

  useEffect(() => {
    try {
      window.localStorage.setItem(WORKSPACE_STORAGE_KEY, JSON.stringify(workspaceLibrary));
    } catch {
      setStatus('Workspace library could not be saved in this browser.');
    }
  }, [workspaceLibrary]);

  function upsertWorkspaceEntry(serverPath, updater) {
    const normalizedServerPath = normalizeServerPath(serverPath) || '__unassigned__';

    setWorkspaceLibrary((prev) => {
      const existingIndex = prev.findIndex((entry) => entry.serverPathKey === normalizedServerPath);
      const existing = existingIndex >= 0 ? prev[existingIndex] : null;
      const nextEntry = updater(
        existing || {
          id: createId('workspace'),
          serverPath: normalizedServerPath === '__unassigned__' ? '' : normalizedServerPath,
          serverPathKey: normalizedServerPath,
          definitions: [],
          testSuites: [],
        }
      );

      const nextLibrary = [...prev];
      if (existingIndex >= 0) {
        nextLibrary[existingIndex] = nextEntry;
      } else {
        nextLibrary.push(nextEntry);
      }

      return nextLibrary.sort((a, b) => a.serverPath.localeCompare(b.serverPath));
    });
  }

  function saveCurrentDefinitionSnapshot() {
    if (!definition) {
      setStatus('Load a valid definition before saving it to the workspace library.');
      return;
    }

    const snapshotId = createId('definition');
    const snapshotName = fileName && fileName !== 'No file selected' ? fileName.replace(/\.json$/i, '') : `Definition ${formatSnapshotTime(new Date().toISOString())}`;
    const definitionSnapshot = {
      id: snapshotId,
      name: snapshotName,
      savedAt: new Date().toISOString(),
      definitionText,
      definition,
      endpoints,
      config: { ...config },
    };

    upsertWorkspaceEntry(currentServerPath, (existing) => ({
      ...existing,
      serverPath: currentServerPath || existing.serverPath,
      serverPathKey: existing.serverPathKey,
      definitions: [definitionSnapshot, ...(existing.definitions || [])],
      testSuites: existing.testSuites || [],
    }));

    setSelectedDefinitionSnapshotId(snapshotId);
    setStatus(`Saved definition snapshot under ${currentWorkspaceRootName}.`);
  }

  function saveCurrentTestSuiteSnapshot() {
    if (testSuite.length === 0) {
      setStatus('Generate or import a test suite before saving it to the workspace library.');
      return;
    }

    const snapshotId = createId('suite');
    const snapshotName = `Test Suite ${formatSnapshotTime(new Date().toISOString())}`;
    const suiteSnapshot = {
      id: snapshotId,
      name: snapshotName,
      savedAt: new Date().toISOString(),
      testSuite,
      config: { ...config },
    };

    upsertWorkspaceEntry(currentServerPath, (existing) => ({
      ...existing,
      serverPath: currentServerPath || existing.serverPath,
      serverPathKey: existing.serverPathKey,
      definitions: existing.definitions || [],
      testSuites: [suiteSnapshot, ...(existing.testSuites || [])],
    }));

    setSelectedTestSuiteSnapshotId(snapshotId);
    setStatus(`Saved test suite snapshot under ${currentWorkspaceRootName}.`);
  }

  async function saveCurrentLibraryEntryToFile() {
    if (!definition || testSuite.length === 0) {
      setStatus('Load a definition and test suite before exporting a library entry JSON file.');
      return;
    }

    const savedAt = new Date().toISOString();
    const normalizedPath = normalizeServerPath(currentServerPath || config.serverPath);
    const extractedEndpoints = endpoints.length > 0 ? endpoints : extractEndpoints(definition);
    const definitionName = fileName && fileName !== 'No file selected' ? fileName.replace(/\.json$/i, '') : `Definition ${formatSnapshotTime(savedAt)}`;

    const payload = {
      type: LIBRARY_FILE_TYPE,
      version: LIBRARY_FILE_VERSION,
      savedAt,
      serverPath: normalizedPath,
      entry: {
        serverPath: normalizedPath,
        definitionSnapshot: {
          id: createId('definition'),
          name: definitionName,
          savedAt,
          definitionText,
          definition,
          endpoints: extractedEndpoints,
          config: { ...config },
        },
        testSuiteSnapshot: {
          id: createId('suite'),
          name: `Test Suite ${formatSnapshotTime(savedAt)}`,
          savedAt,
          testSuite,
          config: { ...config },
        },
      },
    };

    const fileNameBase = sanitizeFileNamePart(getWorkspaceRootName(normalizedPath));
    const suggestedName = `${fileNameBase}.swiftapi-library.json`;
    const jsonText = JSON.stringify(payload, null, 2);
    const blob = new Blob([jsonText], { type: 'application/json' });

    try {
      if (typeof window.showSaveFilePicker === 'function') {
        const handle = await window.showSaveFilePicker({
          suggestedName,
          types: [
            {
              description: 'SwiftAPI Library JSON',
              accept: {
                'application/json': ['.json'],
              },
            },
          ],
        });

        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();

        setStatus(`Saved library entry JSON for ${getWorkspaceRootName(normalizedPath)}.`);
        return;
      }
    } catch (error) {
      if (error?.name === 'AbortError') {
        setStatus('Save cancelled.');
        return;
      }
    }

    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = suggestedName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);

    setStatus('Save picker is not supported in this browser. File was downloaded using browser defaults.');
  }

  function triggerLibraryImport() {
    if (libraryInputRef.current) {
      libraryInputRef.current.click();
    }
  }

  async function handleLibraryImport(event) {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) {
      return;
    }

    let importedFiles = 0;
    let skippedFiles = 0;

    for (const file of files) {
      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        const importedEntries = parseLibraryFilePayload(parsed, file.name.replace(/\.json$/i, ''));

        if (importedEntries.length === 0) {
          skippedFiles += 1;
          continue;
        }

        setWorkspaceLibrary((prev) => {
          const byKey = new Map(prev.map((entry) => [entry.serverPathKey, { ...entry }]));

          importedEntries.forEach((entry) => {
            const existing = byKey.get(entry.serverPathKey);
            if (!existing) {
              byKey.set(entry.serverPathKey, {
                ...entry,
                definitions: [...(entry.definitions || [])],
                testSuites: [...(entry.testSuites || [])],
              });
              return;
            }

            byKey.set(entry.serverPathKey, {
              ...existing,
              serverPath: existing.serverPath || entry.serverPath,
              definitions: [...(entry.definitions || []), ...(existing.definitions || [])],
              testSuites: [...(entry.testSuites || []), ...(existing.testSuites || [])],
            });
          });

          return Array.from(byKey.values()).sort((a, b) => a.serverPath.localeCompare(b.serverPath));
        });

        importedFiles += 1;
      } catch {
        skippedFiles += 1;
      }
    }

    event.target.value = '';
    if (importedFiles === 0) {
      setStatus('No valid library JSON files were loaded.');
      return;
    }

    if (skippedFiles > 0) {
      setStatus(`Loaded ${importedFiles} library file(s). Skipped ${skippedFiles} unsupported file(s).`);
      return;
    }

    setStatus(`Loaded ${importedFiles} library file(s) into workspace library.`);
  }

  function loadDefinitionSnapshot(snapshot) {
    setDefinitionText(snapshot.definitionText || '');
    setDefinition(snapshot.definition || null);
    setFileName(snapshot.name || 'Saved definition');
    setParseError('');
    setConfig((prev) => ({
      ...prev,
      ...(snapshot.config || {}),
    }));
    const snapshotEndpoints = Array.isArray(snapshot.endpoints) ? snapshot.endpoints : [];
    setEndpoints(snapshotEndpoints);
    setEndpointSelections(buildEndpointSelectionMap(snapshotEndpoints));
    setEndpointSelectionEnabled(false);
    setSelectedDefinitionSnapshotId(snapshot.id);
    setEndpointFilter('');
    setStatus(`Loaded definition snapshot ${snapshot.name}.`);
  }

  function loadTestSuiteSnapshot(snapshot) {
    setTestSuite(Array.isArray(snapshot.testSuite) ? snapshot.testSuite : []);
    setConfig((prev) => ({
      ...prev,
      ...(snapshot.config || {}),
    }));
    setSelectedTestSuiteSnapshotId(snapshot.id);
    setTestSuiteNameDrafts((prev) => ({
      ...prev,
      [snapshot.id]: snapshot.name || 'Test Suite',
    }));
    setCollapsedTestGroups({});
    setStatus(`Loaded test suite snapshot ${snapshot.name}.`);
  }

  function handleTestSuiteNameDraftChange(snapshotId, value) {
    setTestSuiteNameDrafts((prev) => ({
      ...prev,
      [snapshotId]: value,
    }));
  }

  function saveTestSuiteSnapshotName(serverPathKey, snapshotId) {
    const draftName = (testSuiteNameDrafts[snapshotId] || '').trim();
    if (!draftName) {
      setStatus('Test suite name cannot be empty.');
      return;
    }

    let updated = false;

    setWorkspaceLibrary((prev) =>
      prev.map((entry) => {
        if (entry.serverPathKey !== serverPathKey) {
          return entry;
        }

        const nextSuites = (entry.testSuites || []).map((suite) => {
          if (suite.id !== snapshotId) {
            return suite;
          }

          updated = true;
          return {
            ...suite,
            name: draftName,
          };
        });

        return {
          ...entry,
          testSuites: nextSuites,
        };
      })
    );

    if (updated) {
      setStatus(`Saved test suite name as ${draftName}.`);
    }
  }

  function removeTestSuiteSnapshot(serverPathKey, snapshotId) {
    const workspace = workspaceLibrary.find((entry) => entry.serverPathKey === serverPathKey);
    const snapshot = workspace?.testSuites?.find((suite) => suite.id === snapshotId);
    if (!workspace || !snapshot) {
      return;
    }

    const confirmed = window.confirm(`Remove test suite ${snapshot.name}?`);
    if (!confirmed) {
      return;
    }

    setWorkspaceLibrary((prev) =>
      prev.map((entry) => {
        if (entry.serverPathKey !== serverPathKey) {
          return entry;
        }

        return {
          ...entry,
          testSuites: (entry.testSuites || []).filter((suite) => suite.id !== snapshotId),
        };
      })
    );

    setTestSuiteNameDrafts((prev) => {
      const next = { ...prev };
      delete next[snapshotId];
      return next;
    });

    if (selectedTestSuiteSnapshotId === snapshotId) {
      setSelectedTestSuiteSnapshotId('');
    }

    setStatus(`Removed test suite ${snapshot.name}.`);
  }

  function removeWorkspaceLibrary(serverPathKey) {
    const target = workspaceLibrary.find((entry) => entry.serverPathKey === serverPathKey);
    if (!target) {
      return;
    }

    const label = getWorkspaceRootName(target.serverPath);
    const confirmed = window.confirm(`Remove opened library for ${label}? This removes all saved definitions and test suites under this server path.`);
    if (!confirmed) {
      return;
    }

    const removedDefinitionIds = new Set((target.definitions || []).map((item) => item.id));
    const removedSuiteIds = new Set((target.testSuites || []).map((item) => item.id));

    setWorkspaceLibrary((prev) => prev.filter((entry) => entry.serverPathKey !== serverPathKey));

    if (removedDefinitionIds.has(selectedDefinitionSnapshotId)) {
      setSelectedDefinitionSnapshotId('');
    }

    if (removedSuiteIds.has(selectedTestSuiteSnapshotId)) {
      setSelectedTestSuiteSnapshotId('');
    }

    setStatus(`Removed opened library for ${label}.`);
  }

  function renderEndpointRow(item) {
    return (
      <li key={`${item.path}:${item.methods.join(',')}`} className="hierarchy-row">
        <span className="hierarchy-path">{item.path}</span>
      </li>
    );
  }

  function renderSuiteRow(item) {
    return (
      <li key={item.id} className="hierarchy-row hierarchy-row-suite">
        <span className="hierarchy-path">{item.path}</span>
        <span className="hierarchy-method">{String(item.method || '').toUpperCase()}</span>
      </li>
    );
  }

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
    setEndpoints([]);
    setEndpointSelections({});
    setEndpointSelectionEnabled(false);
    setTestSuite([]);
    setEndpointFilter('');

    const suggestedServer = resolveServerFromDefinition(parsed.parsed, config.serverPath);
    setConfig((prev) => ({
      ...prev,
      serverPath: suggestedServer,
    }));
    setStatus('Definition loaded. You can extract endpoints or generate the test suite.');
  }

  function clearExtractDefinition() {
    setEndpoints([]);
    setEndpointSelections({});
    setEndpointSelectionEnabled(false);
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
    setEndpointSelections(buildEndpointSelectionMap(extracted));
    setEndpointSelectionEnabled(false);
    setStatus(`Extracted ${extracted.length} endpoints.`);
  }

  function handleSelectTestSuite() {
    if (!definition) {
      setStatus('Load a valid definition before selecting endpoint methods.');
      return;
    }

    if (!hasExtractedEndpoints) {
      setStatus('Extract endpoints before selecting endpoint methods.');
      return;
    }

    setEndpointSelectionEnabled(true);
    setStatus('Endpoint method checkboxes are enabled. Choose methods and click Generate Test Suite.');
  }

  function isMethodSelected(path, method) {
    const key = endpointMethodKey(path, method);
    if (!(key in endpointSelections)) {
      return true;
    }

    return Boolean(endpointSelections[key]);
  }

  function setAllMethodSelections(checked) {
    setEndpointSelections((prev) => {
      const next = { ...prev };
      endpoints.forEach((item) => {
        item.methods.forEach((method) => {
          next[endpointMethodKey(item.path, method)] = checked;
        });
      });
      return next;
    });
  }

  function toggleMethodSelection(path, method, checked) {
    const key = endpointMethodKey(path, method);
    setEndpointSelections((prev) => ({
      ...prev,
      [key]: checked,
    }));
  }

  function handleGenerateTestSuite() {
    if (!definition) {
      setStatus('Load a valid definition before generating tests.');
      return;
    }

    if (!hasExtractedEndpoints) {
      setStatus('Extract endpoints before generating tests.');
      return;
    }

    if (!endpointSelectionEnabled) {
      setStatus('Click Select Endpoints to enable endpoint selection before generating tests.');
      return;
    }

    if (selectedEndpointMethodCount === 0) {
      setStatus('Select at least one endpoint method before generating tests.');
      return;
    }

    const extracted = endpoints;
    const selections = endpointSelections;

    const selectedKeys = new Set(
      Object.entries(selections)
        .filter(([, isSelected]) => Boolean(isSelected))
        .map(([key]) => key)
    );

    const rows = generateTestSuite(definition, config, selectedKeys);
    setTestSuite(rows);
    setCollapsedTestGroups({});

    if (rows.length === 0) {
      setStatus('No endpoint methods are selected. Select one or more checkboxes and generate again.');
      return;
    }

    setStatus(`Generated ${rows.length} test rows based on selected endpoints.`);
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

      <div className="workspace-layout">
        <main className="workspace-main">
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

                  <div className="definition-top-actions">
                    <button onClick={handleExtractEndpoints}>Extract Endpoints</button>
                    <button onClick={handleSelectTestSuite} disabled={!hasExtractedEndpoints}>
                      Select Endpoints
                    </button>
                    <button onClick={handleGenerateTestSuite} disabled={!canGenerateTestSuite}>
                      Generate TestSuite
                    </button>
                  </div>

                  <div className="action-row">
                    <button onClick={clearExtractDefinition} className="ghost">
                      Clear Extract
                    </button>
                    <button onClick={clearTestSuite} className="ghost">
                      Clear TestSuite
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
              <div className="panel-head-actions">
                <button
                  type="button"
                  className="ghost"
                  onClick={() => setAllMethodSelections(true)}
                  disabled={!endpointSelectionEnabled || endpointCount === 0}
                >
                  Select all
                </button>
                <button
                  type="button"
                  className="ghost"
                  onClick={() => setAllMethodSelections(false)}
                  disabled={!endpointSelectionEnabled || endpointCount === 0}
                >
                  Deselect all
                </button>
                <button
                  className="ghost icon-toggle"
                  onClick={() => toggleSection('endpoints')}
                  title={collapsedSections.endpoints ? 'Expand section' : 'Minimize section'}
                  aria-label={collapsedSections.endpoints ? 'Expand section' : 'Minimize section'}
                >
                  {collapsedSections.endpoints ? '+' : '-'}
                </button>
              </div>
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
                                const renderMethodCell = (methodName) => {
                                  if (!hasMethod(methodName)) {
                                    return '';
                                  }

                                  return (
                                    <input
                                      type="checkbox"
                                      className="endpoint-method-checkbox"
                                      checked={isMethodSelected(item.path, methodName)}
                                      disabled={!endpointSelectionEnabled}
                                      onChange={(event) => toggleMethodSelection(item.path, methodName, event.target.checked)}
                                      aria-label={`${methodName.toUpperCase()} ${item.path}`}
                                    />
                                  );
                                };

                                return (
                                  <tr key={`${group.basePath}:${item.path}`}>
                                    <td>{item.path}</td>
                                    <td>{renderMethodCell('get')}</td>
                                    <td>{renderMethodCell('post')}</td>
                                    <td>{renderMethodCell('delete')}</td>
                                    <td>{renderMethodCell('put')}</td>
                                    <td>{renderMethodCell('patch')}</td>
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
        </main>

        <aside className="workspace-sidebar panel">
          <div className="panel-head">
            <h2>Workspace Library</h2>
          </div>

          <p className="muted workspace-intro">Explorer-style saved definitions and test suites grouped by server path.</p>

          <div className="workspace-actions">
            <button onClick={saveCurrentDefinitionSnapshot} disabled={!definition}>
              Save Loaded Definition
            </button>
            <button onClick={saveCurrentTestSuiteSnapshot} disabled={testSuite.length === 0} className="ghost">
              Save Test Suite
            </button>
            <button onClick={saveCurrentLibraryEntryToFile} disabled={!definition || testSuite.length === 0}>
              Save Entry JSON
            </button>
            <button onClick={triggerLibraryImport} className="ghost">
              Load Library JSON
            </button>
            <input
              ref={libraryInputRef}
              type="file"
              accept=".json,application/json"
              multiple
              onChange={handleLibraryImport}
              style={{ display: 'none' }}
            />
          </div>

          {workspaceLibrary.length === 0 ? (
            <p className="muted tree-empty">No saved workspaces yet.</p>
          ) : (
            <div className="workspace-list">
              {workspaceLibrary.map((workspace) => (
                <details key={workspace.id} className="workspace-root" open>
                  <summary>
                    <span className="workspace-root-name">{getWorkspaceRootName(workspace.serverPath)}</span>
                    <button
                      type="button"
                      className="ghost workspace-root-remove"
                      onClick={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                        removeWorkspaceLibrary(workspace.serverPathKey);
                      }}
                    >
                      Remove
                    </button>
                  </summary>

                  <div className="workspace-root-body">
                    <div className="tree-branch">
                      <h4>Loaded Entry Points</h4>
                      {workspace.definitions?.length ? (
                        <div className="snapshot-list">
                          {workspace.definitions.map((snapshot) => (
                            <SnapshotHierarchy
                              key={snapshot.id}
                              snapshot={snapshot}
                              groups={groupByBasePath(snapshot.endpoints || [], (item) => item.path)}
                              renderRow={renderEndpointRow}
                              emptyMessage="No endpoints saved in this snapshot."
                              selected={selectedDefinitionSnapshotId === snapshot.id}
                              onLoad={() => loadDefinitionSnapshot(snapshot)}
                            />
                          ))}
                        </div>
                      ) : (
                        <p className="muted tree-empty">No saved definitions for this server path.</p>
                      )}
                    </div>

                    <div className="tree-branch">
                      <h4>Test Suites</h4>
                      {workspace.testSuites?.length ? (
                        <div className="snapshot-list">
                          {workspace.testSuites.map((snapshot) => (
                            <SnapshotHierarchy
                              key={snapshot.id}
                              snapshot={snapshot}
                              groups={groupByBasePath(snapshot.testSuite || [], (item) => item.path)}
                              renderRow={renderSuiteRow}
                              emptyMessage="No test rows saved in this snapshot."
                              selected={selectedTestSuiteSnapshotId === snapshot.id}
                              onLoad={() => loadTestSuiteSnapshot(snapshot)}
                              renderControls={() => (
                                <div className="snapshot-suite-controls">
                                  <label className="snapshot-suite-name-field">
                                    Name
                                    <input
                                      value={testSuiteNameDrafts[snapshot.id] ?? snapshot.name ?? ''}
                                      onChange={(event) => handleTestSuiteNameDraftChange(snapshot.id, event.target.value)}
                                      placeholder="Test Suite"
                                    />
                                  </label>
                                  <div className="snapshot-suite-actions">
                                    <button
                                      type="button"
                                      className="ghost"
                                      onClick={() => saveTestSuiteSnapshotName(workspace.serverPathKey, snapshot.id)}
                                    >
                                      Save Name
                                    </button>
                                    <button
                                      type="button"
                                      className="ghost"
                                      onClick={() => removeTestSuiteSnapshot(workspace.serverPathKey, snapshot.id)}
                                    >
                                      Remove Suite
                                    </button>
                                  </div>
                                </div>
                              )}
                            />
                          ))}
                        </div>
                      ) : (
                        <p className="muted tree-empty">No saved test suites for this server path.</p>
                      )}
                    </div>
                  </div>
                </details>
              ))}
            </div>
          )}
        </aside>
      </div>
    </div>
  );
}

export default App;
