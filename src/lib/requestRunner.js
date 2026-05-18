function parseNameValueLines(input) {
  const result = {};
  if (!input) {
    return result;
  }

  String(input)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      const separator = line.indexOf(':');
      if (separator < 0) {
        return;
      }

      const key = line.slice(0, separator).trim();
      const value = line.slice(separator + 1).trim();
      if (key) {
        result[key] = value;
      }
    });

  return result;
}

function applyUrlValues(urlTemplate, values) {
  return urlTemplate.replace(/\{([^}]+)\}/g, (_, token) => {
    const value = values[token];
    return value === undefined ? `{${token}}` : encodeURIComponent(value);
  });
}

function parseBody(rawBody) {
  const trimmed = String(rawBody || '').trim();
  if (!trimmed) {
    return undefined;
  }

  try {
    return JSON.stringify(JSON.parse(trimmed));
  } catch {
    return rawBody;
  }
}

function headersToText(headers) {
  const lines = [];
  headers.forEach((value, key) => {
    lines.push(`${key}: ${value}`);
  });
  return lines.join('\n');
}

export async function executeSingleRow(row, authToken) {
  const paramValues = parseNameValueLines(row.paramValues);
  const headerValues = parseNameValueLines(row.headerValues);
  const resolvedUrl = applyUrlValues(row.urlTemplate, paramValues);

  if (authToken && !headerValues.Authorization) {
    headerValues.Authorization = `Bearer ${authToken}`;
  }

  const requestInit = {
    method: row.method,
    headers: headerValues,
  };

  if (!['GET', 'HEAD'].includes(row.method)) {
    const body = parseBody(row.body);
    if (body !== undefined) {
      requestInit.body = body;
      if (!requestInit.headers['Content-Type']) {
        requestInit.headers['Content-Type'] = 'application/json';
      }
    }
  }

  const start = performance.now();

  try {
    const response = await fetch(resolvedUrl, requestInit);
    const text = await response.text();
    const elapsed = performance.now() - start;

    return {
      ...row,
      urlTemplate: resolvedUrl,
      responseCode: String(response.status),
      responseHeaders: headersToText(response.headers),
      responseBody: text,
      responseTimeMs: Math.round(elapsed),
    };
  } catch (error) {
    const elapsed = performance.now() - start;
    return {
      ...row,
      urlTemplate: resolvedUrl,
      responseCode: '0',
      responseHeaders: '',
      responseBody: `Request failed: ${error.message}`,
      responseTimeMs: Math.round(elapsed),
    };
  }
}

export async function fetchOAuthToken(config) {
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    scope: config.scope || '',
    client_id: config.clientId || '',
    client_secret: config.clientSecret || '',
  });

  const response = await fetch(config.tokenUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: body.toString(),
  });

  const payload = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || response.statusText || 'Token request failed');
  }

  if (!payload.access_token) {
    throw new Error('Token response does not include access_token');
  }

  return payload.access_token;
}
