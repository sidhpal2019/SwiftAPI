const SUPPORTED_METHODS = ['get', 'post', 'delete', 'put', 'patch'];

function toAbsoluteServerPath(value) {
  if (!value) {
    return '';
  }

  return String(value)
    .replace(/^https?:\/\//i, '')
    .replace(/\/+$/, '');
}

function guessType(param) {
  if (param?.schema?.type) {
    return param.schema.type;
  }

  if (param?.type) {
    return param.type;
  }

  if (param?.schema?.items?.type) {
    return `array<${param.schema.items.type}>`;
  }

  return 'unknown';
}

function extractEnumValues(param) {
  const directEnum = param?.enum;
  if (Array.isArray(directEnum) && directEnum.length > 0) {
    return directEnum.map(String);
  }

  const schemaEnum = param?.schema?.enum;
  if (Array.isArray(schemaEnum) && schemaEnum.length > 0) {
    return schemaEnum.map(String);
  }

  const itemEnum = param?.items?.enum || param?.schema?.items?.enum;
  if (Array.isArray(itemEnum) && itemEnum.length > 0) {
    return itemEnum.map(String);
  }

  return [];
}

function paramKey(param) {
  return `${String(param.name || '').toLowerCase()}|${String(param.in || '').toLowerCase()}`;
}

function normalizeParameter(param) {
  return {
    name: String(param?.name || '').trim(),
    in: String(param?.in || 'unknown').trim(),
    required: Boolean(param?.required),
    type: guessType(param),
    enumValues: extractEnumValues(param),
  };
}

function buildBodyParameter(operation) {
  const requestBody = operation?.requestBody;
  if (!requestBody) {
    return null;
  }

  const contentTypes = Object.keys(requestBody.content || {});
  const firstContentType = contentTypes[0] || 'application/json';
  const schema = requestBody.content?.[firstContentType]?.schema || {};
  const schemaType = schema.type || 'object';

  return {
    name: 'requestBody',
    in: 'body',
    required: Boolean(requestBody.required),
    type: schemaType,
    enumValues: [],
  };
}

function buildSampleValue(param) {
  if (Array.isArray(param.enumValues) && param.enumValues.length > 0) {
    return String(param.enumValues[0]);
  }

  switch (String(param.type).toLowerCase()) {
    case 'integer':
    case 'int32':
    case 'int64':
    case 'number':
      return '1';
    case 'boolean':
      return 'true';
    case 'array':
    case 'array<string>':
      return 'sample1,sample2';
    case 'object':
      return '{}';
    default:
      return 'sample';
  }
}

function parameterToDisplay(param) {
  const enumSuffix = param.enumValues.length > 0 ? ` {${param.enumValues.join('|')}}` : '';
  return `${param.name} (${param.in}): ${String(param.required)} [${param.type}]${enumSuffix}`;
}

function makeQueryTemplate(queryParams) {
  if (queryParams.length === 0) {
    return '';
  }

  const parts = queryParams.map((param) => `${param.name}={${param.name}}`);
  return `?${parts.join('&')}`;
}

function mergePathAndMethodParams(pathParams, methodParams) {
  const merged = new Map();

  pathParams.forEach((param) => {
    merged.set(paramKey(param), param);
  });

  methodParams.forEach((param) => {
    merged.set(paramKey(param), param);
  });

  return Array.from(merged.values());
}

export function parseDefinition(jsonText) {
  try {
    const parsed = JSON.parse(jsonText);
    return { parsed, error: null };
  } catch (error) {
    return { parsed: null, error: `Invalid JSON definition: ${error.message}` };
  }
}

export function resolveServerFromDefinition(definition, fallbackServerPath) {
  if (!definition || typeof definition !== 'object') {
    return toAbsoluteServerPath(fallbackServerPath);
  }

  if (Array.isArray(definition.servers) && definition.servers.length > 0) {
    const first = definition.servers[0]?.url || '';
    return toAbsoluteServerPath(first.replace(/\{[^}]+\}/g, ''));
  }

  if (definition.host) {
    return toAbsoluteServerPath(`${definition.host}${definition.basePath || ''}`);
  }

  return toAbsoluteServerPath(fallbackServerPath);
}

export function extractEndpoints(definition) {
  const paths = definition?.paths || {};

  return Object.keys(paths).map((pathName) => {
    const pathItem = paths[pathName] || {};
    const availableMethods = SUPPORTED_METHODS.filter((method) => Boolean(pathItem[method]));

    return {
      path: pathName,
      methods: availableMethods,
    };
  });
}

export function generateTestSuite(definition, config) {
  const protocol = String(config?.protocol || 'https').replace(':', '');
  const serverPath = resolveServerFromDefinition(definition, config?.serverPath || '');
  const paths = definition?.paths || {};
  const rows = [];

  Object.entries(paths).forEach(([pathName, pathItem]) => {
    const pathParams = Array.isArray(pathItem?.parameters)
      ? pathItem.parameters.map(normalizeParameter)
      : [];

    SUPPORTED_METHODS.forEach((methodName) => {
      const operation = pathItem?.[methodName];
      if (!operation) {
        return;
      }

      const methodParams = Array.isArray(operation.parameters)
        ? operation.parameters.map(normalizeParameter)
        : [];
      const mergedParams = mergePathAndMethodParams(pathParams, methodParams);
      const bodyFromRequestBody = buildBodyParameter(operation);
      if (bodyFromRequestBody) {
        mergedParams.push(bodyFromRequestBody);
      }

      const queryParams = mergedParams.filter((param) => param.in.toLowerCase() === 'query');
      const headerParams = mergedParams.filter((param) => param.in.toLowerCase() === 'header');
      const bodyParams = mergedParams.filter((param) => param.in.toLowerCase() === 'body');

      const paramsDisplay = mergedParams.map(parameterToDisplay).join(', ');
      const paramValues = mergedParams
        .filter((param) => ['path', 'query', 'body'].includes(param.in.toLowerCase()))
        .map((param) => `${param.name}: ${buildSampleValue(param)}`)
        .join('\n');
      const headerValues = headerParams
        .map((param) => `${param.name}: ${buildSampleValue(param)}`)
        .join('\n');

      const defaultBody =
        bodyParams.length > 0 && !paramValues.includes('requestBody')
          ? '{}'
          : operation?.requestBody
            ? JSON.stringify({}, null, 2)
            : '';

      rows.push({
        id: `${pathName}:${methodName}:${rows.length}`,
        protocol,
        serverPath,
        path: pathName,
        method: methodName.toUpperCase(),
        params: paramsDisplay,
        urlTemplate: `${protocol}://${serverPath}${pathName}${makeQueryTemplate(queryParams)}`,
        paramValues,
        headerValues,
        body: defaultBody,
        responseCode: '',
        responseHeaders: '',
        responseBody: '',
        responseTimeMs: '',
      });
    });
  });

  return rows;
}
