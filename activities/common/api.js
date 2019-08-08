'use strict';

const got = require('got');
const HttpAgent = require('agentkeepalive');
const HttpsAgent = HttpAgent.HttpsAgent;

let _activity = null;

function api(path, opts) {
  if (typeof path !== 'string') {
    return Promise.reject(new TypeError(`Expected \`path\` to be a string, got ${typeof path}`));
  }

  opts = Object.assign({
    json: true,
    token: _activity.Context.connector.token,
    endpoint: 'https://graph.microsoft.com',
    agent: {
      http: new HttpAgent(),
      https: new HttpsAgent()
    }
  }, opts);

  opts.headers = Object.assign({
    accept: 'application/json',
    'user-agent': 'adenin Digital Assistant, https://www.adenin.com/digital-assistant/'
  }, opts.headers);

  if (opts.token) opts.headers.Authorization = `Bearer ${opts.token}`;

  const url = /^http(s)\:\/\/?/.test(path) && opts.endpoint ? path : opts.endpoint + path;

  if (opts.stream) return got.stream(url, opts);

  return got(url, opts).catch((err) => {
    throw err;
  });
}

const helpers = [
  'get',
  'post',
  'put',
  'patch',
  'head',
  'delete'
];

api.stream = (url, opts) => api(url, Object.assign({}, opts, {
  json: false,
  stream: true
}));

api.initialize = function (activity) {
  _activity = activity;
};

for (const x of helpers) {
  const method = x.toUpperCase();
  api[x] = (url, opts) => api(url, Object.assign({}, opts, {method}));
  api.stream[x] = (url, opts) => api.stream(url, Object.assign({}, opts, {method}));
}

api.handleError = function (activity, error) {
  if (error.response && error.response.statusCode === 400) {
    // MSGraph might return a 'nice' error message
    // if available use that instead of generic error message
    if (error.body.error && error.body.error.message) {
      activity.Response.ErrorCode = error.response.statusCode;
      activity.Response.Data = {
        ErrorText: error.body.error.message
      };

      return;
    }
  }

  // if OAuth2 tokens are used status 401 should be mapped to 461 auth required
  const authRequiresStatusCodes = [401];

  if (error.response && error.response.statusCode && authRequiresStatusCodes.indexOf(error.response.statusCode) >= 0) {
    error.response.statusCode = 461;
  } else {
    logger.error(error);
  }

  let m = error.message;

  if (error.stack) m = m + ': ' + error.stack;

  activity.Response.ErrorCode = (error.response && error.response.statusCode) || 500;
  activity.Response.Data = {
    ErrorText: m
  };
};

api.convertInsightsItem = function (raw) {
  const type = raw.resourceVisualization.type.toLowerCase();
  let nowIcon = 'now:file';

  switch (type) {
  case 'word':
    nowIcon = 'now:word';
    break;
  case 'powerpoint':
    nowIcon = 'now:powerpoint';
    break;
  case 'excel':
    nowIcon = 'now:excel';
    break;
  case 'pdf':
    nowIcon = 'now:pdf';
    break;
  }

  return {
    id: raw.id,
    title: raw.resourceVisualization.title,
    description: raw.resourceVisualization.previewText,
    type: raw.resourceVisualization.type || raw.resourceVisualization.containerType,
    link: raw.resourceReference.webUrl,
    preview: raw.resourceVisualization.previewImageUrl,
    containerTitle: raw.resourceVisualization.containerDisplayName,
    containerLink: raw.resourceVisualization.containerWebUrl,
    containerType: raw.resourceVisualization.containerType,
    nowIcon: nowIcon,
    date: raw.lastUsed ? (new Date(raw.lastUsed.lastAccessedDateTime)).toISOString() : undefined,
    raw: raw
  };
};

//** filters tickets by provided daterange */
api.filterByDateRange = function (tasks, dateRange) {
  let filteredTasks = [];
  const timeMin = Date.parse(dateRange.startDate);
  const timeMax = Date.parse(dateRange.endDate);

  for (let i = 0; i < tasks.length; i++) {
    const createTime = Date.parse(tasks[i].createdDateTime);
    if (createTime > timeMin && createTime < timeMax) {
      filteredTasks.push(tasks[i]);
    }
  }

  return filteredTasks;
};

//** paginate items[] based on provided pagination */
api.paginateItems = function (items, pagination) {
  let pagiantedItems = [];
  const pageSize = parseInt(pagination.pageSize);
  const offset = (parseInt(pagination.page) - 1) * pageSize;

  if (offset > items.length) return pagiantedItems;

  for (let i = offset; i < offset + pageSize; i++) {
    if (i >= items.length) {
      break;
    }
    pagiantedItems.push(items[i]);
  }
  return pagiantedItems;
};
module.exports = api;
