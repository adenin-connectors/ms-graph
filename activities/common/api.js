'use strict';

const logger = require('@adenin/cf-logger');
const got = require('got');
const HttpAgent = require('agentkeepalive');
const HttpsAgent = HttpAgent.HttpsAgent;
const cfActivity = require('@adenin/cf-activity');

const endpoint = 'https://graph.microsoft.com';

let _activity = null;

function api(path, opts) {
  if (typeof path !== 'string') {
    return Promise.reject(new TypeError(`Expected \`path\` to be a string, got ${typeof path}`));
  }

  opts = Object.assign({
    json: true,
    token: _activity.Context.connector.token,
    endpoint: endpoint,
    agent: {
      http: new HttpAgent(),
      https: new HttpsAgent()
    }
  }, opts);

  opts.headers = Object.assign({
    accept: 'application/json',
    'user-agent': 'adenin Digital Assistant, https://www.adenin.com/digital-assistant/'
  }, opts.headers);

  if (opts.token) {
    opts.headers.Authorization = `Bearer ${opts.token}`;
  }

  const url = /^http(s)\:\/\/?/.test(path) && opts.endpoint ? path : opts.endpoint + path;

  if (opts.stream) {
    return got.stream(url, opts);
  }

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

api.stream = (url, opts) => got(url, Object.assign({}, opts, {
  json: false,
  stream: true
}));

api.initialize = function (activity) {
  _activity = activity;
};

for (const x of helpers) {
  const method = x.toUpperCase();

  api[x] = (url, opts) => api(url, Object.assign({}, opts, {
    method
  }));

  api.stream[x] = (url, opts) => api.stream(url, Object.assign({}, opts, {
    method
  }));
}

api.handleError = (activity, error) => {
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

  if (error.stack) {
    m = m + ': ' + error.stack;
  }

  activity.Response.ErrorCode = (error.response && error.response.statusCode) || 500;
  activity.Response.Data = {
    ErrorText: m
  };
};

api.isResponseOk = function (activity, response, succssStatusCodes) {
  return cfActivity.isResponseOk(activity, response, succssStatusCodes);
};

api.pagination = function (activity) {
  activity.Response.Data = {};

  const pagination = cfActivity.pagination(activity);

  if (activity.Request.Data && activity.Request.Data.args && activity.Request.Data.args.atAgentAction === 'nextpage') {
    pagination.nextlink = activity.Request.Data.args._nextlink;
  }

  return pagination;
};

module.exports = api;
