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
    token: 'eyJ0eXAiOiJKV1QiLCJub25jZSI6ImlXTDNZVm0tZUNUYXZkMVZWS3llRzFoMUpmeDdFQVA0RzhEMzBMS19GWmciLCJhbGciOiJSUzI1NiIsIng1dCI6InBpVmxsb1FEU01LeGgxbTJ5Z3FHU1ZkZ0ZwQSIsImtpZCI6InBpVmxsb1FEU01LeGgxbTJ5Z3FHU1ZkZ0ZwQSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9jZTRjYzY2MS00NTA2LTRkNDgtOGM2NC1jNWIwOTBhYTQ2ZmIvIiwiaWF0IjoxNTc5MDE1MDg5LCJuYmYiOjE1NzkwMTUwODksImV4cCI6MTU3OTAxODk4OSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFTUUEyLzhPQUFBQTJUcGZOdHhmbTlzWXk4cUgzQ3VsVHVVUmNWUmlWdUx2bm9uMDZhZWhCZUU9IiwiYW1yIjpbInB3ZCJdLCJhcHBfZGlzcGxheW5hbWUiOiJEaWdpdGFsIEFzc2lzdGFudCIsImFwcGlkIjoiNzk4NDI5YzctNjZhNy00NjJiLTliYzctODhlMTI0ZTY4MGIwIiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiJNY05hbGx5IiwiZ2l2ZW5fbmFtZSI6Ik9saXZlciIsImlwYWRkciI6IjgyLjMzLjE3OS41MyIsIm5hbWUiOiJPbGl2ZXIgTWNOYWxseSIsIm9pZCI6ImNjZWFjOGVlLTU2MjctNGJkMy05YzFiLWU2ODRhMTdlZTNiYiIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzN0ZGRUFDQzg5MDkxIiwic2NwIjoiQ2FsZW5kYXJzLlJlYWQgQ2FsZW5kYXJzLlJlYWRXcml0ZS5TaGFyZWQgRGlyZWN0b3J5LlJlYWQuQWxsIEZpbGVzLlJlYWQgR3JvdXAuUmVhZC5BbGwgUGVvcGxlLlJlYWQuQWxsIFNpdGVzLlJlYWQuQWxsIFRhc2tzLlJlYWRXcml0ZSBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCBwcm9maWxlIG9wZW5pZCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IlI0RWdZNU5ZMUxfeHgxcW9BOE9iNU1HNXZtR1NVT1pXZXNDMTYwNkRBMzAiLCJ0aWQiOiJjZTRjYzY2MS00NTA2LTRkNDgtOGM2NC1jNWIwOTBhYTQ2ZmIiLCJ1bmlxdWVfbmFtZSI6Im9saXZlci5tY25hbGx5QGFkZW5pbi5jb20iLCJ1cG4iOiJvbGl2ZXIubWNuYWxseUBhZGVuaW4uY29tIiwidXRpIjoiVXVDOTc1T2g1VUNDNEtpT08zVFBBQSIsInZlciI6IjEuMCIsInhtc19zdCI6eyJzdWIiOiIyb2UxaGxMTWtSelJXQUdqaE5ObFdUcFo0YkZrQ1FjZkxRS3VPYTMzc2wwIn0sInhtc190Y2R0IjoxMzg0NDU4MzI1fQ.Vrft15cXHEvcGh7iIZnyfVD73cKLLRFulnNKjWYqTd0rR8PJe-lMDqVB0xGs2LRWlPvi6WtFFmAdfbiLInlelcgXXiMs5jNZClfTG2-Ymt00FXzptoq6DhgRUC6lfakigwFQOOKXn5GX0JWJTHAmozOda_NlXVHUBbmW5LF50xZSNkCjfljQ7qAECnK2qFEMUwg9ZYfm5lm1OQBCfTmPvLGYuwExCV_Y_LeQkJrA-X6l9XMWEBabXtuib-UlgH33f-Zu-dgG1V8Dz60Dxh1RLoVwxGS_YgO62R3wGd-lNtdRV3kDkIXNBoJSFaZ-lBFDNOYq20UFB_ZUVRTJfvg0Pw', //_activity.Context.connector.token,
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

//** filters tickets by provided daterange */
api.filterByDateRange = function (tasks, dateRange) {
  const filteredTasks = [];
  const timeMin = Date.parse(dateRange.startDate);
  const timeMax = Date.parse(dateRange.endDate);

  for (let i = 0; i < tasks.length; i++) {
    const createTime = Date.parse(tasks[i].createdDateTime);

    if (createTime > timeMin && createTime < timeMax) filteredTasks.push(tasks[i]);
  }

  return filteredTasks;
};

//** paginate items[] based on provided pagination */
api.paginateItems = function (items, pagination) {
  const paginatedItems = [];
  const pageSize = parseInt(pagination.pageSize);
  const offset = (parseInt(pagination.page) - 1) * pageSize;

  if (offset > items.length) return paginatedItems;

  for (let i = offset; i < offset + pageSize; i++) {
    if (i >= items.length) break;

    paginatedItems.push(items[i]);
  }

  return paginatedItems;
};

module.exports = api;
