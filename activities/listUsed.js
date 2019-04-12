'use strict';

const api = require('./common/api');

module.exports = async () => {
  try {
    if (!Activity.Request.Query.pageSize) Activity.Request.Query.pageSize = 5;

    const pages = Activity.pagination();
    const skip = (pages.page - 1) * pages.pageSize;
    const top = pages.pageSize;

    const response = await api(`/beta/me/insights/used?$skip=${skip}&$top=${top}`);

    if (Activity.isErrorResponse(response)) return;

    Activity.Response.Data.items = response.body.value;
  } catch (error) {
    api.handleError(Activity, error);
  }
};
