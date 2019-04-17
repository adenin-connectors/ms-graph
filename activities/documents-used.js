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

    Activity.Response.Data.items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      Activity.Response.Data.items.push(api.convertInsightsItem(response.body.value[i]));
    }

    const link = Activity.Response.Data.items[0].link;
    const base = link.substring(0, link.lastIndexOf('.com') + 4);

    Activity.Response.Data.title = T('Recent files from Sharepoint');
    Activity.Response.Data.link = base;
    Activity.Response.Data.linkLabel = T('Go to Sharepoint');
  } catch (error) {
    api.handleError(Activity, error);
  }
};
