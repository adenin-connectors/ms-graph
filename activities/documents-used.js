'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    if (!activity.Request.Query.pageSize) activity.Request.Query.pageSize = 5;

    const pages = $.pagination(activity);
    const skip = (pages.page - 1) * pages.pageSize;
    const top = pages.pageSize;

    api.initialize(activity);

    const response = await api(`/beta/me/insights/used?$skip=${skip}&$top=${top}`);

    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data.items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      activity.Response.Data.items.push(api.convertInsightsItem(response.body.value[i]));
    }

    const link = activity.Response.Data.items[0].link;
    const base = link.substring(0, link.lastIndexOf('.com') + 4);

    activity.Response.Data.title = T(activity, 'Recent Files');
    activity.Response.Data.link = base;
    activity.Response.Data.linkLabel = T(activity, 'Go to OneDrive');
  } catch (error) {
    api.handleError(activity, error);
  }
};
