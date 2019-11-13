'use strict';

const api = require('./common/api');
const helpers = require('./common/helpers');

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
      activity.Response.Data.items.push(helpers.convertInsightsItem(response.body.value[i]));
    }

    activity.Response.Data.title = T(activity, 'Recent Files');

    if (activity.Response.Data.items[0] && activity.Response.Data.items[0].containerLink) {
      activity.Response.Data.link = activity.Response.Data.items[0].containerLink;
      activity.Response.Data.linkLabel = T(activity, 'Go to OneDrive');
    }
  } catch (error) {
    api.handleError(activity, error);
  }
};
