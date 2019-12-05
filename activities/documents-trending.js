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

    const response = await api(`/beta/me/insights/trending?$skip=${skip}&$top=${top}`);

    if ($.isErrorResponse(activity, response)) return;

    const items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      items.push(helpers.convertInsightsItem(response.body.value[i]));
    }

    activity.Response.Data.title = T(activity, 'Trending Files');

    let count = 0;
    let readDate = (new Date(new Date().setDate(new Date().getDate() - 30))).toISOString(); // default read date 30 days in the past

    if (activity.Request.Query.readDate) readDate = activity.Request.Query.readDate;

    for (let i = 0; i < items.length; i++) {
      if (items[i].date > readDate) count++;
    }

    const pagination = $.pagination(activity);

    activity.Response.Data.items = api.paginateItems(items, pagination);

    if (parseInt(pagination.page) === 1 && count > 0) {
      const first = items[0];

      activity.Response.Data.link = first.containerLink;
      activity.Response.Data.linkLabel = T(activity, 'Go to OneDrive');
      activity.Response.Data.thumbnail = 'https://www.adenin.com/assets/images/wp-images/logo/office-365.svg';
      activity.Response.Data.actionable = count > 0;
      activity.Response.Data.value = count;
      activity.Response.Data.date = first.date;
      activity.Response.Data.description = count > 1 ? `You have ${count} new trending files.` : 'You have 1 new trending file.';
      activity.Response.Data.briefing = activity.Response.Data.description + ` The latest is '${first.title}'`;
    } else {
      activity.Response.Data.description = T(activity, 'You have no new trending files.');
    }
  } catch (error) {
    api.handleError(activity, error);
  }
};
