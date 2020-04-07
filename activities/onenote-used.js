'use strict';

const api = require('./common/api');
const helpers = require('./common/helpers');

const onlyOneNote = 'ResourceVisualization/Type eq \'OneNote\'';

module.exports = async (activity) => {
  try {
    if (!activity.Request.Query.pageSize) activity.Request.Query.pageSize = 5;

    const pages = $.pagination(activity);
    const skip = (pages.page - 1) * pages.pageSize;
    const top = pages.pageSize;

    api.initialize(activity);

    const response = await api(`/v1.0/me/insights/used?$top=${top}&$skip=${skip}&$filter=${onlyOneNote}`);
    const items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      items.push(helpers.convertInsightsItem(response.body.value[i]));
    }

    const count = items.length;

    activity.Response.Data = Object.assign(activity.Response.Data, {
      title: T(activity, 'Recent Notebooks'),
      link: 'https://office.com/launch/onenote',
      linkLabel: 'Go to OneNote',
      thumbnail: activity.Context.connector.host.connectorLogoUrl,
      value: count,
      actionable: count > 0,
      items: items.sort($.compare.dateDescending)
    });

    if (parseInt(pages.page) === 1 && count > 0) {
      const first = items[0];

      activity.Response.Data.date = first.date;
      activity.Response.Data.description = count > 1 ? `You have ${count} recent NoteBooks.` : 'You have 1 recent Notebook.';
      activity.Response.Data.briefing = activity.Response.Data.description + ` The last used was '${first.title}'`;
    } else {
      activity.Response.Data.description = T(activity, 'You have no recent Notebooks.');
    }
  } catch (error) {
    $.handleError(activity, error);
  }
};
