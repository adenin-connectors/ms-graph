'use strict';

const crypto = require('crypto');

const api = require('./common/api');
const helpers = require('./common/helpers');

module.exports = async (activity) => {
  try {
    if (!activity.Request.Query.pageSize) activity.Request.Query.pageSize = 5;

    const pages = $.pagination(activity);
    const skip = (pages.page - 1) * pages.pageSize;
    const top = pages.pageSize;

    api.initialize(activity);

    const response = await api(`/beta/me/insights/shared?$skip=${skip}&$top=${top}`);

    if ($.isErrorResponse(activity, response)) return;

    const items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      items.push(convertItem(response.body.value[i]));
    }

    activity.Response.Data.title = T(activity, 'Files Shared with You');

    let count = 0;
    let readDate = (new Date(new Date().setDate(new Date().getDate() - 30))).toISOString(); // default read date 30 days in the past

    if (activity.Request.Query.readDate) readDate = activity.Request.Query.readDate;

    for (let i = 0; i < items.length; i++) {
      if (items[i].date > readDate) {
        count++;
        items[i].isNew = true;
      }
    }

    const pagination = $.pagination(activity);

    activity.Response.Data.items = api.paginateItems(items, pagination);

    if (parseInt(pagination.page) === 1 && count > 0) {
      const first = items[0];

      activity.Response.Data.link = first.containerLink;
      activity.Response.Data.linkLabel = T(activity, 'Go to OneDrive');
      activity.Response.Data.thumbnail = activity.Context.connector.host.connectorLogoUrl;
      activity.Response.Data.actionable = count > 0;
      activity.Response.Data.value = count;
      activity.Response.Data.date = first.date;
      activity.Response.Data.description = count > 1 ? `You have ${count} new files shared with you.` : 'You have 1 new file shared with you.';
      activity.Response.Data.briefing = activity.Response.Data.description + ` The latest is '${first.title}'`;
    } else {
      activity.Response.Data.description = T(activity, 'You have no new files shared with you.');
    }
  } catch (error) {
    api.handleError(activity, error);
  }

  const hash = crypto.createHash('md5').update(JSON.stringify(activity.Response.Data)).digest('hex');

  activity.Response.Data._hash = hash;
};

function convertItem(raw) {
  return {
    id: raw.id,
    title: raw.resourceVisualization.title,
    description: raw.lastShared.sharingSubject,
    type: raw.resourceVisualization.type || raw.resourceVisualization.containerType,
    link: raw.lastShared.sharingReference.webUrl,
    containerTitle: helpers.stripSpecialChars(raw.resourceVisualization.containerDisplayName),
    containerLink: raw.resourceVisualization.containerWebUrl,
    containerType: raw.resourceVisualization.containerType,
    date: new Date(raw.lastShared.sharedDateTime)
  };
}
