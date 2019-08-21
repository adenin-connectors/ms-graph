'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    if (!activity.Request.Query.pageSize) activity.Request.Query.pageSize = 5;

    const pages = $.pagination(activity);
    const skip = (pages.page - 1) * pages.pageSize;
    const top = pages.pageSize;

    api.initialize(activity);

    const response = await api(`/beta/me/insights/shared?$skip=${skip}&$top=${top}`);

    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data.items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      activity.Response.Data.items.push(convertItem(response.body.value[i]));
    }

    activity.Response.Data.title = T(activity, 'Shared with You');

    if (activity.Response.Data.items[0] && activity.Response.Data.items[0].containerLink) {
      activity.Response.Data.link = activity.Response.Data.items[0].containerLink;
      activity.Response.Data.linkLabel = T(activity, 'Go to OneDrive');
    }
  } catch (error) {
    api.handleError(activity, error);
  }
};

function convertItem(raw) {
  return {
    id: raw.id,
    title: raw.resourceVisualization.title,
    description: raw.lastShared.sharingSubject,
    type: raw.resourceVisualization.type || raw.resourceVisualization.containerType,
    link: raw.lastShared.sharingReference.webUrl,
    preview: raw.resourceVisualization.previewImageUrl,
    containerTitle: raw.resourceVisualization.containerDisplayName,
    containerLink: raw.resourceVisualization.containerWebUrl,
    containerType: raw.resourceVisualization.containerType,
    date: new Date(raw.lastShared.sharedDateTime),
    raw: raw
  };
}
