'use strict';

const api = require('./common/api');
const helpers = require('./common/helpers');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const promises = [];

    promises.push(api('/beta/me/insights/shared'));
    promises.push(api('/beta/me/insights/trending'));
    promises.push(api('/beta/me/insights/used'));

    const responses = await Promise.all(promises);
    const map = new Map();

    for (let i = 0; i < responses.length; i++) {
      const response = responses[i];

      if ($.isErrorResponse(activity, response)) return;

      for (let j = 0; j < response.body.value.length && j < 2; j++) {
        const item = response.body.value[j];

        if (map.has(item.id)) {
          map.set(item.id, Object.assign(map.get(item.id), item));
        } else {
          map.set(item.id, item);
        }
      }
    }

    const deduplicated = Array.from(map.values());
    const items = [];

    for (let i = 0; i < deduplicated.length; i++) {
      items.push(convertItem(deduplicated[i]));
    }

    activity.Response.Data.title = T(activity, 'Cloud Files');

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
      activity.Response.Data.description = count > 1 ? `You have ${count} new cloud files.` : 'You have 1 new cloud file.';
      activity.Response.Data.briefing = activity.Response.Data.description + ` The latest is '${first.title}'`;
    } else {
      activity.Response.Data.description = T(activity, 'You have no new cloud files.');
    }
  } catch (error) {
    $.handleError(activity, error);
  }
};

function convertItem(raw) {
  const item = {
    id: raw.id,
    title: raw.resourceVisualization.title,
    description: raw.resourceVisualization.previewText,
    type: raw.resourceVisualization.type || raw.resourceVisualization.containerType,
    link: raw.resourceReference.webUrl,
    preview: raw.resourceVisualization.previewImageUrl,
    containerTitle: helpers.stripSpecialChars(raw.resourceVisualization.containerDisplayName),
    containerLink: raw.resourceVisualization.containerWebUrl,
    containerType: raw.resourceVisualization.containerType
  };

  if (raw.lastShared) {
    item.sharedBy = helpers.stripSpecialChars(raw.lastShared.sharedBy.displayName);
    item.sharedDate = new Date(raw.lastShared.sharedDateTime);
    item.description = raw.lastShared.sharingSubject;
    item.link = raw.lastShared.sharingReference.webUrl;
  }

  if (raw.lastUsed) {
    const now = new Date();

    const twentyFourHoursAgo = new Date(now.setHours(now.getHours() - 24));
    const twoDaysAgo = new Date(now.setDate(now.getDate() - 2));

    const lastOpened = new Date(raw.lastUsed.lastAccessedDateTime);
    const lastModified = new Date(raw.lastUsed.lastModifiedDateTime);

    if (lastOpened > twentyFourHoursAgo) item.lastOpened = lastOpened;
    if (lastModified > twoDaysAgo) item.lastModified = lastModified;
  }

  item.raw = raw;

  return item;
}
