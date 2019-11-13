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
    const processed = [];

    for (let i = 0; i < deduplicated.length; i++) {
      processed.push(convertItem(deduplicated[i]));
    }

    activity.Response.Data.title = T(activity, 'Cloud Files');
    activity.Response.Data.items = processed;
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
