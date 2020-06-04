'use strict';

const crypto = require('crypto');

const api = require('./common/api');
const helpers = require('./common/helpers');

const excludeSites = 'ResourceVisualization/Type ne \'Web\' and ResourceVisualization/Type ne \'spsite\'';
const excludeMail = 'ResourceVisualization/ContainerType ne \'Mail\'';
const onlyMail = 'ResourceVisualization/ContainerType eq \'Mail\' and ResourceVisualization/Type ne \'Text\' and ResourceVisualization/Type ne \'Image\' and ResourceVisualization/Type ne \'Other\'';

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    activity.Request.Query.pageSize = 10;

    const pagination = $.pagination(activity);
    const top = pagination.pageSize;
    const skip = (pagination.page - 1) * pagination.pageSize;

    let action = $.getObjPath(activity, 'Request.Data.args.atAgentAction');

    if (action === 'nextpage') action = activity.Request.Data.args._tab;

    const promises = [];

    switch (action) {
    case 1:
    case 'trending':
      promises.push(api(`/v1.0/me/insights/trending?$top=${top}&skip=${skip}&$filter=${excludeSites} and ${excludeMail}`));
      activity.Response.Data._tab = 1;
      break;
    case 2:
    case 'shared':
      promises.push(api(`/v1.0/me/insights/shared?$top=${top}&skip=${skip}&$filter=${excludeSites} and ${excludeMail}`));
      activity.Response.Data._tab = 2;
      break;
    case 3:
    case 'email':
      promises.push(api(`/v1.0/me/insights/used?$top=${top}&skip=${skip}&$filter=${onlyMail}`));
      promises.push(api(`/v1.0/me/insights/shared?$top=${top}&skip=${skip}&$filter=${onlyMail}`));
      promises.push(api(`/v1.0/me/insights/trending?$top=${top}&skip=${skip}&$filter=${onlyMail}`));
      activity.Response.Data._tab = 3;
      break;
    default:
    case 0:
    case 'used':
      promises.push(api(`/v1.0/me/insights/used?$top=${top}&skip=${skip}&$filter=${excludeSites} and ${excludeMail}`));
      activity.Response.Data._tab = 0;
      break;
    }

    const responses = await Promise.all(promises);
    const map = new Map();

    for (let i = 0; i < responses.length; i++) {
      const response = responses[i];

      if ($.isErrorResponse(activity, response)) return;

      for (let j = 0; j < response.body.value.length; j++) {
        const item = convertItem(response.body.value[j]);

        if (map.has(item.id)) {
          map.set(item.id, Object.assign(map.get(item.id), item));
        } else {
          map.set(item.id, item);
        }
      }
    }

    const items = Array.from(map.values());

    activity.Response.Data.title = T(activity, 'Cloud Files');
    activity.Response.Data.items = items;

    if (parseInt(pagination.page) === 1) {
      const count = items.length;

      activity.Response.Data.link = 'https://office.com/launch/onedrive';
      activity.Response.Data.linkLabel = T(activity, 'Go to OneDrive');
      activity.Response.Data.thumbnail = activity.Context.connector.host.connectorLogoUrl;
      activity.Response.Data.actionable = count > 0;

      if (count > 0) {
        const first = items[0];

        activity.Response.Data.items[0]._isInitial = true;
        activity.Response.Data.value = count;
        activity.Response.Data.date = first.date;
        activity.Response.Data.description = count > 1 ? `You have ${count} new cloud files.` : 'You have 1 new cloud file.';
        activity.Response.Data.briefing = activity.Response.Data.description + ` The latest is '${first.title}'`;
      } else {
        activity.Response.Data.description = T(activity, 'You have no new cloud files.');
      }
    }
  } catch (error) {
    $.handleError(activity, error);
  }

  const hash = crypto.createHash('md5').update(JSON.stringify(activity.Response.Data)).digest('hex');

  activity.Response.Data._hash = hash;
};

function convertItem(raw) {
  const item = {
    id: raw.id,
    title: raw.resourceVisualization.title,
    description: helpers.stripNonAscii(raw.resourceVisualization.previewText),
    type: raw.resourceVisualization.type || raw.resourceVisualization.containerType,
    link: raw.resourceReference.webUrl,
    containerTitle: helpers.stripSpecialChars(raw.resourceVisualization.containerDisplayName).replace('\\', ''),
    containerLink: raw.resourceVisualization.containerWebUrl,
    containerType: raw.resourceVisualization.containerType
  };

  if (raw.lastShared) {
    item.sharedBy = helpers.stripSpecialChars(raw.lastShared.sharedBy.displayName);
    item.sharedDate = new Date(raw.lastShared.sharedDateTime);
    item.date = item.sharedDate;
    item.description = raw.lastShared.sharingSubject;
    item.link = raw.lastShared.sharingReference.webUrl;
  }

  if (raw.lastUsed) {
    const now = new Date();

    const twentyFourHoursAgo = new Date(now.setHours(now.getHours() - 24));
    const twoDaysAgo = new Date(now.setDate(now.getDate() - 2));

    const lastOpened = new Date(raw.lastUsed.lastAccessedDateTime);
    const lastModified = new Date(raw.lastUsed.lastModifiedDateTime);

    if (lastOpened > twentyFourHoursAgo) {
      item.lastOpened = lastOpened;
    }

    if (lastModified > twoDaysAgo) {
      item.lastModified = lastModified;
    }

    if (!item.date) item.date = lastOpened;
    if (!item.date) item.date = lastModified;
  }

  return item;
}
