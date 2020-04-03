'use strict';

const api = require('./common/api');
const helpers = require('./common/helpers');

const onlySites = 'ResourceVisualization/Type eq \'spsite\'';

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api(`/v1.0/me/insights/trending?$top=999&$filter=${onlySites}`);

    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data.title = 'Trending Sites';
    activity.Response.Data.link = 'https://office.com/launch/sharepoint';
    activity.Response.Data.linkLabel = 'Go to Sharepoint';
    activity.Response.Data.items = [];

    if (!response.body.value || !response.body.value.length) return;

    for (let i = 0; i < response.body.value.length; i++) {
      const raw = response.body.value[i];

      // for Cisco and v1 cases we need to make the avatar app.adenin.com always.
      const plainTitle = helpers.stripSpecialChars(raw.resourceVisualization.title);
      const rawAvatar = $.avatarLink(plainTitle);
      const avatar = `https://app.adenin.com/avatar${rawAvatar.substring(rawAvatar.lastIndexOf('/'), rawAvatar.length)}?color=1e4471&size=52&fontSize=64`;

      activity.Response.Data.items.push({
        id: raw.id,
        title: raw.resourceVisualization.title,
        description: raw.resourceVisualization.containerType,
        link: raw.resourceReference.webUrl,
        thumbnail: avatar,
        imageIsAvatar: true
      });
    }
  } catch (error) {
    $.handleError(activity, error);
  }
};
