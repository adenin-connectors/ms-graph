'use strict';

const api = require('./common/api');
const helpers = require('./common/helpers');

const onlySites = 'ResourceVisualization/Type eq \'spsite\'';

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api(`/v1.0/me/insights/trending?$top=999&$filter=${onlySites}`);

    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data = {
      title: 'Trending Sites',
      link: 'https://office.com/launch/sharepoint',
      linkLabel: 'Go to Sharepoint',
      items: [],
      _remainders: []
    };

    if (!response.body.value || !response.body.value.length) return;

    for (let i = 0; i < response.body.value.length; i++) {
      const raw = response.body.value[i];

      // for Cisco and v1 cases we need to make the avatar app.adenin.com always.
      const plainTitle = helpers.stripSpecialChars(raw.resourceVisualization.title);
      const rawAvatar = $.avatarLink(plainTitle);
      const avatar = `https://app.adenin.com/avatar${rawAvatar.substring(rawAvatar.lastIndexOf('/'), rawAvatar.length)}?size=48&fontSize=56`;

      activity.Response.Data.items.push({
        id: raw.id,
        title: raw.resourceVisualization.title,
        description: raw.resourceVisualization.containerType,
        link: raw.resourceReference.webUrl,
        thumbnail: avatar,
        imageIsAvatar: true
      });
    }

    const remainder = 3 - (activity.Response.Data.items.length % 3);

    for (let i = 1; i <= remainder; i++) {
      activity.Response.Data._remainders.push({});
    }
  } catch (error) {
    $.handleError(activity, error);
  }
};
