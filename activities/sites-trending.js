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

    const map = new Map();
    const promises = [];

    for (let i = 0; i < response.body.value.length; i++) {
      const raw = response.body.value[i];

      // for Cisco and v1 cases we need to make the avatar app.adenin.com always.
      const plainTitle = helpers.stripSpecialChars(raw.resourceVisualization.title);
      const rawAvatar = $.avatarLink(plainTitle);
      const avatar = `https://app.adenin.com/avatar${rawAvatar.substring(rawAvatar.lastIndexOf('/'), rawAvatar.length)}?color=1e4471&size=52&fontSize=64`;

      const id = raw.resourceReference.id.replace('sites/', '');

      map.set(id, {
        id: id,
        title: raw.resourceVisualization.title,
        description: raw.resourceVisualization.containerType,
        link: raw.resourceReference.webUrl,
        thumbnail: avatar,
        imageIsAvatar: true
      });

      promises.push(api(`/v1.0/sites/${id}`));
    }

    const results = await Promise.all(promises);
    const items = [];

    for (let i = 0; i < results.length; i++) {
      const site = results[i];

      if ($.isErrorResponse(activity, site)) return;

      const item = map.get(site.body.id);

      item.date = site.body.lastModifiedDateTime;

      items.push(item);
    }

    activity.Response.Data.items = items.sort($.compare.dateDescending);

    const remainder = 3 - (activity.Response.Data.items.length % 3);

    for (let i = 1; i <= remainder; i++) {
      activity.Response.Data._remainders.push(i);
    }
  } catch (error) {
    $.handleError(activity, error);
  }
};
