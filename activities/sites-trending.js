'use strict';

const api = require('./common/api');

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

    if (!response.body.data.value || !response.body.data.value.length) return;

    for (let i = 0; i < response.body.data.value.length; i++) {
      activity.Response.Data.items.push(convertItem(response.body.data.value[i]));
    }
  } catch (error) {
    $.handleError(activity, error);
  }
};

function convertItem(raw) {
  return {
    id: raw.id,
    title: raw.resourceVisualization.title,
    description: raw.resourceVisualization.containerType,
    link: raw.resourceReference.webUrl
  };
}
