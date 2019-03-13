'use strict';

const logger = require('@adenin/cf-logger');
const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const pages = api.pagination(activity);
    const query = activity.Request.Query.query;

    let response;

    if (pages.nextlink) {
      response = await api(pages.nextlink);
    } else {
      response = await api('/v1.0/me/drive/root/search(q=\'' + query + '\')?$top=' + pages.pageSize);
    }

    logger.info('received', response);

    if (!api.isResponseOk(activity, response)) {
      return;
    }

    activity.Response.Data.items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      activity.Response.Data.items.push(convertItem(response.body.value[i]));
    }

    activity.Response.Data._nextlink = response.body['@odata.nextLink'];
  } catch (error) {
    api.handleError(activity, error);
  }
};

function convertItem(_item) {
  return {
    id: _item.id,
    title: _item.name,
    link: _item.webUrl,
    date: (new Date(_item.createdDateTime)).toISOString()
  };
}
