'use strict';

const api = require('./common/api');

module.exports = async () => {
  try {
    const pages = Activity.pagination();
    const query = Activity.Request.Query.query;

    let response;

    if (pages.nextpage) {
      response = await api(pages.nextpage);
    } else {
      response = await api(`/v1.0/me/drive/root/search(q='${query}')?$top=${pages.pageSize}`);
    }

    logger.info('received', response);

    if (!Activity.isResponseOk(response)) return;

    for (let i = 0; i < response.body.value.length; i++) {
      Activity.Response.Data.items.push(convertItem(response.body.value[i]));
    }

    Activity.Response.Data._nextpage = response.body['@odata.nextLink'];
  } catch (error) {
    api.handleError(Activity, error);
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
