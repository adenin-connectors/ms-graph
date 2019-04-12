'use strict';

const querystring = require('querystring');
const api = require('./common/api');

module.exports = async () => {
  try {
    if (Activity.Request.Path) {
      const user = await api(`/v1.0/users/${Activity.Request.Path}`);

      Activity.Response.Data = convertItem(user.body);

      return;
    }

    const pages = Activity.pagination();

    // return empty result if no search term was provided
    if (!Activity.Request.Query.query) return;

    let url = '/v1.0/me/people';

    if (Activity.Request.Query.query) {
      // replace special characters
      const search = Activity.Request.Query.query.replace(/[`~!#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, ' ').trim();

      if (search) url = `${url}?$search=${querystring.escape(search)}`;
    }

    const response = await api(url);

    if (Activity.isErrorResponse(response));

    const startItem = Math.max(pages.page - 1, 0) * pages.pageSize;
    let endItem = startItem + pages.pageSize;

    if (endItem > response.body.value.length) endItem = response.body.value.length;

    for (let i = startItem; i < endItem; i++) {
      Activity.Response.Data.items.push(convertItem(response.body.value[i]));
    }
  } catch (error) {
    api.handleError(Activity, error);
  }
};

function convertItem(_item) {
  const item = {
    id: _item.id,
    name: _item.displayName,
    title: _item.displayName,
    email: _item.userPrincipalName
  };

  if (!_item.userPrincipalName) {
    if (_item.scoredEmailAddresses && _item.scoredEmailAddresses.length > 0) {
      item.email = _item.scoredEmailAddresses[0].address;
    }
  }

  if (item.email) item.id = item.email;

  return item;
}
