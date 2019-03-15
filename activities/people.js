'use strict';

const querystring = require('querystring');
const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    if (activity.Request.Path) {
      const user = await api('/v1.0/users/' + activity.Request.Path);
      activity.Response.Data = convertItem(user.body);
      return;
    }

    let action = 'firstpage';
    let page = parseInt(activity.Request.Query.page, 10) || 1;
    let pageSize = parseInt(activity.Request.Query.pageSize, 10) || 20;

    if (
      activity.Request.Data &&
            activity.Request.Data.args &&
            activity.Request.Data.args.atAgentAction === 'nextpage'
    ) {
      page = parseInt(activity.Request.Data.args._page, 10) || 2;
      pageSize = parseInt(activity.Request.Data.args._pageSize, 10) || 20;
      action = 'nextpage';
    }

    if (page < 0) {
      page = 1;
    }

    if (pageSize < 1 || pageSize > 99) {
      pageSize = 20;
    }

    activity.Response.Data._action = action;
    activity.Response.Data._page = page;
    activity.Response.Data._pageSize = pageSize;

    activity.Response.Data.items = [];

    // return empty result if no search term was provided
    if (!activity.Request.Query.query) {
      return;
    }

    let url = '/v1.0/me/people';
    if (activity.Request.Query.query) {
      // replace special characters
      const search = activity.Request.Query.query.replace(/[`~!#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, ' ').trim();
      if (search) {
        url = url + '?$search=' + querystring.escape(search);
      }
    }

    const response = await api(url);

    if (response.statusCode === 200 && response.body.value && response.body.value.length > 0) {
      const startItem = Math.max(page - 1, 0) * pageSize;
      let endItem = startItem + pageSize;

      if (endItem > response.body.value.length) {
        endItem = response.body.value.length;
      }

      for (let i = startItem; i < endItem; i++) {
        activity.Response.Data.items.push(convertItem(response.body.value[i]));
      }
    } else {
      activity.Response.Data = {
        statusCode: response.statusCode,
        message: 'Bad request or no people found',
        items: []
      };
    }
  } catch (error) {
    api.handleError(activity, error);
  }
};

function convertItem(_item) {
  const r = {
    id: _item.id,
    name: _item.displayName,
    title: _item.displayName,
    email: _item.userPrincipalName
  };

  if (!_item.userPrincipalName) {
    if (_item.scoredEmailAddresses && _item.scoredEmailAddresses.length > 0) {
      r.email = _item.scoredEmailAddresses[0].address;
    }
  }

  if (r.email) {
    r.id = r.email;
  }

  return r;
}
