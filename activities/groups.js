'use strict';

const api = require('./common/api');

module.exports = async () => {
  try {
    const response = await api('/v1.0/me/memberOf');

    if (Activity.isErrorResponse(response)) return;

    Activity.Response.Data.items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      Activity.Response.Data.items.push(convertItem(response.body.value[i]));
    }
  } catch (error) {
    api.handleError(Activity, error);
  }
};

function convertItem(_item) {
  return {
    id: _item.id,
    title: _item.displayName,
    description: _item.description,
    email: _item.mail,
    date: _item.createdDateTime
  };
}
