'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/v1.0/me/memberOf');

    if (activity.isErrorResponse(response)) return;

    activity.Response.Data.items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      activity.Response.Data.items.push(convertItem(response.body.value[i]));
    }
  } catch (error) {
    api.handleError(activity, error);
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
