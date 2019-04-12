'use strict';

const api = require('./common/api');

module.exports = async () => {
  try {
    const response = await api('/v1.0/me/memberOf');

    if (response.statusCode === 200 && response.body.value && response.body.value.length > 0) {
      Activity.Response.Data.items = [];

      for (let i = 0; i < response.body.value.length; i++) {
        Activity.Response.Data.items.push(convertItem(response.body.value[i]));
      }
    } else {
      Activity.Response.Data = {
        statusCode: response.statusCode,
        message: 'Bad request or no group memberships found',
        items: []
      };
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
