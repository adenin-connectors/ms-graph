'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/beta/me/outlook/tasks');

    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data.items = response.body.value.map((raw) => convertTask(raw));
  } catch (error) {
    api.handleError(activity, error);
  }
};

function convertTask(raw) {
  return {
    id: raw.id,
    title: raw.subject,
    description: raw.body.content,
    raw: raw
  };
}
