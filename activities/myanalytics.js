'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/beta/me/analytics/activitystatistics');

    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data = {
      items: []
    };

    if (!response.body.value || !response.body.value.length) return;

    for (let i = 0; i < response.body.value.length; i++) {
      activity.Response.Data.items.push(convertItem(response.body.value[i]));
    }
  } catch (error) {
    // handle generic exception
    $.handleError(activity, error);
  }
};

function convertItem(raw) {
  // convert dates to ISO
  raw.startDate = new Date(raw.startDate);
  raw.endDate = new Date(raw.endDate);

  return raw;
}
