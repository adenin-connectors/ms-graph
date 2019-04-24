'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/v1.0/me');

    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data = {
      success: response && response.statusCode === 200
    };
  } catch (error) {
    api.handleError(activity, error);

    activity.Response.Data.success = false;
  }
};
