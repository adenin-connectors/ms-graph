'use strict';

const api = require('./common/api');

module.exports = async () => {
  try {
    const response = await api('/v1.0/me');

    if (Activity.isErrorResponse(response)) return;

    Activity.Response.Data = {
      success: response && response.statusCode === 200
    };
  } catch (error) {
    api.handleError(Activity, error);

    Activity.Response.Data.success = false;
  }
};
