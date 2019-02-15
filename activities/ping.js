'use strict';

const handleError = require('@adenin/cf-activity').handleError;
const api = require('./common/api');

module.exports = async function (activity) {
    try {
        api.initialize(activity);

        const response = await api('/me');

        activity.Response.Data = {
            success: response && response.statusCode === 200
        };
    } catch (error) {
        handleError(error, activity);
        activity.Response.Data.success = false;
    }
};
