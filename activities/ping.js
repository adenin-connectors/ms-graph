'use strict';

const api = require('./common/api');
const utils = require('./common/utils');

module.exports = async function (activity) {
    try {
        api.initialize(activity);

        const response = await api('/me');

        activity.Response.Data = {
            success: response && response.statusCode === 200
        };
    } catch (error) {
        utils.handleError(error, activity);
        activity.Response.Data.success = false;
    }
};
