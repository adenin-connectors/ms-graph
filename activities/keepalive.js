'use strict';

const utils = require('./common/utils');

module.exports = async (activity) => {
    try {
        activity.Response.Data = new Date().toISOString();
    } catch (error) {
        utils.handleError(error, activity);
    }
};
