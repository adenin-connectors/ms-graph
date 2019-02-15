'use strict';

const handleError = require('@adenin/cf-activity').handleError;
const api = require('./common/api');

module.exports = async (activity) => {
    try {
        api.initialize(activity);

        const response = await api('/me/memberOf');

        if (response.statusCode === 200 && response.body.value && response.body.value.length > 0) {
            activity.Response.Data.items = [];

            for (let i = 0; i < response.body.value.length; i++) {
                activity.Response.Data.items.push(convertItem(response.body.value[i]));
            }
        } else {
            activity.Response.Data = {
                statusCode: response.statusCode,
                message: 'Bad request or no group memberships found',
                items: []
            };
        }
    } catch (error) {
        handleError(error, activity);
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
