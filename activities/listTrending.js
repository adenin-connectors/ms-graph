'use strict';

const {handleError} = require('@adenin/cf-activity');
const api = require('./common/api');

const PAGE_SIZE = 5;

let action = null;
let page = null;
let pageSize = null;

module.exports = async function (activity) {
  try {
    api.initialize(activity);

    configureRange();

    const skip = (page - 1) * pageSize;
    const top = pageSize;

    const response = await api('/beta/me/insights/trending?$skip=' + skip + '&$top=' + top);

    activity.Response.Data.items = response.body.value;

    activity.Response.Data._action = action;
    activity.Response.Data._page = page;
    activity.Response.Data._pageSize = pageSize;
  } catch (error) {
    handleError(error, activity);
  }

  function configureRange() {
    action = 'firstpage';
    page = parseInt(activity.Request.Query.page, 10) || 1;
    pageSize = parseInt(activity.Request.Query.pageSize, 10) || PAGE_SIZE;

    if (
      activity.Request.Data &&
      activity.Request.Data.args &&
      activity.Request.Data.args.atAgentAction === 'nextpage'
    ) {
      action = 'nextpage';
      page = parseInt(activity.Request.Data.args._page, 10) || 2;
      pageSize = parseInt(activity.Request.Data.args._pageSize, 10) || PAGE_SIZE;
    }
  }
};