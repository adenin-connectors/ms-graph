'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);
    const pagination = $.pagination(activity);
    let url = `/beta/me/outlook/tasks?$count=true&$top=${pagination.pageSize}`;
    if (pagination.nextpage) {
      url = pagination.nextpage;
    }
    const response = await api(url);
    if ($.isErrorResponse(activity, response)) return;

    const value = response.body['@odata.count'];

    activity.Response.Data.items = convertResponse(response.body.value);
    activity.Response.Data.title = T(activity, 'Active Tasks');
    activity.Response.Data.link = `https://outlook.office365.com/owa/?modurl=0&path=/tasks`;
    activity.Response.Data.linkLabel = T(activity, 'All Tasks');
    activity.Response.Data.actionable = value > 0;

    if (value > 0) {
      activity.Response.Data.value = value;
      activity.Response.Data.color = 'blue';
      activity.Response.Data.description = value > 1 ? T(activity, "You have {0} tasks.", value)
        : T(activity, "You have 1 task.");
    } else {
      activity.Response.Data.description = T(activity, `You have no tasks.`);
    }
    activity.Response.Data._nextpage = response.body['@odata.nextLink'];
  } catch (error) {
    api.handleError(activity, error);
  }
};

//**maps resposne data to items */
function convertResponse(tasks) {
  let items = [];

  for (let i = 0; i < tasks.length; i++) {
    let raw = tasks[i];
    let item = {
      id: raw.id,
      name: raw.subject,
      description: raw.importance,
      date: raw.createdDateTime,
      link: 'https://to-do.office.com/tasks/',
      raw: raw
    };
    items.push(item);
  }

  return { items };
}