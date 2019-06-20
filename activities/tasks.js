'use strict';
const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);
    let allTasks = [];

    // Since api is currently in beta there is no way filter by daterange in request, this should be
    // FIXED when api is updated

    // Also can't filter for assignee

    let url = `/beta/me/outlook/tasks?$count=true&$top=1000&$orderby=createdDateTime desc`;
    let response = await api(url);
    if ($.isErrorResponse(activity, response)) return;
    allTasks.push(...response.body.value);

    let hasMore = response.body['@odata.nextLink'] || null;
    while (hasMore) {
      url = response.body['@odata.nextLink'];
      response = await api(url);
      if ($.isErrorResponse(activity, response)) return;
      allTasks.push(...response.body.value);

      hasMore = response.body['@odata.nextLink'] || null;
    }

    const daterange = $.dateRange(activity);
    let tasks = api.filterByDateRange(allTasks, daterange);
    tasks = convertResponse(tasks);
    let value = tasks.length;
    let dateToAssign = tasks.length > 0 ? tasks[0].date : null;

    let pagination = $.pagination(activity);
    tasks = api.paginateItems(tasks, pagination);

    // const value = response.body['@odata.count'];

    activity.Response.Data.items = tasks;
    if (parseInt(pagination.page) == 1) {
      activity.Response.Data.title = T(activity, 'Active Tasks');
      activity.Response.Data.link = `https://outlook.office365.com/owa/?modurl=0&path=/tasks`;
      activity.Response.Data.linkLabel = T(activity, 'All Tasks');
      activity.Response.Data.actionable = value > 0;

      if (value > 0) {
        activity.Response.Data.value = value;
        activity.Response.Data.date = new Date(dateToAssign).toISOString();
        activity.Response.Data.color = 'blue';
        activity.Response.Data.description = value > 1 ? T(activity, "You have {0} tasks.", value)
          : T(activity, "You have 1 task.");
      } else {
        activity.Response.Data.description = T(activity, `You have no tasks.`);
      }
    }
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

  return items;
}