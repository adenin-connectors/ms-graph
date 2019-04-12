'use strict';

const api = require('./common/api');
const cfActivity = require('@adenin/cf-activity');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    var dateRange = cfActivity.dateRange(activity,"today");

    const url = `/v1.0/me/calendarview?count=true&startdatetime=${dateRange.startDate}&enddatetime=${dateRange.endDate}`;
    const response = await api(url);
    if (!cfActivity.isResponseOk(activity, response)) {
      return;
    }

    let taskCount = response.body["@odata.count"];

    let taskStatus = {
      title: 'Active Tasks',
      url: 'https://outlook.office365.com/owa/?modurl=0&path=/tasks',
      urlLabel: 'All tasks',
    };

    if (taskCount != 0) {
      taskStatus = {
        ...taskStatus,
        description: `You have ${taskCount > 1 ? taskCount + " tasks" : taskCount + " task"} assigned.`,
        color: 'blue',
        value: taskCount,
        actionable: true
      };
    } else {
      taskStatus = {
        ...taskStatus,
        description: `You have no tasks today.`,
        actionable: false
      };
    }

    activity.Response.Data = taskStatus;

  } catch (error) {
    // handle generic exception
    api.handleError(activity, error);
  }
};