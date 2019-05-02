'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/beta/me/outlook/tasks');

    if ($.isErrorResponse(activity, response)) return;

    let tasks = [];

    if (response.body.value.length) tasks = response.body.value;

    let taskStatus = {
      title: T(activity, 'Active Tasks'),
      link: 'https://outlook.office.com/owa/?path=/tasks',
      linkLabel: T(activity, 'All Tasks')
    };

    const count = tasks.length;

    if (count !== 0) {
      taskStatus = {
        ...taskStatus,
        description: count > 1 ? T(activity, 'You have {0} tasks.', count) : T(activity, 'You have 1 task.'),
        color: 'blue',
        value: count,
        actionable: true
      };
    } else {
      taskStatus = {
        ...taskStatus,
        description: T(activity, 'You have no tasks.'),
        actionable: false
      };
    }

    activity.Response.Data = taskStatus;
  } catch (error) {
    api.handleError(activity, error);
  }
};
