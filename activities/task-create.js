'use strict';

const fs = require('fs');
const path = require('path');
const yaml = require('js-yaml');
const api = require('./common/api');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    let data = {};

    // extract _action from Request
    let _action = getObjPath(activity.Request, 'Data.model._action');

    if (_action) {
      activity.Request.Data.model._action = {};
    } else {
      _action = {};
    }

    switch (activity.Request.Path) {
    case 'create':
    case 'submit': {
      const form = _action.form;
      const body = {
        subject: form.subject,
        body: {
          contentType: 'Text',
          content: form.description
        }
      };

      if (form.starttime) body.startDateTime = {
        dateTime: new Date(form.starttime)
      };

      if (form.duetime) body.dueDateTime = {
        dateTime: new Date(form.duetime)
      };

      const response = await api.post('/beta/me/outlook/tasks', {
        json: true,
        body: body
      });

      if ($.isErrorResponse(activity, response, [200, 201])) return;

      const comment = 'Task created';

      data = getObjPath(activity.Request, 'Data.model');
      data._action = {
        response: {
          success: true,
          message: comment
        }
      };

      break;
    }
    default: {
      // initialize form subject with query parameter (if provided)
      if (activity.Request.Query && activity.Request.Query.query) {
        data = {
          form: {
            subject: activity.Request.Query.query
          }
        };
      }

      const fname = __dirname + path.sep + 'common' + path.sep + 'task-create.form';
      const schema = yaml.safeLoad(fs.readFileSync(fname, 'utf8'));

      data.title = T(activity, 'Create Outlook Task');
      data.formSchema = schema;

      // initialize form subject with query parameter (if provided)
      if (activity.Request.Query && activity.Request.Query.query) {
        data = {
          form: {
            subject: activity.Request.Query.query
          }
        };
      }

      data._actionList = [{
        id: 'create',
        label: T(activity, 'Create Task'),
        settings: {
          actionType: 'a'
        }
      }];

      break;
    }
    }

    // copy response data
    activity.Response.Data = data;
  } catch (error) {
    // handle generic exception
    api.handleError(activity, error);
  }

  function getObjPath(obj, path) {
    if (!path) return obj;
    if (!obj) return null;

    const paths = path.split('.');
    let current = obj;

    for (let i = 0; i < paths.length; ++i) {
      if (!current[paths[i]]) {
        return undefined;
      } else {
        current = current[paths[i]];
      }
    }

    return current;
  }
};
