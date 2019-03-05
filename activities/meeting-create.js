'use strict';

const path = require('path');
const api = require('./common/api');
const yaml = require('js-yaml');
const fs = require('fs');
const logger = require('@adenin/cf-logger');
const cfActivity = require('@adenin/cf-activity');

module.exports = async (activity) => {

    try {
        api.initialize(activity);
        var data = {};

        // extract _action from Request
        var _action = getObjPath(activity.Request, "Data.model._action");
        if (_action) {
            activity.Request.Data.model._action = {};
        } else {
            _action = {};
        }

        switch (activity.Request.Path) {

            case "create":
            case "submit":

                const form = _action.form;
                var body = {};

                // get attendees[]
                var attendees = _action.form.attendees.split(",");

                // ensure that current user is part of attendees
                if (attendees.indexOf(activity.Context.UserEmail) < 0) attendees.push(activity.Context.UserEmail);

                var meetingAttendees = [];

                attendees.forEach(element => {
                    meetingAttendees.push({
                        emailAddress: {
                            address: element
                        },
                        "type": "Required"
                    });
                });


                // auto schedule ?
                if (form.auto != "2") {

                    body.attendees = meetingAttendees;
                    body.meetingDuration = form.duration;

                    var timeslots = form.daterange.split("/");
                    body.timeConstraint = {
                        "timeslots": [{
                            "start": {
                                "dateTime": timeslots[0] + "T00:00:00Z",
                                "timeZone": "UTC"
                            },
                            "end": {
                                "dateTime": timeslots[1] + "T23:59:59Z",
                                "timeZone": "UTC"
                            }
                        }]
                    };

                    var response = await api.post("/v1.0/me/findMeetingTimes", {
                        json: true,
                        body: body
                    });

                }

                // return failure when API returns emptySuggestionsReason
                var comment = response.body.emptySuggestionsReason;
                if (comment) {
                    data = getObjPath(activity.Request, "Data.model");
                    data._action = {
                        response: {
                            success: false,
                            message: comment
                        }
                    };
                    break;
                }

                // use suggested timeslot
                var timeslot = response.body.meetingTimeSuggestions[0].meetingTimeSlot;

                body = {
                    subject: form.subject,
                    body: {
                        contentType: "Text",
                        content: form.description
                    },
                };
                body.attendees = meetingAttendees;
                body.start = timeslot.start;
                body.end = timeslot.end;

                var response = await api.post("/v1.0/me/events", {
                    json: true,
                    body: body
                });

                var comment = "Meeting scheduled";
                data = getObjPath(activity.Request, "Data.model");
                data._action = {
                    response: {
                        success: true,
                        message: comment,
                        datetime: response.body.start.dateTime
                    }
                };
                break;


            default:
                // initialize form subject with query parameter (if provided)
                var fname = activity.Context.ScriptFolder + path.sep + '/common/meeting-create.form';
                var schema = yaml.safeLoad(fs.readFileSync(fname, 'utf8'));

                // provide lookup url for attendees 
                schema.properties.attendees.url = activity.Context.connector.baseurl + "/people";
                schema.properties.daterange.hide = false;
                schema.properties.starttime.hide = true;

                // initialize start date to next hour
                var d = new Date();
                d.setHours(d.getHours() + 1);
                d.setMinutes(0, 0, 0);

                // initialize tomorrow
                var dt = new Date();
                dt.setHours(0, 0, 0, 0);
                dt.setDate(dt.getDate() + 1);

                var dtn = new Date(dt);
                dtn.setDate(dtn.getDate() + 30);


                // return form schema
                data.formSchema = schema;

                // set form default value
                data.form = {};
                data.form.starttime = d.toISOString();
                data.form.auto = '1';
                data.form.daterange = dt.toISOString().substring(0, 10) + "/" + dtn.toISOString().substring(0, 10);

                if (activity.Request.Query && activity.Request.Query.query) {
                    data.form.subject = activity.Request.Query.query;
                }

                break;
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

        var paths = path.split('.'),
            current = obj;

        for (var i = 0; i < paths.length; ++i) {
            if (current[paths[i]] == undefined) {
                return undefined;
            } else {
                current = current[paths[i]];
            }
        }
        return current;
    }
};
