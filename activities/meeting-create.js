'use strict';

const path = require('path');
const api = require('./common/api');
const yaml = require('js-yaml');
const fs = require('fs');
const moment = require('moment-timezone');


module.exports = async (activity) => {

    try {
        api.initialize(activity);
        var data = {};
        var form = {};
        var body = {};

        // extract _action from Request
        var _action = getObjPath(activity.Request, "Data.model._action");
        if (_action) {
            activity.Request.Data.model._action = {};
        } else {
            _action = {};
        }

        if (activity.Request.Path) {

            form = _action.form;

            // get attendees[]
            var attendees = _action.form.attendees.split(",");

            // ensure that current user is part of attendees
            if (activity.Context.UserEmail && attendees.indexOf(activity.Context.UserEmail) < 0) attendees.push(activity.Context.UserEmail);

            var meetingAttendees = [];

            attendees.forEach(element => {
                meetingAttendees.push({
                    emailAddress: {
                        address: element
                    },
                    "type": "Required"
                });
            });
        }


        switch (activity.Request.Path) {
            
            case "find": // --- find meeting times ----------

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


                data._actionList = [{
                    id: "find",
                    label: "Find Meeting Time",
                    settings: {
                        actionType: "api"
                    }
                }, {
                    id: "schedule",
                    label: "Schedule Meeting",
                    settings: {
                        actionType: "a"
                    }
                }];


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

                var schema = readSchema(activity);

                schema.properties.meetingtimes.xvaluelist = [];
                var tz = activity.Context.UserTimezone||"America/New_York";
                
                // convert returned suggestions
                for(var mi=0;mi<response.body.meetingTimeSuggestions.length;mi++) {
                  var meetingTimeSuggestion = response.body.meetingTimeSuggestions[mi];                
                    schema.properties.meetingtimes.xvaluelist.push({
                        //value: meetingTimeSuggestion.meetingTimeSlot.start.dateTime + "|" + meetingTimeSuggestion.meetingTimeSlot.end.dateTime,
                        value: JSON.stringify(meetingTimeSuggestion.meetingTimeSlot),
                        title: moment(meetingTimeSuggestion.meetingTimeSlot.start.dateTime+"Z").tz(tz).format('lll')
                    });
                };

                schema.properties.meetingtimes.hide = schema.properties.meetingtimes.xvaluelist.length == 0;
                if (schema.properties.meetingtimes.xvaluelist.length > 0) form.meetingtimes = schema.properties.meetingtimes.xvaluelist[0].value;

                // return form schema & form value
                data._formSchema = schema;
                data.form = form;

                break;


            case "schedule": // --- schedule meeting ----------

                // use suggested timeslot
                var timeslot = JSON.parse(_action.form.meetingtimes);

                // get attendees[]
                var attendees = _action.form.attendees.split(",");

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

                var schema = readSchema(activity);
                schema.properties.meetingtimes.hide = true;

                // return form schema
                data._formSchema = schema;

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



                // set form default value
                data.form = {};
                data.form.starttime = d.toISOString();
                data.form.auto = '1';
                data.form.daterange = dt.toISOString().substring(0, 10) + "/" + dtn.toISOString().substring(0, 10);


                // initialize form subject with query parameter (if provided)
                if (activity.Request.Query && activity.Request.Query.query) {
                    data.form.subject = activity.Request.Query.query;
                }

                // initialize actions                
                data._actionList = [{
                    id: "find",
                    label: "Find Meeting Time",
                    settings: {
                        actionType: "api"
                    }
                }]


                break;
        }

        // copy response data
        activity.Response.Data = data;


    } catch (error) {
        // handle generic exception
        api.handleError(activity, error);
    }


    function readSchema(activity) {

        var fname = activity.Context.ScriptFolder + path.sep + '/common/meeting-create.form';
        var schema = yaml.safeLoad(fs.readFileSync(fname, 'utf8'));

        // provide lookup url for attendees 
        schema.properties.attendees.url = activity.Context.connector.baseurl + "/people-search";

        return schema;
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
