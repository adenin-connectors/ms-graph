'use strict';

const api = require('./common/api');

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

            case "find":
            case "submit":

                const form = _action.form;
                var body = {};

                // get attendees[]
                var attendees = _action.form.attendees.split(",");

                var meetingAttendees = [];

                attendees.forEach(element => {
                    meetingAttendees.push({
                        emailAddress: {
                            address: element
                        },
                        "type": "Required"
                    });
                });


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



                // return failure when API returns emptySuggestionsReason
                data._action = {};
                data.response = true;
                data.message = response.body.emptySuggestionsReason;
                data.meetingTimeSuggestions = response.body.meetingTimeSuggestions;

                break;


            default:

                // initialize date range
                var dt = new Date();
                dt.setHours(0, 0, 0, 0);
                dt.setDate(dt.getDate() + 1);

                var dtn = new Date(dt);
                dtn.setDate(dtn.getDate() + 30);

                // set form default value
                data.form = {};
                data.form.daterange = dt.toISOString().substring(0, 10) + "/" + dtn.toISOString().substring(0, 10);

                if (activity.Request.Query && activity.Request.Query.query) {
                    data.form.subject = activity.Request.Query.query;
                }

                if (!data._card) data._card = {};
                data._card.actionList = [];
                var action = {
                    id: "find",
                    label: "Find Meeting Time",
                    settings: {
                        actionType: "api",
                        buttonType: 1
                    }
                };

                data._card.actionList.push(action);
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
