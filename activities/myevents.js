'use strict';
const api = require('./common/api');
const moment = require('moment-timezone');

module.exports = async (activity) => {
  try {
    api.initialize(activity);
    const dateRange = $.dateRange(activity, 'today');
    const pagination = $.pagination(activity);
    // eslint-disable-next-line max-len
    let url = `/v1.0/me/calendarview?count=true&startdatetime=${dateRange.startDate}&enddatetime=${dateRange.endDate}&$top=${pagination.pageSize}`;
    if (pagination.nextpage) {
      url = pagination.nextpage;
    }
    const response = await api(url);
    if ($.isErrorResponse(activity, response)) return;

    activity.Response.Data.items = convertResponse(response);
    activity.Response.Data.title = T(activity, 'Events Today');
    activity.Response.Data.link = `https://outlook.office365.com/calendar/view/month`;
    activity.Response.Data.linkLabel = T(activity, 'All events');
    const value = response.body['@odata.count'];
    activity.Response.Data.actionable = value > 0;

    if (value > 0) {
      const nextEvent = getNexEvent(response.body.value);

      const eventFormatedTime = getEventFormatedTimeAsString(activity, nextEvent);
      const eventPluralorNot = value > 1 ? T(activity, "events scheduled") : T(activity, "event scheduled");
      const description = T(activity, `You have {0} {1} today. The next event '{2}' starts {3}`, value, eventPluralorNot, nextEvent.subject, eventFormatedTime);

      activity.Response.Data.value = value;
      activity.Response.Data.color = 'blue';
      activity.Response.Data.description = description;
    } else {
      activity.Response.Data.description = T(activity, `You have no events today.`);
    }
    if (response.body["@odata.nextLink"]) activity.Response.Data._nextpage = response.body["@odata.nextLink"];
  } catch (error) {
    $.handleError(activity, error);
  }
};
// convert response from /issues endpoint to
function convertResponse(response) {
  const items = [];
  const events = response.body.value;

  // iterate through each issue and extract id, title, etc. into a new array
  for (let i = 0; i < events.length; i++) {
    const raw = events[i];
    const item = {
      id: raw.id,
      title: raw.subject,
      description: raw.bodyPreview,
      date: raw.createdDateTime,
      link: `https://outlook.office365.com/calendar/item/${raw.id}`,
      raw: raw
    };

    items.push(item);
  }

  return { items };
}
/**filters out first upcoming event in google calendar*/
function getNexEvent(events) {
  let nextEvent = null;
  let nextEventMilis = 0;

  for (let i = 0; i < events.length; i++) {
    let tempDate = Date.parse(events[i].start.dateTime);

    if (nextEventMilis == 0) {
      nextEventMilis = tempDate;
      nextEvent = events[i];
    }

    if (nextEventMilis > tempDate) {
      nextEventMilis = tempDate;
      nextEvent = events[i];
    }
  }

  return nextEvent;
}
//** checks if event is in less then hour, today or tomorrow and returns formated string accordingly */
function getEventFormatedTimeAsString(activity, nextEvent) {
  const eventTime = moment(nextEvent.start.dateTime)
    .tz(activity.Context.UserTimezone)
    .locale(activity.Context.UserLocale);
  const timeNow = moment(new Date());

  const diffInHrs = eventTime.diff(timeNow, 'hours');

  if (diffInHrs == 0) {
    //events that start in less then 1 hour
    let diffInMins = eventTime.diff(timeNow, 'minutes');
    return T(activity, `in {0} minutes.`, diffInMins);
  } else {
    //events that start in more than 1 hour
    const diffInDays = eventTime.diff(timeNow, 'days');

    let datePrefix = '';
    let momentDate = '';
    if (diffInDays == 1) {
      //events that start tomorrow
      datePrefix = 'tomorrow ';
    } else if (diffInDays > 1) {
      //events that start day after tomorrow and later
      datePrefix = 'on ';
      momentDate = eventTime.format('LL') + " ";
    }

    return T(activity, `{0}{1}{2}{3}.`, T(activity, datePrefix), momentDate, T(activity, "at "), eventTime.format('LT'));
  }
}