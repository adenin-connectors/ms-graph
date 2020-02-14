'use strict';

const crypto = require('crypto');
const moment = require('moment-timezone');

const api = require('./common/api');
const helpers = require('./common/helpers');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/v1.0/me/events');

    if ($.isErrorResponse(activity, response)) return;

    moment.tz.setDefault(activity.Context.UserTimezone);

    const now = moment().utc();
    const items = [];

    let count = 0;
    let firstFutureIndex = null;

    for (let i = 0; i < response.body.value.length; i++) {
      let raw = response.body.value[i];
      let eventDate = moment.tz(raw.start.dateTime, raw.start.timeZone).tz(activity.Context.UserTimezone).utc();

      if (raw.recurrence && !now.isSame(eventDate, 'date')) {
        raw = await resolveRecurrence(raw.id);

        if (!raw) continue;

        eventDate = moment.tz(raw.start.dateTime, raw.start.timeZone).tz(activity.Context.UserTimezone).utc();
      }

      const item = convertItem(raw, activity);

      const endDate = moment.tz(raw.end.dateTime, raw.end.timeZone).tz(activity.Context.UserTimezone).utc();
      const overAnHourAgo = now.clone().minutes(now.minutes() - 61);

      // comment out code that allows multi-day events to pass through until we decide if/how to display them
      //if (endDate.isAfter(overAnHourAgo) && (today.isSame(eventDate, 'date') || today.isAfter(eventDate, 'date'))) items.push(item);

      if (now.isSame(eventDate, 'date') && endDate.isAfter(overAnHourAgo)) {
        items.push(item);

        if (eventDate.isAfter(now)) {
          count++;

          if (!firstFutureIndex) firstFutureIndex = items.length - 1;
        }
      }
    }

    const pagination = $.pagination(activity);
    const paginatedItems = api.paginateItems(items, pagination);

    activity.Response.Data.items = paginatedItems;
    activity.Response.Data._hash = crypto.createHash('md5').update(JSON.stringify(paginatedItems)).digest('hex');

    if (parseInt(pagination.page) === 1) {
      activity.Response.Data.title = T(activity, 'Events Today');
      activity.Response.Data.link = 'https://outlook.office365.com/mail/inbox';
      activity.Response.Data.linkLabel = T(activity, 'All events');
      activity.Response.Data.thumbnail = 'https://www.adenin.com/assets/images/wp-images/logo/office-365.svg';
      activity.Response.Data.actionable = count > 0;

      if (count > 0) {
        const first = activity.Response.Data.items[firstFutureIndex];

        activity.Response.Data.value = paginatedItems.length;
        activity.Response.Data.date = first.date;
        activity.Response.Data.description = paginatedItems.length > 1 ? `You have ${paginatedItems.length} events today.` : 'You have 1 event today.';

        activity.Response.Data.briefing = activity.Response.Data.description + ` The next is '${first.title}' at ${moment(first.date).format('LT')}`;
      } else {
        activity.Response.Data.description = T(activity, 'You have no events today.');
      }
    }
  } catch (error) {
    api.handleError(activity, error);
  }

  async function resolveRecurrence(eventId) {
    try {
      const start = moment().utc().startOf('day').format();
      const end = moment().utc().endOf('day').format();

      const endpoint = `/v1.0/me/events/${eventId}/instances?startDateTime=${start}&endDateTime=${end}`;

      const response = await api(endpoint);

      if ($.isErrorResponse(activity, response)) return null;

      return response.body.value[0]; // can only recur once per day
    } catch (error) {
      api.handleError(activity, error);
    }
  }

  function convertItem(raw) {
    const item = {
      id: raw.id,
      title: raw.subject,
      link: raw.webLink,
      isCancelled: raw.isCancelled,
      isRecurring: raw.recurrence ? true : false
    };

    item.description = raw.bodyPreview.replace(/\r/g, '');
    item.description = item.description.replace(/\n/g, '<br/>');

    item.date = moment.tz(raw.start.dateTime, raw.start.timeZone).tz(activity.Context.UserTimezone).format();
    item.endDate = moment.tz(raw.end.dateTime, raw.end.timeZone).tz(activity.Context.UserTimezone).format();
    item.duration = moment.duration(moment(raw.end.dateTime).diff(moment(raw.start.dateTime))).humanize();

    if (raw.location && raw.location.coordinates && raw.location.coordinates.latitude) {
      item.location = {
        link: `https://www.google.com/maps/search/?api=1&query=${raw.location.coordinates.latitude},${raw.location.coordinates.longitude}`,
        title: raw.location.displayName
      };
    } else if (raw.location.displayName && !raw.onlineMeetingUrl) {
      const url = parseUrl(raw.location.displayName);

      if (url !== null) {
        item.onlineMeetingUrl = url;
        item.location = null;
      }
    } else {
      item.location = null;
    }

    if (!item.onlineMeetingUrl) {
      const url = parseUrl(item.description);

      if (url !== null) item.onlineMeetingUrl = url;
    }

    item.thumbnail = $.avatarLink(raw.organizer.emailAddress.name, raw.organizer.emailAddress.address);
    item.imageIsAvatar = true;

    item.organizer = {
      avatar: item.thumbnail,
      email: raw.organizer.emailAddress.address,
      name: helpers.stripSpecialChars(raw.organizer.emailAddress.name)
    };

    item.attendees = [];

    if (raw.attendees.length > 0) {
      for (let j = 0; j < raw.attendees.length; j++) {
        const attendee = {
          email: raw.attendees[j].emailAddress.address,
          name: helpers.stripSpecialChars(raw.attendees[j].emailAddress.name),
          avatar: $.avatarLink(raw.attendees[j].emailAddress.name, raw.attendees[j].emailAddress.address)
        };

        if (attendee.email === item.organizer.email) {
          attendee.response = 'organizer';
        } else if (raw.attendees[j].status.response !== 'none') {
          attendee.response = raw.attendees[j].status.response === 'accepted' ? 'accepted' : 'declined';
        }

        item.attendees.push(attendee);
      }
    } else {
      item.attendees = null;
    }

    switch (raw.responseStatus.response) {
    case 'none':
      break;
    case 'organizer':
      item.response = {
        status: 'organizer'
      };

      break;
    default:
      item.response = {
        status: raw.responseStatus.response === 'accepted' ? 'accepted' : 'declined',
        date: raw.responseStatus.time
      };
    }

    return item;
  }
};

const urlRegex = /[-a-zA-Z0-9@:%_\+.~#?&//=]{2,256}\.[a-z]{2,4}\b(\/[-a-zA-Z0-9@:%_\+.~#?&//=]*)?/gi;

function parseUrl(text) {
  text = text.replace(/\n|\r/g, ' ');

  if (text.search(urlRegex) !== -1) {
    let url = text.substring(text.search(urlRegex), text.length);

    if (url.indexOf(' ') !== -1) url = url.substring(0, url.indexOf(' '));
    if (!url.match(/^[a-zA-Z]+:\/\//)) url = 'https://' + url;

    return url;
  }

  return null;
}
