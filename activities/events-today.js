'use strict';

const moment = require('moment-timezone');

const api = require('./common/api');
const helpers = require('./common/helpers');

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/v1.0/me/events');

    if ($.isErrorResponse(activity, response)) return;

    moment.tz.setDefault(activity.Context.UserTimezone);

    const today = moment().tz(activity.Context.UserTimezone).utc();
    const items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      let raw = response.body.value[i];

      const eventDate = moment(raw.start.dateTime).tz(raw.start.timeZone).utc();

      if (raw.recurrence && !today.isSame(eventDate, 'date')) {
        raw = await resolveRecurrence(raw.id);

        if (!raw) continue;
      }

      const item = convertItem(raw);

      const endDate = moment(raw.end.dateTime).tz(raw.end.timeZone).utc();
      const overAnHourAgo = today.clone().minutes(today.minutes() - 61);

      if (today.isSame(eventDate, 'date') && endDate.isAfter(overAnHourAgo)) items.push(item);
    }

    const value = items.length;

    const pagination = $.pagination(activity);

    activity.Response.Data.items = paginateItems(items, pagination);

    if (parseInt(pagination.page) === 1) {
      activity.Response.Data.title = T(activity, 'Events Today');
      activity.Response.Data.link = 'https://outlook.office365.com/mail/inbox';
      activity.Response.Data.linkLabel = T(activity, 'All events');
      activity.Response.Data.thumbnail = 'https://www.adenin.com/assets/images/wp-images/logo/office-365.svg';
      activity.Response.Data.actionable = value > 0;

      if (value > 0) {
        const first = activity.Response.Data.items[0];

        activity.Response.Data.value = value;
        activity.Response.Data.date = first.date;
        activity.Response.Data.description = value > 1 ? `You have ${value} events today.` : 'You have 1 event today.';

        const when = moment().to(moment(first.date));

        activity.Response.Data.briefing = activity.Response.Data.description + ` The next is '${first.title}' ${when}`;
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
};

function convertItem(raw) {
  const item = {
    id: raw.id,
    title: raw.subject,
    link: raw.webLink,
    isCancelled: raw.isCancelled,
    raw: raw
  };

  item.description = raw.bodyPreview.replace(/\r/g, '');
  item.description = item.description.replace(/\n/g, '<br/>');

  item.date = moment(raw.start.dateTime).tz(raw.start.timeZone).utc().format();
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

  item.organizer = {
    avatarProperties: parseAvatarProperties(raw.organizer.emailAddress.name),
    email: raw.organizer.emailAddress.address,
    name: helpers.stripSpecialChars(raw.organizer.emailAddress.name)
  };

  item.attendees = [];

  if (raw.attendees.length > 0) {
    for (let j = 0; j < raw.attendees.length; j++) {
      const attendee = {
        email: raw.attendees[j].emailAddress.address,
        name: helpers.stripSpecialChars(raw.attendees[j].emailAddress.name),
        avatarProperties: parseAvatarProperties(raw.attendees[j].emailAddress.name)
      };

      item.attendees.push(attendee);
    }
  } else {
    item.attendees = null;
  }

  if (raw.responseStatus.response !== 'none') {
    item.response = {
      status: raw.responseStatus.response === 'accepted' ? 'accepted' : 'declined',
      date: raw.responseStatus.time
    };
  }

  item.showDetails = false;

  return item;
}

const colors = ['blue', 'green', 'orange', 'pink', 'purple', 'red', 'teal'];

function parseAvatarProperties(name) {
  let colorCode = 1;

  for (let k = 0; k < name.length; k++) {
    colorCode += name.charCodeAt(k);
  }

  const names = name.split(' ');
  let initials = '';

  // stop after two initials
  for (let k = 0; k < names.length && k < 2; k++) {
    initials += names[k].charAt(0);
  }

  return {
    initials: initials.toUpperCase(),
    color: colors[colorCode % colors.length]
  };
}

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

function paginateItems(items, pagination) {
  const paginatedItems = [];
  const pageSize = parseInt(pagination.pageSize);
  const offset = (parseInt(pagination.page) - 1) * pageSize;

  if (offset > items.length) return paginatedItems;

  for (let i = offset; i < offset + pageSize; i++) {
    if (i >= items.length) break;

    paginatedItems.push(items[i]);
  }

  return paginatedItems;
}
