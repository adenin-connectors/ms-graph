'use strict';

const moment = require('moment');
const api = require('./common/api');

const urlRegex = /[-a-zA-Z0-9@:%_\+.~#?&//=]{2,256}\.[a-z]{2,4}\b(\/[-a-zA-Z0-9@:%_\+.~#?&//=]*)?/gi;

const dateAscending = (a, b) => {
  a = new Date(a.date);
  b = new Date(b.date);

  return a < b ? -1 : (a > b ? 1 : 0);
};

module.exports = async (activity) => {
  try {
    api.initialize(activity);

    const response = await api('/v1.0/me/events');

    if ($.isErrorResponse(activity, response)) return;

    const today = new Date();
    const items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      let raw = response.body.value[i];

      const rawDate = new Date(raw.start.dateTime);

      if (raw.recurrence && (today.setHours(0, 0, 0, 0) !== rawDate.setHours(0, 0, 0, 0))) {
        raw = await resolveRecurrence(raw.id);

        if (!raw) continue;
      }

      const item = convertItem(raw);
      const eventDate = new Date(item.date);

      if (today.setHours(0, 0, 0, 0) === eventDate.setHours(0, 0, 0, 0)) items.push(item);
    }

    if (items.length > 0) {
      activity.Response.Data.items = items.sort(dateAscending);
    } else {
      activity.Response.Data = {
        items: [],
        message: 'No events found for current date'
      };
    }
  } catch (error) {
    api.handleError(activity, error);
  }

  async function resolveRecurrence(eventId) {
    try {
      const today = new Date();

      const start = (new Date(today.setHours(0, 0, 0, 0))).toISOString();
      const end = (new Date(today.setHours(23, 59, 0, 0))).toISOString();

      const endpoint = `/v1.0/me/events/${eventId}/instances?startDateTime=${start}&endDateTime=${end}`;

      const response = await api(endpoint);

      if ($.isErrorResponse(activity, response)) return null;

      return response.body.value[0]; // can only recur once per day
    } catch (error) {
      api.handleError(activity, error);
    }
  }
};

function convertItem(_item) {
  const item = _item;

  item.date = new Date(_item.start.dateTime).toISOString();

  const _duration = moment.duration(moment(_item.end.dateTime).diff(moment(_item.start.dateTime)));

  let duration = '';

  if (_duration._data.years > 0) duration += _duration._data.years + 'y ';
  if (_duration._data.months > 0) duration += _duration._data.months + 'mo ';
  if (_duration._data.days > 0) duration += _duration._data.days + 'd ';
  if (_duration._data.hours > 0) duration += _duration._data.hours + 'h ';
  if (_duration._data.minutes > 0) duration += _duration._data.minutes + 'm';

  item.duration = duration.trim();

  if (item.location && item.location.coordinates && item.location.coordinates.latitude) {
    const baseUrl = 'https://www.google.com/maps/search/?api=1&query=';

    item.location.link = `${baseUrl}${item.location.coordinates.latitude},${item.location.coordinates.longitude}`;
  } else if (item.location.displayName && !item.onlineMeetingUrl) {
    const url = parseUrl(item.location.displayName);

    if (url !== null) {
      item.onlineMeetingUrl = url;
      item.location = null;
    }
  } else {
    item.location = null;
  }

  if (!item.onlineMeetingUrl) {
    const url = parseUrl(item.bodyPreview);

    if (url !== null) item.onlineMeetingUrl = url;
  }

  // Disable avatars until workable solution found

  /*const basePhotoUri = 'https://outlook.office.com/owa/service.svc/s/GetPersonaPhoto?email=';

  const photoSize = '&size=HR64x64';

  item.organizer.photo = basePhotoUri + _item.organizer.emailAddress.address + photoSize;

  for (let i = 0; i < _item.attendees.length; i++) {
      item.attendees[i].photo = basePhotoUri + _item.attendees[i].emailAddress.address + photoSize;
  }*/

  item.organizer.initials = parseInitials(_item.organizer.emailAddress.name);

  if (_item.attendees.length > 0) {
    for (let j = 0; j < _item.attendees.length; j++) {
      item.attendees[j].initials = parseInitials(_item.attendees[j].emailAddress.name);
    }
  } else {
    item.attendees = null;
  }

  item.showDetails = false;

  return item;
}

function parseInitials(name) {
  const names = name.split(' ');
  let initials = '';

  // stop after two initials
  for (let k = 0; k < names.length && k < 2; k++) {
    initials += names[k].charAt(0);
  }

  return initials.toUpperCase();
}

function parseUrl(text) {
  if (text.search(urlRegex) !== -1) {
    let url = text.substring(text.search(urlRegex), text.length);

    if (url.indexOf(' ') !== -1) url = url.substring(0, url.indexOf(' '));
    if (!url.match(/^[a-zA-Z]+:\/\//)) url = 'https://' + url;

    return url;
  }

  return null;
}
