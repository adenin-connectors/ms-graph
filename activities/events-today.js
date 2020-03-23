'use strict';

const crypto = require('crypto');
const moment = require('moment-timezone');

const api = require('./common/api');
const helpers = require('./common/helpers');

module.exports = async (activity) => {
  moment.tz.setDefault(activity.Context.UserTimezone);
  const now = moment().utc();

  try {
    api.initialize(activity);

    const response = await api(`/v1.0/me/calendarView?startDateTime=${now.clone().startOf('day').format()}&endDateTime=${now.clone().endOf('day').format()}`);

    if ($.isErrorResponse(activity, response)) return;

    let items = [];
    let pastCount = 0;
    let allDayCount = 0;

    for (let i = 0; i < response.body.value.length; i++) {
      const raw = response.body.value[i];

      if (raw.isCancelled) continue;

      const item = convertItem(raw, activity);

      const startTime = moment.tz(raw.start.dateTime, raw.start.timeZone).tz(activity.Context.UserTimezone).utc();
      const endTime = moment.tz(raw.end.dateTime, raw.end.timeZone).tz(activity.Context.UserTimezone).utc();
      const halfAnHourAgo = now.clone().minutes(now.minutes() - 30);

      if ((now.isSame(startTime, 'date') || item.isAllDay) && endTime.isAfter(halfAnHourAgo)) {
        if (endTime.isBefore(now)) {
          pastCount++;
          item.isPast = true;
        }

        if (item.isAllDay) allDayCount++;

        items.push(item);
      }
    }

    items = items.sort($.compare.dateAscending);

    const pagination = $.pagination(activity);
    const paginatedItems = api.paginateItems(items, pagination);

    activity.Response.Data.items = paginatedItems;
    activity.Response.Data._hash = crypto.createHash('md5').update(JSON.stringify(paginatedItems)).digest('hex');

    const value = paginatedItems.length - pastCount;

    if (parseInt(pagination.page) === 1) {
      activity.Response.Data.title = T(activity, 'Events Today');
      activity.Response.Data.link = 'https://outlook.office365.com/mail/inbox';
      activity.Response.Data.linkLabel = T(activity, 'All events');
      activity.Response.Data.thumbnail = 'https://www.adenin.com/assets/images/wp-images/logo/office-365.svg';
      activity.Response.Data.actionable = value > 0;
      activity.Response.Data.integration = 'Outlook';
      activity.Response.Data.pastCount = pastCount;
      activity.Response.Data.allDayCount = allDayCount;

      if (value > 0) {
        const first = paginatedItems[pastCount];

        activity.Response.Data.value = value;
        activity.Response.Data.date = first.date;
        activity.Response.Data.description = value > 1 ? `You have ${value} events today.` : 'You have 1 event today.';

        activity.Response.Data.briefing = activity.Response.Data.description + ` The next is '${first.title}' at ${moment(first.date).format('LT')}`;
      } else {
        activity.Response.Data.date = now.startOf('day');
        activity.Response.Data.description = T(activity, 'You have no events today.');
      }
    }
  } catch (error) {
    api.handleError(activity, error);
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

    const duration = moment.duration(moment(raw.end.dateTime).diff(moment(raw.start.dateTime)));

    if (duration.asDays() >= 1) item.isAllDay = true;

    item.duration = duration.humanize();

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
      const url = parseUrl(raw.body.content);

      if (url !== null) item.onlineMeetingUrl = url;
    }

    const organizerName = helpers.stripSpecialChars(raw.organizer.emailAddress.name);

    item.thumbnail = $.avatarLink(organizerName, raw.organizer.emailAddress.address);
    item.imageIsAvatar = true;
    item.organizer = {
      avatar: item.thumbnail,
      email: raw.organizer.emailAddress.address,
      name: organizerName
    };

    item.attendees = [];

    if (raw.attendees.length > 0) {
      for (let j = 0; j < raw.attendees.length; j++) {
        const attendeeName = helpers.stripSpecialChars(raw.attendees[j].emailAddress.name);
        const attendee = {
          email: raw.attendees[j].emailAddress.address,
          name: attendeeName,
          avatar: $.avatarLink(attendeeName, raw.attendees[j].emailAddress.address)
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

  const matches = text.match(urlRegex);

  if (!matches.length) return null;

  for (let i = 0; i < matches.length; i++) {
    const match = matches[i];

    if (
      match.includes('webex') ||
      match.includes('gotomeet') ||
      match.includes('gotomeeting') ||
      match.includes('zoom') ||
      match.includes('skype') ||
      match.includes('teams')
    ) {
      return match;
    }
  }

  return matches[0];
}
