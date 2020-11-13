'use strict';

const api = require('./common/api');

module.exports = async (activity) => {
  try {
    if (!activity.Request.Query.pageSize) activity.Request.Query.pageSize = 5;

    const pages = $.pagination(activity);
    const skip = (pages.page - 1) * pages.pageSize;
    const top = pages.pageSize;

    api.initialize(activity);

    const response = await api(`/v1.0/me/onenote/notebooks?$top=${top}&$skip=${skip}`);
    const items = [];

    for (let i = 0; i < response.body.value.length; i++) {
      const raw = response.body.value[i];

      items.push({
        id: raw.id,
        title: raw.displayName,
        type: 'OneNote',
        date: raw.lastModifiedDateTime,
        lastModified: raw.lastModifiedDateTime,
        link: raw.links.oneNoteWebUrl.href,
        hideType: true
      });
    }

    const count = items.length;

    activity.Response.Data = Object.assign(activity.Response.Data, {
      title: T(activity, 'My Recent Notebooks'),
      link: 'https://office.com/launch/onenote',
      linkLabel: 'Go to OneNote',
      thumbnail: activity.Context.connector.host.connectorLogoUrl,
      value: count,
      items: items.sort($.compare.dateDescending),
      _card: {
        type: 'cloud-files'
      }
    });

    if (parseInt(pages.page) === 1 && count > 0) {
      activity.Response.Data.items[0]._isInitial = true;

      const first = items[0];

      activity.Response.Data.date = first.date;
      activity.Response.Data.description = count > 1 ? `You have ${count} recent NoteBooks.` : 'You have 1 recent Notebook.';
      activity.Response.Data.briefing = activity.Response.Data.description + ` The last used was '${first.title}'`;
    } else {
      activity.Response.Data.description = T(activity, 'You have no recent Notebooks.');
    }
  } catch (error) {
    $.handleError(activity, error);
  }
};
