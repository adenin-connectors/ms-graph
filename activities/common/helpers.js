'use strict';

const specialCharRegex = /[^a-zA-z\s]/;

module.exports = {
  stripSpecialChars: (input) => {
    const index = input.search(specialCharRegex);

    if (index === -1) return input;

    return input.substring(0, index).trim();
  },
  convertInsightsItem: function (raw) {
    return {
      id: raw.id,
      title: raw.resourceVisualization.title,
      description: raw.resourceVisualization.previewText,
      type: raw.resourceVisualization.type || raw.resourceVisualization.containerType,
      link: raw.resourceReference.webUrl,
      preview: raw.resourceVisualization.previewImageUrl,
      containerTitle: this.stripSpecialChars(raw.resourceVisualization.containerDisplayName),
      containerLink: raw.resourceVisualization.containerWebUrl,
      containerType: raw.resourceVisualization.containerType,
      date: raw.lastUsed ? (new Date(raw.lastUsed.lastAccessedDateTime)).toISOString() : undefined,
      raw: raw
    };
  }
};
