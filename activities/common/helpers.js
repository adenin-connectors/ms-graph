'use strict';

const specialCharRegex = /[^a-zA-z\s]/;

module.exports = {
  stripSpecialChars: (input) => {
    const index = input.search(specialCharRegex);

    if (index === -1) return input;

    return input.substring(0, index);
  }
};
