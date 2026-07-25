'use strict';

const path = require('path');

function validateArchiveEntries(entries) {
  if(!Array.isArray(entries) || !entries.length) throw new Error('Archive is empty');
  entries.forEach((entry) => {
    const value = String(entry || '').replace(/\\/g, '/');
    if(!value || value[0] === '/' || /^[A-Za-z]:/.test(value)) throw new Error('Archive contains an absolute path');
    const normalized = path.posix.normalize(value);
    if(normalized === '..' || normalized.startsWith('../') || normalized.includes('/../')) throw new Error('Archive path traversal rejected');
  });
  return true;
}

module.exports = {validateArchiveEntries};
