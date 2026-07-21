(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.PrayerTemplateProfiles = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const DEFAULT_TEMPLATE_ID = 'prayer';

  const profiles = {
    prayer: {
      id: 'prayer',
      label: '禱告會 PPT 產生器',
      selectorLabel: '禱告會',
      coverTitle: '禱告會',
      filenamePrefix: '禱告會',
      draftKey: 'lkc-prayer-ppt-draft',
      eventTypeName: '禱告會',
      activeSectionId: 'silence',
      defaultBackgroundColor: '#ffffff', // Light background by default
      sections: [
        ['silence', '請安靜心、等候神', 'content', {
          title: '請安靜心、等候神',
          body: ''
        }],
        ['hymn-1', '詩歌一 (華語)', 'praise', {
          title: '',
          body: ''
        }],
        ['scripture', '經文 (台語)', 'bible', {
          title: '經文 (台語)',
          bibleQuery: '',
          bibleVersion: 'tghg',
          bibleRecords: []
        }],
        ['thanksgiving', '獻上感謝讚美', 'list', {
          title: '獻上感謝讚美的 pray.',
          body: ''
        }],
        ['repentance', '悔改認罪', 'list-bible', {
          title: '悔改認罪 pray',
          bibleQuery: '',
          bibleVersion: 'tghg',
          bibleRecords: [],
          body: ''
        }],
        ['world', '為世界 pray', 'list', {
          title: '為世界 pray',
          body: ''
        }],
        ['nation', '為國家、社會 pray', 'list-bible', {
          title: '為國家、社會 pray',
          bibleQuery: '',
          bibleVersion: 'tghg',
          bibleRecords: [],
          body: ''
        }],
        ['church', '為教會 pray', 'list', {
          title: '為教會 pray.',
          body: ''
        }],
        ['members', '為教會肢體 pray', 'list', {
          title: '為教會肢體 pray.',
          body: ''
        }],
        ['oneself', '為自己 pray', 'list', {
          title: '為自己 pray.',
          body: ''
        }],
        ['verse', 'pray 金句', 'bible', {
          title: 'pray 金句',
          bibleQuery: '',
          bibleVersion: 'tghg',
          bibleRecords: []
        }],
        ['hymn-2', '詩歌二 (台語)', 'praise', {
          title: '',
          body: ''
        }],
        ['benediction', '結束 pray：祈禱文', 'content', {
          title: '結束 pray: 祈禱文',
          body: ''
        }]
      ]
    }
  };

  function clone(value) {
    return value == null ? value : JSON.parse(JSON.stringify(value));
  }

  function getTemplateProfile(templateId) {
    return profiles.prayer;
  }

  function createTemplateModel(profileInput) {
    const profile = getTemplateProfile();
    const model = {};
    profile.sections.forEach(([id, label, type, defaults]) => {
      const source = clone(defaults || {});
      model[id] = {
        label,
        type,
        title: source.title || label,
        kicker: source.kicker || '',
        body: source.body || '',
        bibleQuery: source.bibleQuery || '',
        bibleVersion: source.bibleVersion || 'tghg',
        bibleRecords: source.bibleRecords || [],
        ...source
      };
    });
    return model;
  }

  return {
    DEFAULT_TEMPLATE_ID,
    profiles,
    getTemplateProfile,
    createTemplateModel
  };
});
