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
      defaultBackgroundColor: '#111111', // Sleek dark mode by default for prayer meetings
      sections: [
        ['silence', '請安靜心、等候神', 'content', {
          title: '請安靜心、等候神',
          body: '請放下身心重擔，因有人在為你禱告，如安歇在神的翅膀蔭下.'
        }],
        ['hymn-1', '詩歌一 (華語)', 'praise', {
          title: '有人在為你禱告 (華語)',
          body: ''
        }],
        ['scripture', '經文 (台語)', 'bible', {
          title: '經文 (台語)',
          bibleQuery: '羅 8:26; 弗 6:18; 提前 2:1; 路 4:38',
          bibleVersion: 'tghg',
          bibleRecords: []
        }],
        ['thanksgiving', '獻上感謝讚美', 'list', {
          title: '獻上感謝讚美的 pray.',
          body: 'a. 讚美上帝的美好 --- 公義、恩慈、良善、信實、真理 .....\nb. 讚美上帝的作為 --- 創造、拯救、成全、陪伴、同工 .....\nc. 讚美上帝對我的好 --- 救贖的主、慈愛的天父、真理的引導者.'
        }],
        ['repentance', '悔改認罪', 'list-bible', {
          title: '悔改認罪 pray',
          bibleQuery: '路 5:32',
          bibleVersion: 'tghg',
          bibleRecords: [],
          body: 'a. 因為得罪人心中憂傷嗎?\nb. 因為沒有能 glory 主名而羞愧嗎?\nc. 因為工作、事工沒有做好而煩悶嗎?\nd. 因為靈命停滯、靈修不穩定徬徨嗎?'
        }],
        ['world', '為世界 pray', 'list', {
          title: '為世界 pray',
          body: 'a. 為戰爭止息，全球和平.\nb. 為極權國家的轉變與不再擴張.\nc. 為各地天災人禍、地震、瘟疫...所帶給人們的痛苦.'
        }],
        ['nation', '為國家、社會 pray', 'list-bible', {
          title: '為國家、社會 pray',
          bibleQuery: '申 11:12-15',
          bibleVersion: 'tghg',
          bibleRecords: [],
          body: 'a. 為上帝掌管天候與土地，賜下春雨秋雨使國家五穀豐登、人民安居樂業 pray.\nb. 為國家安全、國際外交合宜關係 pray.\nc. 為各級政府各項施政都秉持行公義、存憐憫的態度 pray.'
        }],
        ['church', '為教會 pray', 'list', {
          title: '為教會 pray.',
          body: 'a. 求主為林口教會預備牧者 pray.\nb. 為卓牧師夫婦願意繼續委身幫忙林口教會各項牧會事工來感謝 pray.\nc. 為教會各事工小組的計劃 pray.\nd. 為各團契、小組的肢體、同工、帶領者的靈命、外展的熱忱 and 智慧，及彼此的包容、愛的關懷、合一的扶持來 pray.\ne. 為林口教會牧師、長執、同工身、心、靈健壯守望 pray.\nf. 為教會落實異象，發揮光、鹽影響力 pray.\ng. 為教會近期各項事工 pray:\n(1) 今天 7/19 日舉行全教會培靈會，為講員蔡安祿牧師的分享及一切行程出入平安，為弟兄姊妹靈命得更新、得造就、得激勵、得深化.\n(2) 7/26 日基督精兵營將安排薇依牧師前來報告及募款.\n(3) 8/2 日父親節主日.'
        }],
        ['members', '為教會肢體 pray', 'list', {
          title: '為教會肢體 pray.',
          body: 'a. 為弟兄姊妹能從信徒轉變成門徒.\nb. 為孩童暑期的照顧，及青年暑期生活有內容、升學選擇有智慧.\n(c) 為社青進入職場的適應，及兵役的平安.\n(d) 為長輩每一天的健康、喜樂、平安.\n(e) 為肢體身、心、靈有病痛、軟弱、憂慮的.\n(f) 為失喪家人的兄姊 pray.\n(g) 為肢體有失業及經濟弱勢的 pray.\n(h) 為住院或在家療養 of 弟兄姊妹.'
        }],
        ['oneself', '為自己 pray', 'list', {
          title: '為自己 pray.',
          body: 'a. 為自己及家人的身、心、靈剛強健壯.\nb. 為自己及家人能先尋求祂的國 and 祂的義.\nc. 為自己及家人隨時隨地高舉神的名，為主作見證.\ne. 為自己及家人關係密切和諧.\nf. 為自己及家人各項需要能蒙神供應有餘.\ng. 為未信主的家人蒙恩得救.'
        }],
        ['verse', 'pray 金句', 'bible', {
          title: 'pray 金句',
          bibleQuery: '耶利米書 29:12',
          bibleVersion: 'tghg',
          bibleRecords: []
        }],
        ['hymn-2', '詩歌二 (台語)', 'praise', {
          title: '有人在為你禱告 (台語)',
          body: ''
        }],
        ['benediction', '結束 pray：祈禱文', 'content', {
          title: '結束 pray: 祈禱文',
          body: '主禱文'
        }]
      ]
    }
  };

  function clone(value) {
    return value == null ? value : JSON.parse(JSON.stringify(value));
  }

  function getTemplateProfile(templateId) {
    return profiles.prayer; // Always return prayer profile
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
