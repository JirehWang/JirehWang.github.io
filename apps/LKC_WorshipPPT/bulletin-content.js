(function(root, factory) {
  const production = typeof module === 'object' && module.exports
    ? require('./slide-production.js')
    : root.TaiwaneseWorshipSlideProduction;
  const api = factory(production || {});
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipBulletinContent = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(production) {
  const clean = value => String(value == null ? '' : value).trim();
  const REPORT_FONT_SIZE = 48;
  const REPORT_BOX_WIDTH = 84;
  const REPORT_MAX_LINES = 5;

  function buildBulletinCloudUrl(endpoint, kind, date) {
    const prefix = kind === 'praise' ? 'praise_songs_' : 'reports_';
    const separator = String(endpoint).includes('?') ? '&' : '?';
    return `${endpoint}${separator}action=load&key=${encodeURIComponent(prefix + clean(date))}`;
  }

  function normalizeReports(data) {
    const source = data && typeof data === 'object' ? data : {};
    const prayer = source.prayer && typeof source.prayer === 'object' ? source.prayer : {};
    return {
      announcements: (Array.isArray(source.announcements) ? source.announcements : []).map(clean).filter(Boolean),
      churchNews: (Array.isArray(source.churchNews) ? source.churchNews : []).map(clean).filter(Boolean),
      prayer: {
        homeRest: clean(prayer.homeRest),
        hospital: clean(prayer.hospital),
        other: clean(prayer.other)
      }
    };
  }

  function reportLines(value) {
    const text = clean(value);
    if (!text) return [];
    const wrapped = typeof production.wrapTextForBox === 'function'
      ? production.wrapTextForBox(text, {
        fontSize: REPORT_FONT_SIZE,
        boxWidth: REPORT_BOX_WIDTH,
        bold: true
      })
      : text;
    return String(wrapped).split('\n');
  }

  function paginateReportEntries(entries, title) {
    const pages = [];
    let currentLines = [];
    const flush = () => {
      if (!currentLines.length) return;
      pages.push({ kind: 'report', title, body: currentLines.join('\n') });
      currentLines = [];
    };

    (entries || []).forEach(entry => {
      let lines = reportLines(entry.text);
      if (!lines.length) return;
      const gap = currentLines.length ? 1 : 0;
      if (lines.length + gap <= REPORT_MAX_LINES - currentLines.length) {
        if (gap) currentLines.push('');
        currentLines.push(...lines);
        return;
      }

      flush();
      if (lines.length <= REPORT_MAX_LINES) {
        currentLines.push(...lines);
        return;
      }

      currentLines.push(...lines.splice(0, REPORT_MAX_LINES));
      flush();
      while (lines.length) {
        currentLines.push(entry.continuation || '（續）');
        currentLines.push(...lines.splice(0, REPORT_MAX_LINES - 1));
        if (lines.length) flush();
      }
    });
    flush();
    return pages;
  }

  function buildReportPages(data) {
    const reports = normalizeReports(data);
    const pages = [];
    const appendNumberedPages = (items, title) => {
      pages.push(...paginateReportEntries(items.map((text, index) => ({
        text: `${index + 1}. ${text}`,
        continuation: `${index + 1}.（續）`
      })), title));
    };
    appendNumberedPages(reports.announcements, '報告－本會消息');
    appendNumberedPages(reports.churchNews, '報告－教界消息');
    const prayerParts = [
      ['在家調養兄姐：', reports.prayer.homeRest],
      ['住院：', reports.prayer.hospital],
      ['其他代禱：', reports.prayer.other]
    ].filter(([, value]) => value).map(([label, value]) => `${label}${value}`);
    if (prayerParts.length) {
      pages.push(...paginateReportEntries(prayerParts.map(text => ({ text, continuation: '（續）' })), '報告－關懷代禱'));
    }
    return pages;
  }

  function applyReportsToModel(model, data) {
    if (!model || !model.announcements) return model;
    const reports = normalizeReports(data);
    model.announcements.announcements = reports.announcements;
    model.announcements.churchNews = reports.churchNews;
    model.announcements.prayer = reports.prayer;
    model.announcements.pptPages = buildReportPages(reports);
    model.announcements.includeSectionTitle = true;
    return model;
  }

  function applyPraiseToModel(model, data) {
    if (!model || !model.praise || !data) return model;
    model.praise.title = clean(data.title) || '讚美';
    model.praise.kicker = clean(data.kicker) || '聖歌隊';
    model.praise.body = clean(data.lyrics);
    return model;
  }

  async function loadCloudRecord(endpoint, kind, date, fetchImpl) {
    const response = await fetchImpl(buildBulletinCloudUrl(endpoint, kind, date));
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const json = await response.json();
    return json && json.success && json.data
      ? { state: 'loaded', data: json.data }
      : { state: 'missing', data: null };
  }

  return {
    buildBulletinCloudUrl,
    normalizeReports,
    reportLines,
    paginateReportEntries,
    buildReportPages,
    applyReportsToModel,
    applyPraiseToModel,
    loadCloudRecord
  };
});
