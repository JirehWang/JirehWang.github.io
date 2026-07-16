(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipBulletinContent = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const clean = value => String(value == null ? '' : value).trim();

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

  function buildReportPages(data, announcementsPerPage = 2) {
    const reports = normalizeReports(data);
    const pageSize = Math.max(1, Number(announcementsPerPage) || 2);
    const pages = [];
    const appendNumberedPages = (items, title) => {
      for (let index = 0; index < items.length; index += pageSize) {
        const body = items
          .slice(index, index + pageSize)
          .map((text, offset) => `${index + offset + 1}. ${text}`)
          .join('\n\n');
        pages.push({ kind: 'report', title, body });
      }
    };
    appendNumberedPages(reports.announcements, '報告－本會消息');
    appendNumberedPages(reports.churchNews, '報告－教界消息');
    const prayerParts = [
      ['在家調養兄姐：', reports.prayer.homeRest],
      ['住院：', reports.prayer.hospital],
      ['其他代禱：', reports.prayer.other]
    ].filter(([, value]) => value).map(([label, value]) => `${label}${value}`);
    if (prayerParts.length) {
      pages.push({ kind: 'report', title: '報告－關懷代禱', body: prayerParts.join('\n\n') });
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
    buildReportPages,
    applyReportsToModel,
    applyPraiseToModel,
    loadCloudRecord
  };
});
