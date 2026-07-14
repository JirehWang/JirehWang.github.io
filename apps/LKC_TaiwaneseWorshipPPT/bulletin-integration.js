(function(root) {
  const api = root.TaiwaneseWorshipBulletinContent;
  const endpoint = 'https://script.google.com/macros/s/AKfycbyLLQZsz_XZqhWVwaT_8hcvfQc8fSWztAncEmBUk7lnzGr-TcP33uzS-weUG_cavgEn/exec';

  root.loadBulletinPptContent = async function(date) {
    const [reportsResult, praiseResult] = await Promise.all([
      api.loadCloudRecord(endpoint, 'reports', date, root.fetch.bind(root)),
      api.loadCloudRecord(endpoint, 'praise', date, root.fetch.bind(root))
    ]);
    if (reportsResult.state === 'loaded') api.applyReportsToModel(model, reportsResult.data);
    if (praiseResult.state === 'loaded') api.applyPraiseToModel(model, praiseResult.data);
    return {
      reports: reportsResult,
      praise: praiseResult,
      reportPageCount: reportsResult.state === 'loaded' ? model.announcements.pptPages.length + 1 : 0,
      praisePageCount: praiseResult.state === 'loaded'
        ? 1 + String(model.praise.body || '').split(/\n\s*\n/).filter(Boolean).length
        : 0
    };
  };

  root.describeBulletinPptContent = function(result) {
    if (!result || result.error) return `週報資料讀取失敗：${result && result.error ? result.error : '未知錯誤'}`;
    const parts = [];
    parts.push(result.reports.state === 'loaded' ? `報告 ${result.reportPageCount} 頁` : '報告無資料');
    parts.push(result.praise.state === 'loaded' ? `讚美 ${result.praisePageCount} 頁` : '讚美無資料');
    return parts.join('、');
  };

  const previousEditor = editor;
  editor = function() {
    if (active !== 'announcements') return previousEditor();
    const item = model.announcements;
    const form = document.getElementById('editor-form');
    const reports = api.normalizeReports({ announcements: item.announcements, prayer: item.prayer });
    form.innerHTML = [
      '<div class="inline-note">依禮拜日期從週報系統帶入，並依原 PPT 分成「本會消息」與「關懷代禱」。仍可在此手動修改。</div>',
      field('本會消息（每則消息之間空一行）', 'reportAnnouncements', reports.announcements.join('\n\n'), 'textarea'),
      '<div class="inline-note">關懷代禱</div>',
      field('在家調養兄姐', 'prayerHomeRest', reports.prayer.homeRest, 'textarea'),
      field('住院', 'prayerHospital', reports.prayer.hospital, 'textarea'),
      field('其他代禱', 'prayerOther', reports.prayer.other, 'textarea')
    ].join('');

    form.querySelectorAll('[data-key]').forEach(element => {
      element.oninput = () => {
        api.applyReportsToModel(model, {
          announcements: form.querySelector('[data-key="reportAnnouncements"]').value.split(/\n\s*\n/),
          prayer: {
            homeRest: form.querySelector('[data-key="prayerHomeRest"]').value,
            hospital: form.querySelector('[data-key="prayerHospital"]').value,
            other: form.querySelector('[data-key="prayerOther"]').value
          }
        });
        preview();
        flow();
      };
    });
  };
  render();
})(window);
