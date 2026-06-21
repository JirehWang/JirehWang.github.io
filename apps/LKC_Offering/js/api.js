/**
 * api.js - LKC 奉獻管理系統 API 封裝
 */

const OfferingAPI = {
  /**
   * 安全查詢個人年度奉獻明細 (會友端)
   * @param {string} code - 會友代號
   * @param {string} name - 姓名
   * @param {string} year - 查詢年份
   */
  async queryMemberOffering(code, name, year) {
    try {
      const response = await window.churchAPI('queryMemberOffering', { code, name, year });
      if (response && response.error) {
        return { success: false, error: response.error };
      }
      if (response && response.data) {
        return response.data;
      }
      return { success: false, error: '後端未回傳有效資料' };
    } catch (err) {
      console.error('[OfferingAPI] queryMemberOffering 失敗:', err);
      return { success: false, error: err.message || '連線逾時，請稍後再試' };
    }
  },

  /**
   * 批次登錄當週奉獻資料 (管理端)
   * @param {string} date - 登錄日期
   * @param {Array} items - [{ code, type, amount, notes }]
   */
  async adminAddOfferings(date, items) {
    try {
      const response = await window.churchAPI('adminAddOfferings', { date, items });
      if (response && response.error) {
        return { success: false, error: response.error };
      }
      if (response && response.data) {
        return response.data;
      }
      return { success: false, error: '後端未回傳有效資料' };
    } catch (err) {
      console.error('[OfferingAPI] adminAddOfferings 失敗:', err);
      return { success: false, error: err.message || '連線逾時，請稍後再試' };
    }
  }
};
