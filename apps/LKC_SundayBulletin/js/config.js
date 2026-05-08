// 設定檔 - 教會週報管理系統

const CONFIG = {

  GAS_SYNC_URL: 'https://script.google.com/macros/s/AKfycbyLLQZsz_XZqhWVwaT_8hcvfQc8fSWztAncEmBUk7lnzGr-TcP33uzS-weUG_cavgEn/exec',
  SHARED_TOKEN: 'ChurchApp-2026',

  LKC1958_GAS_URL:       'https://script.google.com/macros/s/AKfycbx4268IkgwQm2Es0gjDHLU_U9nKJrRMR1-xzbbtuaq08lePLgAQ2wnDRrCeHdy9jNhh/exec',
  LKWORSHIP_GAS_URL:     'https://script.google.com/macros/s/AKfycbyk_6tUucVg-U4rRQjYHvk632teZyxufDkNX_X1WRUXPMGgsTaemVXD_mv9kBDjuSwOnA/exec',
  LKCSCHEDULE_GAS_URL:   'https://script.google.com/macros/s/AKfycbwiYYWgKxmLRAEaE_pbp_kWyAzlRPcwYVQfvmJVamRJvosvt5wTTkvwebbFBkP8rMqX/exec',
  LKGROUP_GAS_URL:       'https://script.google.com/macros/s/AKfycbzfaWh_ooRTGijLV_7lYFUHFm83oL6DvYt9rt6ze5mDXhtwLv8ymxLX_PGuDTHzmNwe/exec',
  LKC_ATTENDANCE_GAS_URL:'https://script.google.com/macros/s/AKfycbyJbzjHIeFFRbqT-Ttk2OAPYfF-qDKYES8dJiu4sJCR4t2Fq9PTtbALwuiJDBxh55kR/exec',

  CHURCH_NAME: '台灣基督長老教會林口教會',
  CHURCH_NAME_EN: 'Linkou Presbyterian Church',

  BANK_ACCOUNT: `教會奉獻帳戶：彰化銀行林口分行
戶名：台灣基督長老教會林口教會
銀行代碼：009　帳號：9689-51-29395500
若有匯款請通知教會辦公室，謝謝！`,

  AUTO_SAVE_INTERVAL: 60000,
  MAX_DRAFTS: 10,
  DRAFT_KEY_PREFIX: 'bulletin_draft_',
  VERSION: '1.0.0',

  // LKGroup API 動態取得小組列表；此為網路失敗時的鈴援列表
  TW_GROUPS: [
    '牧養組總聚', '葡萄樹A組', '葡萄樹B組', '葡萄樹聯合',
    '芥菜種A組', '芥菜種B組', '芥菜種C組', '芥菜種聯合',
    '棕樹A組', '棕樹B組', '稭子小組', '提摩太小組',
    '橄欖樹小組', '恩典團契', '學青小組',
    '以斯帖小組', '松年團契', '香柏樹小組'
  ],

  SUNDAY_SCHOOL_CLASSES: [
    '幼小班', '初小班', '中小班', '高小班',
    '青少年班', '成人A班', '成人B班'
  ]
};
