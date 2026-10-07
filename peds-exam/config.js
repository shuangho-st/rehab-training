/* ================= 設定（部署時只需要改這裡） =================
 * API_URL：把 gas/前後測後端.gs 部署成「網頁應用程式」後得到的網址（…/exec）。
 *          留空時學員仍可作答、看到成績，但成績不會集中到試算表。
 * HOSPITALS：學員報到時可選的醫院。
 * LEVELS：學員可選的職級。
 * ============================================================== */
window.PEDS_CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbwxzu2VqIe7ugZlfMjGxPVPGF7HuANR7cBfIA5hGow4ZS0I4nIeQBUeVRCyClOrdKO-cg/exec',
  HOSPITALS: ['北醫附醫', '雙和醫院', '萬芳醫院', '其他'],
  LEVELS: ['PGY', 'R1', 'R2', 'R3', 'R4', 'CR', 'Fellow', '主治醫師', '醫學生', '其他'],
  BANKS: 'data/banks.json'
};
