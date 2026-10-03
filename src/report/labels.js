// Chinese wording used in the Excel report, plus the small pure formatters
// that turn journal / summary values into one-line texts. No I/O here.
//
// Status codes come from src/campaign.js (STATUS, REMIND_STATUS); skip reasons
// and verify issue types from the interface spec. An unknown code is shown
// as-is, so a newer journal never breaks the report.

import { formatUsd } from '../select/amount.js';

/** Issue status of a recipient (Excel "状态" column). */
export const STATUS_LABELS = Object.freeze({
  pending: '待发放',
  in_progress: '进行中',
  unknown: '需人工核对',
  created: '已建卡未打tag',
  done: '已完成',
  failed: '失败',
  skipped: '发放前跳过',
});

/** Display order of the statuses in the summary and the explanation sheet. */
export const STATUS_ORDER = Object.freeze(['pending', 'in_progress', 'unknown', 'created', 'done', 'failed', 'skipped']);

/** Row fill per status (ARGB). Pending stays unfilled (white) so the grid lines remain visible. */
export const STATUS_FILLS = Object.freeze({
  pending: null,
  in_progress: 'FFDDEBF7', // light blue
  unknown: 'FFFCE4D6', // light orange
  created: 'FFFFF2CC', // light yellow
  done: 'FFE2EFDA', // light green
  failed: 'FFFFD7D7', // light red
  skipped: 'FFEDEDED', // light grey
});

// From REMIND_1_DATE on, issue only records leftovers and fixes tags for the real
// campaign (fix spec D2): the statuses that wait for a new card say so.
const NO_NEW_CARDS_FROM_REMIND_1 = '正式活动从 REMIND_1_DATE 起 issue 不再建新卡。';

/** What each status means (sheet "说明"). */
export const STATUS_HELP = Object.freeze({
  pending: `还没有处理；或核对确认上次没有建成卡，下次运行 issue 会重新处理。${NO_NEW_CARDS_FROM_REMIND_1}`,
  in_progress: 'issue 正在处理这个人。只在 issue 运行时显示；issue 不在运行时，停在这一步的人显示为“需人工核对”。',
  unknown: 'Shopify 的回复丢失，还不知道卡是否建好。下次运行 issue 或 verify 会先去 Shopify 查找这张卡，找到就补记，超过 10 分钟仍找不到才确认没建成。',
  created: '卡已建好（Shopify 已自动发出首封邮件），tag 还没打上。下次运行 issue 会只补打 tag，不会再建卡。',
  done: '已建卡并打上 tag。',
  failed: `Shopify 明确拒绝建卡，原因见“备注/错误”。只有 issue --retry-failed 才会重试。${NO_NEW_CARDS_FROM_REMIND_1}`,
  skipped: '发放前复查发现此人已不符合条件，不发卡，原因见“备注/错误”。',
});

// 'not-subscribed' covers every marketing state other than SUBSCRIBED (not
// subscribed, unsubscribed, pending, invalid, none); the bracketed detail names it.
const NOT_SUBSCRIBED_LABEL = '未订阅营销邮件';

/** Pre-flight skip reasons written by issue (journal op "skip"). */
export const ISSUE_SKIP_LABELS = Object.freeze({
  'customer-deleted': '客户已删除',
  'no-email': '邮箱没了',
  'relay-email': '邮箱域名在排除名单里',
  'not-subscribed': NOT_SUBSCRIBED_LABEL,
  'already-tagged': '已有发放 tag',
  'ordered-since-snapshot': '导出后下过单',
  'address-ordered-since-snapshot': '同地址账户导出后下过单',
});

/** Reminder skip reasons written by remind (journal op "remind.skip"). */
export const REMIND_SKIP_LABELS = Object.freeze({
  'no-card': '没有卡',
  'multiple-cards': '有多张卡',
  'card-disabled': '卡已停用',
  'card-expired': '卡已过期',
  used: '已用过卡',
  'customer-deleted': '客户已删除',
  'no-email': '没有邮箱',
  'not-subscribed': NOT_SUBSCRIBED_LABEL,
});

/**
 * The tag remind puts on a customer after each reminder of `round`: "<SENT_TAG>-R<round>",
 * e.g. OCT26RTPROMO-R1 (test campaign: OCT26RTPROMO-TEST-R1). The same name as
 * roundTag() in src/remind.js; test/excel.test.js checks that the two agree.
 */
export function roundTagName(sentTag, round) {
  return `${sentTag}-R${round}`;
}

/** "第 N 次提醒" cell of a reminder found sent by its round tag (journal op remind.found, source 'tag'). */
export const REMIND_SENT_BY_TAG = '已发（按 Shopify 上的本轮 tag 补记）';
/** Added to a sent reminder's cell when tagging the customer with the round tag failed (remind.tag.fail). */
export const REMIND_TAG_MISSING_SUFFIX = '；本轮 tag 没打上，下次运行会补打';

/** Reminder cell texts explained on the "说明" sheet. */
export const REMIND_HELP = Object.freeze([
  ['（空白）', '这一轮还没有处理这个人。'],
  ['已发 MM-DD HH:MM', '已让 Shopify 重发礼品卡邮件，时间是店铺时间；发出后给客户打上本轮 tag（见“本轮 tag”）。每一轮每人最多一封。'],
  [REMIND_SENT_BY_TAG, '本地日志里这一轮没有“已发”的记录（例如日志丢失，或被更早的备份覆盖），但客户在 Shopify 上带着本轮 tag，所以算作这一轮已发，不会再发。发送时间不详。'],
  [`已发 MM-DD HH:MM${REMIND_TAG_MISSING_SUFFIX}`, '提醒已经发出，但给客户打本轮 tag 失败了：本地日志记着已发，这一轮不会再给他发。下次真实运行同一轮时会先补打这个 tag（见“补打本轮 tag”），补打成功后这里只显示“已发 MM-DD HH:MM”。'],
  ['跳过：原因', `这一轮不符合提醒条件，例如已用过卡、${NOT_SUBSCRIBED_LABEL}、卡已停用或过期。再次运行同一轮会重新判断。`],
  ['失败：原因', 'Shopify 拒绝发送，原因写在后面。不会自动重发；原因解决后可用 remind --retry-failed 重发。'],
  ['结果不明', 'Shopify 的回复丢失，不知道邮件是否发出。不会自动重发（重发一定会多一封邮件）；确认后可用 remind --retry-unknown 补发。'],
  ['进行中', 'remind 正在处理这个人。只在 remind 运行时显示；remind 不在运行时显示为“结果不明”。'],
]);

/**
 * The per-round tags explained on the "说明" sheet, with this campaign's tag names
 * (`sentTag` = the campaign's SENT_TAG): [item, text] pairs. A function, not a constant:
 * it needs the tag, and it quotes OP_LABELS (declared below; read only when called).
 */
export function roundTagHelp(sentTag) {
  return [
    ['本轮 tag', `第 1 次提醒是 ${roundTagName(sentTag, 1)}，第 2 次提醒是 ${roundTagName(sentTag, 2)}，两轮互不影响。每发出一封提醒，就给客户打上这一轮的 tag。`],
    ['带本轮 tag 的人', `每次运行 remind（预演也一样）都先查 Shopify 上带本轮 tag 的客户：他们一律算这一轮已发过，不会再发，加 --retry-unknown 或 --retry-failed 也不会。所以即使本地日志丢失或被旧备份覆盖，这一轮也不会给同一个人再发一封；真实运行会把他们在本地日志里补记为“${REMIND_SENT_BY_TAG}”。`],
    ['补打本轮 tag', `打本轮 tag 失败不影响已经发出的提醒（本地日志记着已发），“操作日志”里记一行“${OP_LABELS['remind.tag.fail']}”。下次真实运行同一轮时会先补打；发出满 10 分钟、Shopify 上却查不到本轮 tag 的人（例如 tag 在后台被删掉）也会补打。补打成功，或发现 Shopify 上其实已带这个 tag，记一行“${OP_LABELS['remind.tag.ok']}”。发提醒后直接打上 tag 的不单独记一行，只在运行记录里计数。`],
  ];
}

/** Customer kind of a recipient. */
export const KIND_LABELS = Object.freeze({ ordered: '有下单', never: '从没下单', test: '测试' });

/** Shopify email marketing states ('NONE' = the commands' stand-in for a missing state). */
export const MARKETING_LABELS = Object.freeze({
  SUBSCRIBED: '已订阅',
  NOT_SUBSCRIBED: '未订阅',
  UNSUBSCRIBED: '已退订',
  PENDING: '待确认',
  INVALID: '无效',
  REDACTED: '已隐去',
  NONE: '没有营销状态',
});

/** Notes for recipients whose amount is not based on their very last order. */
export const BASIS_NOTES = Object.freeze({
  'last-cancelled': '最近一笔已取消，按更早的付费订单计算',
  'last-test': '最近一笔是测试订单，按更早的付费订单计算',
  'last-zero': '最近一笔是 $0，按更早的付费订单计算',
});
export const NEVER_NOTES = Object.freeze({
  'only-cancelled-or-zero': '订单都已取消或都是 $0，按从没下单处理',
});

/** Journal ops (sheet "操作日志"). */
export const OP_LABELS = Object.freeze({
  'create.start': '开始建卡',
  'create.ok': '建卡成功',
  'create.fail': '建卡被拒',
  'create.rejected': '建卡未执行（可重试）',
  'create.unknown': '建卡结果不明',
  'reconcile.found': '核对：找到卡',
  'reconcile.none': '核对：确认没建成',
  'tag.ok': '打 tag 成功',
  'tag.fail': '打 tag 失败',
  skip: '发放前跳过',
  'remind.start': '开始提醒',
  'remind.ok': '提醒已发',
  'remind.fail': '提醒失败',
  'remind.rejected': '提醒未发出（可重试）',
  'remind.unknown': '提醒结果不明',
  'remind.skip': '提醒跳过',
  'remind.found': '按本轮 tag 补记已发',
  'remind.tag.fail': '打本轮 tag 失败',
  'remind.tag.ok': '本轮 tag 已打上', // written by the repair only: tag added again, or found in Shopify after all
});

/** Where a reconcile.found entry came from. */
export const RECONCILE_SOURCE_LABELS = Object.freeze({
  'issue-reconcile': 'issue 开始前核对',
  'issue-inline': '建卡结果不明后立即核对',
  preflight: '发放前复查',
  verify: 'verify 核对',
});

/** verify.json issue types (sheet "核对"). */
export const VERIFY_TYPE_LABELS = Object.freeze({
  'missing-in-shopify': '日志有卡，Shopify 没有',
  'not-in-journal': 'Shopify 有卡，日志没有（已补记）',
  'not-in-selection': '卡的客户不在名单里',
  'duplicate-cards': '同一客户有多张卡',
  'amount-mismatch': '金额不一致',
  'card-disabled': '卡已停用',
  'tag-missing': '有卡但没有 tag',
  'tag-without-card': '有 tag 但没有卡',
  'still-unknown': '仍需人工核对',
  'resolved-unknown': '需人工核对的已查明',
});

/** selection.snapshot.source. */
export const SNAPSHOT_SOURCE_LABELS = Object.freeze({
  bulk: '全店导出',
  paginated: '普通分页',
  nodes: '按客户 ID 读取',
});

/** run.end summary `stopped` values written by remind. */
export const STOP_LABELS = Object.freeze({
  rejected: '已停止：Shopify 没有接受发送请求',
  'too-many-failures': '已停止：连续多次没有发送成功',
  aborted: '已中断（Ctrl+C）',
  error: '已停止：出错',
});

/** Fixed English notes some journal entries carry (anything else is shown as-is). */
export const NOTE_LABELS = Object.freeze({
  'already tagged in Shopify': 'Shopify 上已带这个 tag',
});

/** usage.json unmatched[].reason codes that may appear (anything else is shown as-is). */
export const UNMATCHED_REASON_LABELS = Object.freeze({
  'no-gift-card-id': '回执里没有礼品卡 ID',
  'no-receipt-id': '回执里没有礼品卡 ID',
  'missing-gift-card-id': '回执里没有礼品卡 ID',
  'no-receipt': '没有回执',
  'bad-receipt': '回执无法解析',
});

/** Command exit codes (sheet "说明"). */
export const EXIT_CODE_HELP = Object.freeze([
  [0, '成功'],
  [1, '中途停止或失败（原因见命令输出）'],
  [2, '用法错误：参数不对、缺少 --limit、还没到活动日期等，或已过 REMIND_1_DATE（issue 只补记和补打 tag，不建新卡）、已过 REMIND_2_DATE（不能再发第 1 次提醒）'],
  [130, '按 Ctrl+C 中断（做完当前这个人后退出）'],
]);

/**
 * Why 礼品卡抵扣 and 已用金额 can differ (使用报告, 汇总, 说明): the first is
 * taken at checkout (usage.js F1), the second from the card balances.
 */
export const USAGE_SPLIT_NOTE = '礼品卡抵扣按结账时计算，之后退回卡里的钱不扣；已用金额 = 原金额 − 现在余额，已扣除退回卡里的钱。所以有退款退回卡时两者不同，差额就是退回卡里的金额。';

// ---------------------------------------------------------------------------
// Small lookups
// ---------------------------------------------------------------------------

/** dict[key] when it is the dictionary's own entry (never an inherited Object member such as "constructor"). */
function own(dict, key) {
  return Object.hasOwn(dict, key) ? dict[key] : undefined;
}

/** label from a dictionary, or the raw code (empty for null/undefined). */
export function labelOf(dict, code) {
  if (code === null || code === undefined || code === '') return '';
  return own(dict, code) ?? String(code);
}

// A journal detail that is an instant: Shopify's ISO timestamps (always with a zone).
const ISO_INSTANT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})$/;
const GID = /^gid:\/\/shopify\/(\w+)\/(\d+)(?:\?.*)?$/;
const GID_TYPE_LABELS = Object.freeze({ Customer: '客户', Order: '订单', GiftCard: '礼品卡' });

/**
 * A skip detail from the journal as people read it: an ISO instant becomes
 * store-local "YYYY-MM-DD HH:MM 店铺时间" (via `stamp`, the workbook's clock;
 * no brackets, because the detail itself is shown in brackets), a customer
 * gid "客户 <number>", a marketing state its Chinese label.
 * Anything else (already Chinese, an order name, a domain, ...) stays as it is.
 * @param {*} detail
 * @param {{ stamp?: (iso: string) => string }} [options]
 */
export function detailText(detail, { stamp } = {}) {
  if (detail === null || detail === undefined) return '';
  const s = String(detail).trim();
  if (!s) return '';
  if (stamp && ISO_INSTANT.test(s) && Number.isFinite(Date.parse(s))) {
    const local = stamp(s);
    if (local) return `${local} 店铺时间`;
  }
  const g = GID.exec(s);
  if (g) return own(GID_TYPE_LABELS, g[1]) ? `${GID_TYPE_LABELS[g[1]]} ${g[2]}` : g[2];
  return own(MARKETING_LABELS, s) ?? s;
}

/** Issue skip reason with its optional detail, e.g. "导出后下过单（#1234）"; `options` as for detailText. */
export function issueSkipText(reason, detail, options) {
  const label = labelOf(ISSUE_SKIP_LABELS, reason);
  const d = detailText(detail, options);
  return d ? `${label}（${d}）` : label;
}

/** Reminder skip reason with its optional detail; `options` as for detailText. */
export function remindSkipText(reason, detail, options) {
  const label = labelOf(REMIND_SKIP_LABELS, reason);
  const d = detailText(detail, options);
  return d ? `${label}（${d}）` : label;
}

/** A fixed English journal note in Chinese; other notes unchanged. */
export function noteText(note) {
  if (note === null || note === undefined) return '';
  return own(NOTE_LABELS, note) ?? String(note);
}

// ---------------------------------------------------------------------------
// Shopify error messages (built by src/shopify.js) → Chinese
// ---------------------------------------------------------------------------

/** The " [CODE]" that throwIfUserErrors puts after each userError message ("…; " between them). */
const USER_ERROR_CODE = / \[([A-Z][A-Z0-9_]*)\](?=; |$)/g;

/** "prefix：tail", or the prefix alone when there is no tail. */
const withTail = (prefix, tail = '') => (tail.trim() ? `${prefix}：${tail}` : prefix);

// [message form, Chinese text]. Only our own English around Shopify's words is
// translated: Shopify's message text stays as it is.
const SHOPIFY_ERROR_FORMS = [
  // throwIfUserErrors: "<mutation> rejected: <field>: <message> [<CODE>]; ..."
  [/^[a-z][A-Za-z0-9]* rejected: ([\s\S]*)$/, (m) => withTail('Shopify 拒绝', m[1].replace(USER_ERROR_CODE, '（$1）'))],
  [/^Network error calling Shopify(?:: ([\s\S]*))?$/, (m) => withTail('网络错误', m[1])],
  [/^Shopify HTTP (\d+)(?:: ([\s\S]*))?$/, (m) => withTail(`Shopify 返回 HTTP ${m[1]}`, m[2])],
  [/^(?:Throttled by Shopify|HTTP 429 from Shopify) (\d+) times in a row; giving up$/, (m) => `Shopify 限流：连续 ${m[1]} 次被拒，已放弃`],
  [/^GraphQL error(?:: ([\s\S]*))?$/, (m) => withTail('Shopify 查询错误', m[1])],
  [/^Shopify returned non-JSON\b/, () => 'Shopify 返回了无法解析的内容'],
  [/^Shopify response contained no data$/, () => 'Shopify 没有返回数据'],
  // src/customers.js addTag (also remind's round tag) and src/giftcards.js: a mutation answer without its payload
  [/^[a-z][A-Za-z0-9]* returned no payload$/, () => 'Shopify 没有返回结果'],
];

/** One Shopify error message in Chinese, or null when it is not one of the forms above. */
function shopifyErrorText(message) {
  for (const [form, text] of SHOPIFY_ERROR_FORMS) {
    const m = form.exec(message);
    if (m) return text(m);
  }
  return null;
}

/**
 * A journal error as people read it in the workbook (发放名单 备注/错误, the
 * reminder cells, 操作日志 错误), e.g.
 *   "giftCardCreate rejected: input: Customer is invalid [INVALID]" → "Shopify 拒绝：input: Customer is invalid（INVALID）"
 *   "Network error calling Shopify: fetch failed" → "网络错误：fetch failed"
 * Anything else (already Chinese, another error) is unchanged. The commands
 * also embed a Shopify message in their own Chinese text, after "：" and joined
 * with "；" (issue: "<error>；查卡也失败了：<error>"); each of those is translated too.
 * @param {*} message
 */
export function errorText(message) {
  if (message === null || message === undefined) return '';
  return String(message)
    .split('；')
    .map((part) => {
      const whole = shopifyErrorText(part);
      if (whole !== null) return whole;
      const colon = part.indexOf('：');
      const tail = colon >= 0 ? shopifyErrorText(part.slice(colon + 1)) : null;
      return tail === null ? part : `${part.slice(0, colon + 1)}${tail}`;
    })
    .join('；');
}

// ---------------------------------------------------------------------------
// run.end summaries → one line
// ---------------------------------------------------------------------------

// Keys the commands put in run.end summaries (src/select/index.js, issue.js,
// remind.js, usage.js, verify.js), plus a few generic ones. An unknown key is
// shown verbatim, so a new counter still appears in the run history.
const SUMMARY_KEY_LABELS = {
  // select
  recipients: '入选',
  totalCents: '合计',
  source: '导出方式',
  snapshotReused: '复用了 24 小时内导出的客户数据',
  // issue
  attempted: '尝试',
  created: '建卡',
  tagged: '打 tag',
  tagFixed: '补打 tag',
  reconciled: '核对补记',
  skipped: '跳过',
  failed: '失败',
  rejected: '被拒（可重试）',
  unknown: '结果不明',
  amountCents: '金额',
  tagFailed: '打 tag 失败',
  reconciledNone: '确认未建成',
  stillUnknown: '仍需人工核对',
  newCardsRefused: '没有建新卡（待建卡）', // people still waiting: on/after REMIND_1_DATE issue makes no new card (live and dry run)
  repairOnly: '只补记和补打 tag', // issue --repair-only (true: the label alone)
  stoppedByDate: '按日期停止', // issue / remind round 1 stopped mid-run at this store date ('YYYY-MM-DD')
  // remind
  eligible: '符合条件',
  planned: '计划',
  sent: '已发',
  retriedUnknown: '补发（之前结果不明）',
  retriedFailed: '重发（之前失败）',
  alreadySent: '之前已发',
  previouslyFailed: '之前失败',
  waitingUnknown: '结果不明待核对',
  notIssued: '未建卡',
  alreadySentByTag: '按 tag 认定已发', // carries the round tag in Shopify, the journal had no send: counted as sent
  roundTagged: '打本轮 tag',
  roundTagFailed: '打本轮 tag 失败',
  roundTagFixed: '补打本轮 tag', // sent earlier but missing the round tag: tagged again at the start of a live run
  roundTagMissing: '缺本轮 tag', // dry run: how many a live run would tag again
  // usage
  issuedCards: '发出的卡',
  usedCards: '已用的卡',
  usedCents: '已用金额',
  orders: '订单',
  ordersTotalCents: '订单总额',
  giftCardCents: '礼品卡抵扣',
  customerPaidCents: '顾客另外支付',
  unmatched: '需人工核对的付款',
  // verify
  cardCount: '卡',
  taggedCount: '带 tag 的客户',
  issueCount: '问题',
  fixedCount: '补记日志',
  counts: '分类',
  // generic
  total: '总数',
  selected: '入选',
  notSelected: '未入选',
  candidates: '候选',
  processed: '处理',
  done: '完成',
  pending: '待发放',
  resolved: '已查明',
  retried: '重试',
  reminded: '已提醒',
  cards: '卡',
  used: '已用',
  payments: '付款',
  issues: '问题',
  limit: '上限',
  from: '从序号',
  to: '到序号',
  issuedCents: '面额',
  stoppedReason: '停止原因',
  stopReason: '停止原因',
  interrupted: '中断',
  reason: '原因',
  message: '说明',
  file: '文件',
};

// Already shown in their own columns of the run history (批次/轮次, 预演/实际).
const SUMMARY_SKIP_KEYS = new Set(['batch', 'round', 'dryRun']);

// A run's mode reads first ("只补记和补打 tag，补打 tag 2"), wherever the command put the key.
const SUMMARY_LEAD_KEYS = new Set(['repairOnly']);

const MAX_SUMMARY_LENGTH = 500;

/** The command a summary came from when the caller does not say (only issue and remind summaries need it). */
function summaryCommand(summary) {
  if ('eligible' in summary || 'alreadySent' in summary || 'round' in summary) return 'remind';
  if ('tagFixed' in summary || 'attempted' in summary || 'seqFrom' in summary) return 'issue';
  return null;
}

function skipLabel(code, command) {
  const [first, second] = command === 'remind' ? [REMIND_SKIP_LABELS, ISSUE_SKIP_LABELS] : [ISSUE_SKIP_LABELS, REMIND_SKIP_LABELS];
  return own(first, code) ?? own(second, code) ?? String(code);
}

function keyLabel(key, ctx) {
  if (key === 'attempted' && ctx.dryRun && ctx.command === 'issue') return '将建卡';
  if (key === 'planned' && ctx.command === 'remind') return ctx.dryRun ? '将发' : '本次要发';
  return own(SUMMARY_KEY_LABELS, key) ?? key;
}

/** Label of a key inside a nested object: `skipped` → skip reasons, `counts` → verify issue types. */
function innerLabel(parent, key, ctx) {
  if (parent === 'skipped') return skipLabel(key, ctx.command);
  if (parent === 'counts') return labelOf(VERIFY_TYPE_LABELS, key);
  return keyLabel(key, ctx);
}

function summaryValue(key, value, depth, ctx) {
  if (value === null || value === undefined || value === '') return '';
  // false flags and zero counters mean "nothing happened": left out
  if (typeof value === 'boolean') return value ? '是' : '';
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value === 0) return '';
    return /cents$/i.test(key) ? formatUsd(value) : value.toLocaleString('en-US');
  }
  if (typeof value === 'string') return key === 'source' ? labelOf(SNAPSHOT_SOURCE_LABELS, value) : value;
  if (Array.isArray(value)) {
    const items = value.slice(0, 10).map((v) => summaryValue('', v, depth + 1, ctx)).filter(Boolean);
    return items.length ? `${items.join('、')}${value.length > 10 ? ` 等 ${value.length} 项` : ''}` : '';
  }
  if (typeof value === 'object') {
    if (depth >= 2) return '…';
    const inner = summaryParts(value, depth + 1, ctx, key);
    return inner.length ? `（${inner.join('，')}）` : '';
  }
  return String(value);
}

function summaryParts(obj, depth, ctx, parent = null) {
  const parts = [];
  const top = depth === 0;
  let seqShown = false;
  const entries = Object.entries(obj);
  // stable sort: lead keys first, everything else in the command's order
  if (top) entries.sort(([a], [b]) => Number(SUMMARY_LEAD_KEYS.has(b)) - Number(SUMMARY_LEAD_KEYS.has(a)));
  for (const [key, value] of entries) {
    if (top && SUMMARY_SKIP_KEYS.has(key)) continue;
    if (top && (key === 'seqFrom' || key === 'seqTo')) {
      // one "序号 3–5" (at the first of the two keys) instead of two counters
      if (seqShown) continue;
      seqShown = true;
      const from = Number.isFinite(obj.seqFrom) ? obj.seqFrom : null;
      const to = Number.isFinite(obj.seqTo) ? obj.seqTo : null;
      if (from !== null && to !== null) parts.push(from === to ? `序号 ${from}` : `序号 ${from}–${to}`);
      else if (from !== null || to !== null) parts.push(`序号 ${from ?? to}`);
      continue;
    }
    if (top && key === 'stopped' && typeof value === 'string' && value) {
      parts.push(own(STOP_LABELS, value) ?? `停止：${value}`);
      continue;
    }
    if (top && key === 'error' && typeof value === 'string' && value) {
      parts.push(`出错：${errorText(value)}`);
      continue;
    }
    const text = summaryValue(key, value, depth, ctx);
    if (!text) continue;
    const label = parent ? innerLabel(parent, key, ctx) : keyLabel(key, ctx);
    if (typeof value === 'boolean') parts.push(label);
    else parts.push(text.startsWith('（') ? `${label}${text}` : `${label} ${text}`);
  }
  return parts;
}

/**
 * run.end summary (any shape) → one Chinese line for the run history, e.g.
 * "尝试 20，建卡 20，打 tag 20，金额 $215.40，序号 1–20". A summary's own `text`
 * is shown as-is; zero counters are left out.
 * @param {*} summary
 * @param {{ command?: string|null, dryRun?: boolean }} [run] the run it belongs to
 *   (command and dry-run flag of its run.start); inferred from the summary when missing
 */
export function summaryText(summary, run = {}) {
  if (summary === null || summary === undefined) return '';
  let text;
  if (typeof summary === 'object' && !Array.isArray(summary)) {
    if (typeof summary.text === 'string' && summary.text.trim()) {
      // verify writes e.g. "失败：Network error calling Shopify: …": translate an embedded Shopify error.
      text = summary.text.trim().replace(/^(失败：)(.+)$/s, (_, head, rest) => `${head}${errorText(rest)}`);
    } else {
      const ctx = {
        command: run.command ?? summaryCommand(summary),
        dryRun: typeof run.dryRun === 'boolean' ? run.dryRun : summary.dryRun === true,
      };
      const parts = summaryParts(summary, 0, ctx);
      text = parts.length ? parts.join('，') : Object.keys(summary).length ? '各项都是 0' : '';
    }
  } else {
    text = summaryValue('', summary, 0, { command: run.command ?? null, dryRun: !!run.dryRun });
  }
  return text.length > MAX_SUMMARY_LENGTH ? `${text.slice(0, MAX_SUMMARY_LENGTH - 1)}…` : text;
}
