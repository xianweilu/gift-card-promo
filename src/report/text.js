// Every text the campaign workbook shows, per language: TEXT.zh (the Chinese edition,
// gift-card-promo-<id>.xlsx) and TEXT.en (the English edition, gift-card-promo-<id>-en.xlsx).
// Both packs have the same keys; templated texts are functions. src/report/excel.js and
// src/report/labels.js take every visible word, number format with words and separator
// from the active pack, nothing from inline literals.
//
// T.stored.* translate what the data files hold in Chinese (selection.json, journal.jsonl,
// verify.json, usage.json, run summaries). The Chinese pack shows the stored value unchanged;
// the English pack renders or translates it, and shows unknown text unchanged (never throws).
//
// CONSOLE holds the console messages of the workbook writer: the console stays Chinese for
// both editions, so they are not part of a pack.
//
// No I/O here. This module must not import src/report/labels.js (labels.js imports it).

import { describeTiers } from '../select/amount.js';
import { channelLabel } from '../select/rules.js';
import { ruleLabelEn, notSelectedReasonsEn, formulaEn, channelLabelEn, storedTextEn } from './translate-en.js';

const freeze = (o) => Object.freeze(o);

/** Deep-freeze plain objects and arrays (functions are left as they are). */
function deepFreeze(value) {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const v of Object.values(value)) deepFreeze(v);
    Object.freeze(value);
  }
  return value;
}

// ---------------------------------------------------------------------------
// Chinese (the original wording; test/excel-golden.test.js proves it unchanged)
// ---------------------------------------------------------------------------

/** "prefix：tail", or the prefix alone when there is no tail. */
const withTail = (prefix, tail = '') => (tail.trim() ? `${prefix}：${tail}` : prefix);

const SHEETS = freeze({
  usageReport: '使用报告',
  summary: '汇总',
  recipients: '发放名单',
  notSelected: '未入选',
  duplicates: '同地址重复',
  usageDetail: '使用明细',
  verify: '核对',
  journal: '操作日志',
  help: '说明',
});

/** Issue status of a recipient (Excel "状态" column). */
const STATUS_LABELS = freeze({
  pending: '待发放',
  in_progress: '进行中',
  unknown: '需人工核对',
  created: '已建卡未打tag',
  done: '已完成',
  failed: '失败',
  skipped: '发放前跳过',
});

// From REMIND_1_DATE on, issue only records leftovers and fixes tags for the real
// campaign (fix spec D2): the statuses that wait for a new card say so.
const NO_NEW_CARDS_FROM_REMIND_1 = '正式活动从 REMIND_1_DATE 起 issue 不再建新卡。';

/** What each status means (sheet "说明"). */
const STATUS_HELP = freeze({
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
const ISSUE_SKIP_LABELS = freeze({
  'customer-deleted': '客户已删除',
  'no-email': '邮箱没了',
  'relay-email': '邮箱域名在排除名单里',
  'not-subscribed': NOT_SUBSCRIBED_LABEL,
  'already-tagged': '已有发放 tag',
  'ordered-since-snapshot': '导出后下过单',
  'address-ordered-since-snapshot': '同地址账户导出后下过单',
});

/** Reminder skip reasons written by remind (journal op "remind.skip"). */
const REMIND_SKIP_LABELS = freeze({
  'no-card': '没有卡',
  'multiple-cards': '有多张卡',
  'card-disabled': '卡已停用',
  'card-expired': '卡已过期',
  used: '已用过卡',
  'customer-deleted': '客户已删除',
  'no-email': '没有邮箱',
  'not-subscribed': NOT_SUBSCRIBED_LABEL,
});

/** "第 N 次提醒" cell of a reminder found sent by its round tag (journal op remind.found, source 'tag'). */
const REMIND_SENT_BY_TAG = '已发（按 Shopify 上的本轮 tag 补记）';
/** Added to a sent reminder's cell when tagging the customer with the round tag failed (remind.tag.fail). */
const REMIND_TAG_MISSING_SUFFIX = '；本轮 tag 没打上，下次运行会补打';

/** Reminder cell texts explained on the "说明" sheet. */
const REMIND_HELP = freeze([
  freeze(['（空白）', '这一轮还没有处理这个人。']),
  freeze(['已发 MM-DD HH:MM', '已让 Shopify 重发礼品卡邮件，时间是店铺时间；发出后给客户打上本轮 tag（见“本轮 tag”）。每一轮每人最多一封。']),
  freeze([REMIND_SENT_BY_TAG, '本地日志里这一轮没有“已发”的记录（例如日志丢失，或被更早的备份覆盖），但客户在 Shopify 上带着本轮 tag，所以算作这一轮已发，不会再发。发送时间不详。']),
  freeze([`已发 MM-DD HH:MM${REMIND_TAG_MISSING_SUFFIX}`, '提醒已经发出，但给客户打本轮 tag 失败了：本地日志记着已发，这一轮不会再给他发。下次真实运行同一轮时会先补打这个 tag（见“补打本轮 tag”），补打成功后这里只显示“已发 MM-DD HH:MM”。']),
  freeze(['跳过：原因', `这一轮不符合提醒条件，例如已用过卡、${NOT_SUBSCRIBED_LABEL}、卡已停用或过期。再次运行同一轮会重新判断。`]),
  freeze(['失败：原因', 'Shopify 拒绝发送，原因写在后面。不会自动重发；原因解决后可用 remind --retry-failed 重发。']),
  freeze(['结果不明', 'Shopify 的回复丢失，不知道邮件是否发出。不会自动重发（重发一定会多一封邮件）；确认后可用 remind --retry-unknown 补发。']),
  freeze(['进行中', 'remind 正在处理这个人。只在 remind 运行时显示；remind 不在运行时显示为“结果不明”。']),
]);

/** Customer kind of a recipient. */
const KIND_LABELS = freeze({ ordered: '有下单', never: '从没下单', test: '测试' });

/** Shopify email marketing states ('NONE' = the commands' stand-in for a missing state). */
const MARKETING_LABELS = freeze({
  SUBSCRIBED: '已订阅',
  NOT_SUBSCRIBED: '未订阅',
  UNSUBSCRIBED: '已退订',
  PENDING: '待确认',
  INVALID: '无效',
  REDACTED: '已隐去',
  NONE: '没有营销状态',
});

/** Notes for recipients whose amount is not based on their very last order. */
const BASIS_NOTES = freeze({
  'last-cancelled': '最近一笔已取消，按更早的付费订单计算',
  'last-test': '最近一笔是测试订单，按更早的付费订单计算',
  'last-zero': '最近一笔是 $0，按更早的付费订单计算',
});
const NEVER_NOTES = freeze({
  'only-cancelled-or-zero': '订单都已取消或都是 $0，按从没下单处理',
});

/** Journal ops (sheet "操作日志"). */
const OP_LABELS = freeze({
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
const RECONCILE_SOURCE_LABELS = freeze({
  'issue-reconcile': 'issue 开始前核对',
  'issue-inline': '建卡结果不明后立即核对',
  preflight: '发放前复查',
  verify: 'verify 核对',
});

/** verify.json issue types (sheet "核对"). */
const VERIFY_TYPE_LABELS = freeze({
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
const SNAPSHOT_SOURCE_LABELS = freeze({
  bulk: '全店导出',
  paginated: '普通分页',
  nodes: '按客户 ID 读取',
});

/** run.end summary `stopped` values written by remind. */
const STOP_LABELS = freeze({
  rejected: '已停止：Shopify 没有接受发送请求',
  'too-many-failures': '已停止：连续多次没有发送成功',
  aborted: '已中断（Ctrl+C）',
  error: '已停止：出错',
});

/** Fixed English notes some journal entries carry (anything else is shown as-is). */
const NOTE_LABELS = freeze({
  'already tagged in Shopify': 'Shopify 上已带这个 tag',
});

/** usage.json unmatched[].reason codes that may appear (anything else is shown as-is). */
const UNMATCHED_REASON_LABELS = freeze({
  'no-gift-card-id': '回执里没有礼品卡 ID',
  'no-receipt-id': '回执里没有礼品卡 ID',
  'missing-gift-card-id': '回执里没有礼品卡 ID',
  'no-receipt': '没有回执',
  'bad-receipt': '回执无法解析',
});

/** Command exit codes (sheet "说明"). */
const EXIT_CODE_HELP = freeze([
  freeze([0, '成功']),
  freeze([1, '中途停止或失败（原因见命令输出）']),
  freeze([2, '用法错误：参数不对、缺少 --limit、还没到活动日期等，或已过 REMIND_1_DATE（issue 只补记和补打 tag，不建新卡）、已过 REMIND_2_DATE（不能再发第 1 次提醒）']),
  freeze([130, '按 Ctrl+C 中断（做完当前这个人后退出）']),
]);

/**
 * Why 礼品卡抵扣 and 已用金额 can differ (使用报告, 汇总, 说明): the first is
 * taken at checkout (usage.js F1), the second from the card balances.
 */
const USAGE_SPLIT_NOTE = '礼品卡抵扣按结账时计算，之后退回卡里的钱不扣；已用金额 = 原金额 − 现在余额，已扣除退回卡里的钱。所以有退款退回卡时两者不同，差额就是退回卡里的金额。';

/** Customer / order / gift card numbers in a skip detail ("客户 42"). */
const GID_TYPE_LABELS = freeze({ Customer: '客户', Order: '订单', GiftCard: '礼品卡' });

// Keys the commands put in run.end summaries (src/select/index.js, issue.js,
// remind.js, usage.js, verify.js), plus a few generic ones. An unknown key is
// shown verbatim, so a new counter still appears in the run history.
const SUMMARY_KEY_LABELS = freeze({
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
  // verify's own one-line `text`: shown as the whole line when the pack keeps it (Chinese), so
  // this label only appears when a pack ignores it (English) and the line is built from the counters.
  text: '摘要',
});

const zh = {
  lang: 'zh',

  // ---- dictionaries (src/report/labels.js exports the Chinese ones under these names)
  STATUS_LABELS,
  STATUS_HELP,
  ISSUE_SKIP_LABELS,
  REMIND_SKIP_LABELS,
  REMIND_SENT_BY_TAG,
  REMIND_TAG_MISSING_SUFFIX,
  REMIND_HELP,
  KIND_LABELS,
  MARKETING_LABELS,
  BASIS_NOTES,
  NEVER_NOTES,
  OP_LABELS,
  RECONCILE_SOURCE_LABELS,
  VERIFY_TYPE_LABELS,
  SNAPSHOT_SOURCE_LABELS,
  STOP_LABELS,
  NOTE_LABELS,
  UNMATCHED_REASON_LABELS,
  EXIT_CODE_HELP,
  USAGE_SPLIT_NOTE,
  GID_TYPE_LABELS,
  SUMMARY_KEY_LABELS,

  /**
   * The per-round tags explained on the "说明" sheet ([item, text] pairs); tag1 / tag2 are the
   * campaign's round tags (labels.js roundTagName), the texts quote OP_LABELS and REMIND_SENT_BY_TAG.
   */
  roundTagHelp: (tag1, tag2) => [
    ['本轮 tag', `第 1 次提醒是 ${tag1}，第 2 次提醒是 ${tag2}，两轮互不影响。每发出一封提醒，就给客户打上这一轮的 tag。`],
    ['带本轮 tag 的人', `每次运行 remind（预演也一样）都先查 Shopify 上带本轮 tag 的客户：他们一律算这一轮已发过，不会再发，加 --retry-unknown 或 --retry-failed 也不会。所以即使本地日志丢失或被旧备份覆盖，这一轮也不会给同一个人再发一封；真实运行会把他们在本地日志里补记为“${REMIND_SENT_BY_TAG}”。`],
    ['补打本轮 tag', `打本轮 tag 失败不影响已经发出的提醒（本地日志记着已发），“操作日志”里记一行“${OP_LABELS['remind.tag.fail']}”。下次真实运行同一轮时会先补打；发出满 10 分钟、Shopify 上却查不到本轮 tag 的人（例如 tag 在后台被删掉）也会补打。补打成功，或发现 Shopify 上其实已带这个 tag，记一行“${OP_LABELS['remind.tag.ok']}”。发提醒后直接打上 tag 的不单独记一行，只在运行记录里计数。`],
  ],

  // ---- shared words, separators and formats
  sep: freeze({
    list: '；', // notes, reasons, tier ranges, the parts of an error
    comma: '，', // the parts of a run summary or of a journal result
    enum: '、', // items of an enumeration
    colon: '：', // a text and the Shopify message it introduces (errorText)
  }),
  yes: '是',
  no: '否',
  notSet: '未设置',
  timeZone: freeze({ default: '洛杉矶时间', other: (tz) => `${tz} 时间` }),
  /** "第 N 次提醒": the reminder rounds (汇总, 发放名单 titles, 操作日志 批次/轮次). */
  roundName: (round) => `第 ${round} 次提醒`,
  /** Excel number formats with words (custom formats keep the cells numeric). */
  fmt: freeze({
    cards: '#,##0" 张"',
    orders: '#,##0" 笔"',
    share: '"占 "0.0%',
    shareOfValue: '"占面额 "0.0%',
  }),
  /** Prefixes of money cells: "已用 $1.00" (negative: "已用 -$1.00"). */
  moneyPrefix: freeze({
    faceValue: '面额 ',
    used: '已用 ',
    ordersTotal: '订单总额 ',
    perOrder: '平均每单 ',
    giftCardPaid: '其中礼品卡抵扣 ',
    customerPaid: '顾客另外支付 ',
  }),
  /** A channel (an order's sourceName) as shown: "web" → "网店", unknown codes as they are. */
  channelLabel,
  /** The tier ranges, e.g. ["≤ $10.77 → $10.77", "$10.78–$15.33 → $15.33", "≥ $15.34 → $19.77"]. */
  describeTiers,

  sheets: SHEETS,

  // ---- 使用报告
  usageReport: freeze({
    title: (asOf, tz) => `使用报告（截至 ${asOf}，${tz}）`,
    unknownTime: '未知时间',
    issued: '发出礼品卡',
    used: '已经使用',
    orders: '带来订单',
    byTier: freeze(['按档位', '发出', '已使用', '使用率', '已用金额']),
    byKind: freeze(['按客户类型', '发出', '已使用', '使用率', '已用金额']),
    daily: freeze(['每日', '当天新用的卡', '当天订单', '当天订单金额', '累计使用的卡', '累计使用率']),
    topProducts: freeze(['卖得最多的 10 个商品', '数量', '金额']),
    footer: (asOf, tz) => `截至 ${asOf}（${tz}）。每天运行一次 usage 更新本表；卡的余额小于原金额就算已使用。`,
  }),

  // ---- 汇总
  summary: freeze({
    notice: '本文件由程序生成，修改无效，每次运行会覆盖',
    subtitle: (campaignId, isTest, at, tz) => `活动 ${campaignId} · ${isTest ? '测试活动' : '正式活动'} · 生成于 ${at}（${tz}）`,
    // campaign parameters
    params: '活动参数',
    campaignId: '活动 ID（CAMPAIGN_ID）',
    listType: '名单类型',
    listTypeTest: '测试活动：只取 TEST_CUSTOMER_IDS 里的客户，不套用筛选规则，金额固定',
    listTypeLive: '正式活动：按筛选规则选出',
    createdAt: '名单生成时间',
    exportedAt: '客户数据导出时间',
    exportedNote: (source, count) => `${source}，共 ${count} 个客户`,
    sourceUnknown: '来源未知',
    testAmount: '测试固定金额',
    cutoff: (months) => `近 ${months} 个月下单的起算日（cutoff）`,
    cutoffNote: '这一天 0 点（店铺时间）以后有有效订单的人不发',
    minAge: '注册满',
    days: (n) => `${n} 天`,
    requireSubscribed: '要求已订阅营销邮件',
    percent: '金额比例',
    percentValue: (p) => `${p}%`,
    percentNote: '基数 = 上次有效订单总额 × 比例，四舍五入到分，再对档位',
    tiers: '档位（按基数）',
    average: '有下单入选者上次订单的平均值',
    averageNote: '从没下单的人用“平均值 × 比例”对档位',
    neverAmount: '从没下单的人发放金额',
    expiry: '礼品卡到期日',
    noExpiry: '不设到期日',
    expiryNote: '到期日当天仍可使用',
    expiryChanged: (env, frozen) => `到期日当天仍可使用。注意：.env 的 GIFT_CARD_EXPIRES_ON 现在是 ${env}，但建卡仍用生成名单时的 ${frozen}`,
    launchDate: '首封邮件日期（正式建卡）',
    remind1Date: '第一次提醒',
    remind2Date: '第二次提醒',
    /** A date of .env that differs from the one frozen in the list ('' = 未设置, filled in by the caller). */
    dateChanged: (then, now) => `生成名单时为 ${then}，现在按 .env 为 ${now}`,
    sentTag: '发放 tag（SENT_TAG）',
    excludeTags: '排除的 tag（不分大小写）',
    excludeDomains: '排除的邮箱域名',
    excludeSources: '排除的订单渠道',
    /** "Sellbrite（205641）": a channel whose name differs from its code. */
    sourceWithCode: (label, code) => `${label}（${code}）`,
    currency: '礼品卡币种',
    note: '礼品卡内部备注（note）',
    templateSuffix: '邮件和礼品卡页面模板后缀',
    // funnel
    funnel: '筛选漏斗',
    funnelHeader: freeze(['条件（按顺序判断，第一个不满足的记为排除原因）', '排除人数', '“未入选”表']),
    allCustomers: '全店客户',
    ruleRow: (n, label) => `${n}. ${label}`,
    countedOnly: '只计数',
    listed: '逐行列出',
    finalRecipients: '最终入选',
    funnelNote: (listed, unlisted) => `“未入选”表逐行列出 ${listed} 人；只因第 1–3 条（没邮箱、平台中转或占位邮箱、未订阅营销邮件）被排除的 ${unlisted} 人只计数，不逐行列出。`,
    byDomain: '第 2 条：按邮箱域名',
    byMarketing: '第 3 条：按营销状态',
    byTag: '第 5 条：按 tag',
    byChannel: '第 9 条：按渠道',
    people: '人数',
    emptyKey: '（空）',
    // same-address dedupe
    dedupe: '同地址去重',
    groups: '重复组数',
    accounts: '涉及的候选账户',
    removed: '少发人数（每组只保留一个账户）',
    bulk: '疑似批量注册（一个地址 10 个以上账户）',
    bulkNote: '仍按规则每组发一张，详见“同地址重复”表',
    largest: '最大的几组（账户数）',
    noAddress: '没有可比对地址、不参与去重的入选者',
    // amounts
    amounts: '金额',
    amountsHeader: freeze(['档位', '有下单', '从没下单', '人数', '合计金额']),
    testTier: (amount) => `测试固定金额 ${amount}`,
    total: '合计',
    whyCancelled: (n) => `最近一笔已取消 ${n} 人`,
    whyTest: (n) => `最近一笔是测试订单 ${n} 人`,
    whyZero: (n) => `最近一笔是 $0 ${n} 人`,
    orderedTotal: '有下单的人合计',
    neverTotal: '从没下单的人合计',
    personCount: (n) => `${n} 人`,
    fromEarlier: '按更早的付费订单计算',
    onlyCancelled: '订单都已取消或都是 $0，按从没下单处理',
    median: '有下单的人基数的中位数',
    // issuing progress
    progress: '发放进度',
    progressHeader: freeze(['状态', '人数', '金额']),
    lastBatch: '已发放到第几批',
    nextSeq: '下一个待发放的序号',
    nothingPending: '没有待发放的人',
    skipReasonsHeader: freeze(['发放前跳过的原因', '人数']),
    reasonMissing: '原因未记录',
    progressNote: '金额：已建卡的按卡的金额，其余按名单金额。“进行中”只在 issue 运行时出现；issue 不在运行时，这些人算作“需人工核对”。',
    // reminders
    reminders: '提醒',
    remindersHeader: freeze(['轮次', '提醒日期', '已发', '跳过', '失败', '结果不明', '进行中', '还没处理（已建卡的人）', '本轮 tag']),
    remindSkipHeader: freeze(['提醒跳过的原因', '第 1 次', '第 2 次']),
    remindDateChanged: (round, note) => `第 ${round} 次提醒日期：${note}（remind 按现在的日期判断哪天能发）`,
    audienceTest: '已建卡、卡还没用过（余额等于原金额）、没停用没过期的人（测试活动不看营销订阅状态）',
    audienceLive: '已建卡、卡还没用过（余额等于原金额）、没停用没过期、仍订阅营销邮件的人',
    remindNote: (audience) => `提醒只发给${audience}；每一轮每人最多一封：发出后给客户打上本轮 tag，带本轮 tag 的人这一轮不会再发，即使本地日志丢失。跳过的人再次运行同一轮会重新判断。`,
    // usage
    usage: '使用情况',
    asOf: '截至',
    issuedCards: '发出礼品卡（张）',
    issuedValue: '发出面额',
    usedCards: '已经使用（张）',
    usedValue: '已用金额',
    orders: '带来订单（笔）',
    ordersTotal: '订单总额',
    perOrder: '平均每单',
    giftCardPaid: '礼品卡抵扣',
    customerPaid: '顾客另外支付',
    topProducts: '卖得最多的商品',
    dailyTrend: '每日趋势',
    seeUsageReport: '见“使用报告”表',
    // run history
    runs: '运行记录',
    runsHeader: freeze(['开始时间', '命令', '预演/实际', '批次/轮次', '数量上限', '退出码', '结束时间', '结果摘要']),
    running: '运行中',
    notEnded: '没有正常结束（可能被中断）',
    dryRun: '预演',
    live: '实际',
    noRuns: '还没有运行记录',
  }),

  // ---- 发放名单: title and help (说明 sheet) of every column; functions take the tag they name
  recipients: freeze({
    testTier: '测试',
    columns: freeze({
      seq: freeze({ title: '序号', help: '发放顺序：有下单的人按上次下单日期从近到远，后面是从没下单的人，按注册时间从新到旧。名单生成后不再改变。' }),
      status: freeze({ title: '状态', help: '发放状态，见“发放状态”。整行的颜色和状态对应。' }),
      batch: freeze({ title: '批次', help: '第几批 issue 发放（或跳过）的这个人。' }),
      customerId: freeze({ title: '客户 ID', help: '点击打开 Shopify 后台的客户页面。' }),
      name: freeze({ title: '姓名', help: '导出时的客户姓名。' }),
      email: freeze({ title: '邮箱', help: '导出时的默认邮箱。' }),
      marketing: freeze({ title: '营销状态', help: '导出时的邮件营销状态。' }),
      kind: freeze({ title: '客户类型', help: '有下单 / 从没下单 / 测试。' }),
      basisOrder: freeze({ title: '上次有效订单号', help: '金额依据的订单：最近一笔未取消、非测试、付过钱的订单。点击打开订单。' }),
      basisDate: freeze({ title: '上次下单日期', help: '这笔订单的下单日期（店铺时间）。' }),
      channel: freeze({ title: '最近订单渠道', help: '最近一笔有效订单（未取消、非测试，$0 也算）的来源渠道，第 9 条按它判断；没有有效订单时留空。' }),
      daysAgo: freeze({ title: '距今天数', help: '上次有效订单的下单日期到客户数据导出那天相隔的天数。' }),
      basisTotal: freeze({ title: '上次订单总额', help: '下单时的总价，含运费和税，不扣后来的退款。' }),
      formula: freeze({ title: '金额算式', help: '基数 = 上次订单总额（从没下单的人用平均值）× 比例，四舍五入到分，再对档位。' }),
      tier: freeze({ title: '档位', help: '发放金额所在的档位；测试活动显示“测试”。' }),
      amount: freeze({ title: '礼品卡金额', help: '卡的面额：已建卡的按卡的金额，其余按名单金额。' }),
      groupSize: freeze({ title: '同地址账户数', help: '同一地址上参与去重的候选账户数（含本人）。没有可比对地址的人留空。' }),
      groupOthers: freeze({ title: '同地址其他账户', help: '同地址其他候选账户的客户 ID，这些账户不发。' }),
      hasTag: freeze({
        title: (sentTag) => `是否有 ${sentTag}`,
        help: (sentTag) =>
          `以最近一次从 Shopify 刷新 tag（export --refresh 或 verify）的结果为准：刷新时客户带 ${sentTag} 显示“是”，不带显示“否”，`
          + '所以在后台删掉的 tag 刷新后会显示“否”。刷新之后才打上 tag 的人，按本地日志显示“是”；从没刷新过时只看本地日志。'
          + '刷新读的是 Shopify 的搜索结果，刚打上的 tag 可能要过几分钟才查得到，所以打 tag 后 10 分钟内的刷新不会把本地日志的“是”改成“否”。',
      }),
      giftCardId: freeze({ title: '礼品卡 ID', help: '点击打开 Shopify 后台的礼品卡页面。' }),
      last4: freeze({ title: '卡号后 4 位', help: '卡号的最后 4 位，用来和后台核对。程序从不读取完整卡号。' }),
      createdAt: freeze({ title: '建卡时间', help: '店铺时间。' }),
      taggedAt: freeze({ title: '打 tag 时间', help: '店铺时间。' }),
      remind1: freeze({ title: '第 1 次提醒', help: (roundTag) => `见“提醒状态”。这一轮的 tag 是 ${roundTag}。` }),
      remind2: freeze({ title: '第 2 次提醒', help: (roundTag) => `见“提醒状态”。这一轮的 tag 是 ${roundTag}。` }),
      usedAmount: freeze({ title: '已使用金额', help: '最近一次运行 usage 时这张卡已用掉的金额。' }),
      balance: freeze({ title: '剩余余额', help: '最近一次运行 usage 时这张卡的余额。' }),
      usedOrders: freeze({ title: '使用的订单', help: '用这张卡付过款的订单号。' }),
      notes: freeze({ title: '备注/错误', help: '跳过或失败的原因、Shopify 的错误信息，以及金额依据的特殊情况。' }),
      city: freeze({ title: '城市', help: '客户默认地址的城市。' }),
      province: freeze({ title: '州', help: '客户默认地址的州。' }),
      zip: freeze({ title: '邮编', help: '客户默认地址的邮编。' }),
      orderCount: freeze({ title: '订单数', help: '导出时 Shopify 记录的订单数。' }),
      amountSpent: freeze({ title: '累计消费', help: '导出时 Shopify 记录的累计消费。' }),
      accountCreated: freeze({ title: '注册日期', help: '客户账户的创建日期。' }),
    }),
    /** 发放名单 第 N 次提醒 cells (see REMIND_HELP); `at` is "MM-DD HH:MM" store time. */
    remind: freeze({
      sentAt: (at) => `已发 ${at}`,
      sent: '已发',
      skippedFor: (reason) => `跳过：${reason}`,
      skipped: '跳过',
      failedWith: (error) => `失败：${error}`,
      failed: '失败',
      unknown: '结果不明',
      inProgress: '进行中',
    }),
    /** 发放名单 备注/错误 (joined with sep.list). */
    notes: freeze({
      interrupted: '建卡请求发出后没有记录结果（运行被中断）；下次运行 issue 会先去 Shopify 查找这张卡',
      reconciledFrom: (source) => `卡由${source}在 Shopify 找到后补记`,
      amountDiffers: (cardAmount, listAmount) => `卡的金额 ${cardAmount} 与名单金额 ${listAmount} 不同`,
      bulk: (size) => `同地址 ${size} 个账户，疑似批量注册`,
    }),
  }),

  // ---- 未入选
  notSelected: freeze({
    header: freeze(['客户 ID', '姓名', '邮箱', '营销状态', '主要原因', '全部原因', '上次下单日期', '订单数', '相关账户', '组号', 'tags']),
  }),

  // ---- 同地址重复
  duplicates: freeze({
    header: freeze(['组号', '组内账户数', '规范化地址', '是否保留', '客户 ID', '姓名', '邮箱', '上次付费下单日期', '订单数', '注册日期', '疑似批量注册']),
    kept: '保留',
    notSent: '不发',
    activeTitle: (months, count) => `因同地址账户近 ${months} 个月下过单而不发的人（${count} 人）`,
    activeHeader: freeze(['客户 ID', '姓名', '邮箱', '活跃账户', '活跃账户下单时间']),
  }),

  // ---- 使用明细
  usageDetail: freeze({
    header: freeze(['订单号', '下单时间', '客户 ID', '卡号后 4 位', '卡面额', '本单用卡金额', '订单总额', '已取消', '买了什么']),
    noUnmatched: '没有需要人工核对的礼品卡付款',
    unmatchedTitle: (count) => `需要人工核对的礼品卡付款：回执里没有本活动礼品卡 ID（${count} 笔）`,
    unmatchedHeader: freeze(['订单号', '下单时间', '客户 ID', '付款时间', '金额', '原因']),
  }),

  // ---- 核对
  verify: freeze({
    title: (at, tz, cards, sentTag, tagged) => `核对时间 ${at}（${tz}）；本活动的卡 ${cards} 张；带 ${sentTag} 的客户 ${tagged} 个`,
    unknownTime: '未知',
    countItem: (label, n) => `${label} ${n}`,
    found: (n, counts) => `发现 ${n} 个问题${counts ? `：${counts}` : ''}`,
    none: '没有发现问题',
    header: freeze(['问题类型', '客户 ID', '礼品卡 ID', '日志记录', 'Shopify 实际', '建议操作']),
  }),

  // ---- 操作日志
  journal: freeze({
    header: freeze(['时间', '命令', '批次/轮次', '客户 ID', '动作', '结果/说明', '礼品卡 ID', '错误']),
    amount: (usd) => `金额 ${usd}`,
    last4: (last4) => `卡号后 4 位 ${last4}`,
    retry: '重新发送（之前失败或结果不明）',
    /** Instead of the round tag's name when an entry has no round. */
    roundTagFallback: '本轮 tag',
    alreadyTagged: (tag) => `Shopify 上已带 ${tag}`,
    tagMissing: (tag) => `${tag} 没打上，下次运行会补打`,
    tagRepaired: (tag) => `已补打 ${tag}`,
  }),

  // ---- 说明
  help: freeze({
    header: freeze(['项目', '说明']),
    sheetsSection: '各表说明',
    sheets: freeze([
      freeze([SHEETS.usageReport, '运行 usage 后出现，排在第一张：发出多少卡、用了多少、带来多少订单和金额、按档位和客户类型的使用率、每日趋势、卖得最多的商品。']),
      freeze([SHEETS.summary, '活动参数、筛选漏斗、金额、发放进度、提醒、使用情况和运行记录。']),
      freeze([SHEETS.recipients, '每个入选者一行，按发放顺序（序号）。整行颜色表示发放状态。前 3 列和表头冻结。']),
      freeze([SHEETS.notSelected, '被排除的人和原因。只因第 1–3 条（没邮箱、平台中转或占位邮箱、未订阅营销邮件）被排除的人不逐行列出，只在“汇总”里计数。']),
      freeze([SHEETS.duplicates, '同一地址的多个候选账户只保留一个（按组着色）；下面另列因同地址账户近期下过单而不发的人。']),
      freeze([SHEETS.usageDetail, '运行 usage 后出现：每笔用本活动礼品卡付款一行；下面另列回执里没有礼品卡 ID、需要人工核对的付款。']),
      freeze([SHEETS.verify, '运行 verify 后出现：本地日志和 Shopify 对不上的地方，以及建议操作。']),
      freeze([SHEETS.journal, '每次写操作的记录（建卡、打 tag、跳过、提醒），按时间顺序。发提醒后直接打上本轮 tag 的不单独记一行，见“补打本轮 tag”。']),
      freeze([SHEETS.help, '本页。']),
    ]),
    columnsSection: '发放名单各列',
    statusSection: '发放状态（发放名单“状态”列和整行颜色）',
    remindSection: '提醒状态（“第 1 次提醒”“第 2 次提醒”两列）',
    remindSkipReasons: '提醒跳过的原因',
    issueSkipReasons: '发放前跳过的原因',
    roundTagSection: '本轮 tag（每一轮每人最多一封提醒）',
    amountSection: '金额规则',
    testCampaign: '测试活动',
    testAmount: (usd) => `每人固定 ${usd}，不按档位。`,
    base: '基数',
    baseText: (percent) => `有下单的人：上次有效订单总额 × ${percent}%，四舍五入到分。上次有效订单是最近一笔未取消、非测试、付过钱的订单；最近一笔已取消、是测试订单或是 $0，就继续往前找。`,
    never: '从没下单的人',
    neverText: (percent) => `平均值 × ${percent}%：平均值是本次所有“有下单”入选者上次订单总额的平均数（先四舍五入到分）。订单都已取消或都是 $0 的人也按从没下单处理。`,
    /** One row per tier: [item, text]; `i` is 0-based, `range` one of describeTiers(). */
    tierRow: (i, range) => [`档位 ${i + 1}`, `基数 ${range}`],
    roundFirst: '先四舍五入再对档位',
    roundFirstText: '例：上次订单 $107.74 → 基数 $10.77 → 档位 $10.77；$107.75 → 基数 $10.78 → 档位 $15.33。',
    usageSection: '使用情况的金额（“使用报告”和“汇总”）',
    usageAmounts: '礼品卡抵扣和已用金额',
    otherSection: '其他',
    time: '时间',
    timeText: (timeZone, label) => `所有时间都是店铺时区（${timeZone}，${label}）。`,
    links: '链接',
    linksText: '客户 ID、订单号、礼品卡 ID 可以点击，打开 Shopify 后台对应的页面。',
    readOnly: '只读',
    readOnlyText: '本文件由程序生成，每次运行命令都会重新生成；手动修改会被覆盖，程序也从不读取这份 Excel。',
    filter: '筛选',
    filterText: '表格已加保护（无密码），可以使用筛选和调整列宽。',
    exitCode: (code) => `退出码 ${code}`,
  }),

  // ---- labels.js formatters
  /** A skip detail that is a store-local time ("2026-10-04 11:30 店铺时间"; no brackets: the detail is bracketed). */
  storeTime: (local) => `${local} 店铺时间`,
  /** A skip reason with its detail: "导出后下过单（#1234）". */
  withDetail: (label, detail) => `${label}（${detail}）`,
  /** Shopify error messages built by src/shopify.js, in this language (labels.js errorText parses them). */
  shopifyErrors: freeze({
    /** "<mutation> rejected: <rest>"; the " [CODE]" after each user error is already replaced by userErrorCode. */
    rejected: (rest) => withTail('Shopify 拒绝', rest),
    userErrorCode: (code) => `（${code}）`,
    network: (detail) => withTail('网络错误', detail),
    http: (status, detail) => withTail(`Shopify 返回 HTTP ${status}`, detail),
    throttled: (count) => `Shopify 限流：连续 ${count} 次被拒，已放弃`,
    graphql: (detail) => withTail('Shopify 查询错误', detail),
    nonJson: 'Shopify 返回了无法解析的内容',
    noData: 'Shopify 没有返回数据',
    noPayload: 'Shopify 没有返回结果',
  }),
  /** run.end summaries → one line (labels.js summaryText). */
  runSummary: freeze({
    /** issue's `attempted` in a dry run; remind's `planned` in a dry run / a live run. */
    attemptedDry: '将建卡',
    plannedDry: '将发',
    plannedLive: '本次要发',
    /** A `true` inside a list. */
    yes: '是',
    /** After the first 10 items of a longer list. */
    more: (count) => ` 等 ${count} 项`,
    /** A nested object's parts, already joined with sep.comma. */
    nested: (parts) => `（${parts}）`,
    /** A key's label and its value text ("尝试 7"; a nested object goes right after the label). */
    labelled: (label, text) => (text.startsWith('（') ? `${label}${text}` : `${label} ${text}`),
    seqOne: (n) => `序号 ${n}`,
    seqRange: (from, to) => `序号 ${from}–${to}`,
    stopped: (value) => `停止：${value}`,
    error: (text) => `出错：${text}`,
    allZero: '各项都是 0',
  }),

  // ---- texts stored in the data files (Chinese: shown exactly as stored)
  stored: freeze({
    /**
     * Name of a funnel rule. `rule` = a selection.json funnel.byRule[] entry { n, code, label, count };
     * `params` = selection.params (minAccountAgeDays, inactiveMonths, ...). Chinese: rule.label as stored.
     */
    ruleLabel: (rule, params) => rule.label,
    /**
     * Main reason of a selection.json notSelected[] entry (primaryText; entries of newer lists also
     * have reasons: [{ n, code, detail }]). `info` = { params, timeZone } of the list.
     */
    primaryReason: (entry, info) => entry.primaryText,
    /** All reasons of a notSelected[] entry, one text (allReasons; Chinese joins them with '；'). */
    allReasons: (entry, info) => entry.allReasons,
    /**
     * Amount formula of a selection.json recipients[] entry (formula), rebuilt from kind, basis.totalCents,
     * rawCents, amountCents with selection.averageCents and selection.params.giftPercent. Chinese: as stored.
     */
    formula: (recipient, selection) => recipient.formula,
    /** A funnel.channels key: the channel label select stored (Chinese labels, or a code). */
    channel: (key) => key,
    /**
     * A text our own code wrote into the journal: a note (reconcile.none, tag.ok, remind.tag.ok, ...), a skip
     * detail ('邮箱格式无效', '余额 $5.33 / 面额 $15.33', '到期日 2026-10-19', ...), an error, a run summary
     * string; also each '；' part of an error and the text before '：' when a Shopify message follows it.
     * Unknown text is returned unchanged.
     */
    text: (s) => s,
    /** One text of a verify.json issue; `field` = 'journal' | 'shopify' | 'action' (English: the *En field). */
    verify: (issue, field) => issue[field],
    /**
     * A run summary's own one-line `text` (verify writes it, in Chinese): the text to show, or null to build
     * the line from the summary's counters instead.
     */
    summaryText: (text) => text,
    /** A product name from usage.json (topProducts / line items): '（无名称）' is usage's own stand-in. */
    productName: (name) => name,
  }),
};

// ---------------------------------------------------------------------------
// English (the same keys and shapes as the Chinese pack; test/excel-en.test.js compares them).
// Concise operational wording, ASCII punctuation. Stored Chinese texts are rendered by
// ./translate-en.js.
// ---------------------------------------------------------------------------

/** "prefix: tail", or the prefix alone when there is no tail. */
const withTailEn = (prefix, tail = '') => (tail.trim() ? `${prefix}: ${tail}` : prefix);

const SHEETS_EN = freeze({
  usageReport: 'Usage report',
  summary: 'Summary',
  recipients: 'Recipients',
  notSelected: 'Not selected',
  duplicates: 'Same address',
  usageDetail: 'Usage details',
  verify: 'Verify',
  journal: 'Activity log',
  help: 'Help',
});

const STATUS_LABELS_EN = freeze({
  pending: 'Pending',
  in_progress: 'In progress',
  unknown: 'Needs review',
  created: 'Card created, not tagged',
  done: 'Done',
  failed: 'Failed',
  skipped: 'Skipped before issuing',
});

const NO_NEW_CARDS_FROM_REMIND_1_EN = 'From REMIND_1_DATE on, the live campaign\'s issue runs create no new cards.';

const STATUS_HELP_EN = freeze({
  pending: `Not processed yet, or a check confirmed the last attempt created no card; the next issue run processes this person again. ${NO_NEW_CARDS_FROM_REMIND_1_EN}`,
  in_progress: 'issue is processing this person right now. Shown only while issue is running; when it is not, a row stuck here shows "Needs review".',
  unknown: 'Shopify\'s answer was lost, so it is unknown whether the card exists. The next issue or verify run first looks for the card in Shopify: found, it is recorded; not found after 10 minutes, the attempt is confirmed as not created.',
  created: 'The card exists (Shopify has sent the first email automatically) but the tag is still missing. The next issue run only adds the tag; it does not create another card.',
  done: 'Card created and customer tagged.',
  failed: `Shopify refused to create the card; the reason is in "Notes/errors". Only issue --retry-failed retries it. ${NO_NEW_CARDS_FROM_REMIND_1_EN}`,
  skipped: 'The pre-flight check found this person no longer qualifies; no card is created. The reason is in "Notes/errors".',
});

const NOT_SUBSCRIBED_LABEL_EN = 'Not subscribed to marketing emails';

const ISSUE_SKIP_LABELS_EN = freeze({
  'customer-deleted': 'Customer deleted',
  'no-email': 'Email missing',
  'relay-email': 'Email domain on the exclusion list',
  'not-subscribed': NOT_SUBSCRIBED_LABEL_EN,
  'already-tagged': 'Already has the sent tag',
  'ordered-since-snapshot': 'Ordered after the export',
  'address-ordered-since-snapshot': 'Same-address account ordered after the export',
});

const REMIND_SKIP_LABELS_EN = freeze({
  'no-card': 'No card',
  'multiple-cards': 'More than one card',
  'card-disabled': 'Card disabled',
  'card-expired': 'Card expired',
  used: 'Card already used',
  'customer-deleted': 'Customer deleted',
  'no-email': 'No email',
  'not-subscribed': NOT_SUBSCRIBED_LABEL_EN,
});

const REMIND_SENT_BY_TAG_EN = 'Sent (recorded from the round tag in Shopify)';
const REMIND_TAG_MISSING_SUFFIX_EN = '; round tag missing, the next run adds it';

const REMIND_HELP_EN = freeze([
  freeze(['(blank)', 'This person has not been processed in this round yet.']),
  freeze(['Sent MM-DD HH:MM', 'Shopify was told to resend the gift card email at this store time; the customer then got this round\'s tag (see "Round tag"). At most one email per person per round.']),
  freeze([REMIND_SENT_BY_TAG_EN, 'The local journal has no "sent" record for this round (lost, or overwritten by an older backup), but the customer carries this round\'s tag in Shopify, so the reminder counts as sent and is not sent again. The time is unknown.']),
  freeze([`Sent MM-DD HH:MM${REMIND_TAG_MISSING_SUFFIX_EN}`, 'The reminder went out but tagging the customer with the round tag failed: the journal records it as sent, so this round never emails them again. The next live run of the same round adds the tag first (see "Adding a missing round tag"); once added, the cell shows only "Sent MM-DD HH:MM".']),
  freeze(['Skipped: reason', `Not eligible this round, e.g. card already used, ${NOT_SUBSCRIBED_LABEL_EN.toLowerCase()}, card disabled or expired. Running the same round again re-evaluates them.`]),
  freeze(['Failed: reason', 'Shopify refused to send; the reason follows. Not resent automatically; once fixed, remind --retry-failed resends.']),
  freeze(['Outcome unknown', 'Shopify\'s answer was lost, so it is unknown whether the email went out. Not resent automatically (a resend would always add an email); once checked, remind --retry-unknown resends.']),
  freeze(['In progress', 'remind is processing this person right now. Shown only while remind is running; otherwise the cell shows "Outcome unknown".']),
]);

const KIND_LABELS_EN = freeze({ ordered: 'Ordered before', never: 'Never ordered', test: 'Test' });

const MARKETING_LABELS_EN = freeze({
  SUBSCRIBED: 'Subscribed',
  NOT_SUBSCRIBED: 'Not subscribed',
  UNSUBSCRIBED: 'Unsubscribed',
  PENDING: 'Pending',
  INVALID: 'Invalid',
  REDACTED: 'Redacted',
  NONE: 'No status',
});

const BASIS_NOTES_EN = freeze({
  'last-cancelled': 'Latest order cancelled, based on an earlier paid order',
  'last-test': 'Latest order is a test order, based on an earlier paid order',
  'last-zero': 'Latest order is $0, based on an earlier paid order',
});
const NEVER_NOTES_EN = freeze({
  'only-cancelled-or-zero': 'All orders cancelled or $0, treated as never ordered',
});

const OP_LABELS_EN = freeze({
  'create.start': 'Card creation started',
  'create.ok': 'Card created',
  'create.fail': 'Card refused',
  'create.rejected': 'Card not created (can retry)',
  'create.unknown': 'Card outcome unknown',
  'reconcile.found': 'Check: card found',
  'reconcile.none': 'Check: confirmed not created',
  'tag.ok': 'Tag added',
  'tag.fail': 'Tag failed',
  skip: 'Skipped before issuing',
  'remind.start': 'Reminder started',
  'remind.ok': 'Reminder sent',
  'remind.fail': 'Reminder failed',
  'remind.rejected': 'Reminder not sent (can retry)',
  'remind.unknown': 'Reminder outcome unknown',
  'remind.skip': 'Reminder skipped',
  'remind.found': 'Recorded as sent from the round tag',
  'remind.tag.fail': 'Round tag failed',
  'remind.tag.ok': 'Round tag added',
});

const RECONCILE_SOURCE_LABELS_EN = freeze({
  'issue-reconcile': 'check at issue start',
  'issue-inline': 'check right after the unknown outcome',
  preflight: 'pre-flight check',
  verify: 'verify',
});

const VERIFY_TYPE_LABELS_EN = freeze({
  'missing-in-shopify': 'Card in the journal, not in Shopify',
  'not-in-journal': 'Card in Shopify, not in the journal (recorded now)',
  'not-in-selection': 'Card\'s customer not on the list',
  'duplicate-cards': 'Several cards for one customer',
  'amount-mismatch': 'Amount differs',
  'card-disabled': 'Card disabled',
  'tag-missing': 'Card without the tag',
  'tag-without-card': 'Tag without a card',
  'still-unknown': 'Still needs review',
  'resolved-unknown': 'Needs-review row settled',
});

const SNAPSHOT_SOURCE_LABELS_EN = freeze({
  bulk: 'Bulk export of the whole store',
  paginated: 'Page by page',
  nodes: 'By customer ID',
});

const STOP_LABELS_EN = freeze({
  rejected: 'Stopped: Shopify did not accept the send request',
  'too-many-failures': 'Stopped: too many consecutive send failures',
  aborted: 'Interrupted (Ctrl+C)',
  error: 'Stopped: error',
});

const NOTE_LABELS_EN = freeze({
  'already tagged in Shopify': 'Already carries this tag in Shopify',
});

const UNMATCHED_REASON_LABELS_EN = freeze({
  'no-gift-card-id': 'No gift card ID in the receipt',
  'no-receipt-id': 'No gift card ID in the receipt',
  'missing-gift-card-id': 'No gift card ID in the receipt',
  'no-receipt': 'No receipt',
  'bad-receipt': 'Receipt could not be parsed',
});

const EXIT_CODE_HELP_EN = freeze([
  freeze([0, 'Success']),
  freeze([1, 'Stopped midway or failed (the reason is in the command output)']),
  freeze([2, 'Usage error: bad arguments, missing --limit, campaign date not reached, or past REMIND_1_DATE (issue only records cards and adds tags, creates none) or past REMIND_2_DATE (reminder 1 can no longer be sent)']),
  freeze([130, 'Interrupted with Ctrl+C (exits after finishing the current person)']),
]);

const USAGE_SPLIT_NOTE_EN = 'Paid by gift card is taken at checkout and ignores money refunded back to the card; used amount = original amount - current balance, so refunds back to the card are deducted. After a refund to a card the two differ, by the amount refunded to the card.';

const GID_TYPE_LABELS_EN = freeze({ Customer: 'customer', Order: 'order', GiftCard: 'gift card' });

const SUMMARY_KEY_LABELS_EN = freeze({
  // select
  recipients: 'selected',
  totalCents: 'total',
  source: 'export method',
  snapshotReused: 'reused customer data exported within 24 hours',
  // issue
  attempted: 'attempted',
  created: 'cards created',
  tagged: 'tagged',
  tagFixed: 'tags added later',
  reconciled: 'recorded from Shopify',
  skipped: 'skipped',
  failed: 'failed',
  rejected: 'refused (can retry)',
  unknown: 'outcome unknown',
  amountCents: 'amount',
  tagFailed: 'tag failed',
  reconciledNone: 'confirmed not created',
  stillUnknown: 'still needing review',
  newCardsRefused: 'no new card (still pending)',
  repairOnly: 'records and tags only',
  stoppedByDate: 'stopped by date',
  // remind
  eligible: 'eligible',
  planned: 'planned',
  sent: 'sent',
  retriedUnknown: 'resent (outcome was unknown)',
  retriedFailed: 'resent (had failed)',
  alreadySent: 'sent earlier',
  previouslyFailed: 'failed earlier',
  waitingUnknown: 'unknown outcome awaiting review',
  notIssued: 'no card yet',
  alreadySentByTag: 'counted as sent from the tag',
  roundTagged: 'round tag added',
  roundTagFailed: 'round tag failed',
  roundTagFixed: 'round tag added later',
  roundTagMissing: 'round tag missing',
  // usage
  issuedCards: 'cards issued',
  usedCards: 'cards used',
  usedCents: 'used amount',
  orders: 'orders',
  ordersTotalCents: 'orders total',
  giftCardCents: 'paid by gift card',
  customerPaidCents: 'paid by customer',
  unmatched: 'payments needing review',
  // verify
  cardCount: 'cards',
  taggedCount: 'tagged customers',
  issueCount: 'issues',
  fixedCount: 'journal entries added',
  counts: 'by type',
  // generic
  total: 'total',
  selected: 'selected',
  notSelected: 'not selected',
  candidates: 'candidates',
  processed: 'processed',
  done: 'done',
  pending: 'pending',
  resolved: 'settled',
  retried: 'retried',
  reminded: 'reminded',
  cards: 'cards',
  used: 'used',
  payments: 'payments',
  issues: 'issues',
  limit: 'limit',
  from: 'from seq',
  to: 'to seq',
  issuedCents: 'face value',
  stoppedReason: 'stop reason',
  stopReason: 'stop reason',
  interrupted: 'interrupted',
  reason: 'reason',
  message: 'message',
  file: 'file',
  text: 'summary',
});

/** describeTiers in English: a single tier reads "any → $10.00". */
const describeTiersEn = (tiersCents) => describeTiers(tiersCents).map((t) => t.replace(/^任意 →/, 'any →'));

const en = {
  lang: 'en',

  STATUS_LABELS: STATUS_LABELS_EN,
  STATUS_HELP: STATUS_HELP_EN,
  ISSUE_SKIP_LABELS: ISSUE_SKIP_LABELS_EN,
  REMIND_SKIP_LABELS: REMIND_SKIP_LABELS_EN,
  REMIND_SENT_BY_TAG: REMIND_SENT_BY_TAG_EN,
  REMIND_TAG_MISSING_SUFFIX: REMIND_TAG_MISSING_SUFFIX_EN,
  REMIND_HELP: REMIND_HELP_EN,
  KIND_LABELS: KIND_LABELS_EN,
  MARKETING_LABELS: MARKETING_LABELS_EN,
  BASIS_NOTES: BASIS_NOTES_EN,
  NEVER_NOTES: NEVER_NOTES_EN,
  OP_LABELS: OP_LABELS_EN,
  RECONCILE_SOURCE_LABELS: RECONCILE_SOURCE_LABELS_EN,
  VERIFY_TYPE_LABELS: VERIFY_TYPE_LABELS_EN,
  SNAPSHOT_SOURCE_LABELS: SNAPSHOT_SOURCE_LABELS_EN,
  STOP_LABELS: STOP_LABELS_EN,
  NOTE_LABELS: NOTE_LABELS_EN,
  UNMATCHED_REASON_LABELS: UNMATCHED_REASON_LABELS_EN,
  EXIT_CODE_HELP: EXIT_CODE_HELP_EN,
  USAGE_SPLIT_NOTE: USAGE_SPLIT_NOTE_EN,
  GID_TYPE_LABELS: GID_TYPE_LABELS_EN,
  SUMMARY_KEY_LABELS: SUMMARY_KEY_LABELS_EN,

  roundTagHelp: (tag1, tag2) => [
    ['Round tag', `Reminder 1 uses ${tag1}, reminder 2 uses ${tag2}; the rounds are independent. Every reminder sent tags the customer with that round's tag.`],
    ['People carrying the round tag', `Every remind run (dry runs too) first looks up the customers carrying this round's tag in Shopify: they count as already reminded this round and are never sent again, not even with --retry-unknown or --retry-failed. So even when the local journal is lost or overwritten by an old backup, nobody gets a second email in the same round; a live run records them in the local journal as "${REMIND_SENT_BY_TAG_EN}".`],
    ['Adding a missing round tag', `A failed round tag does not affect the reminder already sent (the local journal records it as sent); the "${SHEETS_EN.journal}" sheet gets a row "${OP_LABELS_EN['remind.tag.fail']}". The next live run of the same round adds the tag first; people whose reminder went out 10 minutes ago or more but who have no round tag in Shopify (e.g. the tag was removed in the admin) get it again too. When the tag is added, or turns out to be in Shopify after all, a row "${OP_LABELS_EN['remind.tag.ok']}" is written. A tag added right after sending gets no row of its own; it is only counted in the run history.`],
  ],

  sep: freeze({
    list: '; ',
    comma: ', ',
    enum: ', ',
    colon: ': ',
  }),
  yes: 'Yes',
  no: 'No',
  notSet: 'not set',
  timeZone: freeze({ default: 'Los Angeles time', other: (tz) => `${tz} time` }),
  roundName: (round) => `Reminder ${round}`,
  fmt: freeze({
    cards: '#,##0" cards"',
    orders: '#,##0" orders"',
    share: '"share "0.0%',
    shareOfValue: '"of face value "0.0%',
  }),
  moneyPrefix: freeze({
    faceValue: 'face value ',
    used: 'used ',
    ordersTotal: 'orders total ',
    perOrder: 'average per order ',
    giftCardPaid: 'paid by gift card ',
    customerPaid: 'paid by customer ',
  }),
  channelLabel: channelLabelEn,
  describeTiers: describeTiersEn,

  sheets: SHEETS_EN,

  // ---- Usage report
  usageReport: freeze({
    title: (asOf, tz) => `Usage report (as of ${asOf}, ${tz})`,
    unknownTime: 'unknown time',
    issued: 'Gift cards issued',
    used: 'Used',
    orders: 'Orders brought in',
    byTier: freeze(['By tier', 'Issued', 'Used', 'Usage rate', 'Used amount']),
    byKind: freeze(['By customer type', 'Issued', 'Used', 'Usage rate', 'Used amount']),
    daily: freeze(['Daily', 'Cards first used that day', 'Orders that day', 'Order total that day', 'Cards used to date', 'Usage rate to date']),
    topProducts: freeze(['Top 10 products', 'Quantity', 'Amount']),
    footer: (asOf, tz) => `As of ${asOf} (${tz}). Run usage once a day to refresh this sheet; a card counts as used once its balance is below the original amount.`,
  }),

  // ---- Summary
  summary: freeze({
    notice: 'Generated by the program; edits are lost, every run overwrites this file',
    subtitle: (campaignId, isTest, at, tz) => `Campaign ${campaignId} · ${isTest ? 'test campaign' : 'live campaign'} · generated ${at} (${tz})`,
    params: 'Campaign parameters',
    campaignId: 'Campaign ID (CAMPAIGN_ID)',
    listType: 'List type',
    listTypeTest: 'Test campaign: only the customers in TEST_CUSTOMER_IDS, no selection rules, fixed amount',
    listTypeLive: 'Live campaign: selected by the rules',
    createdAt: 'List generated at',
    exportedAt: 'Customer data exported at',
    exportedNote: (source, count) => `${source}, ${count} ${count === 1 ? "customer" : "customers"}`,
    sourceUnknown: 'source unknown',
    testAmount: 'Fixed test amount',
    cutoff: (months) => `Cutoff date for "ordered in the last ${months} months"`,
    cutoffNote: 'Anyone with a valid order from 00:00 (store time) on this day gets no card',
    minAge: 'Minimum account age',
    days: (n) => `${n} ${n === 1 ? 'day' : 'days'}`,
    requireSubscribed: 'Requires marketing email subscription',
    percent: 'Gift percentage',
    percentValue: (p) => `${p}%`,
    percentNote: 'Base = last valid order total × percentage, rounded to the cent, then mapped to a tier',
    tiers: 'Tiers (by base)',
    average: 'Average last order of the selected customers who ordered',
    averageNote: 'Never-ordered customers use "average × percentage" mapped to a tier',
    neverAmount: 'Amount for never-ordered customers',
    expiry: 'Gift card expiry date',
    noExpiry: 'no expiry',
    expiryNote: 'The card still works on the expiry date itself',
    expiryChanged: (env, frozen) => `The card still works on the expiry date itself. Note: GIFT_CARD_EXPIRES_ON in .env is now ${env}, but cards are still created with ${frozen}, frozen when the list was generated`,
    launchDate: 'First email date (live card creation)',
    remind1Date: 'Reminder 1',
    remind2Date: 'Reminder 2',
    dateChanged: (then, now) => `${then} when the list was generated, now ${now} per .env`,
    sentTag: 'Sent tag (SENT_TAG)',
    excludeTags: 'Excluded tags (case-insensitive)',
    excludeDomains: 'Excluded email domains',
    excludeSources: 'Excluded order channels',
    sourceWithCode: (label, code) => `${label} (${code})`,
    currency: 'Gift card currency',
    note: 'Gift card internal note',
    templateSuffix: 'Email and gift card page template suffix',
    // funnel
    funnel: 'Selection funnel',
    funnelHeader: freeze(['Rule (checked in order; the first one failed is the exclusion reason)', 'Excluded', `"${SHEETS_EN.notSelected}" sheet`]),
    allCustomers: 'All customers',
    ruleRow: (n, label) => `${n}. ${label}`,
    countedOnly: 'counted only',
    listed: 'listed row by row',
    finalRecipients: 'Final recipients',
    funnelNote: (listed, unlisted) => `The "${SHEETS_EN.notSelected}" sheet lists ${listed} ${listed === 1 ? "person" : "people"} row by row; the ${unlisted} excluded only by rules 1-3 (no email, relay or placeholder email, not subscribed to marketing emails) are counted, not listed.`,
    byDomain: 'Rule 2: by email domain',
    byMarketing: 'Rule 3: by marketing status',
    byTag: 'Rule 5: by tag',
    byChannel: 'Rule 9: by channel',
    people: 'People',
    emptyKey: '(empty)',
    // same-address dedupe
    dedupe: 'Same-address dedupe',
    groups: 'Duplicate groups',
    accounts: 'Candidate accounts involved',
    removed: 'Accounts not sent (one kept per group)',
    bulk: 'Suspected bulk sign-ups (10+ accounts at one address)',
    bulkNote: `Still one card per group by the rule; see the "${SHEETS_EN.duplicates}" sheet`,
    largest: 'Largest groups (accounts)',
    noAddress: 'Recipients without a comparable address (not deduplicated)',
    // amounts
    amounts: 'Amounts',
    amountsHeader: freeze(['Tier', 'Ordered before', 'Never ordered', 'People', 'Total amount']),
    testTier: (amount) => `Fixed test amount ${amount}`,
    total: 'Total',
    whyCancelled: (n) => `latest order cancelled: ${n}`,
    whyTest: (n) => `latest order a test order: ${n}`,
    whyZero: (n) => `latest order $0: ${n}`,
    orderedTotal: 'Total for customers who ordered',
    neverTotal: 'Total for never-ordered customers',
    personCount: (n) => `${n} ${String(n) === '1' ? 'person' : 'people'}`,
    fromEarlier: 'Based on an earlier paid order',
    onlyCancelled: 'All orders cancelled or $0, treated as never ordered',
    median: 'Median base of customers who ordered',
    // issuing progress
    progress: 'Issuing progress',
    progressHeader: freeze(['Status', 'People', 'Amount']),
    lastBatch: 'Last batch issued',
    nextSeq: 'Next pending seq',
    nothingPending: 'nobody pending',
    skipReasonsHeader: freeze(['Reasons skipped before issuing', 'People']),
    reasonMissing: 'reason not recorded',
    progressNote: 'Amount: created cards by the card\'s amount, the rest by the list amount. "In progress" appears only while issue is running; otherwise those people count as "Needs review".',
    // reminders
    reminders: 'Reminders',
    remindersHeader: freeze(['Round', 'Reminder date', 'Sent', 'Skipped', 'Failed', 'Outcome unknown', 'In progress', 'Not processed yet (people with a card)', 'Round tag']),
    remindSkipHeader: freeze(['Reasons reminders were skipped', 'Round 1', 'Round 2']),
    remindDateChanged: (round, note) => `Reminder ${round} date: ${note} (remind uses the current date to decide when it can send)`,
    audienceTest: 'people with a card that is unused (balance equals the original amount), enabled and not expired (a test campaign ignores the marketing subscription)',
    audienceLive: 'people with a card that is unused (balance equals the original amount), enabled and not expired, who are still subscribed to marketing emails',
    remindNote: (audience) => `Reminders go only to ${audience}; at most one per person per round: after sending, the customer gets this round's tag, and anyone carrying it is not sent again this round, even if the local journal is lost. Skipped people are re-evaluated when the same round runs again.`,
    // usage
    usage: 'Usage',
    asOf: 'As of',
    issuedCards: 'Gift cards issued (cards)',
    issuedValue: 'Face value issued',
    usedCards: 'Used (cards)',
    usedValue: 'Used amount',
    orders: 'Orders brought in',
    ordersTotal: 'Orders total',
    perOrder: 'Average per order',
    giftCardPaid: 'Paid by gift card',
    customerPaid: 'Paid by customer',
    topProducts: 'Top products',
    dailyTrend: 'Daily trend',
    seeUsageReport: `See the "${SHEETS_EN.usageReport}" sheet`,
    // run history
    runs: 'Run history',
    runsHeader: freeze(['Started', 'Command', 'Dry run/live', 'Batch/round', 'Limit', 'Exit code', 'Ended', 'Summary']),
    running: 'running',
    notEnded: 'did not end normally (possibly interrupted)',
    dryRun: 'dry run',
    live: 'live',
    noRuns: 'no runs yet',
  }),

  // ---- Recipients
  recipients: freeze({
    testTier: 'Test',
    columns: freeze({
      seq: freeze({ title: 'Seq', help: 'Issuing order: customers who ordered first, by last order date from newest to oldest, then never-ordered customers by sign-up date from newest to oldest. Fixed once the list is generated.' }),
      status: freeze({ title: 'Status', help: 'Issue status, see "Issue status". The row colour matches the status.' }),
      batch: freeze({ title: 'Batch', help: 'The issue batch that issued (or skipped) this person.' }),
      customerId: freeze({ title: 'Customer ID', help: 'Click to open the customer in the Shopify admin.' }),
      name: freeze({ title: 'Name', help: 'Customer name at export time.' }),
      email: freeze({ title: 'Email', help: 'Default email at export time.' }),
      marketing: freeze({ title: 'Marketing', help: 'Email marketing status at export time.' }),
      kind: freeze({ title: 'Customer type', help: 'Ordered before / Never ordered / Test.' }),
      basisOrder: freeze({ title: 'Last valid order', help: 'The order the amount is based on: the latest order that is not cancelled, not a test and was paid. Click to open the order.' }),
      basisDate: freeze({ title: 'Last order date', help: 'Date of that order (store time).' }),
      channel: freeze({ title: 'Latest order channel', help: 'Channel of the most recent valid order (not cancelled, not a test; $0 counts), the one rule 9 judges; empty without a valid order.' }),
      daysAgo: freeze({ title: 'Days ago', help: 'Days between the last valid order and the customer data export.' }),
      basisTotal: freeze({ title: 'Last order total', help: 'Total at the time of the order, shipping and tax included, later refunds not deducted.' }),
      formula: freeze({ title: 'Amount formula', help: 'Base = last order total (the average for never-ordered customers) × percentage, rounded to the cent, then mapped to a tier.' }),
      tier: freeze({ title: 'Tier', help: 'The tier of the gift amount; a test campaign shows "Test".' }),
      amount: freeze({ title: 'Gift card amount', help: 'Face value of the card: created cards by the card\'s amount, the rest by the list amount.' }),
      groupSize: freeze({ title: 'Accounts at address', help: 'Candidate accounts at the same address that took part in the dedupe (this one included). Empty without a comparable address.' }),
      groupOthers: freeze({ title: 'Other accounts at address', help: 'Customer IDs of the other candidate accounts at this address; they get no card.' }),
      hasTag: freeze({
        title: (sentTag) => `Has ${sentTag}`,
        help: (sentTag) =>
          `Based on the latest tag refresh from Shopify (export --refresh or verify): "Yes" when the customer carried ${sentTag} at the refresh, "No" when not, `
          + 'so a tag removed in the admin shows "No" after a refresh. People tagged after the refresh show "Yes" from the local journal; without any refresh only the journal counts. '
          + 'The refresh reads Shopify\'s search results, which may take a few minutes to show a new tag, so a refresh within 10 minutes of tagging does not turn the journal\'s "Yes" into "No".',
      }),
      giftCardId: freeze({ title: 'Gift card ID', help: 'Click to open the gift card in the Shopify admin.' }),
      last4: freeze({ title: 'Card last 4', help: 'Last 4 characters of the card code, to match against the admin. The program never reads full card codes.' }),
      createdAt: freeze({ title: 'Card created', help: 'Store time.' }),
      taggedAt: freeze({ title: 'Tagged', help: 'Store time.' }),
      remind1: freeze({ title: 'Reminder 1', help: (roundTag) => `See "Reminder status". This round's tag is ${roundTag}.` }),
      remind2: freeze({ title: 'Reminder 2', help: (roundTag) => `See "Reminder status". This round's tag is ${roundTag}.` }),
      usedAmount: freeze({ title: 'Used amount', help: 'Amount spent from this card at the latest usage run.' }),
      balance: freeze({ title: 'Balance', help: 'Balance of this card at the latest usage run.' }),
      usedOrders: freeze({ title: 'Orders using the card', help: 'Orders paid with this card.' }),
      notes: freeze({ title: 'Notes/errors', help: 'Why a row was skipped or failed, Shopify error messages, and special cases of the amount basis.' }),
      city: freeze({ title: 'City', help: 'City of the customer\'s default address.' }),
      province: freeze({ title: 'State', help: 'State of the customer\'s default address.' }),
      zip: freeze({ title: 'ZIP', help: 'ZIP of the customer\'s default address.' }),
      orderCount: freeze({ title: 'Orders', help: 'Order count in Shopify at export time.' }),
      amountSpent: freeze({ title: 'Total spent', help: 'Total spent in Shopify at export time.' }),
      accountCreated: freeze({ title: 'Signed up', help: 'Creation date of the customer account.' }),
    }),
    remind: freeze({
      sentAt: (at) => `Sent ${at}`,
      sent: 'Sent',
      skippedFor: (reason) => `Skipped: ${reason}`,
      skipped: 'Skipped',
      failedWith: (error) => `Failed: ${error}`,
      failed: 'Failed',
      unknown: 'Outcome unknown',
      inProgress: 'In progress',
    }),
    notes: freeze({
      interrupted: 'The card request was sent but no outcome was recorded (the run was interrupted), the next issue run first looks for the card in Shopify',
      reconciledFrom: (source) => `Card found in Shopify (${source}) and recorded`,
      amountDiffers: (cardAmount, listAmount) => `Card amount ${cardAmount} differs from the list amount ${listAmount}`,
      bulk: (size) => `${size} accounts at this address, suspected bulk sign-ups`,
    }),
  }),

  // ---- Not selected
  notSelected: freeze({
    header: freeze(['Customer ID', 'Name', 'Email', 'Marketing', 'Main reason', 'All reasons', 'Last order date', 'Orders', 'Related account', 'Group', 'Tags']),
  }),

  // ---- Same address
  duplicates: freeze({
    header: freeze(['Group', 'Accounts in group', 'Normalized address', 'Kept', 'Customer ID', 'Name', 'Email', 'Last paid order date', 'Orders', 'Signed up', 'Suspected bulk']),
    kept: 'Kept',
    notSent: 'Not sent',
    activeTitle: (months, count) => `Not sent because a same-address account ordered in the last ${months} months (${count} ${count === 1 ? 'person' : 'people'})`,
    activeHeader: freeze(['Customer ID', 'Name', 'Email', 'Active account', 'Active account\'s order time']),
  }),

  // ---- Usage details
  usageDetail: freeze({
    header: freeze(['Order', 'Order time', 'Customer ID', 'Card last 4', 'Card face value', 'Card amount on this order', 'Order total', 'Cancelled', 'Items']),
    noUnmatched: 'No gift card payments need review',
    unmatchedTitle: (count) => `Gift card payments needing review: no campaign gift card ID in the receipt (${count})`,
    unmatchedHeader: freeze(['Order', 'Order time', 'Customer ID', 'Payment time', 'Amount', 'Reason']),
  }),

  // ---- Verify
  verify: freeze({
    title: (at, tz, cards, sentTag, tagged) => `Verified ${at} (${tz}); ${cards} campaign cards; ${tagged} customers carrying ${sentTag}`,
    unknownTime: 'unknown',
    countItem: (label, n) => `${label} ${n}`,
    found: (n, counts) => `${n} ${n === 1 ? 'issue' : 'issues'} found${counts ? `: ${counts}` : ''}`,
    none: 'No issues found',
    header: freeze(['Issue type', 'Customer ID', 'Gift card ID', 'Journal', 'Shopify', 'Suggested action']),
  }),

  // ---- Activity log
  journal: freeze({
    header: freeze(['Time', 'Command', 'Batch/round', 'Customer ID', 'Action', 'Result/notes', 'Gift card ID', 'Error']),
    amount: (usd) => `amount ${usd}`,
    last4: (last4) => `card last 4 ${last4}`,
    retry: 'Resent (had failed or outcome unknown)',
    roundTagFallback: 'the round tag',
    alreadyTagged: (tag) => `Already carries ${tag} in Shopify`,
    tagMissing: (tag) => `${tag} not added; the next run adds it`,
    tagRepaired: (tag) => `${tag} added`,
  }),

  // ---- Help
  help: freeze({
    header: freeze(['Item', 'Explanation']),
    sheetsSection: 'Sheets',
    sheets: freeze([
      freeze([SHEETS_EN.usageReport, 'Appears after a usage run, as the first sheet: cards issued, cards used, orders and revenue brought in, usage rates by tier and customer type, the daily trend and the top products.']),
      freeze([SHEETS_EN.summary, 'Campaign parameters, selection funnel, amounts, issuing progress, reminders, usage and run history.']),
      freeze([SHEETS_EN.recipients, 'One row per recipient, in issuing order (seq). The row colour shows the issue status. The first 3 columns and the header are frozen.']),
      freeze([SHEETS_EN.notSelected, `Excluded people and why. People excluded only by rules 1-3 (no email, relay or placeholder email, not subscribed to marketing emails) are not listed row by row, only counted in "${SHEETS_EN.summary}".`]),
      freeze([SHEETS_EN.duplicates, 'Several candidate accounts at one address keep only one (coloured by group); below, the people not sent because a same-address account ordered recently.']),
      freeze([SHEETS_EN.usageDetail, 'Appears after a usage run: one row per payment with a campaign gift card; below, payments whose receipt has no gift card ID and need review.']),
      freeze([SHEETS_EN.verify, 'Appears after a verify run: where the local journal and Shopify disagree, with suggested actions.']),
      freeze([SHEETS_EN.journal, 'Every write operation (card creation, tagging, skips, reminders) in time order. A round tag added right after a reminder gets no row of its own; see "Adding a missing round tag".']),
      freeze([SHEETS_EN.help, 'This page.']),
    ]),
    columnsSection: `${SHEETS_EN.recipients} columns`,
    statusSection: `Issue status ("Status" column and row colour of ${SHEETS_EN.recipients})`,
    remindSection: 'Reminder status ("Reminder 1" and "Reminder 2" columns)',
    remindSkipReasons: 'Reasons reminders were skipped',
    issueSkipReasons: 'Reasons people are skipped before issuing',
    roundTagSection: 'Round tag (at most one reminder per person per round)',
    amountSection: 'Amount rules',
    testCampaign: 'Test campaign',
    testAmount: (usd) => `Fixed ${usd} per person, no tiers.`,
    base: 'Base',
    baseText: (percent) => `Customers who ordered: last valid order total × ${percent}%, rounded to the cent. The last valid order is the latest order that is not cancelled, not a test and was paid; when the latest order is cancelled, a test order or $0, the next older one is used.`,
    never: 'Never-ordered customers',
    neverText: (percent) => `Average × ${percent}%: the average of the last order totals of every selected customer who ordered (rounded to the cent first). People whose orders are all cancelled or all $0 count as never ordered too.`,
    tierRow: (i, range) => [`Tier ${i + 1}`, `base ${range}`],
    roundFirst: 'Round first, then map to a tier',
    roundFirstText: 'Example: last order $107.74 → base $10.77 → tier $10.77; $107.75 → base $10.78 → tier $15.33.',
    usageSection: `Usage amounts ("${SHEETS_EN.usageReport}" and "${SHEETS_EN.summary}")`,
    usageAmounts: 'Paid by gift card vs. used amount',
    otherSection: 'Other',
    time: 'Times',
    timeText: (timeZone, label) => `All times are in the store's time zone (${timeZone}, ${label}).`,
    links: 'Links',
    linksText: 'Customer IDs, order numbers and gift card IDs are links to the matching Shopify admin pages.',
    readOnly: 'Read-only',
    readOnlyText: 'Generated by the program and rebuilt by every command; manual edits are overwritten, and the program never reads this workbook.',
    filter: 'Filtering',
    filterText: 'The sheets are protected (no password); filtering and resizing columns still work.',
    exitCode: (code) => `Exit code ${code}`,
  }),

  // ---- labels.js formatters
  storeTime: (local) => `${local} store time`,
  withDetail: (label, detail) => `${label} (${detail})`,
  shopifyErrors: freeze({
    rejected: (rest) => withTailEn('Shopify refused', rest),
    userErrorCode: (code) => ` (${code})`,
    network: (detail) => withTailEn('Network error', detail),
    http: (status, detail) => withTailEn(`Shopify returned HTTP ${status}`, detail),
    throttled: (count) => `Shopify rate limit: refused ${count} times in a row, gave up`,
    graphql: (detail) => withTailEn('Shopify query error', detail),
    nonJson: 'Shopify returned an unreadable response',
    noData: 'Shopify returned no data',
    noPayload: 'Shopify returned no result',
  }),
  runSummary: freeze({
    attemptedDry: 'would create',
    plannedDry: 'would send',
    plannedLive: 'to send',
    yes: 'yes',
    more: (count) => ` and more (${count} in all)`,
    nested: (parts) => `(${parts})`,
    labelled: (label, text) => `${label} ${text}`,
    seqOne: (n) => `seq ${n}`,
    seqRange: (from, to) => `seq ${from}-${to}`,
    stopped: (value) => `stopped: ${value}`,
    error: (text) => `error: ${text}`,
    allZero: 'all counters are 0',
  }),

  // ---- texts stored in the data files (English: rendered or translated by ./translate-en.js)
  stored: freeze({
    ruleLabel: (rule, params) => ruleLabelEn(rule?.code, params, rule?.label),
    primaryReason: (entry, info) => notSelectedReasonsEn(entry, 'primary', info),
    allReasons: (entry, info) => notSelectedReasonsEn(entry, 'all', info),
    formula: (recipient, selection) => formulaEn(recipient, selection),
    channel: (key) => channelLabelEn(key) || key,
    text: (s) => storedTextEn(s),
    /** The *En field verify writes next to journal / shopify / action; an older file's Chinese text unchanged. */
    verify: (issue, field) => {
      const v = issue?.[`${field}En`];
      return typeof v === 'string' && v !== '' ? v : issue?.[field];
    },
    /** verify's own Chinese line is ignored: the line is built from the summary's counters. */
    summaryText: (text) => null, // the stored text is never shown, whatever it says
    productName: (name) => storedTextEn(name),
  }),
};

/** The text packs: TEXT.zh (Chinese edition) and TEXT.en (English edition), same keys. */
export const TEXT = deepFreeze({ zh, en });

/**
 * The pack of `lang` ('zh' | 'en'; default 'zh'), or `lang` itself when it already is a pack
 * (an object), so callers can pass either. Throws for an unknown language.
 */
export function textFor(lang = 'zh') {
  if (lang !== null && typeof lang === 'object') return lang;
  const pack = Object.hasOwn(TEXT, lang ?? 'zh') ? TEXT[lang ?? 'zh'] : null;
  if (!pack) throw new Error(`no workbook texts for language "${lang}"`);
  return pack;
}

// ---------------------------------------------------------------------------
// Console messages of the workbook writer (Chinese for both editions)
// ---------------------------------------------------------------------------

const OPEN_IN_EXCEL = 'Excel 正打开此文件，请关闭后重新打开才能看到最新内容';
const OPEN_IN_EXCEL_EN = '英文版 Excel 正打开此文件，请关闭后重新打开才能看到最新内容';

export const CONSOLE = Object.freeze({
  openInExcel: OPEN_IN_EXCEL,
  openInExcelEn: OPEN_IN_EXCEL_EN,
  /** The same about an --out copy, naming it. */
  openInExcelAt: (file) => `${OPEN_IN_EXCEL}：${file}`,
  openInExcelEnAt: (file) => `${OPEN_IN_EXCEL_EN}：${file}`,
  noSelection: '还没有名单，请先运行 select',
  /** Put before a warning about the English workbook. */
  englishPrefix: '英文版 Excel：',
  englishFailed: (reason) => `英文版 Excel 没有生成：${reason}`,
  rowLimit: (sheet, max) => `“${sheet}”超过 Excel 的行数上限（${max} 行），后面的行没有写入`,
  unreadable: (label, message) => `${label} 无法读取，本次 Excel 不包含它的内容：${message}`,
  replaceFailed: (target, code) => `无法替换 ${target}（${code}）：文件可能被 Excel 锁住，请关闭后重新运行 export`,
  copyFailed: (target, message) => `另存 Excel 到 ${target} 失败：${message}`,
  copyFailedEn: (target, message) => `另存英文版 Excel 到 ${target} 失败：${message}`,
  /** Why a copy may not go to the other edition's own workbook. */
  copyOntoOther: '目标是另一份语言版本的 Excel，不能覆盖',
  badTimeZone: (tz, fallback) => `时区 "${tz}" 无效，Excel 里的时间改用 ${fallback}`,
  foreignTags: (tag, sentTag) => `tags.json 记录的是 tag "${tag}"，不是本活动的 "${sentTag}"，已忽略`,
  tagsWithoutIds: 'tags.json 里没有客户列表（ids），已忽略',
});
