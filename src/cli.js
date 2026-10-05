// Command-line parsing for index.js, kept separate so it can be tested
// without running a command. Arguments are validated before the
// configuration is loaded, so a typo is a usage error (exit 2) straight away.
import { parseArgs } from 'node:util';

export const USAGE = `用法：node index.js <命令> [选项]

  check                                检查 token、权限和配置，不改任何东西
  select [--refresh]                   第 1 步：导出全店客户 → 筛选 → 同地址去重 → 算金额 → 生成 Excel（只读）
                                       --refresh 强制重新导出（默认 24 小时内复用上次导出的数据）
  issue [--limit N] [--retry-failed] [--repair-only]
                                       第 2 步：建卡（Shopify 自动发首封邮件）并打 tag；默认只预演
                                       --retry-failed 只重试"失败"的人
                                       --repair-only 只补记和补打 tag，不建新卡（不需要 --limit）
  usage                                每日使用情况：卡有没有被用、用在哪笔订单、买了什么；给用过卡的人打 <SENT_TAG>-USED（只读卡和订单）
  verify                               与 Shopify 核对已建的卡和 tag（只读）
  export [--refresh] [--out <路径>]     重新生成 Excel；--refresh 先从 Shopify 刷新 tag 列；--out 另存一份
  preview [first|original] [--seq N] [--open]
                                       本地预览邮件（不连 Shopify）；不写邮件名 = 两种都生成
                                       --seq 用名单里第 N 号客户的名字和金额；--open 在浏览器打开

issue 默认只预演。真实运行要在命令前加 DRY_RUN=false，例如：
  DRY_RUN=false node index.js issue --limit 20

提醒邮件改用 Shopify Email 发送：发送前先跑 usage（给用过卡的人打 <SENT_TAG>-USED），收件人条件
  customer_tags CONTAINS '<SENT_TAG>' AND NOT customer_tags CONTAINS '<SENT_TAG>-USED'

测试活动：把 CAMPAIGN_ID、SENT_TAG、TEST_CUSTOMER_IDS 写进 .env.test（见 .env.example），再用
  node --env-file=.env.test index.js <命令>`;

/** Options accepted by each command (node:util parseArgs format). */
export const COMMANDS = {
  check: { options: {} },
  select: { options: { refresh: { type: 'boolean' } } },
  issue: { options: { limit: { type: 'string' }, 'retry-failed': { type: 'boolean' }, 'repair-only': { type: 'boolean' } } },
  usage: { options: {} },
  verify: { options: {} },
  export: { options: { refresh: { type: 'boolean' }, out: { type: 'string' } } },
  preview: { options: { seq: { type: 'string' }, open: { type: 'boolean' } }, positionals: 1 },
};

export const PREVIEW_VARIANTS = ['first', 'original'];

export class UsageError extends Error {
  constructor(message) {
    super(message);
    this.name = 'UsageError';
  }
}

export function positiveInt(value, flag) {
  const n = Number(value);
  if (!/^\d+$/.test(String(value)) || !Number.isSafeInteger(n) || n < 1) throw new UsageError(`${flag} 必须是正整数，收到 "${value}"`);
  return n;
}

/**
 * Parse argv (without "node index.js") into { command, ...options }.
 * Returns { command: 'help' } for -h/--help. Throws UsageError (or the
 * TypeError node:util raises for an unknown or incomplete option).
 */
export function parseCommand(argv) {
  const [command, ...rest] = argv;
  if (command === undefined || ['help', '--help', '-h'].includes(command)) return { command: 'help' };
  const spec = COMMANDS[command];
  if (!spec) throw new UsageError(`未知命令 "${command}"`);
  const { values, positionals } = parseArgs({
    args: rest,
    options: { ...spec.options, help: { type: 'boolean', short: 'h' } },
    strict: true,
    allowPositionals: !!spec.positionals,
  });
  if (positionals.length > (spec.positionals ?? 0)) throw new UsageError(`多余的参数：${positionals.slice(spec.positionals ?? 0).join(' ')}`);
  if (values.help) return { command: 'help' };

  switch (command) {
    case 'select':
      return { command, refresh: !!values.refresh };
    case 'issue': {
      const retryFailed = !!values['retry-failed'];
      const repairOnly = !!values['repair-only'];
      // --repair-only never creates a card; --retry-failed exists to create them.
      if (repairOnly && retryFailed) throw new UsageError('--repair-only 和 --retry-failed 不能一起用：--repair-only 只补记和补打 tag，不建卡');
      return { command, limit: values.limit === undefined ? undefined : positiveInt(values.limit, '--limit'), retryFailed, repairOnly };
    }
    case 'export':
      if (values.out !== undefined && !values.out.trim()) throw new UsageError('--out 后面要写文件路径');
      return { command, refresh: !!values.refresh, out: values.out ?? null };
    case 'preview': {
      const variant = positionals[0] ?? 'all';
      if (variant !== 'all' && !PREVIEW_VARIANTS.includes(variant)) {
        throw new UsageError(`preview 的邮件只能是 ${PREVIEW_VARIANTS.join('、')}，收到 "${variant}"`);
      }
      return { command, variant, seq: values.seq === undefined ? null : positiveInt(values.seq, '--seq'), open: !!values.open };
    }
    default:
      return { command };
  }
}
