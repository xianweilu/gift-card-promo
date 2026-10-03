#!/usr/bin/env node
// Command-line entry point. Each subcommand is one step of the campaign that a
// person reviews before running the next; every step rewrites the Excel report.
// Exit codes: 0 ok · 1 stopped or failed · 2 usage error / date guard · 130 interrupted.
import fs from 'node:fs';
import path from 'node:path';
import { loadConfig, ROOT_DIR } from './src/config.js';
import { USAGE, parseCommand } from './src/cli.js';

/**
 * First Ctrl+C asks issue/remind to stop after the current customer (they then
 * rewrite the Excel); a second one exits at once. The journal records every
 * step before it happens, so the next run reconciles whatever was cut off.
 */
function interruptible() {
  const controller = new AbortController();
  let presses = 0;
  const onSignal = () => {
    presses += 1;
    if (presses === 1) {
      console.error('\n收到停止信号：做完当前这个人就停止，然后更新 Excel。再按一次 Ctrl+C 立即退出（下次运行会自动核对中断的那个人）。');
      controller.abort();
    } else {
      console.error('立即退出。');
      process.exit(130);
    }
  };
  process.on('SIGINT', onSignal);
  process.on('SIGTERM', onSignal);
  return {
    signal: controller.signal,
    dispose() {
      process.off('SIGINT', onSignal);
      process.off('SIGTERM', onSignal);
    },
  };
}

/** issue and remind print their own dry-run / live banner; this only makes a test campaign obvious. */
function modeBanner(config) {
  if (config.testCustomerIds.length) {
    console.info(`测试活动：CAMPAIGN_ID=${config.campaignId}，SENT_TAG=${config.sentTag}，只处理 TEST_CUSTOMER_IDS 里的 ${config.testCustomerIds.length} 个客户`);
  }
}

/** `check`: token, scopes, shop settings and the campaign configuration. Read-only. */
async function runCheck(config) {
  const { connect } = await import('./src/connect.js');
  const { gql } = await import('./src/shopify.js');
  const { SHOP_INFO } = await import('./src/queries.js');
  const { campaignPaths } = await import('./src/campaign.js');
  const { campaignNote } = await import('./src/giftcards.js');

  const { scopes } = await connect(config, { log: console });
  console.info(`已连接 ${config.shop}.myshopify.com（API ${config.apiVersion}），权限：${scopes.join(', ')}`);
  const { shop } = await gql(SHOP_INFO);

  const problems = [];
  const warnings = [];
  if (shop.currencyCode !== config.giftCardCurrency) problems.push(`GIFT_CARD_CURRENCY=${config.giftCardCurrency} 和店铺货币 ${shop.currencyCode} 不一致`);
  if (config.timezone && config.timezone !== shop.ianaTimezone) warnings.push(`TIMEZONE=${config.timezone} 和店铺时区 ${shop.ianaTimezone} 不一致，日期按 TIMEZONE 计算`);
  if (!config.giftCardExpiresOn) warnings.push('GIFT_CARD_EXPIRES_ON 没有设置：卡永不过期，邮件里的到期日是空的');
  if (!config.giftCardTemplateSuffix) {
    warnings.push('GIFT_CARD_TEMPLATE_SUFFIX 没有设置：客户会收到原版邮件，而不是活动邮件');
  } else {
    for (const name of ['gift-card-created.subject.liquid', 'gift-card-created.body.liquid']) {
      const file = path.join(ROOT_DIR, 'notifications', name);
      const text = fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '';
      if (!text.includes(`'${config.giftCardTemplateSuffix}'`)) {
        problems.push(`notifications/${name} 没有判断 template_suffix == '${config.giftCardTemplateSuffix}'，活动卡会收到原版邮件`);
      }
    }
  }
  if (!config.testCustomerIds.length) {
    if (/test/i.test(config.giftCardNote)) warnings.push(`GIFT_CARD_NOTE="${config.giftCardNote}" 含 test，正式活动建议改成正式名称`);
    for (const [key, value] of [['LAUNCH_DATE', config.launchDate], ['REMIND_1_DATE', config.remind1Date], ['REMIND_2_DATE', config.remind2Date]]) {
      if (!value) warnings.push(`${key} 没有设置`);
    }
  }

  const paths = campaignPaths(config);
  console.info(`店铺时区 ${shop.ianaTimezone}，货币 ${shop.currencyCode}`);
  console.info(`活动 ${config.campaignId}${config.testCustomerIds.length ? '（测试活动）' : ''}，发放 tag ${config.sentTag}，DRY_RUN=${config.dryRun}`);
  console.info(`日期：首封 ${config.launchDate || '未设置'}，第一次提醒 ${config.remind1Date || '未设置'}，第二次提醒 ${config.remind2Date || '未设置'}，到期日 ${config.giftCardExpiresOn || '未设置'}`);
  console.info(`卡的内部备注："${campaignNote(config.giftCardNote, config.campaignId)}"，template suffix："${config.giftCardTemplateSuffix}"`);
  console.info(`活动数据目录：${paths.dir}`);
  console.info(`Excel：${paths.excel}`);
  for (const w of warnings) console.warn(`注意：${w}`);
  for (const p of problems) console.error(`问题：${p}`);
  if (!problems.length) console.info('检查通过');
  return problems.length ? 1 : 0;
}

async function dispatch(args, config) {
  const log = console;
  switch (args.command) {
    case 'check':
      return runCheck(config);

    case 'select': {
      const { runSelect } = await import('./src/select/index.js');
      return (await runSelect({ config, refresh: args.refresh, log })).exitCode;
    }

    case 'issue': {
      const { runIssue } = await import('./src/issue.js');
      const stop = interruptible();
      try {
        return (await runIssue({
          config,
          limit: args.limit,
          retryFailed: args.retryFailed,
          repairOnly: args.repairOnly,
          log,
          signal: stop.signal,
        })).exitCode;
      } finally {
        stop.dispose();
      }
    }

    case 'remind': {
      const { runRemind } = await import('./src/remind.js');
      const stop = interruptible();
      try {
        return (await runRemind({
          config,
          round: args.round,
          limit: args.limit,
          retryUnknown: args.retryUnknown,
          retryFailed: args.retryFailed,
          log,
          signal: stop.signal,
        })).exitCode;
      } finally {
        stop.dispose();
      }
    }

    case 'usage': {
      const { runUsage } = await import('./src/usage.js');
      return (await runUsage({ config, log })).exitCode;
    }

    case 'verify': {
      const { runVerify } = await import('./src/verify.js');
      return (await runVerify({ config, log })).exitCode;
    }

    case 'export': {
      const { runExport } = await import('./src/export-command.js');
      // runExport resolves --out itself (~, relative paths, a folder, a missing .xlsx extension).
      return (await runExport({ config, refresh: args.refresh, out: args.out, log })).exitCode;
    }

    case 'preview': {
      const { runPreview } = await import('./src/preview.js');
      return (await runPreview({ config, variant: args.variant, seq: args.seq, open: args.open, log })).exitCode;
    }

    default:
      throw new Error(`未知命令 "${args.command}"`);
  }
}

async function main(argv) {
  let args;
  try {
    args = parseCommand(argv);
  } catch (err) {
    console.error(err.message);
    console.error(`\n${USAGE}`);
    return 2;
  }
  if (args.command === 'help') {
    console.log(USAGE);
    return argv.length ? 0 : 2;
  }

  let config;
  try {
    config = loadConfig();
  } catch (err) {
    console.error(`配置错误：${err.message}`);
    return 1;
  }
  modeBanner(config);

  try {
    return await dispatch(args, config);
  } catch (err) {
    console.error(`出错：${err.message}`);
    return 1;
  }
}

// Set exitCode instead of calling process.exit() so buffered output is never cut off.
process.exitCode = await main(process.argv.slice(2));
