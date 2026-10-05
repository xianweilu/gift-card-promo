// Golden comparison of the Chinese workbook: the scenarios of test/report-fixture.js (frozen
// inputs in test/fixtures/report-scenarios.json) must produce exactly the cells recorded in
// test/fixtures/excel-golden-zh.json: every sheet, row, value, formula, number format and style,
// plus the warnings writeReport returns. The golden was first recorded from the workbook code as it
// was before the Chinese texts moved into src/report/text.js (2026-10-02); it was re-recorded on
// 2026-10-05 when the reminders moved to Shopify Email (the reminder columns, the 提醒 section and
// the round-tag texts left; the 用卡 tag column, line and help came). Since then it proves the
// Chinese edition does not change by accident.
//
// Only when the Chinese output is meant to change (a reviewed wording change), record it again:
//   UPDATE_EXCEL_GOLDEN=1 node --test test/excel-golden.test.js
// and review the diff of test/fixtures/excel-golden-zh.json before committing it.
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

import { writeReport } from '../src/report/excel.js';
import {
  FIXTURES_DIR, SCENARIOS_FILE, buildScenarioInputs, loadScenarios, setupScenario, snapshotWorkbook, formatJson,
} from './report-fixture.js';

const GOLDEN_FILE = path.join(FIXTURES_DIR, 'excel-golden-zh.json');

/** The scenarios file: one input file per line group, each recipient / entry on one line. */
function formatScenarios(scenarios) {
  const depth = { selection: 2, verify: 2, usage: 2 };
  const body = scenarios.map((s) => {
    const fields = Object.keys(s).map((k) => `      ${JSON.stringify(k)}: ${formatJson(s[k], depth[k] ?? 1, '      ')}`);
    return `    {\n${fields.join(',\n')}\n    }`;
  });
  return `{\n  "version": 1,\n  "scenarios": [\n${body.join(',\n')}\n  ]\n}\n`;
}

if (process.env.UPDATE_REPORT_SCENARIOS === '1') {
  test('record the report scenarios (UPDATE_REPORT_SCENARIOS=1)', async () => {
    fs.mkdirSync(FIXTURES_DIR, { recursive: true });
    fs.writeFileSync(SCENARIOS_FILE, formatScenarios(await buildScenarioInputs()));
  });
} else if (process.env.UPDATE_EXCEL_GOLDEN === '1') {
  test('record the Chinese golden workbook (UPDATE_EXCEL_GOLDEN=1)', async () => {
    const golden = { version: 1, scenarios: {} };
    for (const scenario of loadScenarios()) {
      const s = setupScenario(scenario, writeReport);
      try {
        const result = await s.write();
        golden.scenarios[scenario.name] = { warnings: result.warnings, sheets: await snapshotWorkbook(result.file) };
      } finally {
        s.cleanup();
      }
    }
    fs.writeFileSync(GOLDEN_FILE, `${formatJson(golden, 6)}\n`);
  });
} else {
  describe('excel golden (Chinese workbook unchanged)', () => {
    const golden = JSON.parse(fs.readFileSync(GOLDEN_FILE, 'utf8'));
    const scenarios = loadScenarios();

    test('the golden covers every scenario', () => {
      assert.deepEqual(Object.keys(golden.scenarios), scenarios.map((s) => s.name));
    });

    for (const scenario of scenarios) {
      test(`${scenario.name}: every cell of the Chinese workbook is as recorded`, async () => {
        const expected = golden.scenarios[scenario.name];
        const s = setupScenario(scenario, writeReport);
        try {
          const result = await s.write();
          assert.equal(result.file, s.paths.excel);
          assert.deepEqual(result.warnings, expected.warnings, 'the warnings (console texts) are unchanged');
          const sheets = await snapshotWorkbook(result.file);
          assert.deepEqual(sheets.map((x) => x.name), expected.sheets.map((x) => x.name), 'sheet names and order');
          const problems = [];
          sheets.forEach((sheet, i) => {
            const want = expected.sheets[i];
            for (const key of ['columns', 'views', 'autoFilter', 'protection']) {
              if (JSON.stringify(sheet[key]) !== JSON.stringify(want[key])) problems.push(`${sheet.name} ${key}: ${JSON.stringify(sheet[key])} ≠ ${JSON.stringify(want[key])}`);
            }
            if (sheet.rows.length !== want.rows.length) problems.push(`${sheet.name}: ${sheet.rows.length} rows ≠ ${want.rows.length}`);
            const n = Math.max(sheet.rows.length, want.rows.length);
            for (let r = 0; r < n; r += 1) {
              const got = JSON.stringify(sheet.rows[r] ?? null);
              const exp = JSON.stringify(want.rows[r] ?? null);
              if (got !== exp) problems.push(`${sheet.name} row ${r + 1}:\n    got      ${got}\n    expected ${exp}`);
            }
          });
          assert.deepEqual(problems.slice(0, 15), [], `${problems.length} difference(s)`);
        } finally {
          s.cleanup();
        }
      });
    }
  });
}
