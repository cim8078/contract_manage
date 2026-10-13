#!/usr/bin/env node
/**
 * 「潭合同」浏览器逐视图冒烟检查（前端渲染回归）
 * ------------------------------------------------------------------
 * 目的：在真实浏览器中逐个打开左侧菜单视图，检查：
 *   ① 登录页是否渲染（#app 有内容、存在密码框）；
 *   ② 登录后每个视图的工作区是否渲染出**预期内容**（按视图关键字校验）；
 *   ③ 全过程是否出现控制台报错 / 未捕获异常（pageerror）。
 * 用于补足“只验接口与服务端 HTML、没验前端渲染”的盲区：前端一旦出现
 * 未声明的 ref、模板作用域错误、v-if/v-else 配对错误等，本脚本会立即指出**是哪个视图**、
 * 以及具体错误信息（曾据此定位“整页空白”“项目合同视图空白”两次问题）。
 *
 * 用法（项目根目录执行）：
 *   node tools/browser_smoke.mjs --base http://127.0.0.1:5000
 *   node tools/browser_smoke.mjs --base http://127.0.0.1:5000 --user admin --password 123456 --headed
 *
 * 参数：
 *   --base      被测地址（默认 http://127.0.0.1:5000）
 *   --user      登录用户名（默认 admin）
 *   --password  登录口令（默认 123456）
 *   --browser   浏览器通道：msedge（默认，Windows 自带）/ chrome / chromium
 *   --headed    显示浏览器窗口（默认无头）
 *
 * 依赖：本机 Node.js；playwright-core 优先取项目内 node_modules，
 *       否则自动在 npm 的 npx 缓存中查找（执行过 playwright-cli 即已存在）。
 * 退出码：0 全部通过；1 存在失败视图；2 环境/脚本异常。
 */

import { existsSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const argv = process.argv.slice(2);
const opt = (name, def) => {
  const i = argv.indexOf('--' + name);
  return i >= 0 && argv[i + 1] ? argv[i + 1] : def;
};
const base = opt('base', 'http://127.0.0.1:5000').replace(/\/+$/, '');
const userName = opt('user', 'admin');
const password = opt('password', '123456');
const channel = opt('browser', 'msedge');
const headed = argv.includes('--headed');

/** 逐视图关键字：切到该视图后，页面文本中必须出现其中之一 */
const VIEWS = [
  { name: '首页', keywords: ['总合同数', '合同到期提醒'] },
  { name: '项目合同', keywords: ['合同金额', '新增合同'] },
  { name: '资产合同', keywords: ['资产合同台账', '年租金'] },
  { name: '查询统计', keywords: ['统计维度', '按部门'] },
  { name: '用户管理', keywords: ['用户管理', '新增用户'] },
  { name: '部门管理', keywords: ['部门管理', '新增部门'] },
  { name: '日志审计', keywords: ['操作日志'] },
  { name: '短信记录', keywords: ['短信发送记录', '点对点发送'] },
  { name: '系统设置', keywords: ['会话空闲超时'] },
];

function findPlaywrightCore() {
  const cands = [path.join(process.cwd(), 'node_modules', 'playwright-core')];
  const cache = process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'npm-cache', '_npx');
  if (cache && existsSync(cache)) {
    for (const d of readdirSync(cache)) cands.push(path.join(cache, d, 'node_modules', 'playwright-core'));
  }
  return cands.find((p) => existsSync(path.join(p, 'package.json'))) || null;
}

const problems = [];
const say = (s) => process.stdout.write(s + '\n');

const pwPath = findPlaywrightCore();
if (!pwPath) {
  say('✗ 未找到 playwright-core。请任选其一：');
  say('   · 先执行一次 `npx --package @playwright/cli playwright-cli --version` 生成 npx 缓存，再运行本脚本；');
  say('   · 或在项目内 `npm i -D playwright-core playwright`。');
  process.exit(2);
}
const pwMod = await import(pathToFileURL(path.join(pwPath, 'index.js')).href);
const pw = pwMod && pwMod.chromium ? pwMod : (pwMod.default || {});
const chromium = pw.chromium;
if (!chromium) {
  say('✗ playwright-core 导入成功但未取到 chromium（版本不匹配？请更新 playwright-core）。');
  process.exit(2);
}

let browser;
try {
  say(`· 启动浏览器（channel=${channel}${headed ? '，可见' : '，无头'}）→ ${base}`);
  browser = await chromium.launch({ channel, headless: !headed });
  const page = await browser.newPage();
  const errs = [];
  page.on('console', (m) => { if (m.type() === 'error') errs.push('console.error: ' + String(m.text()).slice(0, 200)); });
  page.on('pageerror', (e) => errs.push('pageerror: ' + String(e && e.message).slice(0, 200)));

  await page.goto(base + '/', { waitUntil: 'load' });
  await page.waitForTimeout(1200);

  // ① 登录页渲染
  const loginLen = await page.evaluate(() => (document.getElementById('app')?.innerHTML || '').length);
  const hasPwd = (await page.locator('input[type=password]').count()) > 0;
  const loginOk = loginLen > 100;
  if (!loginOk) problems.push(`登录页未渲染（#app 内容仅 ${loginLen} 字，${errs[0] || '无控制台错误' }）`);
  say(`[${loginOk ? 'PASS' : 'FAIL'}] 登录页渲染（#app ${loginLen} 字，密码框 ${hasPwd ? '有' : '无'}）`);

  // ② 登录（已是登录态则跳过）
  if (hasPwd) {
    await page.getByPlaceholder('请输入用户名').fill(userName);
    await page.getByPlaceholder('请输入密码').fill(password);
    await page.getByRole('button', { name: /登\s*录/ }).click();
    await page.waitForTimeout(2500);
    say(`· 已提交登录（${userName}）`);
  } else {
    say('· 未检测到登录表单（可能已是登录态），直接检查各视图');
  }

  // ③ 逐视图切换
  let checked = 0;
  for (const v of VIEWS) {
    const item = page.locator('.menu-item', { hasText: v.name }).first();
    if ((await item.count()) === 0) continue;
    errs.length = 0;
    await item.click();
    await page.waitForTimeout(1600);
    const text = await page.evaluate(() => (document.body.innerText || '').replace(/\s+/g, ' '));
    const hit = v.keywords.some((k) => text.includes(k));
    const ok = errs.length === 0 && hit;
    checked += 1;
    if (!ok) {
      const why = errs.length ? errs.join(' / ') : `工作区未出现预期内容（${v.keywords.join(' 或 ')}）`;
      problems.push(`${v.name}：${why}`);
      say(`[FAIL] ${v.name}（文本 ${text.length} 字）→ ${why}`);
    } else {
      say(`[PASS] ${v.name}（文本 ${text.length} 字）`);
    }
  }
  if (checked === 0) problems.push('未找到任何左侧菜单项（页面可能未进入主界面）');

  say('');
  if (problems.length) {
    say(`✗ 冒烟未通过：共 ${problems.length} 项`);
    problems.forEach((p, i) => say(`   ${i + 1}) ${p}`));
  } else {
    say(`✓ 冒烟通过：登录页与 ${checked} 个视图均正常渲染，且无控制台错误。`);
  }
} catch (e) {
  problems.push('脚本异常：' + (e && e.message));
  say('✗ ' + problems[problems.length - 1]);
} finally {
  if (browser) { try { await browser.close(); } catch (_) {} }
}
process.exit(problems.length ? 1 : 0);
