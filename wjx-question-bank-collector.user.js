// ==UserScript==
// @name         问卷星通用题库采集 · Excel导出
// @namespace    local.wjx.question-bank
// @version      2.0.1
// @license      MIT
// @homepageURL  https://github.com/DreamYuande/wjx-question-bank-collector
// @supportURL   https://github.com/DreamYuande/wjx-question-bank-collector/issues
// @description  问卷分别保存；识别单选、多选和判断，排除身份信息；自定刷新间隔并导出六列Excel。
// @match        https://*.wjx.com/vm/*
// @match        https://*.wjx.cn/vm/*
// @match        https://*.wjx.com/vj/*
// @match        https://*.wjx.cn/vj/*
// @run-at       document-idle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @noframes
// ==/UserScript==

(() => {
  'use strict';
  function identifySurvey(rawUrl) {
    const url = new URL(rawUrl);
    if (url.protocol !== 'https:' || !/(^|\.)wjx\.(com|cn)$/i.test(url.hostname)) return null;
    const match = url.pathname.match(/^\/(vm|vj)\/([a-z0-9_-]+)\.aspx$/i);
    return match ? { id:match[2], source:`${url.origin}${url.pathname}` } : null;
  }
  const SURVEY = identifySurvey(location.href);
  if (!SURVEY) return;
  // 保持旧版存储键的规则；在原脚本中直接覆盖升级，可继续读取原来的学生手册题库。
  const BANK_KEY = `wjx.${SURVEY.id}.bank.v1`;
  const RUN_KEY = `wjx.${SURVEY.id}.run.v1`;
  const LOCK_KEY = `wjx.${SURVEY.id}.lock.v1`;
  const OWNER_KEY = `wjx.${SURVEY.id}.owner.v1`;
  const SOURCE = SURVEY.source;
  const TITLE = String(document.querySelector('#htitle')?.textContent || document.title || `问卷_${SURVEY.id}`).replace(/\s+/g,' ').trim();
  const DEFAULTS = { interval: 8, idleLimit: 30, roundLimit: 500, target: 0 };
  const clean = s => String(s || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
  const normalized = s => clean(s).normalize('NFKC');
  const textOf = node => {
    if (!node) return '';
    const copy = node.cloneNode(true);
    copy.querySelectorAll('.qtypetip,.req,.topicnumber,script,style,input,textarea,select').forEach(e => e.remove());
    return clean(copy.textContent);
  };
  const imagesOf = node => Array.from(node?.querySelectorAll('img') || [])
    .map(img => ({ url: img.currentSrc || img.src || '', alt: clean(img.alt) }))
    .filter(img => /^https?:\/\//i.test(img.url));
  function questionKey(q) {
    const optionKeys = q.options.map(o => JSON.stringify([normalized(o.text), o.images.map(i => i.url).sort()]));
    // 保留数字、标点、重复选项；只忽略页面题号、空白差异和选项排列顺序。
    return JSON.stringify([q.type, normalized(q.stem), q.images.map(i => i.url).sort(), optionKeys.sort()]);
  }
  function isIdentityField(stem) {
    const text = normalized(stem).replace(/\s/g,'');
    return /^(?:请(?:填写|输入|选择))?(?:(?:您|你|本人|考生|受访者)的?)?(?:基本信息|个人信息|身份信息|姓名|真实姓名|学号|工号|学工号|性别|年龄|手机号|手机号码|联系电话|身份证号|身份证号码|邮箱|电子邮箱|身份|学院|学院部门|部门|班级|专业|学校|所在学院|所属学院|所在部门|所属部门|所在班级|所在学校|所属学校)[:：?？]*$/.test(text);
  }
  function optionPrefix(text) {
    const match = clean(text).match(/^(?:[（(]([A-Za-z])[)）]|([A-Za-z])[.．、:：)）])\s*(.*)$/);
    return match ? {letter:(match[1] || match[2]).toUpperCase(), text:match[3]} : null;
  }
  function normalizeOptions(options) {
    const prefixes = options.map(o => optionPrefix(o.text));
    const letters = new Set(prefixes.filter(Boolean).map(p => p.letter));
    const complete = prefixes.every(Boolean) && letters.size === options.length &&
      options.every((_,index) => letters.has(String.fromCharCode(65+index)));
    return options.map((o,index) => ({...o, rawText:o.rawText ?? o.text,
      text:complete ? prefixes[index].text : o.text}));
  }
  function sectionOf(field) {
    for (let node=field.previousElementSibling; node; node=node.previousElementSibling) {
      if (node.classList.contains('cutfield')) return clean(node.textContent);
    }
    return '';
  }
  function detectQuestionType(field, options) {
    if (field.getAttribute('type') === '4') return '多选题';
    const hint = clean(field.querySelector('.qtypetip')?.textContent) + sectionOf(field);
    if (/判断/.test(hint)) return '判断题';
    if (options.length === 2) {
      const pair = options.map(o => normalized(o.text).toLowerCase()).sort().join('|');
      if (['正确|错误','对|错','false|true','×|√'].includes(pair)) return '判断题';
    }
    return '单选题';
  }
  function scanPage(root = document) {
    const questions = [];
    let skippedIdentity = 0, unsupported = 0;
    for (const field of root.querySelectorAll('.field[topic][type]')) {
      const stemNode = field.querySelector('.topichtml');
      if (isIdentityField(textOf(stemNode))) { skippedIdentity++; continue; }
      if (!['3','4'].includes(field.getAttribute('type'))) { unsupported++; continue; }
      const options = normalizeOptions(Array.from(field.querySelectorAll('.ui-radio,.ui-checkbox')).map((row, index) => {
        const label = row.querySelector('.label,label');
        return { label: String.fromCharCode(65 + index), text: textOf(label), images: imagesOf(label) };
      }));
      const question = {
        type: detectQuestionType(field,options),
        stem: textOf(stemNode), images: imagesOf(stemNode), options,
        topic: field.getAttribute('topic') || '',
        pageNumber: clean(field.querySelector('.topicnumber')?.textContent),
        section: sectionOf(field)
      };
      if ((question.stem || question.images.length) && options.length >= 2 && options.every(o => o.text || o.images.length)) questions.push(question);
      else unsupported++;
    }
    return {questions,skippedIdentity,unsupported};
  }
  function extractQuestions(root = document) {
    return scanPage(root).questions;
  }
  function newBank() {
    return { schemaVersion: 1, surveyId: SURVEY.id, source: SOURCE, title: TITLE,
      settings: { ...DEFAULTS }, questions: [], batches: 0, updatedAt: null };
  }
  function loadBank() {
    const raw = GM_getValue(BANK_KEY, '');
    if (!raw) return newBank();
    const bank = JSON.parse(raw);
    if (bank.schemaVersion !== 1 || !Array.isArray(bank.questions) || bank.questions.some(q => !q.key || !Array.isArray(q.options))) {
      throw new Error('已有题库格式异常，已停止；请先通过油猴备份脚本数据。');
    }
    if ((bank.surveyId && bank.surveyId !== SURVEY.id) || (identifySurvey(bank.source)?.id !== SURVEY.id)) {
      throw new Error('存储中的题库与当前问卷不一致，已停止，避免混入其他问卷。');
    }
    bank.surveyId = SURVEY.id;
    bank.settings = { ...DEFAULTS, ...bank.settings };
    return bank;
  }
  function saveBank(bank) {
    const raw = JSON.stringify(bank);
    GM_setValue(BANK_KEY, raw);
    if (GM_getValue(BANK_KEY, '') !== raw) throw new Error('题库保存校验失败，已停止刷新。');
  }
  function mergeQuestions(bank, batch, timestamp) {
    const known = new Map(bank.questions.map(q => [q.key, q]));
    const seenThisBatch = new Set();
    let added = 0;
    for (const q of batch) {
      const key = questionKey(q);
      if (seenThisBatch.has(key)) continue;
      seenThisBatch.add(key);
      let entry = known.get(key);
      if (!entry) {
        entry = { id: bank.questions.length + 1, key, type: q.type, stem: q.stem,
          images: q.images, options: q.options, section:q.section || '', answer: null, answerStatus: '页面未提供答案',
          topics: [], firstSeen: timestamp, lastSeen: timestamp, seen: 0 };
        bank.questions.push(entry); known.set(key, entry); added++;
      }
      entry.lastSeen = timestamp; entry.seen++;
      if (q.topic && !entry.topics.includes(q.topic)) entry.topics.push(q.topic);
    }
    bank.batches++; bank.updatedAt = timestamp;
    return added;
  }
  function stopReason(run, settings, count) {
    if (settings.target && count >= settings.target) return `达到目标 ${settings.target} 题`;
    if (run.idle >= settings.idleLimit) return `连续 ${run.idle} 轮无新题，已暂停（不代表全部采齐）`;
    if (run.rounds >= settings.roundLimit) return `达到本次 ${settings.roundLimit} 轮上限`;
    return '';
  }
  // CSV 防止题干被 Excel 当作公式执行；JSON/TXT 保留原始文本。
  function csvCell(value) {
    let s = String(value ?? '');
    if (/^[\s\uFEFF]*[=+@-]/.test(s)) s = "'" + s;
    return '"' + s.replace(/"/g, '""') + '"';
  }
  function exportText(bank) {
    return `${bank.title}\n共 ${bank.questions.length} 题；按首次采集顺序排列。\n答案未提供。图片仅保存链接。\n\n` + bank.questions.map(q =>
      `${q.id}. 【${q.type}】${q.stem}\n` + q.images.map(i => `题干图片：${i.url}\n`).join('') +
      q.options.map(o => `${o.label}. ${o.text}${o.images.map(i => ` [图片：${i.url}]`).join('')}`).join('\n') + '\n答案：未提供'
    ).join('\n\n');
  }
  function exportRows(bank) {
    const rows = [['序号', '题型', '类别', '题目', '选项', '答案']];
    const typeOrder = {'单选题':0,'多选题':1,'判断题':2};
    const ordered = [...bank.questions].sort((a,b) => (typeOrder[a.type] ?? 3) - (typeOrder[b.type] ?? 3));
    for (const [i,q] of ordered.entries()) rows.push([i+1, q.type, bank.exportCategory || '',
      q.stem + q.images.map(img => ` [图片：${img.url}]`).join(''),
      q.options.map(o => o.text + o.images.map(img => ` [图片：${img.url}]`).join('')).join('|'),
      q.answer ?? '']);
    return rows;
  }
  function exportCsv(bank) {
    return '\uFEFF' + exportRows(bank).map(row => row.map(csvCell).join(',')).join('\r\n');
  }
  // XLSX_TEMPLATE：由表格工具生成的空白六列工作簿，运行时无需下载任何第三方库。
  const XLSX_TEMPLATE = {"files":{"xl/workbook.xml":"<?xml version=\"1.0\" encoding=\"utf-8\"?><x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheets><x:sheet name=\"题库\" sheetId=\"1\" r:id=\"R678ceb0237dd49ad\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" /></x:sheets></x:workbook>","xl/styles.xml":"<?xml version=\"1.0\" encoding=\"utf-8\"?><x:styleSheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:fonts count=\"3\"><x:font><x:sz val=\"11\" /><x:name val=\"Carlito\" /></x:font><x:font><x:sz val=\"11\" /><x:name val=\"宋体\" /></x:font><x:font><x:b /><x:sz val=\"11\" /><x:name val=\"宋体\" /></x:font></x:fonts><x:fills count=\"2\"><x:fill><x:patternFill patternType=\"none\" /></x:fill><x:fill><x:patternFill patternType=\"gray125\" /></x:fill></x:fills><x:borders count=\"2\"><x:border><x:left /><x:right /><x:top /><x:bottom /><x:diagonal /></x:border><x:border><x:left style=\"thin\"><x:color rgb=\"FF000000\" /></x:left><x:right style=\"thin\"><x:color rgb=\"FF000000\" /></x:right><x:top style=\"thin\"><x:color rgb=\"FF000000\" /></x:top><x:bottom style=\"thin\"><x:color rgb=\"FF000000\" /></x:bottom><x:diagonal /></x:border></x:borders><x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs><x:cellXfs count=\"8\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" /><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" /><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><x:alignment vertical=\"center\" /></x:xf><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><x:alignment horizontal=\"center\" vertical=\"center\" /></x:xf><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><x:alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\" /></x:xf><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><x:alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\" /></x:xf><x:xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><x:alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\" /></x:xf><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><x:alignment horizontal=\"left\" vertical=\"center\" wrapText=\"1\" /></x:xf></x:cellXfs><x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" /></x:cellStyles></x:styleSheet>","xl/theme/theme1.xml":"<?xml version=\"1.0\" encoding=\"utf-8\"?><a:theme name=\"ChatGPT\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:clrScheme name=\"ChatGPT\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\" /></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\" /></a:lt1><a:dk2><a:srgbClr val=\"0E2841\" /></a:dk2><a:lt2><a:srgbClr val=\"E8E8E8\" /></a:lt2><a:accent1><a:srgbClr val=\"156082\" /></a:accent1><a:accent2><a:srgbClr val=\"E97132\" /></a:accent2><a:accent3><a:srgbClr val=\"196B24\" /></a:accent3><a:accent4><a:srgbClr val=\"0F9ED5\" /></a:accent4><a:accent5><a:srgbClr val=\"A02B93\" /></a:accent5><a:accent6><a:srgbClr val=\"4EA72E\" /></a:accent6><a:hlink><a:srgbClr val=\"467886\" /></a:hlink><a:folHlink><a:srgbClr val=\"96607D\" /></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" /><a:ea typeface=\"Calibri Light\" /><a:cs typeface=\"Calibri Light\" /></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" /><a:ea typeface=\"Calibri\" /><a:cs typeface=\"Calibri\" /></a:minorFont></a:fontScheme><a:fmtScheme name=\"ChatGPT\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"67000\" /><a:lumMod val=\"110000\" /><a:satMod val=\"105000\" /></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"73000\" /><a:lumMod val=\"105000\" /><a:satMod val=\"103000\" /></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"81000\" /><a:lumMod val=\"105000\" /><a:satMod val=\"109000\" /></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\" /></a:gradFill><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"94000\" /><a:lumMod val=\"102000\" /><a:satMod val=\"103000\" /></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:shade val=\"100000\" /><a:lumMod val=\"100000\" /><a:satMod val=\"110000\" /></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"78000\" /><a:lumMod val=\"99000\" /><a:satMod val=\"120000\" /></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\" /></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:prstDash val=\"solid\" /></a:ln><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:prstDash val=\"solid\" /></a:ln><a:ln w=\"25400\"><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:prstDash val=\"solid\" /></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst /></a:effectStyle><a:effectStyle><a:effectLst /></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\" /></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\" /></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\" /><a:satMod val=\"170000\" /></a:schemeClr></a:solidFill><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\" /><a:shade val=\"98000\" /><a:lumMod val=\"102000\" /><a:satMod val=\"150000\" /></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\" /><a:shade val=\"90000\" /><a:lumMod val=\"103000\" /><a:satMod val=\"130000\" /></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\" /><a:satMod val=\"120000\" /></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\" /></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>","xl/sharedStrings.xml":"<?xml version=\"1.0\" encoding=\"utf-8\"?><x:sst xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />","_rels/.rels":"﻿<?xml version=\"1.0\" encoding=\"utf-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"R21651b78bd854c77\" /></Relationships>","xl/_rels/workbook.xml.rels":"﻿<?xml version=\"1.0\" encoding=\"utf-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"/xl/styles.xml\" Id=\"Rc38347877bab4bf4\" /><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"/xl/theme/theme1.xml\" Id=\"R8ff7fb9b877b465b\" /><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"/xl/sharedStrings.xml\" Id=\"Rbc98d95b9dff4cca\" /><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/sheet1.xml\" Id=\"R678ceb0237dd49ad\" /></Relationships>","[Content_Types].xml":"﻿<?xml version=\"1.0\" encoding=\"utf-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" /><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" /><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" /><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\" /><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" /><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" /></Types>"},"prefix":"<?xml version=\"1.0\" encoding=\"utf-8\"?><x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetViews><x:sheetView workbookViewId=\"0\"><x:pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\" /></x:sheetView></x:sheetViews><x:sheetFormatPr defaultRowHeight=\"15\" /><x:cols><x:col min=\"1\" max=\"1\" width=\"7.633333206176758\" hidden=\"0\" customWidth=\"1\" /><x:col min=\"2\" max=\"2\" width=\"9.091666221618652\" hidden=\"0\" customWidth=\"1\" /><x:col min=\"3\" max=\"3\" width=\"10.091666221618652\" hidden=\"0\" customWidth=\"1\" /><x:col min=\"4\" max=\"4\" width=\"60.63333511352539\" hidden=\"0\" customWidth=\"1\" /><x:col min=\"5\" max=\"5\" width=\"60.63333511352539\" hidden=\"0\" customWidth=\"1\" /><x:col min=\"6\" max=\"6\" width=\"9.091666221618652\" hidden=\"0\" customWidth=\"1\" /></x:cols><x:sheetData>","suffix":"</x:sheetData><x:pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\" /></x:worksheet>","headerStyle":[6,6,6,6,6,6],"bodyStyle":[5,5,5,7,7,5]};
  function xmlText(value) {
    return String(value).replace(/_x[0-9a-f]{4}_/gi, match => '_x005F_' + match.slice(1))
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, c => `_x${c.charCodeAt(0).toString(16).padStart(4,'0')}_`)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  }
  function zipStored(files) {
    const encoder = new TextEncoder();
    const parts = [], central = [];
    let offset = 0, centralSize = 0;
    const crcTable = Array.from({length:256}, (_,n) => {
      let c=n; for(let j=0;j<8;j++) c=(c&1) ? (0xedb88320^(c>>>1)) : (c>>>1); return c>>>0;
    });
    for (const [name, text] of Object.entries(files)) {
      const path=encoder.encode(name), data=encoder.encode(text);
      let crc=0xffffffff;
      for (const byte of data) crc=crcTable[(crc^byte)&255]^(crc>>>8);
      crc=(crc^0xffffffff)>>>0;
      const local=new Uint8Array(30), l=new DataView(local.buffer);
      l.setUint32(0,0x04034b50,true); l.setUint16(4,20,true); l.setUint16(6,0x0800,true);
      l.setUint16(12,33,true); l.setUint32(14,crc,true); l.setUint32(18,data.length,true);
      l.setUint32(22,data.length,true); l.setUint16(26,path.length,true);
      const dir=new Uint8Array(46), d=new DataView(dir.buffer);
      d.setUint32(0,0x02014b50,true); d.setUint16(4,20,true); d.setUint16(6,20,true);
      d.setUint16(8,0x0800,true); d.setUint16(14,33,true); d.setUint32(16,crc,true);
      d.setUint32(20,data.length,true); d.setUint32(24,data.length,true); d.setUint16(28,path.length,true);
      d.setUint32(42,offset,true);
      parts.push(local,path,data); central.push(dir,path);
      offset+=local.length+path.length+data.length; centralSize+=dir.length+path.length;
      if (offset+centralSize>0xffffffff) throw new Error('Excel文件过大，请改用JSON导出。');
    }
    const end=new Uint8Array(22), e=new DataView(end.buffer), count=Object.keys(files).length;
    e.setUint32(0,0x06054b50,true); e.setUint16(8,count,true); e.setUint16(10,count,true);
    e.setUint32(12,centralSize,true); e.setUint32(16,offset,true);
    return new Blob([...parts,...central,end], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  function exportXlsx(bank) {
    const rows=exportRows(bank);
    if (rows.length>1048576) throw new Error('题目数量超过Excel单表行数限制，请导出JSON。');
    const data=rows.map((row,index) => {
      const rowNumber=index+1;
      const cells=row.map((value,col) => {
        if (String(value).length>32767) throw new Error(`第${rowNumber}行内容超过Excel单元格长度限制，请导出JSON。`);
        const ref=String.fromCharCode(65+col)+rowNumber;
        const style=index===0 ? XLSX_TEMPLATE.headerStyle[col] : XLSX_TEMPLATE.bodyStyle[col];
        if(typeof value==='number') return `<x:c r="${ref}" s="${style}" t="n"><x:v>${value}</x:v></x:c>`;
        return `<x:c r="${ref}" s="${style}" t="inlineStr"><x:is><x:t xml:space="preserve">${xmlText(value)}</x:t></x:is></x:c>`;
      }).join('');
      const lineCount=(value,width) => String(value).split('\n').reduce((n,line) => n+Math.max(1,Math.ceil(Array.from(line).reduce((k,c)=>k+(c.charCodeAt(0)>255?2:1),0)/width)),0);
      const lines=Math.max(lineCount(row[2],9),lineCount(row[3],57),lineCount(row[4],57));
      // 混排的英文单词和中文标点会提前折行，长文本多留一行，避免末行被裁切。
      const height=index===0 ? 24 : Math.min(409,Math.max(30,(lines+(lines>1?1:0))*16+10));
      return `<x:row r="${rowNumber}" ht="${height}" customHeight="1">${cells}</x:row>`;
    }).join('');
    return zipStored({...XLSX_TEMPLATE.files, 'xl/worksheets/sheet1.xml':XLSX_TEMPLATE.prefix+data+XLSX_TEMPLATE.suffix});
  }
  function validateInterval(raw) {
    const value=Number(raw);
    if (String(raw).trim()==='' || !Number.isFinite(value) || value<0) throw new Error('刷新间隔请填写0或更大的秒数，支持小数。');
    return value;
  }
  function filenameFor(bank, kind, timestamp = new Date().toISOString()) {
    const name=String(bank.title || '问卷题库').replace(/[<>:"/\\|?*\u0000-\u001F]/g,'_').replace(/[. ]+$/g,'').slice(0,80);
    return `${name}_${SURVEY.id}_${bank.questions.length}题_${timestamp.replace(/[:.]/g,'-')}.${kind}`;
  }
  // UI 与刷新控制。题库保存在油猴存储，本次运行状态放在当前标签页。
  let bank, run, owner;
  try {
    bank = loadBank();
    owner = sessionStorage.getItem(OWNER_KEY) || crypto.randomUUID();
    sessionStorage.setItem(OWNER_KEY, owner);
    run = JSON.parse(sessionStorage.getItem(RUN_KEY) || 'null') || { active: false, rounds: 0, idle: 0 };
  } catch (e) { alert(`题库采集器：${e.message}`); return; }
  let timer = null, heartbeat = null, ticker = null, ready = false, captured = null, busy = false;
  let batch = [], due = 0, pageInfo = {skippedIdentity:0,unsupported:0};
  if (document.getElementById('wjx-local-bank-collector')) return;
  const host = document.createElement('div');
  host.id = 'wjx-local-bank-collector';
  host.style.cssText = 'position:fixed;right:12px;bottom:12px;z-index:2147483646';
  const ui = host.attachShadow({ mode: 'open' });
  ui.innerHTML = `<style>
    *{box-sizing:border-box}section{font:13px/1.55 system-ui,"Microsoft YaHei",sans-serif;width:290px;max-width:calc(100vw - 24px);color:#24334a;background:#fff;border:1px solid #ccd8e6;border-radius:14px;box-shadow:0 6px 26px #1233;overflow:hidden}
    header{padding:11px 14px;background:#edf4ff;display:flex;justify-content:space-between;align-items:center;font-weight:700}main{padding:12px 14px}p{margin:0 0 8px;white-space:pre-line}small{display:block;color:#66748a;margin-top:8px}label{display:flex;justify-content:space-between;align-items:center;margin:5px 0}input{width:85px;border:1px solid #cbd5e1;border-radius:5px;padding:3px}button{cursor:pointer;border:1px solid #c8d6e8;border-radius:6px;background:#f5f8ff;color:#253e62;padding:6px 9px;font:inherit}button:disabled{opacity:.45;cursor:default}.primary{background:#275fd0;color:#fff;border-color:#275fd0}.row{display:flex;gap:6px;flex-wrap:wrap;margin:9px 0}#status{color:#315ca3}details{margin-top:8px}summary{cursor:pointer}#mini{padding:0 6px}section.fold main{display:none}
    </style><section><header><span>题库采集 · <b id="total">0</b> 题</span><button id="mini" type="button">收起</button></header><main>
    <p id="survey"></p><p id="status">等待题目加载…</p><p id="stats"></p>
    <label>刷新间隔（秒，自定）<input id="interval" type="number" min="0" step="any"></label>
    <div class="row"><button id="start" class="primary" type="button" disabled>开始刷新采集</button><button id="pause" type="button">暂停</button><button id="capture" type="button" disabled>采集本页</button></div>
    <div class="row"><button id="xlsx" class="primary" type="button">导出 Excel</button><button id="json" type="button">导出 JSON</button><button id="csv" type="button">导出 CSV</button><button id="txt" type="button">导出文本</button></div>
    <details><summary>采集设置</summary>
    <label>导出类别（可留空）<input id="category" type="text" placeholder="自行填写"></label>
    <label>无新题暂停（轮）<input id="idleLimit" type="number" min="1" max="10000"></label>
    <label>本次轮数上限<input id="roundLimit" type="number" min="1" max="10000"></label>
    <label>目标题数（0=未知）<input id="target" type="number" min="0" max="100000"></label>
    <button id="save" type="button">保存设置</button></details>
    <small>按问卷分别保存单选/多选/判断题，跳过常见身份信息。答案留空。同一问卷请只运行一个标签页。无新题暂停不等于全部采齐。</small>
    </main></section>`;
  document.body.append(host);
  const $ = id => ui.getElementById(id);
  $('survey').textContent = `${TITLE}\n问卷编号：${SURVEY.id}`;
  const status = msg => { $('status').textContent = msg; };
  function render() {
    $('total').textContent = bank.questions.length;
    const counts = batch.reduce((a,q) => { a[q.type]=(a[q.type] || 0)+1; return a; },{});
    $('stats').textContent = `本页 ${batch.length} 题：单选${counts['单选题'] || 0} / 多选${counts['多选题'] || 0} / 判断${counts['判断题'] || 0}\n跳过身份信息${pageInfo.skippedIdentity}项，其他未支持${pageInfo.unsupported}项\n本次${run.rounds}轮，连续无新题${run.idle}轮`;
    $('start').disabled = !ready || run.active || busy;
    $('capture').disabled = !ready || busy || run.active;
    for (const key of Object.keys(DEFAULTS)) $(key).disabled = run.active;
    $('category').disabled = run.active;
  }
  function writeRun() { sessionStorage.setItem(RUN_KEY, JSON.stringify(run)); }
  function claimLock() {
    const lock = GM_getValue(LOCK_KEY, null);
    if (lock && lock.owner !== owner && lock.until > Date.now()) throw new Error('另一个标签页正在采集；请先暂停它，或关闭后等待60秒。');
    GM_setValue(LOCK_KEY, { owner, until: Date.now() + 60000 });
    if (GM_getValue(LOCK_KEY, null)?.owner !== owner) throw new Error('未能取得采集锁，请只保留一个采集标签页。');
  }
  function clearTimers() {
    clearTimeout(timer); clearInterval(heartbeat); clearInterval(ticker);
    timer = heartbeat = ticker = null;
  }
  function pause(message = '已暂停，题库已保留') {
    run.active = false; clearTimers();
    try {
      writeRun();
      if (GM_getValue(LOCK_KEY, null)?.owner === owner) GM_setValue(LOCK_KEY, null);
    } catch (e) { message += `；状态保存失败：${e.message}`; }
    status(message); render();
  }
  function guarded(action) { return async () => { try { await action(); } catch (e) { pause(e.message); } }; }
  function collect() {
    pageInfo = scanPage(); batch = pageInfo.questions;
    if (!batch.length) throw new Error('当前页面没有可采集的选择题，已停止刷新。');
    const signature = JSON.stringify(batch);
    if (captured?.signature === signature) return captured;
    claimLock();
    bank = loadBank();
    const added = mergeQuestions(bank, batch, new Date().toISOString());
    saveBank(bank);
    captured = { added, total: bank.questions.length, signature };
    render(); return captured;
  }
  function readSettings() {
    const limits = { idleLimit: [1,10000], roundLimit: [1,10000], target: [0,100000] };
    const result = {interval:validateInterval($('interval').value)};
    for (const [key,[min,max]] of Object.entries(limits)) {
      const value = Number($(key).value);
      if (!Number.isInteger(value) || value < min || value > max) throw new Error(`设置无效：${key} 应为 ${min}—${max} 的整数。`);
      result[key] = value;
    }
    return result;
  }
  function schedule() {
    const delay = bank.settings.interval * 1000;
    due = Date.now() + delay;
    const update = () => status(`已保存，本轮新增 ${captured.added} 题\n${Math.max(0, (due-Date.now())/1000).toFixed(1)} 秒后刷新；可随时暂停`);
    update(); ticker = setInterval(update, 1000);
    heartbeat = setInterval(guarded(() => claimLock()), 15000);
    const checkDue = guarded(() => {
      if (!run.active) return;
      if (Date.now()<due) { timer=setTimeout(checkDue,Math.min(due-Date.now(),2147483647)); return; }
      claimLock(); writeRun(); location.reload();
    });
    timer = setTimeout(checkDue,Math.min(delay,2147483647));
  }
  function automaticRound() {
    const result = collect();
    run.rounds++; run.idle = result.added ? 0 : run.idle + 1;
    writeRun(); render();
    const reason = stopReason(run, bank.settings, bank.questions.length);
    if (reason) pause(reason); else schedule();
  }
  function download(kind) {
    const latest = loadBank();
    const category = $('category').value.trim();
    if (category !== (latest.exportCategory || '')) {
      latest.exportCategory = category; saveBank(latest); bank = latest;
    }
    const contents = kind === 'xlsx' ? null : kind === 'json' ? JSON.stringify({ ...latest, exportedAt: new Date().toISOString(),
      completeness: '未确认完整；连续无新题不能证明已采齐', images: '仅保存图片链接，未下载图片' }, null, 2)
      : kind === 'csv' ? exportCsv(latest) : exportText(latest);
    const mime = kind === 'json' ? 'application/json' : kind === 'csv' ? 'text/csv' : 'text/plain';
    const url = URL.createObjectURL(kind==='xlsx' ? exportXlsx(latest) : new Blob([contents], { type: mime + ';charset=utf-8' }));
    const link = document.createElement('a'); link.href = url;
    link.download = filenameFor(latest,kind);
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
  for (const key of Object.keys(DEFAULTS)) $(key).value = bank.settings[key];
  $('category').value = bank.exportCategory || '';
  $('mini').onclick = () => { const folded = ui.querySelector('section').classList.toggle('fold'); $('mini').textContent = folded ? '展开' : '收起'; };
  $('pause').onclick = () => pause();
  $('save').onclick = guarded(() => {
    if (run.active) throw new Error('请先暂停，再修改设置。');
    bank = loadBank(); bank.settings = readSettings(); bank.exportCategory=$('category').value.trim(); saveBank(bank); status('设置已保存');
  });
  $('start').onclick = guarded(() => {
    if (!ready || run.active) return;
    claimLock(); bank = loadBank(); bank.settings = readSettings(); bank.exportCategory=$('category').value.trim(); saveBank(bank);
    run = { active: true, rounds: 0, idle: 0 }; writeRun(); automaticRound();
  });
  $('capture').onclick = guarded(() => {
    const result = collect(); pause(`本页已保存，新增 ${result.added} 题，累计 ${result.total} 题`);
  });
  for (const kind of ['xlsx','json','csv','txt']) $(kind).onclick = guarded(() => download(kind));
  GM_registerMenuCommand('导出Excel题库', guarded(() => download('xlsx')));
  GM_registerMenuCommand('暂停题库采集', () => pause());
  GM_registerMenuCommand('导出题库 JSON', guarded(() => download('json')));
  // 用户在原答卷输入或选择答案时停止刷新，避免丢失填写进度。
  for (const event of ['input','change']) document.addEventListener(event, e => {
    if (e.isTrusted && !e.composedPath().includes(host) && run.active) pause('检测到你在填写答卷，已暂停刷新。');
  }, true);
  document.addEventListener('click', e => {
    if (e.isTrusted && run.active && e.target instanceof Element && e.target.closest('.field .ui-radio,.field .ui-checkbox')) {
      pause('检测到你在选择答案，已暂停刷新。');
    }
  }, true);
  async function waitForQuestions() {
    let previous = '', stable = 0;
    for (let i = 0; i < 60; i++) {
      pageInfo = scanPage(); batch = pageInfo.questions;
      const signature = JSON.stringify(batch);
      stable = batch.length && signature === previous ? stable + 1 : 0;
      if (stable >= 3) return true;
      previous = signature;
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    return false;
  }
  guarded(async () => {
    render();
    ready = await waitForQuestions();
    if (!ready) { pause('30秒内未检测到此版支持的稳定选择题；如有验证或登录提示，请手动完成后刷新。填空、矩阵等题型暂不支持。'); return; }
    render();
    if (run.active) automaticRound();
    else status(`已识别本页 ${batch.length} 题\n点击“开始刷新采集”，或只采集本页。`);
  })();
})();
