import assert from 'node:assert/strict';
import fs from 'node:fs';

const html = fs.readFileSync('frontend/index.html', 'utf8');

assert.match(
  html,
  /\.guest-table-wrap\s*\{[\s\S]*overflow:\s*auto;/,
  '宾客名单需要独立滚动容器',
);
assert.match(
  html,
  /<div class="guest-table-wrap">\s*<table class="list" id="guests-table">/,
  '宾客表格应放在滚动容器内',
);
assert.match(
  html,
  /\.guest-table-wrap thead th\s*\{[\s\S]*position:\s*sticky;[\s\S]*top:\s*0;/,
  '宾客表格整行表头应保持 sticky',
);
assert.match(
  html,
  /\.guest-table-wrap table\.list\s*\{[\s\S]*overflow:\s*visible;/,
  '宾客表格不应裁剪 sticky 表头',
);
