import assert from 'node:assert/strict';
import fs from 'node:fs';

const html = fs.readFileSync('frontend/index.html', 'utf8');

function assertOrder(source, snippets, message) {
  let cursor = -1;
  for (const snippet of snippets) {
    const index = source.indexOf(snippet, cursor + 1);
    assert.notEqual(index, -1, `${message}：缺少 ${snippet}`);
    assert.ok(index > cursor, `${message}：${snippet} 顺序不正确`);
    cursor = index;
  }
}

assert.match(
  html,
  /<th class="sortable" data-sort="note" onclick="sortGuests\('note'\)">备注<\/th>/,
  '宾客名单备注列应支持点击排序',
);

assert.match(
  html,
  /<th class="sortable" data-sort="name" onclick="sortGuests\('name'\)">姓名<\/th>\s*<th class="sortable" data-sort="party_size" onclick="sortGuests\('party_size'\)">预计人数<\/th>\s*<th class="sortable" data-sort="note" onclick="sortGuests\('note'\)">备注<\/th>\s*<th>家属姓名<\/th>\s*<th>确认人数<\/th>/,
  '宾客名单业务列顺序应为姓名、预计人数、备注、家属姓名、确认人数',
);

assertOrder(
  html,
  [
    '<td><b>${esc(g.name)}</b></td>',
    '<td>${g.party_size}</td>',
    '<td>${esc(g.note)}</td>',
    '<td>${esc(g.family_names)',
    "<td>${g.confirm_status === '已确认'",
  ],
  '宾客名单行数据顺序应与表头一致',
);
