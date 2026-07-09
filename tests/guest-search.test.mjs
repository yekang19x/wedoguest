import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const html = fs.readFileSync('frontend/index.html', 'utf8');
const script = html.match(/<script>([\s\S]*?)\/\* ================= 页签 ================= \*\//)[1];
const context = {
  document: { getElementById: () => ({}) },
  fetch: async () => ({ ok: true, json: async () => ({}) }),
  setTimeout,
  clearTimeout,
  Error,
};
vm.createContext(context);
vm.runInContext(script, context);

const guest = {
  name: '张伟',
  family_names: '李娜、王强',
  note: '大学同学',
  table_no: 'A12',
};

assert.equal(context.guestMatchesSearch(guest, 'zw'), true, '姓名首字母 zw 应匹配张伟');
assert.equal(context.guestMatchesSearch(guest, 'ln'), true, '家属首字母 ln 应匹配李娜');
assert.equal(context.guestMatchesSearch(guest, 'wq'), true, '家属首字母 wq 应匹配王强');
assert.equal(context.guestMatchesSearch(guest, '大学'), true, '保留原备注搜索');
assert.equal(context.guestMatchesSearch(guest, 'a12'), true, '保留原桌号搜索');
assert.equal(context.guestMatchesSearch(guest, 'zz'), false, '无关首字母不应匹配');
