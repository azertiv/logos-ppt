const { test } = require('node:test');
const assert = require('node:assert/strict');
const { layout } = require('../public/grid-window');

test('1740-item grid mounts only visible rows plus overscan at every density', () => {
  for (let columns = 1; columns <= 6; columns++) {
    const initial = layout({count:1740,columns,width:332,top:0,viewport:800});
    const visited = new Set();
    for (let top = 0; top < initial.height; top += 400) {
      const window = layout({count:1740,columns,width:332,top,viewport:800});
      assert.ok(window.end - window.start < 150);
      assert.ok(window.start >= 0 && window.end <= 1740);
      for (let i = window.start; i < window.end; i++) visited.add(i);
    }
    assert.equal(visited.size, 1740);
  }
});
test('window clamps empty lists, far-beyond-bottom scroll and width changes', () => {
  assert.equal(layout({count:0,columns:6,width:332,top:0,viewport:800}).height, 0);
  const last = layout({count:17,columns:6,width:332,top:99999,viewport:800});
  assert.equal(last.end, 17);
  assert.equal(last.start, 12);
  assert.ok(layout({count:1740,columns:3,width:700,top:0,viewport:800}).size > 200);
});
