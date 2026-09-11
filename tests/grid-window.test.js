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

test('fractional card widths never exceed the available width from one to six columns', () => {
  for (const width of [176, 203.5, 216, 284, 332, 483.25]) {
    for (let columns = 1; columns <= 6; columns++) {
      const result = layout({count:1740,columns,width,top:7600,viewport:586});
      assert.ok(result.size * columns + (columns - 1) * 14 <= width);
    }
  }
});

test('one viewport of overscan covers rapid scrolling while keeping mounted cards bounded', () => {
  for (const width of [204, 284, 484]) {
    for (let columns = 1; columns <= 6; columns++) {
      for (const top of [0, 7600, 30000]) {
        const result = layout({count:1740,columns,width,top,viewport:586,overscanPixels:586});
        assert.ok(result.end - result.start < 330);
        if (top < result.height) {
          assert.ok(result.start / columns * result.stride <= Math.max(0, top - 586));
          assert.ok(Math.ceil(result.end / columns) * result.stride >= Math.min(result.height, top + 1172));
        }
      }
    }
  }
});
