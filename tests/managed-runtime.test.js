const { test } = require('node:test');
const assert = require('node:assert/strict');
const { PassThrough } = require('node:stream');
const { managedChannel } = require('../bridge/managed-runtime');

test('managed startup accepts a split private message and stops when its owner exits', async () => {
  const pipe = new PassThrough(), channel = managedChannel(pipe);
  let stops = 0; channel.on('stop', () => stops++);
  const secret = 'a'.repeat(64), message = JSON.stringify({pairingToken:secret})+'\n';
  pipe.write(message.slice(0, 20)); pipe.write(message.slice(20));
  assert.equal(await channel.ready, secret);
  pipe.end(); await new Promise(resolve => setImmediate(resolve));
  assert.equal(stops, 1); assert.equal(channel.isClosed(), true);
});

test('managed startup rejects malformed, missing and oversized credentials', async () => {
  for (const value of ['{}\n', '{"pairingToken":"short"}\n', 'not-json\n', 'x'.repeat(8193)]) {
    const pipe = new PassThrough(), channel = managedChannel(pipe);
    const rejected = assert.rejects(channel.ready, /interrompu/);
    pipe.write(value); await rejected; pipe.destroy();
  }
});

test('managed stop commands are idempotent', async () => {
  const pipe = new PassThrough(), channel = managedChannel(pipe);
  let stops = 0; channel.on('stop', () => stops++);
  pipe.write(JSON.stringify({pairingToken:'b'.repeat(64)})+'\n'); await channel.ready;
  pipe.write('{"command":"stop"}\n{"command":"stop"}\n'); pipe.end();
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(stops, 1);
});

test('the actual managed server restarts with the same supplied code and exits with its parent pipe', {timeout:20000}, async t => {
  const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),net=require('node:net');
  const {spawn}=require('node:child_process');
  const directory=fs.mkdtempSync(path.join(os.tmpdir(),'pictos-managed-'));
  t.after(()=>fs.rmSync(directory,{recursive:true,force:true}));
  const probe=net.createServer();await new Promise(resolve=>probe.listen(0,'127.0.0.1',resolve));
  const port=probe.address().port;await new Promise(resolve=>probe.close(resolve));
  const token='c'.repeat(64);
  for(let iteration=0;iteration<2;iteration++){
    const child=spawn(process.execPath,[path.join(__dirname,'../bridge/server.js'),'--managed'],{
      env:{...process.env,PICTOS_DATA_DIR:directory,PICTOS_PORT:String(port),PICTOS_CODEX_PATH:path.join(directory,'no-codex-inference')},
      stdio:['pipe','pipe','pipe'],windowsHide:true
    });
    t.after(()=>{if(child.exitCode===null)child.kill();});
    const exited=new Promise(resolve=>child.once('exit',code=>resolve(code)));
    let output='',errors='';child.stderr.on('data',data=>{errors+=data;});
    const ready=new Promise((resolve,reject)=>{
      child.once('error',reject);
      child.stdout.on('data',data=>{
        output+=data;
        for(const line of output.split('\n')){try{const value=JSON.parse(line);if(value.event==='ready')resolve(value);}catch{}}
      });
      child.once('exit',()=>reject(new Error('Managed server exited before readiness.')));
    });
    child.stdin.write(JSON.stringify({pairingToken:token})+'\n');
    const state=await ready;assert.equal(state.url,'http://127.0.0.1:'+port);
    const response=await fetch(state.url+'/v1/dashboard-ticket',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:'{}'});
    assert.equal(response.status,200);assert.equal(output.includes(token),false);assert.equal(errors.includes(token),false);
    child.stdin.end();assert.equal(await exited,0);
  }
});
