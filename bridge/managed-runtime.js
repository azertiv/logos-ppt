"use strict";
const { EventEmitter } = require("node:events");

// The Windows launcher sends the protected pairing secret through an inherited
// pipe. It never appears in process arguments, the environment or log files.
function managedChannel(input) {
  const channel = new EventEmitter();
  let buffer = "", initialized = false, ended = false;
  let accept, reject;
  channel.ready = new Promise((resolve, fail) => { accept = resolve; reject = fail; });
  const fail = () => {
    if (ended) return;
    ended = true;
    clearTimeout(timer);
    reject(new Error("Le lanceur du compagnon a été interrompu."));
    channel.emit("stop");
  };
  const timer = setTimeout(fail, 10000);
  input.setEncoding("utf8");
  input.on("data", data => {
    if (ended) return;
    buffer += data;
    if (buffer.length > 8192) return fail();
    let boundary;
    while ((boundary = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, boundary); buffer = buffer.slice(boundary + 1);
      let message;
      try { message = JSON.parse(line); } catch { return fail(); }
      if (!initialized) {
        if (!/^[a-f0-9]{64}$/.test(message?.pairingToken || "")) return fail();
        initialized = true; clearTimeout(timer); accept(message.pairingToken);
      } else if (message?.command === "stop") fail();
    }
  });
  input.once("end", fail); input.once("error", fail);
  channel.isClosed = () => ended;
  return channel;
}
module.exports = { managedChannel };
