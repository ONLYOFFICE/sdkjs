// SPDX-License-Identifier: AGPL-3.0-only
// Linux end-to-end test: capture only a dedicated null sink, never a microphone.
// Requires Node.js 22+, xvfb-run, pactl and parec (pulseaudio-utils).
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {spawn, spawnSync} from 'node:child_process';

const [binaryArg] = process.argv.slice(2);
if (!binaryArg) throw new Error('Usage: node audio-playback.mjs /path/to/DesktopEditors');
const binary = path.resolve(binaryArg);
const presentation = fileURLToPath(new URL('./fixtures/next-click-audio.pptx', import.meta.url));
const slides = [0, 1, 0];
const pactl = process.env.PACTL || 'pactl';
const parec = process.env.PAREC || 'parec';
const sink = `onlyoffice_audio_test_${process.pid}`;
function pulse(...args) {
    const result = spawnSync(pactl, args, {encoding:'utf8'});
    assert.equal(result.status, 0, result.stderr || String(result.error));
    return result.stdout.trim();
}
// Find the strongest quarter-second projection onto a known fixture frequency.
function signal(buffers, frequency) {
    const pcm = Buffer.concat(buffers);
    const count = Math.floor(pcm.length / 2);
    let peak = 0;
    for (let start = 0; start + 12000 <= count; start += 12000) {
        let real = 0, imaginary = 0;
        for (let i = 0; i < 12000; i++) {
            const sample = pcm.readInt16LE((start + i) * 2) / 32768;
            const angle = 2 * Math.PI * frequency * i / 48000;
            real += sample * Math.cos(angle);
            imaginary += sample * Math.sin(angle);
        }
        peak = Math.max(peak, 2 * Math.hypot(real, imaginary) / 12000);
    }
    return peak;
}
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
const endpoint = 'http://127.0.0.1:8080/json/list';
let portBusy = false;
try {
    await fetch(endpoint, {signal: AbortSignal.timeout(500)});
    portBusy = true;
} catch {}
assert(!portBusy, 'Port 8080 is already occupied; this test must use its own editor');

const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-audio-test-'));
const log = fs.openSync(path.join(profile, 'editor.log'), 'w');
const appDir = path.dirname(binary);
const moduleId = pulse('load-module', 'module-null-sink', `sink_name=${sink}`, 'rate=48000');
const recorder = spawn(parec, ['--raw', '--format=s16le', '--rate=48000', '--channels=1', `--device=${sink}.monitor`]);
const recorderClosed = new Promise(resolve => {
    recorder.once('close', resolve);
    recorder.once('error', resolve);
});
let capture = [];
let captureError = '';
recorder.stdout.on('data', buffer => capture.push(buffer));
recorder.stderr.on('data', buffer => { captureError += buffer; });
recorder.on('error', error => { captureError += error.message; });
const child = spawn('xvfb-run', ['-a', '-s', '-screen 0 1366x768x24', binary,
    '--ascdesktop-support-debug-info', presentation], {
    cwd: appDir,
    detached: true,
    stdio: ['ignore', log, log],
    env: {...process.env, PULSE_SINK: sink, XDG_DATA_HOME: path.join(profile, 'data'),
        XDG_CONFIG_HOME: path.join(profile, 'config'),
        LD_LIBRARY_PATH: [appDir, path.join(appDir, 'converter'), process.env.LD_LIBRARY_PATH].filter(Boolean).join(':')}
});
let exited = false;
child.on('exit', () => { exited = true; });
child.on('error', () => { exited = true; });
let socket;
try {
    let page;
    for (let attempt = 0; attempt < 90 && !exited; attempt++) {
        try {
            const pages = await (await fetch(endpoint, {signal: AbortSignal.timeout(1000)})).json();
            page = pages.find(p => p.url.includes('doctype=slide'));
            if (page) break;
        } catch {}
        await pause(500);
    }
    assert(page, `Presentation did not start; log: ${profile}/editor.log`);
    socket = new WebSocket(page.webSocketDebuggerUrl);
    const pending = new Map();
    let sequence = 0;
    socket.addEventListener('message', ({data}) => {
        const message = JSON.parse(data);
        if (pending.has(message.id)) {
            const {resolve, timer} = pending.get(message.id);
            clearTimeout(timer);
            pending.delete(message.id);
            resolve(message);
        }
    });
    await new Promise((resolve, reject) => {
        socket.addEventListener('open', resolve, {once:true});
        socket.addEventListener('error', reject, {once:true});
    });
    async function evaluate(expression) {
        const message = await new Promise((resolve, reject) => {
            const id = ++sequence;
            const timer = setTimeout(() => reject(new Error('Editor stopped responding')), 5000);
            pending.set(id, {resolve, timer});
            socket.send(JSON.stringify({id, method:'Runtime.evaluate', params:{expression, returnByValue:true}}));
        });
        assert(!message.error && !message.result.exceptionDetails, JSON.stringify(message));
        return message.result.result.value;
    }
    let ready = false;
    for (let attempt = 0; attempt < 90 && !exited; attempt++) {
        ready = await evaluate(`(()=>{try{const w=document.querySelector('iframe').contentWindow,a=w.Asc.editor;return a.isDocumentLoadComplete && w.PE.getController('Main')._isDocReady && a.WordControl.m_oLogicDocument.Slides.length > ${Math.max(...slides)}}catch(e){return false}})()`);
        if (ready) break;
        await pause(500);
    }
    assert(ready, 'Document did not finish loading');
    await pause(1500);
    await evaluate(`(()=>{const w=document.querySelector('iframe').contentWindow,a=w.Asc.editor;w.audioTestRequests=[];const original=a.callMediaPlayerCommand;a.callMediaPlayerCommand=function(cmd,data){w.audioTestRequests.push({cmd,audio:!!data && !data.isVideo()});return original.call(this,cmd,data)};return true})()`);
    const visited = new Set();
    for (const slide of slides) {
        // GoToSlide resumes a visited slide's animation state. A new slideshow
        // session is needed to replay its click effect from the beginning.
        if (visited.has(slide)) {
            await evaluate(`document.querySelector('iframe').contentWindow.Asc.editor.EndDemonstration()`);
            await pause(700);
        }
        await evaluate(`(()=>{const w=document.querySelector('iframe').contentWindow,a=w.Asc.editor;w.audioTestRequests=[];if(a.WordControl.DemonstrationManager.Mode)a.DemonstrationGoToSlide(${slide});else a.WordControl.GoToPage(${slide});return true})()`);
        await pause(700);
        await evaluate(`(()=>{const w=document.querySelector('iframe').contentWindow;if(!w.Asc.editor.WordControl.DemonstrationManager.Mode)w.PE.getController('Viewport').onPreviewStart(${slide},false,false,true,false);return true})()`);
        await pause(700);
        capture = [];
        await pause(1000);
        assert(signal(capture, 440) < 0.001 && signal(capture, 880) < 0.001,
            'Audio started before Next');
        capture = [];
        await evaluate(`document.querySelector('iframe').contentWindow.Asc.editor.DemonstrationNextSlide()`);
        await pause(4000);
        assert(!captureError, captureError);
        const expected = slide === 0 ? 440 : 880;
        const amplitude = signal(capture, expected);
        const other = signal(capture, expected === 440 ? 880 : 440);
        assert(amplitude > 0.02 && amplitude > other * 10,
            `Expected ${expected} Hz after Next, amplitude=${amplitude}, other=${other}`);
        const state = await evaluate(`(()=>{const w=document.querySelector('iframe').contentWindow,a=w.Asc.editor;return {mode:a.WordControl.DemonstrationManager.Mode,slide:a.WordControl.DemonstrationManager.SlideNum,requests:w.audioTestRequests}})()`);
        assert(!exited && state.mode && state.slide === slide, `Unexpected slideshow state: ${JSON.stringify(state)}`);
        assert(state.requests.some(r => r.audio && r.cmd === 'playFrom(0.0)'), 'Audio effect was not activated');
        visited.add(slide);
        console.log(`PASS: slide ${slide + 1}: silence before Next, ${expected} Hz after Next (amplitude ${amplitude.toFixed(3)})`);
    }
    await evaluate(`document.querySelector('iframe').contentWindow.Asc.editor.EndDemonstration()`);
    await pause(700);
    assert(await evaluate(`!document.querySelector('iframe').contentWindow.Asc.editor.WordControl.DemonstrationManager.Mode`), 'Slideshow did not close');
    console.log(`PASS: slideshow cleanup; diagnostic log: ${profile}/editor.log`);
} finally {
    socket?.close();
    try { process.kill(-child.pid, 'SIGTERM'); } catch (e) { if (e.code !== 'ESRCH') throw e; }
    recorder.kill();
    await recorderClosed;
    pulse('unload-module', moduleId);
    fs.closeSync(log);
}
