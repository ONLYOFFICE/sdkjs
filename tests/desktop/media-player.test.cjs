/* SPDX-License-Identifier: AGPL-3.0-only */
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const {test} = require("node:test");

function createEditor() {
	const commands = [];
	// Model the native player's boundary: only showMediaControl creates a player,
	// only selected media can create it, and playFrom is not a native command.
	const player = {file: null, playing: false, starts: 0};
	const context = {
		Asc: {},
		AscCommon: {c_oEditorId: {Presentation: 2}},
		AscFormat: {IsEqualObjects: (a, b) => JSON.stringify(a) === JSON.stringify(b)},
		AscDesktopEditor: {
			CallMediaPlayerCommand(data) {
				commands.push(JSON.parse(JSON.stringify(data)));
				if (data.Cmd === "hideMediaControl") {
					player.file = null;
					player.playing = false;
				} else if (data.Cmd === "showMediaControl" && !player.file && data.IsSelected) {
					player.file = data.MediaFile;
				} else if (data.Cmd === "play" && player.file) {
					player.playing = true;
					player.starts++;
				} else if ((data.Cmd === "pause" || data.Cmd === "stop") && player.file) {
					player.playing = false;
				} else if (data.Cmd === "resume" && player.file) {
					player.playing = true;
				}
			}
		}
	};
	context.window = context;
	vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../../common/apiBase.js"), "utf8"), context);
	const api = Object.create(context.AscCommon.baseEditorsApi.prototype);
	api.editorId = context.AscCommon.c_oEditorId.Presentation;
	api.WordControl = {GetMediaPlayerData: media => media ? {...media} : null};
	const media = {DrawingId: "audio-5", MediaFile: "media/media2.mp3", IsSelected: false, Volume: 80};
	return {api, media, commands, player, context};
}

test("PowerPoint playFrom(0.0) is translated to the native play command", () => {
	const {api, media} = createEditor();
	api.mediaData = media;
	assert.equal(api.getPlayerData("playFrom(0.0)").Cmd, "play");
});

test("a Next animation plays audio without selecting its icon first", () => {
	const {api, media, player} = createEditor();
	api.callMediaPlayerCommand("playFrom(0.0)", media);
	assert.equal(player.playing, true);
	assert.equal(player.file, media.MediaFile);
	assert.equal(media.IsSelected, false, "playback must not change the drawing selection");
});

test("activating a PowerPoint media command reaches desktop playback", () => {
	const {api, media, player, context} = createEditor();
	context.Asc.editor = api;
	context.AscDFH = {drawingsChangesMap: {}, drawingContentChanges: {}, changesFactory: {}, drawingsConstructorsMap: {}};
	context.AscFormat.CBaseFormatObject = function () {};
	context.AscFormat.CShape = function () {};
	context.AscFormat.InitClass = function (constructor, base, type) {
		constructor.prototype = Object.create(base.prototype);
		constructor.prototype.constructor = constructor;
		constructor.prototype.getObjectType = () => type;
	};
	for (const file of ["common/HistoryCommon.js", "slide/Editor/Format/Timing.js"]) {
		vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../..", file), "utf8"), context);
	}
	const command = new context.AscFormat.CCmd();
	command.type = 0; // p:cmd type="call"
	command.cmd = "playFrom(0.0)";
	command.getTargetObject = () => ({getMediaData: () => media, GetId: () => media.DrawingId});
	command.getRoot = () => null;
	command.setState(1); // TIME_NODE_STATE_ACTIVE, after Next triggers the effect
	assert.equal(player.playing, true);
	assert.equal(player.file, media.MediaFile);
});

test("repeating the same zero-offset effect restarts its audio", () => {
	const {api, media, player} = createEditor();
	api.callMediaPlayerCommand("playFrom(0.0)", media);
	api.callMediaPlayerCommand("playFrom(0.0)", media);
	assert.equal(player.starts, 2);
});

test("ordinary play preserves an initialized player", () => {
	const {api, media, commands, player} = createEditor();
	api.callMediaPlayerCommand("playFrom(0.0)", media);
	api.callMediaPlayerCommand("pause", media);
	commands.length = 0;
	api.callMediaPlayerCommand("play", media);
	assert.deepEqual(commands.map(command => command.Cmd), ["play"]);
	assert.equal(player.playing, true);
});

test("nonzero offsets are not silently restarted at zero", () => {
	const {api, media, commands} = createEditor();
	api.callMediaPlayerCommand("playFrom(2.5)", media);
	assert.deepEqual(commands.map(command => command.Cmd), ["playFrom(2.5)"]);
});

test("the next audio animation replaces the previous media source", () => {
	const {api, media, player} = createEditor();
	api.callMediaPlayerCommand("playFrom(0.0)", media);
	api.callMediaPlayerCommand("playFrom(0.0)", {...media, DrawingId: "audio-7", MediaFile: "media/media4.mp3"});
	assert.equal(player.file, "media/media4.mp3");
	assert.equal(player.playing, true);
});

test("pause, resume and slide cleanup retain their existing commands", () => {
	const {api, media, commands} = createEditor();
	for (const cmd of ["pause", "resume", "stop", "togglePause", "update"]) {
		api.callMediaPlayerCommand(cmd, media);
		assert.equal(commands.at(-1).Cmd, cmd);
	}
	api.hideMediaControl();
	assert.equal(commands.at(-1).Cmd, "hideMediaControl");
	assert.equal(api.mediaData, null);
});

test("duplicate geometry updates are still suppressed", () => {
	const {api, media, commands} = createEditor();
	api.callMediaPlayerCommand("update", media);
	api.callMediaPlayerCommand("update", media);
	assert.equal(commands.length, 1);
});

test("unavailable geometry does not reach the desktop bridge", () => {
	const {api, media, commands} = createEditor();
	api.WordControl.GetMediaPlayerData = () => null;
	api.callMediaPlayerCommand("playFrom(0.0)", media);
	assert.equal(commands.length, 0);
});

test("browser editors without a desktop bridge remain usable", () => {
	const {api, media, context} = createEditor();
	delete context.AscDesktopEditor;
	assert.doesNotThrow(() => api.callMediaPlayerCommand("playFrom(0.0)", media));
});
