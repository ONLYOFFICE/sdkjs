/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function (window, undefined) {

	const WEBP_MIME = "image/webp";
	const KEYFRAME_INTERVAL = 100;
	const MAX_ENCODE_QUEUE = 8;
	const DEFAULT_VIDEO_CODEC = "vp09.00.10.08";
	const WEBP_QUALITY = 0.85;
	const PROBE_WIDTH = 1920;
	const PROBE_HEIGHT = 1080;
	const BITS_PER_PIXEL = 0.07;
	const MIN_VIDEO_BITRATE = 1000000;
	const MAX_VIDEO_BITRATE = 40000000;

	function getVideoBitrate(width, height, fps) {
		const bitrate = Math.round(width * height * fps * BITS_PER_PIXEL);
		if (bitrate < MIN_VIDEO_BITRATE) {
			return MIN_VIDEO_BITRATE;
		}
		return bitrate > MAX_VIDEO_BITRATE ? MAX_VIDEO_BITRATE : bitrate;
	}
	function getFrameRate(settings) {
		return AscCommonSlide.getAnimatedExportFps(settings);
	}
	const DEFAULT_MP4_CODEC = "avc1.42001f";

	function CAnimExportGifWriter(width, height, settings) {
		this.encoder = new AscCommonSlide.CGIFEncoder(width, height, 0,
			AscCommonSlide.isAnimatedExportTransparent(settings));
		this.planner = new AscCommonSlide.CAnimExportDelayPlanner();
	}
	CAnimExportGifWriter.prototype.addFrame = function (frameSource, durationMs) {
		return this.encoder.writeFrame(frameSource.getFrameData(), this.planner.next(durationMs));
	};
	CAnimExportGifWriter.prototype.addDuration = function (durationMs) {
		this.encoder.addDelay(this.planner.next(durationMs));
	};
	CAnimExportGifWriter.prototype.isBusy = function () {
		return false;
	};
	CAnimExportGifWriter.prototype.getErrorText = function () {
		return null;
	};
	CAnimExportGifWriter.prototype.finish = function (callback) {
		let data = null;
		let errorText = null;
		try {
			data = this.encoder.finish();
		}
		catch (error) {
			errorText = String(error.message || error);
		}
		callback(data, errorText);
	};
	CAnimExportGifWriter.prototype.abort = function () {
	};
	CAnimExportGifWriter.prototype.getFrameCount = function () {
		return this.encoder.getFrameCount();
	};
	CAnimExportGifWriter.prototype.getDurationMs = function () {
		return this.planner.getWrittenMs();
	};
	CAnimExportGifWriter.prototype.getExtension = function () {
		return "gif";
	};
	CAnimExportGifWriter.prototype.getMimeType = function () {
		return "image/gif";
	};

	function CAnimExportWebpWriter(width, height, settings) {
		this.width = width;
		this.height = height;
		this.quality = WEBP_QUALITY;
		this.encoder = new AscCommonSlide.CWebPAnimEncoder(width, height, 0);
		this.previous = null;
		this.regionCanvas = null;
		this.errorText = null;
	}
	CAnimExportWebpWriter.isSupported = function () {
		const canvas = document.createElement('canvas');
		canvas.width = 2;
		canvas.height = 2;
		return 0 === canvas.toDataURL(WEBP_MIME).indexOf("data:" + WEBP_MIME);
	};
	CAnimExportWebpWriter.prototype.addFrame = function (frameSource, durationMs) {
		const data = frameSource.getFrameData();
		let rect = {x: 0, y: 0, w: this.width, h: this.height};
		if (this.previous) {
			rect = this.getChangedRect(data);
			if (!rect) {
				this.addDuration(durationMs);
				return false;
			}
		}
		const bytes = this.encodeRegion(frameSource.getCanvas(), rect);
		if (!bytes) {
			this.errorText = "webp encoding is not available";
			return false;
		}
		if (!this.encoder.addFrame(bytes, rect, durationMs, true)) {
			this.errorText = "webp frame has no bitstream";
			return false;
		}
		if (!this.previous) {
			this.previous = new Uint8Array(data.length);
		}
		this.previous.set(data);
		return true;
	};
	CAnimExportWebpWriter.prototype.getChangedRect = function (data) {
		const previous = this.previous;
		let left = this.width;
		let top = this.height;
		let right = -1;
		let bottom = -1;
		for (let y = 0; y < this.height; ++y) {
			const rowStart = y * this.width;
			for (let x = 0; x < this.width; ++x) {
				const offset = (rowStart + x) * 4;
				if (data[offset] !== previous[offset] ||
					data[offset + 1] !== previous[offset + 1] ||
					data[offset + 2] !== previous[offset + 2] ||
					data[offset + 3] !== previous[offset + 3]) {
					if (x < left) { left = x; }
					if (x > right) { right = x; }
					if (y < top) { top = y; }
					bottom = y;
				}
			}
		}
		if (-1 === right) {
			return null;
		}
		left &= ~1;
		top &= ~1;
		return {x: left, y: top, w: right - left + 1, h: bottom - top + 1};
	};
	CAnimExportWebpWriter.prototype.encodeRegion = function (canvas, rect) {
		let source = canvas;
		if (rect.w !== this.width || rect.h !== this.height) {
			if (!this.regionCanvas) {
				this.regionCanvas = document.createElement('canvas');
			}
			if (this.regionCanvas.width !== rect.w || this.regionCanvas.height !== rect.h) {
				this.regionCanvas.width = rect.w;
				this.regionCanvas.height = rect.h;
			}
			const context = this.regionCanvas.getContext('2d');
			context.setTransform(1, 0, 0, 1, 0, 0);
			context.clearRect(0, 0, rect.w, rect.h);
			context.drawImage(canvas, rect.x, rect.y, rect.w, rect.h, 0, 0, rect.w, rect.h);
			source = this.regionCanvas;
		}
		const dataUrl = source.toDataURL(WEBP_MIME, this.quality);
		if (0 !== dataUrl.indexOf("data:" + WEBP_MIME)) {
			return null;
		}
		return AscCommon.Base64.decode(dataUrl.substring(dataUrl.indexOf(",") + 1));
	};
	CAnimExportWebpWriter.prototype.addDuration = function (durationMs) {
		this.encoder.addDuration(durationMs);
	};
	CAnimExportWebpWriter.prototype.isBusy = function () {
		return false;
	};
	CAnimExportWebpWriter.prototype.getErrorText = function () {
		return this.errorText;
	};
	CAnimExportWebpWriter.prototype.finish = function (callback) {
		if (this.errorText) {
			callback(null, this.errorText);
			return;
		}
		let data = null;
		let errorText = null;
		try {
			data = this.encoder.finish();
		}
		catch (error) {
			errorText = String(error.message || error);
		}
		callback(data, errorText);
	};
	CAnimExportWebpWriter.prototype.abort = function () {
	};
	CAnimExportWebpWriter.prototype.getFrameCount = function () {
		return this.encoder.getFrameCount();
	};
	CAnimExportWebpWriter.prototype.getDurationMs = function () {
		return this.encoder.getDurationMs();
	};
	CAnimExportWebpWriter.prototype.getExtension = function () {
		return "webp";
	};
	CAnimExportWebpWriter.prototype.getMimeType = function () {
		return "image/webp";
	};

	function getWebmCodecId(codec) {
		if (0 === codec.indexOf("vp09")) {
			return "V_VP9";
		}
		if (0 === codec.indexOf("av01")) {
			return "V_AV1";
		}
		return "V_VP8";
	}

	function CAnimExportWebmWriter(width, height, settings) {
		this.codec = DEFAULT_VIDEO_CODEC;
		this.muxer = new AscCommonSlide.CWebMMuxer(width, height, getWebmCodecId(this.codec));
		this.timeUs = 0;
		this.frameIndex = 0;
		this.errorText = null;
		this.isClosed = false;
		this.lastCanvas = null;
		this.lastFrameTimeUs = 0;

		const self = this;
		this.encoder = new VideoEncoder({
			output: function (chunk) {
				const data = new Uint8Array(chunk.byteLength);
				chunk.copyTo(data);
				self.muxer.addFrame(data, Math.round(chunk.timestamp / 1000), "key" === chunk.type);
			},
			error: function (error) {
				self.errorText = String(error.message || error);
			}
		});
		try {
			this.encoder.configure({
				codec: this.codec,
				width: width,
				height: height,
				bitrate: getVideoBitrate(width, height, getFrameRate(settings)),
				framerate: getFrameRate(settings)
			});
		}
		catch (error) {
			this.errorText = "video codec " + this.codec + " was rejected: " + (error.message || error.name);
		}
	}
	function probeCodecs(candidates, width, height, callback) {
		if (typeof VideoEncoder === "undefined" || !window.isSecureContext) {
			callback([]);
			return;
		}
		const supported = [];
		let pending = candidates.length;
		const onAnswer = function () {
			--pending;
			if (0 === pending) {
				callback(supported);
			}
		};
		candidates.forEach(function (codec) {
			let promise = null;
			try {
				promise = VideoEncoder.isConfigSupported({
					codec: codec,
					width: width,
					height: height,
					bitrate: 4000000,
					framerate: 25
				});
			}
			catch (error) {
				onAnswer();
				return;
			}
			promise.then(function (answer) {
				if (answer && answer.supported) {
					supported.push(codec);
				}
				onAnswer();
			}).catch(function () {
				onAnswer();
			});
		});
	}
	CAnimExportWebmWriter.isSupported = function () {
		return typeof VideoEncoder !== "undefined" && typeof VideoFrame !== "undefined" && !!window.isSecureContext;
	};
	CAnimExportWebmWriter.prototype.addFrame = function (frameSource, durationMs) {
		if (this.errorText || this.isClosed) {
			return false;
		}
		const frame = new VideoFrame(frameSource.getCanvas(), {
			timestamp: this.timeUs,
			duration: Math.round(durationMs * 1000)
		});
		this.encoder.encode(frame, {keyFrame: 0 === this.frameIndex % KEYFRAME_INTERVAL});
		frame.close();
		++this.frameIndex;
		this.lastCanvas = frameSource.getCanvas();
		this.lastFrameTimeUs = this.timeUs;
		this.timeUs += Math.round(durationMs * 1000);
		return true;
	};
	CAnimExportWebmWriter.prototype.addDuration = function (durationMs) {
		this.timeUs += Math.round(durationMs * 1000);
	};
	CAnimExportWebmWriter.prototype.isBusy = function () {
		if (this.errorText || this.isClosed) {
			return false;
		}
		return this.encoder.encodeQueueSize > MAX_ENCODE_QUEUE;
	};
	CAnimExportWebmWriter.prototype.getErrorText = function () {
		return this.errorText;
	};
	CAnimExportWebmWriter.prototype.finish = function (callback) {
		if (this.errorText) {
			this.abort();
			callback(null, this.errorText);
			return;
		}
		if (0 === this.muxer.getFrameCount() && 0 === this.frameIndex) {
			this.abort();
			callback(null, "no frames were encoded");
			return;
		}
		this.encodeTailFrame();
		const self = this;
		this.encoder.flush().then(function () {
			self.abort();
			self.muxer.setDurationMs(Math.round(self.timeUs / 1000));
			callback(self.errorText ? null : self.muxer.finish(), self.errorText);
		}).catch(function (error) {
			self.abort();
			callback(null, String(error.message || error));
		});
	};
	CAnimExportWebmWriter.prototype.encodeTailFrame = function () {
		if (!this.lastCanvas || this.timeUs <= this.lastFrameTimeUs) {
			return;
		}
		const frame = new VideoFrame(this.lastCanvas, {timestamp: this.timeUs, duration: 0});
		this.encoder.encode(frame, {keyFrame: false});
		frame.close();
	};
	CAnimExportWebmWriter.prototype.abort = function () {
		if (this.isClosed) {
			return;
		}
		this.isClosed = true;
		if ("closed" === this.encoder.state) {
			return;
		}
		try {
			this.encoder.close();
		}
		catch (error) {
			this.errorText = this.errorText || String(error.message || error);
		}
	};
	CAnimExportWebmWriter.prototype.getFrameCount = function () {
		return this.muxer.getFrameCount();
	};
	CAnimExportWebmWriter.prototype.getDurationMs = function () {
		return Math.round(this.timeUs / 1000);
	};
	CAnimExportWebmWriter.prototype.getExtension = function () {
		return "webm";
	};
	CAnimExportWebmWriter.prototype.getMimeType = function () {
		return "video/webm";
	};

	const AVC_LEVELS = [
		{code: 0x1E, macroblocks: 1620},
		{code: 0x1F, macroblocks: 3600},
		{code: 0x20, macroblocks: 5120},
		{code: 0x28, macroblocks: 8192},
		{code: 0x2A, macroblocks: 8704},
		{code: 0x32, macroblocks: 22080},
		{code: 0x33, macroblocks: 36864},
		{code: 0x3C, macroblocks: 139264}
	];

	function getAvcLevelCode(width, height) {
		const macroblocks = Math.ceil(width / 16) * Math.ceil(height / 16);
		for (let index = 0; index < AVC_LEVELS.length; ++index) {
			if (macroblocks <= AVC_LEVELS[index].macroblocks) {
				return AVC_LEVELS[index].code;
			}
		}
		return AVC_LEVELS[AVC_LEVELS.length - 1].code;
	}

	function raiseAvcLevel(codec, width, height) {
		if (11 !== codec.length) {
			return codec;
		}
		const current = parseInt(codec.substring(9), 16);
		const required = getAvcLevelCode(width, height);
		if (!(current < required)) {
			return codec;
		}
		const level = required.toString(16);
		return codec.substring(0, 9) + (level.length < 2 ? "0" + level : level);
	}

	function getMp4Codec(codec, width, height) {
		let result = DEFAULT_MP4_CODEC;
		if (codec && (0 === codec.indexOf("avc1") || 0 === codec.indexOf("avc3") ||
			0 === codec.indexOf("hev1") || 0 === codec.indexOf("hvc1"))) {
			result = codec;
		}
		if (0 === result.indexOf("avc1") || 0 === result.indexOf("avc3")) {
			return raiseAvcLevel(result, width, height);
		}
		return result;
	}

	function CAnimExportMp4Writer(width, height, settings) {
		this.codec = getMp4Codec(DEFAULT_MP4_CODEC, width, height);
		this.muxer = new AscCommonSlide.CMp4Muxer(width, height);
		this.timeUs = 0;
		this.frameIndex = 0;
		this.errorText = null;
		this.isClosed = false;
		this.pendingDurations = [];
		this.lastTimestamp = -1;

		const self = this;
		this.encoder = new VideoEncoder({
			output: function (chunk, metadata) {
				if (metadata && metadata.decoderConfig) {
					self.muxer.setDescription(metadata.decoderConfig.description);
				}
				if (chunk.timestamp < self.lastTimestamp) {
					self.errorText = "the encoder reordered frames, which needs composition offsets";
					return;
				}
				self.lastTimestamp = chunk.timestamp;
				const data = new Uint8Array(chunk.byteLength);
				chunk.copyTo(data);
				const durationMs = self.pendingDurations.length ? self.pendingDurations.shift() : 0;
				self.muxer.addSample(data, durationMs, "key" === chunk.type);
			},
			error: function (error) {
				self.errorText = String(error.message || error);
			}
		});
		try {
			this.encoder.configure({
				codec: this.codec,
				width: width,
				height: height,
				bitrate: getVideoBitrate(width, height, getFrameRate(settings)),
				framerate: getFrameRate(settings),
				avc: {format: "avc"}
			});
		}
		catch (error) {
			this.errorText = "video codec " + this.codec + " was rejected: " + (error.message || error.name);
		}
	}
	CAnimExportMp4Writer.isSupported = function () {
		return typeof VideoEncoder !== "undefined" && typeof VideoFrame !== "undefined" && !!window.isSecureContext;
	};
	CAnimExportMp4Writer.prototype.addFrame = function (frameSource, durationMs) {
		if (this.errorText || this.isClosed) {
			return false;
		}
		const frame = new VideoFrame(frameSource.getCanvas(), {
			timestamp: this.timeUs,
			duration: Math.round(durationMs * 1000)
		});
		this.pendingDurations.push(durationMs);
		this.encoder.encode(frame, {keyFrame: 0 === this.frameIndex % KEYFRAME_INTERVAL});
		frame.close();
		++this.frameIndex;
		this.timeUs += Math.round(durationMs * 1000);
		return true;
	};
	CAnimExportMp4Writer.prototype.addDuration = function (durationMs) {
		this.timeUs += Math.round(durationMs * 1000);
		if (this.pendingDurations.length > 0) {
			this.pendingDurations[this.pendingDurations.length - 1] += durationMs;
			return;
		}
		this.muxer.addDurationToLastSample(durationMs);
	};
	CAnimExportMp4Writer.prototype.isBusy = function () {
		if (this.errorText || this.isClosed) {
			return false;
		}
		return this.encoder.encodeQueueSize > MAX_ENCODE_QUEUE;
	};
	CAnimExportMp4Writer.prototype.getErrorText = function () {
		return this.errorText;
	};
	CAnimExportMp4Writer.prototype.finish = function (callback) {
		if (this.errorText) {
			this.abort();
			callback(null, this.errorText);
			return;
		}
		const self = this;
		this.encoder.flush().then(function () {
			self.abort();
			if (self.errorText) {
				callback(null, self.errorText);
				return;
			}
			if (!self.muxer.hasDescription()) {
				callback(null, "the encoder did not provide a codec description");
				return;
			}
			const data = self.muxer.finish();
			callback(data, data ? null : "no frames were encoded");
		}).catch(function (error) {
			self.abort();
			callback(null, String(error.message || error));
		});
	};
	CAnimExportMp4Writer.prototype.abort = function () {
		if (this.isClosed) {
			return;
		}
		this.isClosed = true;
		if ("closed" === this.encoder.state) {
			return;
		}
		try {
			this.encoder.close();
		}
		catch (error) {
			this.errorText = this.errorText || String(error.message || error);
		}
	};
	CAnimExportMp4Writer.prototype.getFrameCount = function () {
		return this.muxer.getSampleCount();
	};
	CAnimExportMp4Writer.prototype.getDurationMs = function () {
		return Math.round(this.timeUs / 1000);
	};
	CAnimExportMp4Writer.prototype.getExtension = function () {
		return "mp4";
	};
	CAnimExportMp4Writer.prototype.getMimeType = function () {
		return "video/mp4";
	};

	function getSupportedAnimExportFormats(callback) {
		const formats = [Asc.c_oAscAnimatedExportFormat.Gif];
		if (CAnimExportWebpWriter.isSupported()) {
			formats.push(Asc.c_oAscAnimatedExportFormat.WebP);
		}
		const candidates = [];
		if (CAnimExportWebmWriter.isSupported()) {
			candidates.push({format: Asc.c_oAscAnimatedExportFormat.WebM, codec: DEFAULT_VIDEO_CODEC});
		}
		if (CAnimExportMp4Writer.isSupported()) {
			candidates.push({format: Asc.c_oAscAnimatedExportFormat.Mp4, codec: DEFAULT_MP4_CODEC});
		}
		if (0 === candidates.length) {
			callback(formats);
			return;
		}

		let pending = candidates.length;
		const onAnswer = function () {
			if (0 !== --pending) {
				return;
			}
			formats.sort(function (a, b) { return a - b; });
			callback(formats);
		};
		for (let index = 0; index < candidates.length; ++index) {
			const candidate = candidates[index];
			probeCodecs([candidate.codec], PROBE_WIDTH, PROBE_HEIGHT, function (supported) {
				if (supported.length) {
					formats.push(candidate.format);
				}
				onAnswer();
			});
		}
	}

	function isAnimExportFormatSupported(format) {
		if (Asc.c_oAscAnimatedExportFormat.Gif === format) {
			return true;
		}
		if (Asc.c_oAscAnimatedExportFormat.WebP === format) {
			return CAnimExportWebpWriter.isSupported();
		}
		if (Asc.c_oAscAnimatedExportFormat.WebM === format) {
			return CAnimExportWebmWriter.isSupported();
		}
		if (Asc.c_oAscAnimatedExportFormat.Mp4 === format) {
			return CAnimExportMp4Writer.isSupported();
		}
		return false;
	}

	function createAnimExportWriter(width, height, settings) {
		if (Asc.c_oAscAnimatedExportFormat.WebP === settings.format) {
			return new CAnimExportWebpWriter(width, height, settings);
		}
		if (Asc.c_oAscAnimatedExportFormat.WebM === settings.format) {
			return new CAnimExportWebmWriter(width, height, settings);
		}
		if (Asc.c_oAscAnimatedExportFormat.Mp4 === settings.format) {
			return new CAnimExportMp4Writer(width, height, settings);
		}
		return new CAnimExportGifWriter(width, height, settings);
	}

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].CAnimExportGifWriter = CAnimExportGifWriter;
	window["AscCommonSlide"].CAnimExportWebpWriter = CAnimExportWebpWriter;
	window["AscCommonSlide"].CAnimExportWebmWriter = CAnimExportWebmWriter;
	window["AscCommonSlide"].CAnimExportMp4Writer = CAnimExportMp4Writer;
	window["AscCommonSlide"].createAnimExportWriter = createAnimExportWriter;
	window["AscCommonSlide"].isAnimExportFormatSupported = isAnimExportFormatSupported;
	window["AscCommonSlide"].getSupportedAnimExportFormats = getSupportedAnimExportFormats;

})(window);
