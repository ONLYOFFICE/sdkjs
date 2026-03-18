/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function (window, undefined) {

	window["AscCommon"] = window['AscCommon'] || {};

	function GIFDataBase() {
	}
	window["AscCommon"].GIFDataBase = GIFDataBase;
	GIFDataBase.prototype.getWidth = function() {
		return 0;
	};
	GIFDataBase.prototype.getHeight = function() {
		return 0;
	};
	GIFDataBase.prototype.getFramesCount = function() {
		return 0;
	};
	GIFDataBase.prototype.getFrame = function(frameIndex) {
		return null;
	};
	GIFDataBase.prototype.getFrameDelayMs = function(frameIndex) {
		return 0;
	};
	GIFDataBase.prototype.getDurationMs = function() {
		return 0;
	};
	GIFDataBase.prototype.getLoopCount = function() {
		return 0;
	};
	GIFDataBase.prototype.getFrameTimeMs = function(frameIndex) {
		return 0;
	};
	GIFDataBase.prototype.getFrameIndexAtTime = function(timeMs) {
		return 0;
	};



	function GIFDataGIFuct(gifuctData) {
		AscCommon.GIFDataBase.call(this);
		this.isValid = false;
		this.frames = [];
		this.gifuctData = { raw: { lsd: { width: 0, height: 0 } } };

		try {
			if (gifuctData && gifuctData.byteLength > 0) {
				this.gifuctData = new window.GIFUCT.GIF(gifuctData);
				this.frames = this.gifuctData.decompressFrames(true);
				if (this.frames && this.frames.length > 0) {
					this.isValid = true;
				}
			}
		} catch (e) {
		}
	}
	GIFDataGIFuct.prototype = Object.create(AscCommon.GIFDataBase.prototype);
	GIFDataGIFuct.prototype.getWidth = function() {
		return this.gifuctData.raw.lsd.width;
	};
	GIFDataGIFuct.prototype.getHeight = function() {
		return this.gifuctData.raw.lsd.height;
	};
	GIFDataGIFuct.prototype.getFramesCount = function() {
		return this.frames.length;
	};
	GIFDataGIFuct.prototype.getFrame = function(frameIndex) {
		if (frameIndex < 0 || frameIndex >= this.frames.length) {
			return null;
		}
		return this.frames[frameIndex];
	};
	GIFDataGIFuct.prototype.getFrameDelayMs = function(frameIndex) {
		let frame = this.getFrame(frameIndex);
		if (!frame || !frame.delay) {
			return 0;
		}
		return frame.delay;
	};
	GIFDataGIFuct.prototype.getDurationMs = function() {
		let totalDuration = 0;
		let framesCount = this.getFramesCount();
		for (let i = 0; i < framesCount; i++) {
			totalDuration += this.getFrameDelayMs(i);
		}
		return totalDuration;
	};
	GIFDataGIFuct.prototype.getLoopCount = function() {
		return 0;
	};
	GIFDataGIFuct.prototype.getFrameTimeMs = function(frameIndex) {
		let timeMs = 0;
		for (let i = 0; i < frameIndex; i++) {
			timeMs += this.getFrameDelayMs(i);
		}
		return timeMs;
	};
	GIFDataGIFuct.prototype.getFrameIndexAtTime = function(timeMs) {
		let durationMs = this.getDurationMs();
		if (durationMs === 0) {
			return 0;
		}
		let loopCount = this.getLoopCount();
		if (loopCount === 0) {
			timeMs = timeMs % durationMs;
		} else {
			let totalLoopDuration = durationMs * loopCount;
			if (timeMs >= totalLoopDuration) {
				return this.getFramesCount() - 1;
			}
			timeMs = timeMs % durationMs;
		}
		let accumulatedTime = 0;
		let framesCount = this.getFramesCount();
		for (let i = 0; i < framesCount; i++) {
			accumulatedTime += this.getFrameDelayMs(i);
			if (timeMs < accumulatedTime) {
				return i;
			}
		}
		return framesCount - 1;
	};
	window["AscCommon"].GIFDataGIFuct = GIFDataGIFuct;

	function GIFPlayer(gifData, srcImageData) {
		this.gifData = new GIFDataGIFuct(gifData);
		this.srcImageData = srcImageData;
		this.lastDrawFrame = -1;

		this.startTime = null;
	}
	GIFPlayer.prototype.isStarted = function() {
		return this.startTime !== null;
	};
	GIFPlayer.prototype.getCurrentFrameIndex = function() {
		if (!this.isStarted()) {
			return -1;
		}
		let currentTime = (new Date()).getTime();
		let elapsedTime = currentTime - this.startTime;
		return this.gifData.getFrameIndexAtTime(elapsedTime);
	}
	GIFPlayer.prototype.onTick = function() {
		let updated = false;
		if (!this.isStarted()) {
			return updated;
		}

		let currentTime = (new Date()).getTime();
		let elapsedTime = currentTime - this.startTime;
		let currentFrame = this.gifData.getFrameIndexAtTime(elapsedTime);
		if (this.lastDrawFrame !== currentFrame) {
			if (currentFrame < this.lastDrawFrame) {
				this.lastDrawFrame = -1;
			}
			for (let i = this.lastDrawFrame + 1; i <= currentFrame; i++) {
				this.srcImageData.updateFrame(this.getFrameData(i));
			}
			this.lastDrawFrame = currentFrame;
			updated = true;
		}
		return updated;
	};
	GIFPlayer.prototype.getFrameForDraw = function() {
		this.lastDrawFrame = this.getCurrentFrameIndex();
		return this.getFrameData(this.lastDrawFrame);
	};
	GIFPlayer.prototype.getFrameData = function(index) {
		let frame = this.gifData.getFrame(index);
		if (frame) {
			frame.idx = index;
			frame.imageWidth = this.gifData.getWidth();
			frame.imageHeight = this.gifData.getHeight();
		}
		return frame;
	};
	GIFPlayer.prototype.start = function() {
		if (this.isStarted()) {
			return;
		}

		if (!this.gifData.isValid) {
			return;
		}
		this.startTime = (new Date()).getTime();
		this.lastDrawFrame = -1;
		this.onTick();
	};
	GIFPlayer.prototype.stop = function() {
		this.startTime = null;
		this.lastDrawFrame = -1;
		this.srcImageData.clearRawData();
	};
	window["AscCommon"].GIFPlayer = GIFPlayer;


	function GIFAdapter(image) {
		this.blipFill = image;

		this.canvasCtx = null;
		this.imageData = null;

		this.prevDisposal = 0;
		this.prevDims = null;
		this.savedCanvas = null;
	}
	GIFAdapter.prototype.updateFrame = function(rawData) {
		if (!rawData) {
			return;
		}

		if (!rawData.dims || !rawData.patch || !rawData.imageWidth || !rawData.imageHeight) {
			return;
		}

		this.blipFill.rawData = rawData;
		let d = rawData.dims;

		if (d.width <= 0 || d.height <= 0 || d.width > 10000 || d.height > 10000) {
			return;
		}
		if (rawData.imageWidth <= 0 || rawData.imageHeight <= 0 ||
			rawData.imageWidth > 10000 || rawData.imageHeight > 10000) {
			return;
		}
		if (!this.blipFill.canvas) {
			this.blipFill.canvas = document.createElement("canvas");
			this.blipFill.canvas.width = rawData.imageWidth;
			this.blipFill.canvas.height = rawData.imageHeight;
			this.canvasCtx = this.blipFill.canvas.getContext("2d", {
				alpha: true,
				willReadFrequently: false
			});
		}

		if (!this.tempCanvas) {
			this.tempCanvas = document.createElement("canvas");
			this.tempCtx = this.tempCanvas.getContext("2d", {
				alpha: true,
				willReadFrequently: false
			});
		}

		if (this.tempCanvas.width !== d.width || this.tempCanvas.height !== d.height) {
			this.tempCanvas.width = d.width;
			this.tempCanvas.height = d.height;
		}

		if (this.prevDisposal === 2 && this.prevDims) {
			this.canvasCtx.clearRect(this.prevDims.left, this.prevDims.top,
				this.prevDims.width, this.prevDims.height);
		} else if (this.prevDisposal === 3 && this.savedCanvas) {
			this.canvasCtx.clearRect(0, 0, this.blipFill.canvas.width, this.blipFill.canvas.height);
			this.canvasCtx.drawImage(this.savedCanvas, 0, 0);
		}

		if (rawData.idx === 0) {
			this.canvasCtx.clearRect(0, 0, rawData.imageWidth, rawData.imageHeight);
			this.prevDisposal = 0;
			this.prevDims = null;
		}

		if (rawData.disposalType === 3) {
			if (!this.savedCanvas) {
				this.savedCanvas = document.createElement("canvas");
			}
			this.savedCanvas.width = this.blipFill.canvas.width;
			this.savedCanvas.height = this.blipFill.canvas.height;
			let savedCtx = this.savedCanvas.getContext("2d");
			savedCtx.drawImage(this.blipFill.canvas, 0, 0);
		}

		let imageData = new ImageData(rawData.patch, d.width, d.height);
		this.tempCtx.putImageData(imageData, 0, 0);
		this.canvasCtx.drawImage(this.tempCanvas, d.left, d.top);

		this.prevDisposal = rawData.disposalType;
		this.prevDims = { left: d.left, top: d.top, width: d.width, height: d.height };
	};
	GIFAdapter.prototype.clearRawData = function() {
		this.blipFill.rawData = null;
		this.blipFill.canvas = null;
		this.canvasCtx = null;
		this.tempCanvas = null;
		this.tempCtx = null;
		this.prevDisposal = 0;
		this.prevDims = null;
		this.savedCanvas = null;
	};
	window["AscCommon"].GIFAdapter = GIFAdapter;

})(window);
