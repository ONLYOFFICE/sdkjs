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

	const MAX_CLICK_STEPS = 100;

	function createCanvas(width, height) {
		const canvas = document.createElement('canvas');
		canvas.width = width;
		canvas.height = height;
		return canvas;
	}

	function isEndlessTimeNode(node, depth) {
		if (!node || depth > 8) {
			return false;
		}
		if (node.getRepeatCount) {
			const repeatCount = node.getRepeatCount();
			if (repeatCount && repeatCount.isIndefinite()) {
				return true;
			}
		}
		if (node.getRepeatDur) {
			const repeatDur = node.getRepeatDur();
			if (repeatDur && repeatDur.isIndefinite()) {
				return true;
			}
		}
		const children = node.getChildrenTimeNodes ? node.getChildrenTimeNodes() : [];
		for (let index = 0; index < children.length; ++index) {
			if (isEndlessTimeNode(children[index], depth + 1)) {
				return true;
			}
		}
		return false;
	}

	function CAnimExportFrameSource(api, plan) {
		this.api = api;
		this.plan = plan;
		this.presentation = plan.presentation;
		this.width = plan.frameWidth;
		this.height = plan.frameHeight;
		this.rect = {x: 0, y: 0, w: this.width, h: this.height};

		this.hostCanvas = createCanvas(this.width, this.height);
		this.hostOverlay = createCanvas(this.width, this.height);
		this.captureCanvas = createCanvas(this.width, this.height);
		this.captureContext = this.captureCanvas.getContext('2d');
		this.slideSnapshot = createCanvas(this.width, this.height);
		this.transitionTarget = createCanvas(this.width, this.height);

		this.isTransparent = AscCommonSlide.isAnimatedExportTransparent(plan.settings);
		this.hostCanvas.style.backgroundColor = this.isTransparent ? "transparent" : "#000000";

		this.transition = null;
		this.player = null;
		this.slideItem = null;
		this.savedSlidePlayer = null;
		this.savedAnnotations = null;
		this.hasSlideFrame = false;
		this.frameDirty = false;
		this.isTransitionActive = false;

		this.host = this.createHost();
	}
	CAnimExportFrameSource.prototype.createHost = function () {
		const self = this;
		return {
			Canvas: this.hostCanvas,
			Overlay: this.hostOverlay,
			SlideNum: 0,
			OnEndTransition: function () {
				self.isTransitionActive = false;
			},
			GetPrevVisibleSlide: function () {
				return self.host.SlideNum;
			},
			OnRecalculateAnimationFrame: function () {
				self.frameDirty = true;
			},
			OnAnimMainSeqFinished: function () {
			}
		};
	};
	CAnimExportFrameSource.prototype.begin = function () {
		const demonstrationManager = this.api.WordControl.DemonstrationManager;
		if (demonstrationManager) {
			this.savedAnnotations = demonstrationManager.SlideAnnotations;
			demonstrationManager.SlideAnnotations = null;
		}
		this.presentation.ensureRecalculateOnOpen();
	};
	CAnimExportFrameSource.prototype.end = function () {
		this.endSlide();
		this.endTransition();
		const demonstrationManager = this.api.WordControl.DemonstrationManager;
		if (demonstrationManager && this.savedAnnotations) {
			demonstrationManager.SlideAnnotations = this.savedAnnotations;
			this.savedAnnotations = null;
		}
		if (this.transition && this.transition.GLTransition) {
			this.transition.GLTransition.Dispose();
			this.transition.GLTransition = null;
		}
		this.transition = null;
	};
	CAnimExportFrameSource.prototype.getFrameData = function () {
		return this.captureContext.getImageData(0, 0, this.width, this.height).data;
	};
	CAnimExportFrameSource.prototype.getCanvas = function () {
		return this.captureCanvas;
	};
	CAnimExportFrameSource.prototype.beginSlide = function (item) {
		this.endSlide();

		const slide = item.slide;
		const self = this;
		this.presentation.checkSlideRecalculateOnOpen(slide);
		AscFormat.ExecuteNoHistory(function () {
			const timing = slide.timing;
			if (timing) {
				slide.timing = timing.createDuplicate();
			}
			self.player = new AscFormat.CAnimationPlayer(slide, self.host);
			self.player.setVirtualTime(0);
			self.player.start();
			slide.timing = timing;
		}, this, []);

		this.savedSlidePlayer = slide.animationPlayer;
		slide.animationPlayer = this.player;
		this.slideItem = item;
		this.host.SlideNum = item.index;
		this.hasSlideFrame = false;
	};
	CAnimExportFrameSource.prototype.endSlide = function () {
		if (!this.slideItem) {
			return;
		}
		if (this.player) {
			this.player.pause();
		}
		this.slideItem.slide.animationPlayer = this.savedSlidePlayer;
		this.savedSlidePlayer = null;
		this.slideItem = null;
		this.player = null;
	};
	CAnimExportFrameSource.prototype.drawSlideFrame = function () {
		this.captureContext.setTransform(1, 0, 0, 1, 0, 0);
		this.captureContext.clearRect(0, 0, this.width, this.height);
		this.player.drawFrame(this.captureCanvas, this.rect, this.isTransparent ? false : undefined);
	};
	CAnimExportFrameSource.prototype.drawSlideToCache = function (canvas, isBackgroundOnly) {
		const savedWatermark = this.api.watermarkDraw;
		this.api.watermarkDraw = null;
		this.clearCanvas(canvas);
		let layer = isBackgroundOnly ? true : undefined;
		if (this.isTransparent) {
			layer = isBackgroundOnly ? null : false;
		}
		if (null !== layer) {
			this.player.drawFrame(canvas, this.rect, layer);
		}
		this.api.watermarkDraw = savedWatermark;
	};
	CAnimExportFrameSource.prototype.renderSlideFrame = function (timeMs) {
		this.frameDirty = false;
		this.player.setVirtualTime(timeMs);
		this.player.onFrame();
		if (this.hasSlideFrame && !this.frameDirty) {
			return false;
		}
		this.drawSlideFrame();
		this.hasSlideFrame = true;
		return true;
	};
	CAnimExportFrameSource.prototype.nextClickStep = function () {
		return this.player.onNextSlide();
	};
	CAnimExportFrameSource.prototype.isAnimationActive = function () {
		const timing = this.player.timings[0];
		if (!timing) {
			return false;
		}
		const mainSequence = timing.getMainSequence();
		if (!mainSequence) {
			return false;
		}
		const nodes = mainSequence.getChildrenTimeNodes();
		for (let index = 0; index < nodes.length; ++index) {
			if (!nodes[index].isIdle() && !nodes[index].isAtEnd()) {
				return true;
			}
		}
		return false;
	};
	CAnimExportFrameSource.prototype.hasEndlessAnimation = function () {
		for (let index = 0; index < this.player.timings.length; ++index) {
			const root = this.player.timings[index].getTimingRootNode();
			if (isEndlessTimeNode(root, 0)) {
				return true;
			}
		}
		return false;
	};
	CAnimExportFrameSource.prototype.getMaxClickSteps = function () {
		return MAX_CLICK_STEPS;
	};
	CAnimExportFrameSource.prototype.saveSlideSnapshot = function () {
		const context = this.slideSnapshot.getContext('2d');
		context.clearRect(0, 0, this.width, this.height);
		context.drawImage(this.captureCanvas, 0, 0);
	};
	CAnimExportFrameSource.prototype.renderTransitionTarget = function () {
		this.drawSlideToCache(this.transitionTarget);
	};
	CAnimExportFrameSource.prototype.clearSlideSnapshot = function () {
		this.clearCanvas(this.slideSnapshot);
	};
	CAnimExportFrameSource.prototype.renderBackgroundSnapshot = function () {
		this.drawSlideToCache(this.slideSnapshot, true);
	};
	CAnimExportFrameSource.prototype.beginTransition = function (transitionInfo, prevSlideIndex, slideIndex) {
		if (!this.transition) {
			this.transition = new CTransitionAnimation(this.api.WordControl);
			this.transition.DemonstrationObject = this.host;
		}
		const transition = this.transition;
		transition.Type = transitionInfo.type;
		transition.Param = transitionInfo.param;
		transition.Duration = transitionInfo.duration;
		transition.IsBackward = false;
		transition.IsTransparent = this.isTransparent;
		transition.Rect.x = 0;
		transition.Rect.y = 0;
		transition.Rect.w = this.width;
		transition.Rect.h = this.height;

		transition.CacheImage1.Image = this.slideSnapshot;
		transition.CacheImage1.SlideNum = prevSlideIndex;
		transition.CacheImage2.Image = this.transitionTarget;
		transition.CacheImage2.SlideNum = slideIndex;

		this.clearCanvas(this.hostCanvas);
		this.clearCanvas(this.hostOverlay);

		this.isTransitionActive = true;
		transition.VirtualTime = 0;
		transition.Start(false);
		this.stopTransitionTimer();
		return this.isTransitionActive;
	};
	CAnimExportFrameSource.prototype.clearCanvas = function (canvas) {
		const context = canvas.getContext('2d');
		context.setTransform(1, 0, 0, 1, 0, 0);
		context.clearRect(0, 0, this.width, this.height);
	};
	CAnimExportFrameSource.prototype.stopTransitionTimer = function () {
		if (null !== this.transition.TimerId) {
			__cancelFrame(this.transition.TimerId);
		}
	};
	CAnimExportFrameSource.prototype.renderTransitionFrame = function (part) {
		const transition = this.transition;
		transition.VirtualTime = transition.StartTime + part * transition.Duration;
		transition.FrameFunction();
		this.stopTransitionTimer();
		this.composeTransitionFrame();
		return this.isTransitionActive;
	};
	CAnimExportFrameSource.prototype.composeTransitionFrame = function () {
		const context = this.captureContext;
		context.setTransform(1, 0, 0, 1, 0, 0);
		context.clearRect(0, 0, this.width, this.height);
		context.drawImage(this.hostCanvas, 0, 0);
		context.drawImage(this.hostOverlay, 0, 0);
	};
	CAnimExportFrameSource.prototype.endTransition = function () {
		if (!this.transition || !this.isTransitionActive) {
			return;
		}
		this.transition.End(true);
		this.isTransitionActive = false;
	};
	CAnimExportFrameSource.prototype.drawTargetFrame = function () {
		const context = this.captureContext;
		context.setTransform(1, 0, 0, 1, 0, 0);
		context.clearRect(0, 0, this.width, this.height);
		context.drawImage(this.transitionTarget, 0, 0);
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].CAnimExportFrameSource = CAnimExportFrameSource;

})(window);
