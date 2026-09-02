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

	const STAGE_SLIDE_BEGIN = 1;
	const STAGE_TRANSITION = 2;
	const STAGE_SLIDE = 3;
	const STAGE_SLIDE_END = 4;
	const STAGE_DONE = 5;

	const ENDLESS_GUARD_MS = 20000;

	function CAnimatedExporter(api, settings) {
		AscCommon.CActionOnTimerBase.call(this);
		this.SetDoFirstActionOnTimer(true);

		this.api = api;
		this.settings = settings;
		this.plan = null;
		this.frameSource = null;
		this.writer = null;

		this.stage = STAGE_DONE;
		this.slidePosition = 0;
		this.slideTime = 0;
		this.slideIdleTime = 0;
		this.slideStepTime = 0;
		this.clickStep = 0;
		this.noMoreClickSteps = false;
		this.transitionFrame = 0;
		this.transitionFrameCount = 0;
		this.frameCount = 0;
		this.isTruncated = false;
		this.isCanceled = false;
		this.errorId = null;
		this.errorDetails = null;
	}
	CAnimatedExporter.prototype = Object.create(AscCommon.CActionOnTimerBase.prototype);
	CAnimatedExporter.prototype.constructor = CAnimatedExporter;

	CAnimatedExporter.prototype.start = function () {
		this.Begin();
	};
	CAnimatedExporter.prototype.cancel = function () {
		if (!this.IsInProgress()) {
			return false;
		}
		this.isCanceled = true;
		this.Stop();
		if (this.writer) {
			this.writer.abort();
		}
		this.release();
		this.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ExportVideo);
		this.api.sendEvent("asc_onAnimatedExportEnd", null);
		return true;
	};
	CAnimatedExporter.prototype.OnBegin = function () {
		this.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ExportVideo);

		const presentation = this.api.private_GetLogicDocument();
		this.plan = new AscCommonSlide.CAnimExportPlan(presentation, this.settings);
		if (0 === this.plan.getSlideCount()) {
			this.errorId = Asc.c_oAscError.ID.AnimatedExport;
			this.stage = STAGE_DONE;
			return;
		}
		this.frameSource = new AscCommonSlide.CAnimExportFrameSource(this.api, this.plan);
		this.frameSource.begin();
		try {
			this.writer = AscCommonSlide.createAnimExportWriter(this.plan.frameWidth, this.plan.frameHeight,
				this.settings);
		}
		catch (error) {
			this.errorId = Asc.c_oAscError.ID.AnimatedExport;
			this.stage = STAGE_DONE;
			return;
		}
		this.slidePosition = 0;
		this.stage = STAGE_SLIDE_BEGIN;
	};
	CAnimatedExporter.prototype.IsContinue = function () {
		return STAGE_DONE !== this.stage && !this.isCanceled && null === this.errorId;
	};
	CAnimatedExporter.prototype.DoAction = function () {
		if (!this.writer || this.writer.isBusy()) {
			return;
		}
		try {
			switch (this.stage) {
				case STAGE_SLIDE_BEGIN:
					this.doSlideBegin();
					break;
				case STAGE_TRANSITION:
					this.doTransitionFrame();
					break;
				case STAGE_SLIDE:
					this.doSlideFrame();
					break;
				case STAGE_SLIDE_END:
					this.doSlideEnd();
					break;
				default:
					this.stage = STAGE_DONE;
					break;
			}
		}
		catch (error) {
			this.stage = STAGE_DONE;
			this.errorDetails = String(error.stack || error.message || error);
			if (0 === this.frameCount) {
				this.errorId = Asc.c_oAscError.ID.AnimatedExportMemory;
				return;
			}
			this.isTruncated = true;
		}
	};
	CAnimatedExporter.prototype.OnEndTimer = function () {
		this.api.sendEvent("asc_onAnimatedExportProgress", this.getPercent());
	};
	CAnimatedExporter.prototype.getPercent = function () {
		const count = this.plan ? this.plan.getSlideCount() : 0;
		if (count <= 0) {
			return 0;
		}
		return Math.min(99, Math.floor(100 * this.slidePosition / count));
	};
	CAnimatedExporter.prototype.OnEnd = function () {
		if (this.isCanceled) {
			return;
		}
		if (null !== this.errorId) {
			this.finishWithError(this.errorId);
			return;
		}

		const self = this;
		const extension = this.writer.getExtension();
		const mimeType = this.writer.getMimeType();

		this.writer.finish(function (data) {
			if (!data) {
				self.finishWithError(Asc.c_oAscError.ID.AnimatedExport);
				return;
			}

			const result = new Asc.CAscAnimatedExportResult();
			result.format = self.settings.format;
			result.data = data;
			result.width = self.plan.frameWidth;
			result.height = self.plan.frameHeight;
			result.frameCount = self.writer.getFrameCount();
			result.durationMs = self.writer.getDurationMs();
			result.isTruncated = self.isTruncated;

			self.release();
			self.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ExportVideo);

			if (self.settings.autoDownload) {
				self.downloadResult(result, extension, mimeType);
			}
			self.api.sendEvent("asc_onAnimatedExportEnd", result);
		});
	};
	CAnimatedExporter.prototype.finishWithError = function (errorId) {
		if (this.writer) {
			this.writer.abort();
		}
		this.release();
		this.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ExportVideo);
		this.api.sendEvent("asc_onError", errorId, Asc.c_oAscError.Level.NoCritical);
		this.api.sendEvent("asc_onAnimatedExportError", errorId);
	};
	CAnimatedExporter.prototype.canDownloadFromBytes = function () {
		const browser = AscCommon.AscBrowser;
		return !(browser.isAppleDevices && browser.isChrome) && !browser.isAndroidNativeApp;
	};
	CAnimatedExporter.prototype.downloadResult = function (result, extension, mimeType) {
		if (!this.canDownloadFromBytes()) {
			return false;
		}
		const title = AscCommon.changeFileExtention(this.api.documentTitle, extension, Asc.c_nMaxDownloadTitleLen);
		AscCommon.DownloadFileFromBytes(result.data, title, mimeType);
		return true;
	};
	CAnimatedExporter.prototype.release = function () {
		if (this.frameSource) {
			this.frameSource.end();
			this.frameSource = null;
		}
	};
	CAnimatedExporter.prototype.doSlideBegin = function () {
		const item = this.plan.slides[this.slidePosition];
		const isFirst = 0 === this.slidePosition;
		if (!isFirst) {
			this.frameSource.saveSlideSnapshot();
		}
		this.frameSource.beginSlide(item);
		this.slideTime = 0;
		this.slideIdleTime = 0;
		this.slideStepTime = 0;
		this.clickStep = 0;
		this.noMoreClickSteps = false;

		if (item.transition) {
			if (isFirst) {
				this.prepareFirstTransitionSource();
			}
			this.frameSource.renderTransitionTarget();
			const prevIndex = isFirst ? -1 : this.plan.slides[this.slidePosition - 1].index;
			if (this.frameSource.beginTransition(item.transition, prevIndex, item.index)) {
				this.transitionFrame = 0;
				this.transitionFrameCount = this.plan.getSegmentFrameCount(item.transition.duration);
				this.stage = STAGE_TRANSITION;
				return;
			}
		}
		this.stage = STAGE_SLIDE;
	};
	CAnimatedExporter.prototype.prepareFirstTransitionSource = function () {
		this.frameSource.clearSlideSnapshot();
	};
	CAnimatedExporter.prototype.doTransitionFrame = function () {
		const item = this.plan.slides[this.slidePosition];
		const duration = item.transition.duration;
		const part = this.transitionFrame / this.transitionFrameCount;
		const isPlaying = this.frameSource.renderTransitionFrame(part);
		++this.transitionFrame;

		this.emitFrame(duration / this.transitionFrameCount, true);

		if (!isPlaying || this.transitionFrame >= this.transitionFrameCount) {
			this.frameSource.endTransition();
			this.stage = STAGE_SLIDE;
		}
	};
	CAnimatedExporter.prototype.doSlideFrame = function () {
		const item = this.plan.slides[this.slidePosition];
		const step = this.plan.getFrameStepMs();

		const isChanged = this.frameSource.renderSlideFrame(this.slideTime);
		this.emitFrame(step, isChanged);
		this.slideTime += step;

		if (this.frameSource.isAnimationActive()) {
			this.slideIdleTime = 0;
		} else {
			this.slideIdleTime += step;
		}

		const isActive = this.frameSource.isAnimationActive();
		const isFinished = this.noMoreClickSteps || this.frameSource.isMainSequenceFinished();
		if (this.slideTime >= item.slideDuration &&
			((!isActive && isFinished) ||
				this.slideTime >= this.slideStepTime + item.slideDuration + ENDLESS_GUARD_MS)) {
			this.stage = STAGE_SLIDE_END;
			return;
		}
		if (isActive) {
			return;
		}
		if (this.clickStep > 0 && this.slideIdleTime < item.slideDuration) {
			return;
		}
		if (!this.noMoreClickSteps && this.clickStep < this.frameSource.getMaxClickSteps()) {
			if (this.frameSource.nextClickStep()) {
				++this.clickStep;
				this.slideIdleTime = 0;
				this.slideStepTime = this.slideTime;
			}
			else {
				this.noMoreClickSteps = true;
			}
		}
	};
	CAnimatedExporter.prototype.doSlideEnd = function () {
		++this.slidePosition;
		if (this.slidePosition >= this.plan.getSlideCount() || this.isTruncated) {
			this.frameSource.endSlide();
			this.stage = STAGE_DONE;
			return;
		}
		this.stage = STAGE_SLIDE_BEGIN;
	};
	CAnimatedExporter.prototype.isVideoFormat = function () {
		const format = this.settings.format;
		return Asc.c_oAscAnimatedExportFormat.WebM === format || Asc.c_oAscAnimatedExportFormat.Mp4 === format;
	};
	CAnimatedExporter.prototype.emitFrame = function (durationMs, isChanged) {
		if (this.isVideoFormat()) {
			durationMs = this.plan.getFrameStepMs();
		}
		else if (!isChanged) {
			this.writer.addDuration(durationMs);
			return;
		}
		if (this.writer.addFrame(this.frameSource, durationMs)) {
			++this.frameCount;
		}
		if (this.writer.getErrorText()) {
			this.errorId = Asc.c_oAscError.ID.AnimatedExport;
			this.stage = STAGE_DONE;
		}
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].CAnimatedExporter = CAnimatedExporter;

})(window);
