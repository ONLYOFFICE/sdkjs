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

	const MIN_DELAY_CS = 2;
	const CLICK_STEP_DELAY_MS = 1500;

	function CAnimExportDelayPlanner() {
		this.totalMs = 0;
		this.writtenCs = 0;
	}
	CAnimExportDelayPlanner.prototype.next = function (durationMs) {
		this.totalMs += durationMs;
		let delay = Math.round(this.totalMs / 10) - this.writtenCs;
		if (delay < MIN_DELAY_CS) {
			delay = MIN_DELAY_CS;
		}
		this.writtenCs += delay;
		return delay;
	};
	CAnimExportDelayPlanner.prototype.getWrittenMs = function () {
		return this.writtenCs * 10;
	};

	function CAnimExportPlan(presentation, settings) {
		this.presentation = presentation;
		this.settings = settings;
		this.fps = AscCommonSlide.getAnimatedExportFps(settings);
		this.frameWidth = 0;
		this.frameHeight = 0;
		this.slides = [];

		this.calculateFrameSize();
		this.collectSlides();
	}
	CAnimExportPlan.prototype.calculateFrameSize = function () {
		const widthMM = this.presentation.GetWidthMM();
		const heightMM = this.presentation.GetHeightMM();
		const height = AscCommonSlide.getAnimatedExportPresetHeight(this.settings.format, this.settings.preset);
		this.frameHeight = height & ~1;
		this.frameWidth = Math.round(height * widthMM / heightMM) & ~1;
	};
	CAnimExportPlan.prototype.getSlideIndexes = function () {
		const range = this.settings.rangeOptions;
		if (!range) {
			const indexes = [];
			for (let index = 0; index < this.presentation.Slides.length; ++index) {
				if (this.presentation.Slides[index].isVisible()) {
					indexes.push(index);
				}
			}
			return indexes;
		}
		return AscCommonSlide.getRangeSlideIndexes(this.presentation, range, false);
	};
	CAnimExportPlan.prototype.collectSlides = function () {
		const indexes = this.getSlideIndexes();
		for (let position = 0; position < indexes.length; ++position) {
			const slideIndex = indexes[position];
			const slide = this.presentation.Slides[slideIndex];
			if (!slide) {
				continue;
			}
			this.slides.push({
				index: slideIndex,
				slide: slide,
				transition: this.getTransition(slide),
				slideDuration: this.getSlideDuration(slide)
			});
		}
	};
	CAnimExportPlan.prototype.getTransition = function (slide) {
		const transition = slide.transition;
		if (!transition || Asc.c_oAscSlideTransitionTypes.None === transition.TransitionType) {
			return null;
		}
		const duration = transition.TransitionDuration;
		if (!duration || duration <= 0) {
			return null;
		}
		const typeAndOption = transition.getTypeAndOption();
		return {type: typeAndOption.Type, param: typeAndOption.Option, duration: duration};
	};
	CAnimExportPlan.prototype.getSlideDuration = function (slide) {
		if (slide.isAdvanceAfterTransition()) {
			const duration = slide.getAdvanceDuration();
			if (duration > 0) {
				return duration;
			}
		}
		return this.settings.slideDuration;
	};
	CAnimExportPlan.prototype.getSlideCount = function () {
		return this.slides.length;
	};
	CAnimExportPlan.prototype.getSegmentFrameCount = function (durationMs) {
		const count = Math.round(durationMs * this.fps / 1000);
		return count < 1 ? 1 : count;
	};
	CAnimExportPlan.prototype.getFrameStepMs = function () {
		return 1000 / this.fps;
	};
	CAnimExportPlan.prototype.getClickDelayMs = function () {
		return CLICK_STEP_DELAY_MS;
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].CAnimExportPlan = CAnimExportPlan;
	window["AscCommonSlide"].CAnimExportDelayPlanner = CAnimExportDelayPlanner;

})(window);
