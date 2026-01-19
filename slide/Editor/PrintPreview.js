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

(function(undefined) {
	const c_oAscPresentationRangeType = {
		AllSlides:      1,
		SelectedSlides: 2,
		ActiveSlide:    3,
		CustomRange:    4
	};

	const c_oAscSlidesOnPageArrangmentType = {
		Vertical:   1,
		Horizontal: 2
	};

	const c_oAscSlidesOnPagePrintType = {
		FullPageSlides: 1,
		SlideWithNotes: 2,
		Outline:        3,
		Handouts:       4
	};

	function CAscPageRange() {
		this.rangeType = c_oAscPresentationRangeType.AllSlides;
		this.customRange = [];
	}
	CAscPageRange.prototype.asc_setRangeType = function(type) {
		this.rangeType = type;
	};
	CAscPageRange.prototype.asc_getRangeType = function() {
		return this.rangeType;
	};
	CAscPageRange.prototype.asc_setCustomRange = function(range) {
		this.customRange = range;
	};
	CAscPageRange.prototype.asc_getCustomRange = function() {
		return this.customRange;
	};

	function CAscPageOptions() {
		this.orientationType = Asc.c_oAscPageOrientation.PagePortrait;
		this.width = 210;
		this.height = 297;
	}
	CAscPageOptions.prototype.asc_setOrientationType = function(type) {
		this.orientationType = type;
	};
	CAscPageOptions.prototype.asc_getOrientationType = function() {
		return this.orientationType;
	};
	CAscPageOptions.prototype.asc_setWidth = function(width) {
		this.width = width;
	};
	CAscPageOptions.prototype.asc_getWidth = function() {
		return this.width;
	};
	CAscPageOptions.prototype.asc_setHeight = function(height) {
		this.height = height;
	};
	CAscPageOptions.prototype.asc_getHeight = function() {
		return this.height;
	};

	function CAscSlidesOnPagePrintOptions() {
		this.slidesCount = 1;
		this.arrangmentType = c_oAscSlidesOnPageArrangmentType.Horizontal;
		this.printType = c_oAscSlidesOnPagePrintType.FullPageSlides;
	}
	CAscSlidesOnPagePrintOptions.prototype.asc_setSlidesCount = function(slidesCount) {
		this.slidesCount = slidesCount;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_getSlidesCount = function() {
		return this.slidesCount;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_setArrangmentType = function(arrangmentType) {
		this.arrangmentType = arrangmentType;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_getArrangmentType = function() {
		return this.arrangmentType;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_setPrintType = function(printType) {
		this.printType = printType;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_getPrintType = function() {
		return this.printType;
	};

	function CAscSlidesOnPageOptions() {
		this.positionType = null;
		this.isPrintSlideNumbersOnHandouts = false;
		this.isFrameSlides = true;
		this.isScaleToFitPaper = true;
		this.isPrintComments = true;
	}
	CAscSlidesOnPageOptions.prototype.asc_setPositionType = function(positionType) {
		this.positionType = positionType;
	};
	CAscSlidesOnPageOptions.prototype.asc_getPositionType = function() {
		return this.positionType;
	};
	CAscSlidesOnPageOptions.prototype.asc_setIsPrintSlideNumbersOnHandouts = function(isPrintSlideNumbersOnHandouts) {
		this.isPrintSlideNumbersOnHandouts = isPrintSlideNumbersOnHandouts;
	};
	CAscSlidesOnPageOptions.prototype.asc_getIsPrintSlideNumbersOnHandouts = function() {
		return this.isPrintSlideNumbersOnHandouts;
	};
	CAscSlidesOnPageOptions.prototype.asc_setIsFrameSlides = function(isFrameSlides) {
		this.isFrameSlides = isFrameSlides;
	};
	CAscSlidesOnPageOptions.prototype.asc_getIsFrameSlides = function() {
		return this.isFrameSlides;
	};
	CAscSlidesOnPageOptions.prototype.asc_setIsScaleToFitPaper = function(isScaleToFitPaper) {
		this.isScaleToFitPaper = isScaleToFitPaper;
	};
	CAscSlidesOnPageOptions.prototype.asc_getIsScaleToFitPaper = function() {
		return this.isScaleToFitPaper;
	};
	CAscSlidesOnPageOptions.prototype.asc_setIsPrintComments = function(isPrintComments) {
		this.isPrintComments = isPrintComments;
	};
	CAscSlidesOnPageOptions.prototype.asc_getIsPrintComments = function() {
		return this.isPrintComments;
	};

	function CAscSlidePrintOptions() {
		this.isPrintSlideBackground = true;
		this.isPrintHiddenSlides = true;
	}
	CAscSlidePrintOptions.prototype.asc_setIsPrintSlideBackground = function(isPrintSlideBackground) {
		this.isPrintSlideBackground = isPrintSlideBackground;
	};
	CAscSlidePrintOptions.prototype.asc_getIsPrintSlideBackground = function() {
		return this.isPrintSlideBackground;
	};
	CAscSlidePrintOptions.prototype.asc_setIsPrintHiddenSlides = function(isPrintHiddenSlides) {
		this.isPrintHiddenSlides = isPrintHiddenSlides;
	};
	CAscSlidePrintOptions.prototype.asc_getIsPrintHiddenSlides = function() {
		return this.isPrintHiddenSlides;
	};

	function CAscPresentationPrintOptions() {
		this.rangeOptions = new CAscPageRange();
		this.pageOptions = new CAscPageOptions();
		this.slidesOnPageOptions = new CAscSlidesOnPageOptions();
		this.slidePrintOptions = new CAscSlidePrintOptions();
		this.copies = 1;
		this.nativeOptions = null;
	}
	CAscPresentationPrintOptions.prototype.asc_setRangeOptions = function(rangeOptions) {
		this.rangeOptions = rangeOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_getRangeOptions = function() {
		return this.rangeOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_setPageOptions = function(pageOptions) {
		this.pageOptions = pageOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_getPageOptions = function() {
		return this.pageOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_setSlidesOnPageOptions = function(slidesOnPageOptions) {
		this.slidesOnPageOptions = slidesOnPageOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_getSlidesOnPageOptions = function() {
		return this.slidesOnPageOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_setSlidePrintOptions = function(slidePrintOptions) {
		this.slidePrintOptions = slidePrintOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_getSlidePrintOptions = function() {
		return this.slidePrintOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_setCopies = function(copies) {
		this.copies = copies;
	};
	CAscPresentationPrintOptions.prototype.asc_getCopies = function() {
		return this.copies;
	};
	CAscPresentationPrintOptions.prototype.asc_setNativeOptions = function(nativeOptions) {
		this.nativeOptions = nativeOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_getNativeOptions = function() {
		return this.nativeOptions;
	};

	window["Asc"] = window["Asc"] || {};
	window["Asc"]["c_oAscPresentationRangeType"] = window["Asc"].c_oAscPresentationRangeType = c_oAscPresentationRangeType;
	c_oAscPresentationRangeType["AllSlides"] = c_oAscPresentationRangeType.AllSlides;
	c_oAscPresentationRangeType["SelectedSlides"] = c_oAscPresentationRangeType.SelectedSlides;
	c_oAscPresentationRangeType["ActiveSlide"] = c_oAscPresentationRangeType.ActiveSlide;
	c_oAscPresentationRangeType["CustomRange"] = c_oAscPresentationRangeType.CustomRange;

	window["Asc"]["c_oAscSlidesOnPageArrangmentType"] = window["Asc"].c_oAscSlidesOnPageArrangmentType = c_oAscSlidesOnPageArrangmentType;
	c_oAscSlidesOnPageArrangmentType["Vertical"] = c_oAscSlidesOnPageArrangmentType.Vertical;
	c_oAscSlidesOnPageArrangmentType["Horizontal"] = c_oAscSlidesOnPageArrangmentType.Horizontal;

	window["Asc"]["c_oAscSlidesOnPagePrintType"] = window["Asc"].c_oAscSlidesOnPagePrintType = c_oAscSlidesOnPagePrintType;
	c_oAscSlidesOnPagePrintType["FullPageSlides"] = c_oAscSlidesOnPagePrintType.FullPageSlides;
	c_oAscSlidesOnPagePrintType["SlideWithNotes"] = c_oAscSlidesOnPagePrintType.SlideWithNotes;
	c_oAscSlidesOnPagePrintType["Outline"] = c_oAscSlidesOnPagePrintType.Outline;
	c_oAscSlidesOnPagePrintType["Handouts"] = c_oAscSlidesOnPagePrintType.Handouts;

	window["Asc"]["CAscPageRange"] = window["Asc"].CAscPageRange = CAscPageRange;
	CAscPageRange.prototype["asc_setRangeType"] = CAscPageRange.prototype.asc_setRangeType;
	CAscPageRange.prototype["asc_getRangeType"] = CAscPageRange.prototype.asc_getRangeType;
	CAscPageRange.prototype["asc_setCustomRange"] = CAscPageRange.prototype.asc_setCustomRange;
	CAscPageRange.prototype["asc_getCustomRange"] = CAscPageRange.prototype.asc_getCustomRange;

	window["Asc"]["CAscPageOptions"] = window["Asc"].CAscPageOptions = CAscPageOptions;
	CAscPageOptions.prototype["asc_setOrientationType"] = CAscPageOptions.prototype.asc_setOrientationType;
	CAscPageOptions.prototype["asc_getOrientationType"] = CAscPageOptions.prototype.asc_getOrientationType;
	CAscPageOptions.prototype["asc_setWidth"] = CAscPageOptions.prototype.asc_setWidth;
	CAscPageOptions.prototype["asc_getWidth"] = CAscPageOptions.prototype.asc_getWidth;
	CAscPageOptions.prototype["asc_setHeight"] = CAscPageOptions.prototype.asc_setHeight;
	CAscPageOptions.prototype["asc_getHeight"] = CAscPageOptions.prototype.asc_getHeight;

	window["Asc"]["CAscSlidesOnPagePrintOptions"] = window["Asc"].CAscSlidesOnPagePrintOptions = CAscSlidesOnPagePrintOptions;
	CAscSlidesOnPagePrintOptions.prototype["asc_setSlidesCount"] = CAscSlidesOnPagePrintOptions.prototype.asc_setSlidesCount;
	CAscSlidesOnPagePrintOptions.prototype["asc_getSlidesCount"] = CAscSlidesOnPagePrintOptions.prototype.asc_getSlidesCount;
	CAscSlidesOnPagePrintOptions.prototype["asc_setArrangmentType"] = CAscSlidesOnPagePrintOptions.prototype.asc_setArrangmentType;
	CAscSlidesOnPagePrintOptions.prototype["asc_getArrangmentType"] = CAscSlidesOnPagePrintOptions.prototype.asc_getArrangmentType;
	CAscSlidesOnPagePrintOptions.prototype["asc_setPrintType"] = CAscSlidesOnPagePrintOptions.prototype.asc_setPrintType;
	CAscSlidesOnPagePrintOptions.prototype["asc_getPrintType"] = CAscSlidesOnPagePrintOptions.prototype.asc_getPrintType;

	window["Asc"]["CAscSlidesOnPageOptions"] = window["Asc"].CAscSlidesOnPageOptions = CAscSlidesOnPageOptions;
	CAscSlidesOnPageOptions.prototype["asc_setPositionType"] = CAscSlidesOnPageOptions.prototype.asc_setPositionType;
	CAscSlidesOnPageOptions.prototype["asc_getPositionType"] = CAscSlidesOnPageOptions.prototype.asc_getPositionType;
	CAscSlidesOnPageOptions.prototype["asc_setIsPrintSlideNumbersOnHandouts"] = CAscSlidesOnPageOptions.prototype.asc_setIsPrintSlideNumbersOnHandouts;
	CAscSlidesOnPageOptions.prototype["asc_getIsPrintSlideNumbersOnHandouts"] = CAscSlidesOnPageOptions.prototype.asc_getIsPrintSlideNumbersOnHandouts;
	CAscSlidesOnPageOptions.prototype["asc_setIsFrameSlides"] = CAscSlidesOnPageOptions.prototype.asc_setIsFrameSlides;
	CAscSlidesOnPageOptions.prototype["asc_getIsFrameSlides"] = CAscSlidesOnPageOptions.prototype.asc_getIsFrameSlides;
	CAscSlidesOnPageOptions.prototype["asc_setIsScaleToFitPaper"] = CAscSlidesOnPageOptions.prototype.asc_setIsScaleToFitPaper;
	CAscSlidesOnPageOptions.prototype["asc_getIsScaleToFitPaper"] = CAscSlidesOnPageOptions.prototype.asc_getIsScaleToFitPaper;
	CAscSlidesOnPageOptions.prototype["asc_setIsPrintComments"] = CAscSlidesOnPageOptions.prototype.asc_setIsPrintComments;
	CAscSlidesOnPageOptions.prototype["asc_getIsPrintComments"] = CAscSlidesOnPageOptions.prototype.asc_getIsPrintComments;

	window["Asc"]["CAscSlidePrintOptions"] = window["Asc"].CAscSlidePrintOptions = CAscSlidePrintOptions;
	CAscSlidePrintOptions.prototype["asc_setIsPrintSlideBackground"] = CAscSlidePrintOptions.prototype.asc_setIsPrintSlideBackground;
	CAscSlidePrintOptions.prototype["asc_getIsPrintSlideBackground"] = CAscSlidePrintOptions.prototype.asc_getIsPrintSlideBackground;
	CAscSlidePrintOptions.prototype["asc_setIsPrintHiddenSlides"] = CAscSlidePrintOptions.prototype.asc_setIsPrintHiddenSlides;
	CAscSlidePrintOptions.prototype["asc_getIsPrintHiddenSlides"] = CAscSlidePrintOptions.prototype.asc_getIsPrintHiddenSlides;

	window["Asc"]["CAscPresentationPrintOptions"] = window["Asc"].CAscPresentationPrintOptions = CAscPresentationPrintOptions;
	CAscPresentationPrintOptions.prototype["asc_setRangeOptions"] = CAscPresentationPrintOptions.prototype.asc_setRangeOptions;
	CAscPresentationPrintOptions.prototype["asc_getRangeOptions"] = CAscPresentationPrintOptions.prototype.asc_getRangeOptions;
	CAscPresentationPrintOptions.prototype["asc_setPageOptions"] = CAscPresentationPrintOptions.prototype.asc_setPageOptions;
	CAscPresentationPrintOptions.prototype["asc_getPageOptions"] = CAscPresentationPrintOptions.prototype.asc_getPageOptions;
	CAscPresentationPrintOptions.prototype["asc_setSlidesOnPageOptions"] = CAscPresentationPrintOptions.prototype.asc_setSlidesOnPageOptions;
	CAscPresentationPrintOptions.prototype["asc_getSlidesOnPageOptions"] = CAscPresentationPrintOptions.prototype.asc_getSlidesOnPageOptions;
	CAscPresentationPrintOptions.prototype["asc_setSlidePrintOptions"] = CAscPresentationPrintOptions.prototype.asc_setSlidePrintOptions;
	CAscPresentationPrintOptions.prototype["asc_getSlidePrintOptions"] = CAscPresentationPrintOptions.prototype.asc_getSlidePrintOptions;
	CAscPresentationPrintOptions.prototype["asc_setCopies"] = CAscPresentationPrintOptions.prototype.asc_setCopies;
	CAscPresentationPrintOptions.prototype["asc_getCopies"] = CAscPresentationPrintOptions.prototype.asc_getCopies;
	CAscPresentationPrintOptions.prototype["asc_setNativeOptions"] = CAscPresentationPrintOptions.prototype.asc_setNativeOptions;
	CAscPresentationPrintOptions.prototype["asc_getNativeOptions"] = CAscPresentationPrintOptions.prototype.asc_getNativeOptions;
})();
