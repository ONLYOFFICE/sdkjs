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

(function(undefined) {
	const c_oAscPresentationRangeType = {
		AllSlides:      1,
		SelectedSlides: 2,
		ActiveSlide:    3,
		CustomRange:    4
	};

	const c_oAscSlidesOnPageArrangementType = {
		Vertical:   1,
		Horizontal: 2
	};

	const c_oAscSlidesOnPagePrintType = {
		FullPageSlides: 1,
		SlideWithNotes: 2,
		Outline:        3,
		Handouts:       4
	};

	function presentationRangeTypeToString(value) {
		switch (value) {
			case c_oAscPresentationRangeType.AllSlides:      return "AllSlides";
			case c_oAscPresentationRangeType.SelectedSlides: return "SelectedSlides";
			case c_oAscPresentationRangeType.ActiveSlide:    return "ActiveSlide";
			case c_oAscPresentationRangeType.CustomRange:    return "CustomRange";
			default:                                         return null;
		}
	}
	function presentationRangeTypeFromString(str) {
		switch (str) {
			case "AllSlides":      return c_oAscPresentationRangeType.AllSlides;
			case "SelectedSlides": return c_oAscPresentationRangeType.SelectedSlides;
			case "ActiveSlide":    return c_oAscPresentationRangeType.ActiveSlide;
			case "CustomRange":    return c_oAscPresentationRangeType.CustomRange;
			default:               return null;
		}
	}

	function slidesOnPageArrangementTypeToString(value) {
		switch (value) {
			case c_oAscSlidesOnPageArrangementType.Vertical:   return "Vertical";
			case c_oAscSlidesOnPageArrangementType.Horizontal: return "Horizontal";
			default:                                          return null;
		}
	}
	function slidesOnPageArrangementTypeFromString(str) {
		switch (str) {
			case "Vertical":   return c_oAscSlidesOnPageArrangementType.Vertical;
			case "Horizontal": return c_oAscSlidesOnPageArrangementType.Horizontal;
			default:           return null;
		}
	}

	function slidesOnPagePrintTypeToString(value) {
		switch (value) {
			case c_oAscSlidesOnPagePrintType.FullPageSlides: return "FullPageSlides";
			case c_oAscSlidesOnPagePrintType.SlideWithNotes: return "SlideWithNotes";
			case c_oAscSlidesOnPagePrintType.Outline:        return "Outline";
			case c_oAscSlidesOnPagePrintType.Handouts:       return "Handouts";
			default:                                         return null;
		}
	}
	function slidesOnPagePrintTypeFromString(str) {
		switch (str) {
			case "FullPageSlides": return c_oAscSlidesOnPagePrintType.FullPageSlides;
			case "SlideWithNotes": return c_oAscSlidesOnPagePrintType.SlideWithNotes;
			case "Outline":        return c_oAscSlidesOnPagePrintType.Outline;
			case "Handouts":       return c_oAscSlidesOnPagePrintType.Handouts;
			default:               return null;
		}
	}

	function CAscPageRange() {
		this.rangeType = c_oAscPresentationRangeType.AllSlides;
		this.customRange = "";
	}
	CAscPageRange.fromJSON = function (jsonObj) {
		if (!jsonObj) {
			return null;
		}
		const res = new CAscPageRange();
		res.rangeType = presentationRangeTypeFromString(jsonObj["rangeType"]);
		res.customRange = jsonObj["customRange"];
		return res;
	};
	CAscPageRange.prototype.toJSON = function () {
		const res = {};
		res["rangeType"] = presentationRangeTypeToString(this.rangeType);
		res["customRange"] = this.customRange;
		return res;
	};
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
		this.width = 210;
		this.height = 297;
	}
	CAscPageOptions.fromJSON = function (jsonObj) {
		if (!jsonObj) {
			return null;
		}
		const res = new CAscPageOptions();
		res.width = jsonObj["width"];
		res.height = jsonObj["height"];
		return res;
	};
	CAscPageOptions.prototype.toJSON = function () {
		const res = {};
		res["width"] = this.width;
		res["height"] = this.height;
		return res;
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
		this.arrangementType = c_oAscSlidesOnPageArrangementType.Horizontal;
		this.printType = c_oAscSlidesOnPagePrintType.FullPageSlides;
	}
	CAscSlidesOnPagePrintOptions.fromJSON = function (jsonObj) {
		if (!jsonObj) {
			return null;
		}
		const res = new CAscSlidesOnPagePrintOptions();
		res.slidesCount = jsonObj["slidesCount"];
		res.arrangementType = slidesOnPageArrangementTypeFromString(jsonObj["arrangementType"]);
		res.printType = slidesOnPagePrintTypeFromString(jsonObj["printType"]);
		return res;
	};
	CAscSlidesOnPagePrintOptions.prototype.toJSON = function () {
		const res = {};
		res["slidesCount"] = this.slidesCount;
		res["arrangementType"] = slidesOnPageArrangementTypeToString(this.arrangementType);
		res["printType"] = slidesOnPagePrintTypeToString(this.printType);
		return res;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_setSlidesCount = function(slidesCount) {
		this.slidesCount = slidesCount;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_getSlidesCount = function() {
		return this.slidesCount;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_setArrangementType = function(arrangementType) {
		this.arrangementType = arrangementType;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_getArrangementType = function() {
		return this.arrangementType;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_setPrintType = function(printType) {
		this.printType = printType;
	};
	CAscSlidesOnPagePrintOptions.prototype.asc_getPrintType = function() {
		return this.printType;
	};

	function CAscSlidesOnPageOptions() {
		this.isPrintSlideNumbersOnHandouts = false;
		this.isFrameSlides = true;
		this.isScaleToFitPaper = true;
		this.isPrintComments = true;
	}
	CAscSlidesOnPageOptions.fromJSON = function (jsonObj) {
		if (!jsonObj) {
			return null;
		}
		const res = new CAscSlidesOnPageOptions();
		res.isPrintSlideNumbersOnHandouts = jsonObj["isPrintSlideNumbersOnHandouts"];
		res.isFrameSlides = jsonObj["isFrameSlides"];
		res.isScaleToFitPaper = jsonObj["isScaleToFitPaper"];
		res.isPrintComments = jsonObj["isPrintComments"];
		return res;
	};
	CAscSlidesOnPageOptions.prototype.toJSON = function () {
		const res = {};
		res["isPrintSlideNumbersOnHandouts"] = this.isPrintSlideNumbersOnHandouts;
		res["isFrameSlides"] = this.isFrameSlides;
		res["isScaleToFitPaper"] = this.isScaleToFitPaper;
		res["isPrintComments"] = this.isPrintComments;
		return res;
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
	CAscSlidePrintOptions.fromJSON = function (jsonObj) {
		if (!jsonObj) {
			return null;
		}
		const res = new CAscSlidePrintOptions();
		res.isPrintSlideBackground = jsonObj["isPrintSlideBackground"];
		res.isPrintHiddenSlides = jsonObj["isPrintHiddenSlides"];
		return res;
	};
	CAscSlidePrintOptions.prototype.toJSON = function () {
		const res = {};
		res["isPrintSlideBackground"] = this.isPrintSlideBackground;
		res["isPrintHiddenSlides"] = this.isPrintHiddenSlides;
		return res;
	};
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
		this.slidesOnPagePrintOptions = new CAscSlidesOnPagePrintOptions();
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
	CAscPresentationPrintOptions.prototype.asc_setSlidesOnPagePrintOptions = function(slidesOnPagePrintOptions) {
		this.slidesOnPagePrintOptions = slidesOnPagePrintOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_getSlidesOnPagePrintOptions = function() {
		return this.slidesOnPagePrintOptions;
	};
	CAscPresentationPrintOptions.prototype.asc_setCopies = function(copies) {
		this.copies = copies;
	};
	CAscPresentationPrintOptions.prototype.asc_getCopies = function() {
		return this.copies;
	};
	CAscPresentationPrintOptions.fromJSON = function (jsonObj) {
		if (!jsonObj) {
			return null;
		}
		const res = new CAscPresentationPrintOptions();
		res.rangeOptions = CAscPageRange.fromJSON(jsonObj["rangeOptions"]);
		res.pageOptions = CAscPageOptions.fromJSON(jsonObj["pageOptions"]);
		res.slidesOnPageOptions = CAscSlidesOnPageOptions.fromJSON(jsonObj["slidesOnPageOptions"]);
		res.slidePrintOptions = CAscSlidePrintOptions.fromJSON(jsonObj["slidePrintOptions"]);
		res.slidesOnPagePrintOptions = CAscSlidesOnPagePrintOptions.fromJSON(jsonObj["slidesOnPagePrintOptions"]);
		res.copies = jsonObj["copies"];
		return res;
	};
	CAscPresentationPrintOptions.prototype.toJSON = function () {
		const res = {};
		res["rangeOptions"] = this.rangeOptions.toJSON();
		res["pageOptions"] = this.pageOptions.toJSON();
		res["slidesOnPageOptions"] = this.slidesOnPageOptions.toJSON();
		res["slidePrintOptions"] = this.slidePrintOptions.toJSON();
		res["slidesOnPagePrintOptions"] = this.slidesOnPagePrintOptions.toJSON();
		res["copies"] = this.copies;
		return res;
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

	window["Asc"]["c_oAscSlidesOnPageArrangementType"] = window["Asc"].c_oAscSlidesOnPageArrangementType = c_oAscSlidesOnPageArrangementType;
	c_oAscSlidesOnPageArrangementType["Vertical"] = c_oAscSlidesOnPageArrangementType.Vertical;
	c_oAscSlidesOnPageArrangementType["Horizontal"] = c_oAscSlidesOnPageArrangementType.Horizontal;

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
	CAscPageOptions.prototype["asc_setWidth"] = CAscPageOptions.prototype.asc_setWidth;
	CAscPageOptions.prototype["asc_getWidth"] = CAscPageOptions.prototype.asc_getWidth;
	CAscPageOptions.prototype["asc_setHeight"] = CAscPageOptions.prototype.asc_setHeight;
	CAscPageOptions.prototype["asc_getHeight"] = CAscPageOptions.prototype.asc_getHeight;

	window["Asc"]["CAscSlidesOnPagePrintOptions"] = window["Asc"].CAscSlidesOnPagePrintOptions = CAscSlidesOnPagePrintOptions;
	CAscSlidesOnPagePrintOptions.prototype["asc_setSlidesCount"] = CAscSlidesOnPagePrintOptions.prototype.asc_setSlidesCount;
	CAscSlidesOnPagePrintOptions.prototype["asc_getSlidesCount"] = CAscSlidesOnPagePrintOptions.prototype.asc_getSlidesCount;
	CAscSlidesOnPagePrintOptions.prototype["asc_setArrangementType"] = CAscSlidesOnPagePrintOptions.prototype.asc_setArrangementType;
	CAscSlidesOnPagePrintOptions.prototype["asc_getArrangementType"] = CAscSlidesOnPagePrintOptions.prototype.asc_getArrangementType;
	CAscSlidesOnPagePrintOptions.prototype["asc_setPrintType"] = CAscSlidesOnPagePrintOptions.prototype.asc_setPrintType;
	CAscSlidesOnPagePrintOptions.prototype["asc_getPrintType"] = CAscSlidesOnPagePrintOptions.prototype.asc_getPrintType;

	window["Asc"]["CAscSlidesOnPageOptions"] = window["Asc"].CAscSlidesOnPageOptions = CAscSlidesOnPageOptions;
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
	CAscPresentationPrintOptions.prototype["asc_setSlidesOnPagePrintOptions"] = CAscPresentationPrintOptions.prototype.asc_setSlidesOnPagePrintOptions;
	CAscPresentationPrintOptions.prototype["asc_getSlidesOnPagePrintOptions"] = CAscPresentationPrintOptions.prototype.asc_getSlidesOnPagePrintOptions;
	CAscPresentationPrintOptions.prototype["asc_setCopies"] = CAscPresentationPrintOptions.prototype.asc_setCopies;
	CAscPresentationPrintOptions.prototype["asc_getCopies"] = CAscPresentationPrintOptions.prototype.asc_getCopies;
	CAscPresentationPrintOptions.prototype["asc_setNativeOptions"] = CAscPresentationPrintOptions.prototype.asc_setNativeOptions;
	CAscPresentationPrintOptions.prototype["asc_getNativeOptions"] = CAscPresentationPrintOptions.prototype.asc_getNativeOptions;
})();
