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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

	const Asc = window['Asc'];
	const AscCommon = window['AscCommon'];

	// Import
	let prot;
	const c_oAscMouseMoveDataTypes = Asc.c_oAscMouseMoveDataTypes;

	const c_oAscColor = Asc.c_oAscColor;
	const c_oAscFill = Asc.c_oAscFill;
	const c_oAscFillBlipType = Asc.c_oAscFillBlipType;
	const c_oAscChartTypeSettings = Asc.c_oAscChartTypeSettings;
	const c_oAscTickMark = Asc.c_oAscTickMark;
	const c_oAscAxisType = Asc.c_oAscAxisType;
	// ---------------------------------------------------------------------------------------------------------------

	const c_oAscArrUserColors = [10646501, 16749875, 1087211, 103817, 16760641, 16272775, 8765789, 14707685, 48336, 5729515, 16757719, 56805, 10081791, 12884479, 16751001, 6748927, 16762931, 6865407, 15650047, 16737894, 3407768, 16759142, 10852863, 6750176, 16774656, 13926655, 13815039, 3397375, 11927347, 16752947, 9404671, 4980531, 16744678, 3407830, 15919360, 16731553, 52479, 13330175, 16743219, 3386367, 14221056, 16737966, 1896960, 65484, 10970879, 16759296, 16711680, 13496832, 62072, 49906, 16734720, 10682112, 7890687, 16731610, 65406, 38655, 16747008, 59890, 12733951, 15859712, 47077, 15050496, 15224319, 10154496, 58807, 16724950, 1759488, 9981439, 15064320, 15893248, 16724883, 58737, 15007744, 36594, 12772608, 12137471, 6442495, 15039488, 16718470, 14274816, 53721, 16718545, 1625088, 15881472, 13419776, 32985, 16711800, 1490688, 16711884, 8991743, 13407488, 41932, 7978752, 15028480, 52387, 15007927, 12114176, 1421824, 55726, 13041893, 10665728, 30924, 49049, 14241024, 36530, 11709440, 13397504, 45710, 34214];

	function CreateAscColorCustom(r, g, b, auto) {
		let ret = new asc_CColor();
		ret.type = c_oAscColor.COLOR_TYPE_SRGB;
		ret.r = r;
		ret.g = g;
		ret.b = b;
		ret.a = 255;
		ret.Auto = (undefined === auto ? false : auto);
		return ret;
	}

	function CreateAscColor(unicolor) {
		if (null == unicolor || null == unicolor.color) {
			return new asc_CColor();
		}

		let ret = new asc_CColor();
		ret.r = unicolor.RGBA.R;
		ret.g = unicolor.RGBA.G;
		ret.b = unicolor.RGBA.B;
		ret.a = unicolor.RGBA.A;

		let _color = unicolor.color;
		switch (_color.type) {
			case c_oAscColor.COLOR_TYPE_SRGB:
			case c_oAscColor.COLOR_TYPE_SYS: {
				break;
			}
			case c_oAscColor.COLOR_TYPE_PRST:
			case c_oAscColor.COLOR_TYPE_SCHEME: {
				ret.type = _color.type;
				ret.value = _color.id;
				break;
			}
			default:
				break;
		}
		return ret;
	}

	let uuid = [];
	for (let i = 0; i < 256; i++) {
		uuid[i] = (i < 16 ? "0" : "") + (i).toString(16);
	}

	function CreateUUID(isNoDashes) {
		let d0 = Math.random() * 0xffffffff | 0;
		let d1 = Math.random() * 0xffffffff | 0;
		let d2 = Math.random() * 0xffffffff | 0;
		let d3 = Math.random() * 0xffffffff | 0;

		if (isNoDashes) return uuid[d0 & 0xff] + uuid[d0 >> 8 & 0xff] + uuid[d0 >> 16 & 0xff] + uuid[d0 >> 24 & 0xff] +
			uuid[d1 & 0xff] + uuid[d1 >> 8 & 0xff] + uuid[d1 >> 16 & 0x0f | 0x40] + uuid[d1 >> 24 & 0xff] +
			uuid[d2 & 0x3f | 0x80] + uuid[d2 >> 8 & 0xff] + uuid[d2 >> 16 & 0xff] + uuid[d2 >> 24 & 0xff] +
			uuid[d3 & 0xff] + uuid[d3 >> 8 & 0xff] + uuid[d3 >> 16 & 0xff] + uuid[d3 >> 24 & 0xff];
		else
			return uuid[d0 & 0xff] + uuid[d0 >> 8 & 0xff] + uuid[d0 >> 16 & 0xff] + uuid[d0 >> 24 & 0xff] + "-" +
				uuid[d1 & 0xff] + uuid[d1 >> 8 & 0xff] + "-" + uuid[d1 >> 16 & 0x0f | 0x40] + uuid[d1 >> 24 & 0xff] + "-" +
				uuid[d2 & 0x3f | 0x80] + uuid[d2 >> 8 & 0xff] + "-" + uuid[d2 >> 16 & 0xff] + uuid[d2 >> 24 & 0xff] +
				uuid[d3 & 0xff] + uuid[d3 >> 8 & 0xff] + uuid[d3 >> 16 & 0xff] + uuid[d3 >> 24 & 0xff];
	}

	function CreateGUID() {
		function s4() {
			return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1);
		}

		let val = '{' + s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4() + '}';
		val = val.toUpperCase();
		return val;
	}

	function CreateUInt32() {
		return Math.floor(Math.random() * 0x100000000);
	}

	function FixDurableId(val) {
		//numbers greater than 0x7FFFFFFE cause MS Office errors(ST_LongHexNumber by spec)
		let res = val & 0x7FFFFFFF;
		return (0x7FFFFFFF !== res) ? res : res - 1;
	}

	function CreateDurableId() {
		return FixDurableId(CreateUInt32());
	}

	function ExtendPrototype(dst, src) {
		for (let k in src.prototype) {
			dst.prototype[k] = src.prototype[k];
		}
	}

	function isFileBuild() {
		return window["NATIVE_EDITOR_ENJINE"] && !window["IS_NATIVE_EDITOR"] && !window["DoctRendererMode"];
	}

	function checkCanvasInDiv(sDivId) {
		let oDiv = document.getElementById(sDivId);
		if (!oDiv) {
			return null;
		}
		let aChildren = oDiv.children;
		let oCanvas = null, i;
		for (i = 0; i < aChildren.length; ++i) {
			if (aChildren[i].nodeName && aChildren[i].nodeName.toUpperCase() === 'CANVAS') {
				oCanvas = aChildren[i];
				break;
			}
		}
		if (null === oCanvas) {
			let rPR = AscCommon.AscBrowser.retinaPixelRatio;
			oCanvas = document.createElement('canvas');
			oCanvas.style.width = "100%";
			oCanvas.style.height = "100%";
			oDiv.appendChild(oCanvas);
			oCanvas.width = Math.round(oCanvas.clientWidth * rPR);
			oCanvas.height = Math.round(oCanvas.clientHeight * rPR);
		}
		return oCanvas;
	}

	function isValidJs(str) {
		try {
			eval("throw 0;" + str);
		} catch (e) {
			if (e === 0) return true;
		}
		return false;
	}


	function asc_menu_ReadPaddings(_params, _cursor){
		const _paddings = new Asc.asc_CPaddings();
		_paddings.read(_params, _cursor);
		return _paddings;
	}

	function asc_menu_ReadColor(_params, _cursor) {
		const _color = new Asc.asc_CColor();
		_color.read(_params, _cursor);
		return _color;
	}

	function parseJSDoc(jsDoc) {
		// function for parsing jsDoc (for builder method "AddCustomFunction")
		const result = [];

		// A regular expression for searching for JSDoc comments describing parameters, properties, and functions
		const jsDocRegex = /\/\*\*([\s\S]*?)\*\//g;
		const paramRegex = /@param\s+{(\??)(.+?)}\s+(?:(\[)?([\w.]+|\{[\w.]+\}))?(?:\s*=\s*([^@]+?)(?=\s|\)|$))?(?:\s+(.+))?/g;
		const propertyRegex = /@property\s+{(\??)(.+?)}\s+(?:(\[)?([\w.]+|\{[\w.]+\}))?(?:\s*=\s*([^@]+?)(?=\s|\)|$))?(?:\s+(.+))?/g;
	
		let match;
		while ((match = jsDocRegex.exec(jsDoc)) !== null) {
			const parsedData = {};
			const commentBlock = match[1].trim();
	
			// Parsing of function parameters
			const params = [];
			let paramMatch;
			while ((paramMatch = paramRegex.exec(commentBlock)) !== null) {
				const type = paramMatch[2];
				const isOptional = paramMatch[3] !== undefined || paramMatch[1] === '?';
				const name = paramMatch[4];
				const defaultValue = isOptional && paramMatch[5] !== undefined ? paramMatch[5].trim().replace(/\]$/, '') : undefined;
				const description = paramMatch[6] || '';
	
				params.push({
					type: type,
					name: name,
					isOptional: isOptional,
					defaultValue: defaultValue,
					description: description
				});
			}
	
			// Parsing properties
			const properties = [];
			let propertyMatch;
			while ((propertyMatch = propertyRegex.exec(commentBlock)) !== null) {
				const type = propertyMatch[2];
				const isOptional = propertyMatch[3] !== undefined || propertyMatch[1] === '?';
				const name = propertyMatch[4];
				const defaultValue = isOptional && propertyMatch[5] !== undefined ? propertyMatch[5].trim().replace(/\]$/, '') : undefined;
				const description = propertyMatch[6] || '';
	
				properties.push({
					type: type,
					name: name,
					isOptional: isOptional,
					defaultValue: defaultValue,
					description: description
				});
			}
	
			// Parsing what the function returns
			const returnRegex = /@returns?\s+{(.+?)}\s+(.+)/g;
			let returnInfo = null;
			const returnMatch = returnRegex.exec(commentBlock);
			if (returnMatch !== null) {
				returnInfo = {
					type: returnMatch[1],
					description: returnMatch[2]
				};
			}

			const regex = /@nameLocale\s+{([^}]+)}/;
			const nameMatch = regex.exec(commentBlock);
			const nameLocale = {};

			if (nameMatch && nameMatch[1]) {
				const localesAndNames = nameMatch[1].split(/\s*\|\s*/);
				localesAndNames.forEach(function (localeAndName) {
					const splitLocaleAndName = localeAndName.split(':');
					const locale = splitLocaleAndName[0];
					const name = splitLocaleAndName[1];
					nameLocale[locale.trim()] = name.trim();
				});
			}

			const tagRegex = /@(\w+)(?:\s+(.+))?/g;
			const tags = [];
			let tagMatch;
			while ((tagMatch = tagRegex.exec(commentBlock)) !== null) {
				const tagName = tagMatch[1];
				if (!['param', 'property', 'returns', 'return', 'nameLocale'].includes(tagName)) {
					tags[tagName] = 1;
				}
			}

			// Parsing function description
			const descriptionRegex = /\*\s*(.*)/g;
			const descriptionMatch = descriptionRegex.exec(commentBlock);
			const description = descriptionMatch ? descriptionMatch[1].trim() : '';
	
			parsedData.params = params;
			parsedData.properties = properties;
			parsedData.returnInfo = returnInfo;
			parsedData.description = description;
			parsedData.nameLocale = nameLocale;
			parsedData.tags = tags;
			result.push(parsedData);
		}
	
		return result;
	};

	var c_oLicenseResult = {
		Error         : 1,
		Expired       : 2,
		Success       : 3,
		UnknownUser   : 4,
		Connections   : 5,
		ExpiredTrial  : 6,
		SuccessLimit  : 7,
		UsersCount    : 8,
		ConnectionsOS : 9,
		UsersCountOS  : 10,
		ExpiredLimited: 11,
		ConnectionsLiveOS: 12,
		ConnectionsLive: 13,
		UsersViewCount: 14,
		UsersViewCountOS: 15,
		NotBefore: 16
	};

	let c_oRights = {
		None: 0, Edit: 1, Review: 2, Comment: 3, View: 4
	};

	let c_oLicenseMode = {
		None: 0, Trial: 1, Developer: 2, Limited: 4
	};

	let EPluginDataType = {
		none: "none", text: "text", ole: "ole", html: "html", desktop: "desktop"
	};

	/** @constructor */
	function asc_CSignatureLine() {
		this.id = undefined;
		this.guid = "";
		this.signer1 = "";
		this.signer2 = "";
		this.email = "";

		this.instructions = "";
		this.showDate = false;

		this.valid = 0;

		this.image = "";

		this.date = "";
		this.isvisible = false;
		this.isrequested = false;
		
		this.isForm = false;
	}

	asc_CSignatureLine.prototype.correct = function () {
		if (this.id == null) this.id = "0";
		if (this.guid == null) this.guid = "";
		if (this.signer1 == null) this.signer1 = "";
		if (this.signer2 == null) this.signer2 = "";
		if (this.email == null) this.email = "";
		if (this.instructions == null) this.instructions = "";
		if (this.showDate == null) this.showDate = false;
		if (this.valid == null) this.valid = 0;
		if (this.image == null) this.image = "";
		if (this.date == null) this.date = "";
		if (this.isvisible == null) this.isvisible = false;
		if (this.isForm === null) this.isForm = false;
	};
	asc_CSignatureLine.prototype.asc_getId = function () {
		return this.id;
	};
	asc_CSignatureLine.prototype.asc_setId = function (v) {
		this.id = v;
	};
	asc_CSignatureLine.prototype.asc_getGuid = function () {
		return this.guid;
	};
	asc_CSignatureLine.prototype.asc_setGuid = function (v) {
		this.guid = v;
	};
	asc_CSignatureLine.prototype.asc_getSigner1 = function () {
		return this.signer1;
	};
	asc_CSignatureLine.prototype.asc_setSigner1 = function (v) {
		this.signer1 = v;
	};
	asc_CSignatureLine.prototype.asc_getSigner2 = function () {
		return this.signer2;
	};
	asc_CSignatureLine.prototype.asc_setSigner2 = function (v) {
		this.signer2 = v;
	};
	asc_CSignatureLine.prototype.asc_getEmail = function () {
		return this.email;
	};
	asc_CSignatureLine.prototype.asc_setEmail = function (v) {
		this.email = v;
	};
	asc_CSignatureLine.prototype.asc_getInstructions = function () {
		return this.instructions;
	};
	asc_CSignatureLine.prototype.asc_setInstructions = function (v) {
		this.instructions = v;
	};
	asc_CSignatureLine.prototype.asc_getShowDate = function () {
		return this.showDate !== false;
	};
	asc_CSignatureLine.prototype.asc_setShowDate = function (v) {
		this.showDate = v;
	};
	asc_CSignatureLine.prototype.asc_getValid = function () {
		return this.valid;
	};
	asc_CSignatureLine.prototype.asc_setValid = function (v) {
		this.valid = v;
	};
	asc_CSignatureLine.prototype.asc_getDate = function () {
		return this.date;
	};
	asc_CSignatureLine.prototype.asc_setDate = function (v) {
		this.date = v;
	};
	asc_CSignatureLine.prototype.asc_getVisible = function () {
		return this.isvisible;
	};
	asc_CSignatureLine.prototype.asc_setVisible = function (v) {
		this.isvisible = v;
	};
	asc_CSignatureLine.prototype.asc_getRequested = function () {
		return this.isrequested;
	};
	asc_CSignatureLine.prototype.asc_setRequested = function (v) {
		this.isrequested = v;
	};
	asc_CSignatureLine.prototype.asc_getIsForm = function() {
		return this.isForm;
	};
	asc_CSignatureLine.prototype.asc_setIsForm = function(v) {
		this.isForm = v;
	};

	/**
	 * Класс asc_CAscEditorPermissions для прав редакторов
	 * -----------------------------------------------------------------------------
	 *
	 * @constructor
	 * @memberOf Asc
	 */
	function asc_CAscEditorPermissions() {
		this.licenseType = c_oLicenseResult.Error;
		this.licenseMode = c_oLicenseMode.None;
		this.isLight = false;
		this.rights = c_oRights.None;

		this.canCoAuthoring = true;
		this.canReaderMode = true;
		this.canBranding = false;
		this.customization = false;
		this.isAutosaveEnable = true;
		this.AutosaveMinInterval = 300;
		this.isAnalyticsEnable = false;
		this.buildVersion = null;
		this.buildNumber = null;
		this.liveViewerSupport = null;

		this.betaVersion = AscCommon.g_cIsBeta;

		return this;
	}

	asc_CAscEditorPermissions.prototype.asc_getLicenseType = function () {
		return this.licenseType;
	};
	asc_CAscEditorPermissions.prototype.asc_getCanCoAuthoring = function () {
		return this.canCoAuthoring;
	};
	asc_CAscEditorPermissions.prototype.asc_getCanReaderMode = function () {
		return this.canReaderMode;
	};
	asc_CAscEditorPermissions.prototype.asc_getCanBranding = function () {
		return this.canBranding;
	};
	asc_CAscEditorPermissions.prototype.asc_getCustomization = function () {
		return this.customization;
	};
	asc_CAscEditorPermissions.prototype.asc_getIsAutosaveEnable = function () {
		return this.isAutosaveEnable;
	};
	asc_CAscEditorPermissions.prototype.asc_getAutosaveMinInterval = function () {
		return this.AutosaveMinInterval;
	};
	asc_CAscEditorPermissions.prototype.asc_getIsAnalyticsEnable = function () {
		return this.isAnalyticsEnable;
	};
	asc_CAscEditorPermissions.prototype.asc_getIsLight = function () {
		return this.isLight;
	};
	asc_CAscEditorPermissions.prototype.asc_getLicenseMode = function () {
		return this.licenseMode;
	};
	asc_CAscEditorPermissions.prototype.asc_getRights = function () {
		return this.rights;
	};
	asc_CAscEditorPermissions.prototype.asc_getBuildVersion = function () {
		return this.buildVersion;
	};
	asc_CAscEditorPermissions.prototype.asc_getBuildNumber = function () {
		return this.buildNumber;
	};
	asc_CAscEditorPermissions.prototype.asc_getLiveViewerSupport = function () {
		return this.liveViewerSupport;
	};
	asc_CAscEditorPermissions.prototype.asc_getIsBeta = function () {
		return this.betaVersion === 'true';
	};

	asc_CAscEditorPermissions.prototype.setLicenseType = function (v) {
		this.licenseType = v;
	};
	asc_CAscEditorPermissions.prototype.setCanBranding = function (v) {
		this.canBranding = v;
	};
	asc_CAscEditorPermissions.prototype.setCustomization = function (v) {
		this.customization = v;
	};
	asc_CAscEditorPermissions.prototype.setIsLight = function (v) {
		this.isLight = v;
	};
	asc_CAscEditorPermissions.prototype.setLicenseMode = function (v) {
		this.licenseMode = v;
	};
	asc_CAscEditorPermissions.prototype.setRights = function (v) {
		this.rights = v;
	};
	asc_CAscEditorPermissions.prototype.setBuildVersion = function (v) {
		this.buildVersion = v;
	};
	asc_CAscEditorPermissions.prototype.setBuildNumber = function (v) {
		this.buildNumber = v;
	};
	asc_CAscEditorPermissions.prototype.setLiveViewerSupport = function (v) {
		this.liveViewerSupport = v;
	};

	function asc_CAxNumFmt(oAxis) {
		this.formatCode = "General";
		this.sourceLinked = true;
		this.axis = oAxis;
		if (oAxis) {
			let oNumFmt = oAxis.numFmt;
			if (oNumFmt) {
				this.formatCode = oNumFmt.formatCode;
				this.sourceLinked = oNumFmt.sourceLinked;
			}
		}
	}

	asc_CAxNumFmt.prototype.getFormatCode = function () {
		if (this.sourceLinked) {
			if (this.axis) {
				return this.axis.getSourceFormatCode();
			}
			else {
				return "General";
			}
		}
		return this.formatCode || "General";
	};
	asc_CAxNumFmt.prototype.putFormatCode = function (v) {
		this.formatCode = v;
	};
	asc_CAxNumFmt.prototype.getFormatCellsInfo = function () {
		let num_format = AscCommon.oNumFormatCache.get(this.getFormatCode());
		return num_format.getTypeInfo();
	};
	asc_CAxNumFmt.prototype.getSourceLinked = function () {
		return this.sourceLinked;
	};
	asc_CAxNumFmt.prototype.putSourceLinked = function (v) {
		this.sourceLinked = v;
	};
	asc_CAxNumFmt.prototype.isEqual = function (v) {
		if (!v) {
			return false;
		}
		return this.formatCode === v.formatCode && this.sourceLinked === v.sourceLinked;
	};
	asc_CAxNumFmt.prototype.isCorrect = function () {
		return typeof this.formatCode === "string" && this.formatCode.length > 0;
	};

	/** @constructor */
	function asc_ValAxisSettings() {
		this.minValRule = null;
		this.minVal = null;
		this.maxValRule = null;
		this.maxVal = null;
		this.invertValOrder = null;
		this.logScale = null;
		this.logBase = null;

		this.dispUnitsRule = null;
		this.units = null;


		this.showUnitsOnChart = null;
		this.majorTickMark = null;
		this.minorTickMark = null;
		this.tickLabelsPos = null;
		this.crossesRule = null;
		this.crosses = null;
		this.axisType = c_oAscAxisType.val;
		this.show = true;
		this.label = null;
		this.gridlines = null;
		this.numFmt = null;
		this.isRadar = false;
		this.majorUnit = null;
		this.minorUnit = null;
		this.majorUnitRule = false;
		this.minorUnitRule = false;
	}

	asc_ValAxisSettings.prototype.isEqual = function (oPr) {
		if (!oPr) {
			return false;
		}
		if (this.minValRule !== oPr.minValRule) {
			return false;
		}
		if (this.minVal !== oPr.minVal) {
			return false;
		}
		if (this.maxValRule !== oPr.maxValRule) {
			return false;
		}
		if (this.maxVal !== oPr.maxVal) {
			return false;
		}
		if (this.invertValOrder !== oPr.invertValOrder) {
			return false;
		}
		if (this.logScale !== oPr.logScale) {
			return false;
		}
		if (this.logBase !== oPr.logBase) {
			return false;
		}
		if (this.dispUnitsRule !== oPr.dispUnitsRule) {
			return false;
		}
		if (this.units !== oPr.units) {
			return false;
		}
		if (this.showUnitsOnChart !== oPr.showUnitsOnChart) {
			return false;
		}
		if (this.majorTickMark !== oPr.majorTickMark) {
			return false;
		}
		if (this.minorTickMark !== oPr.minorTickMark) {
			return false;
		}
		if (this.tickLabelsPos !== oPr.tickLabelsPos) {
			return false;
		}
		if (this.crossesRule !== oPr.crossesRule) {
			return false;
		}
		if (this.crosses !== oPr.crosses) {
			return false;
		}
		if (this.axisType !== oPr.axisType) {
			return false;
		}
		if (this.show !== oPr.show) {
			return false;
		}
		if (this.label !== oPr.label) {
			return false;
		}
		if (this.gridlines !== oPr.gridlines) {
			return false;
		}
		let bEqualNumFmt = false;
		if (this.numFmt) {
			bEqualNumFmt = this.numFmt.isEqual(oPr.numFmt);
		}
		else {
			bEqualNumFmt = (this.numFmt === oPr.numFmt);
		}
		if (!bEqualNumFmt) {
			return false;
		}
		if (this.minorUnit !== oPr.minorUnit) {
			return false;
		}
		if (this.majorUnit !== oPr.majorUnit) {
			return false;
		}
		if (this.minorUnitRule !== oPr.minorUnitRule) {
			return false;
		}
		if (this.majorUnitRule !== oPr.majorUnitRule) {
			return false;
		}
		return true;
	};
	asc_ValAxisSettings.prototype.putAxisType = function (v) {
		this.axisType = v;
	};
	asc_ValAxisSettings.prototype.putMinValRule = function (v) {
		this.minValRule = v;
	};
	asc_ValAxisSettings.prototype.putMinVal = function (v) {
		this.minVal = v;
	};
	asc_ValAxisSettings.prototype.putMaxValRule = function (v) {
		this.maxValRule = v;
	};
	asc_ValAxisSettings.prototype.putMaxVal = function (v) {
		this.maxVal = v;
	};
	asc_ValAxisSettings.prototype.putInvertValOrder = function (v) {
		this.invertValOrder = v;
	};
	asc_ValAxisSettings.prototype.putLogScale = function (v) {
		this.logScale = v;
	};
	asc_ValAxisSettings.prototype.putLogBase = function (v) {
		this.logBase = v;
	};
	asc_ValAxisSettings.prototype.putUnits = function (v) {
		this.units = v;
	};
	asc_ValAxisSettings.prototype.putShowUnitsOnChart = function (v) {
		this.showUnitsOnChart = v;
	};
	asc_ValAxisSettings.prototype.putMajorTickMark = function (v) {
		this.majorTickMark = v;
	};
	asc_ValAxisSettings.prototype.putMinorTickMark = function (v) {
		this.minorTickMark = v;
	};
	asc_ValAxisSettings.prototype.putTickLabelsPos = function (v) {
		this.tickLabelsPos = v;
	};
	asc_ValAxisSettings.prototype.putCrossesRule = function (v) {
		this.crossesRule = v;
	};
	asc_ValAxisSettings.prototype.putCrosses = function (v) {
		this.crosses = v;
	};
	asc_ValAxisSettings.prototype.putDispUnitsRule = function (v) {
		this.dispUnitsRule = v;
	};
	asc_ValAxisSettings.prototype.getAxisType = function () {
		return this.axisType;
	};
	asc_ValAxisSettings.prototype.getDispUnitsRule = function () {
		return this.dispUnitsRule;
	};
	asc_ValAxisSettings.prototype.getMinValRule = function () {
		return this.minValRule;
	};
	asc_ValAxisSettings.prototype.getMinVal = function () {
		return this.minVal;
	};
	asc_ValAxisSettings.prototype.getMaxValRule = function () {
		return this.maxValRule;
	};
	asc_ValAxisSettings.prototype.getMaxVal = function () {
		return this.maxVal;
	};
	asc_ValAxisSettings.prototype.getInvertValOrder = function () {
		return this.invertValOrder;
	};
	asc_ValAxisSettings.prototype.getLogScale = function () {
		return this.logScale;
	};
	asc_ValAxisSettings.prototype.getLogBase = function () {
		return this.logBase;
	};
	asc_ValAxisSettings.prototype.getUnits = function () {
		return this.units;
	};
	asc_ValAxisSettings.prototype.getShowUnitsOnChart = function () {
		return this.showUnitsOnChart;
	};
	asc_ValAxisSettings.prototype.getMajorTickMark = function () {
		return this.majorTickMark;
	};
	asc_ValAxisSettings.prototype.getMinorTickMark = function () {
		return this.minorTickMark;
	};
	asc_ValAxisSettings.prototype.getTickLabelsPos = function () {
		return this.tickLabelsPos;
	};
	asc_ValAxisSettings.prototype.getCrossesRule = function () {
		return this.crossesRule;
	};
	asc_ValAxisSettings.prototype.getCrosses = function () {
		return this.crosses;
	};
	asc_ValAxisSettings.prototype.setDefault = function () {
		this.putMinValRule(Asc.c_oAscValAxisRule.auto);
		this.putMaxValRule(Asc.c_oAscValAxisRule.auto);
		this.putMajorUnitRule(Asc.c_oAscValAxisRule.auto);
		this.putMinorUnitRule(Asc.c_oAscValAxisRule.auto);
		this.putTickLabelsPos(Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		this.putInvertValOrder(false);
		this.putDispUnitsRule(Asc.c_oAscValAxUnits.none);
		this.putMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		this.putMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		this.putCrossesRule(Asc.c_oAscCrossesRule.auto);
		this.putShow(true);
	};
	asc_ValAxisSettings.prototype.getShow = function () {
		return this.show;
	};
	asc_ValAxisSettings.prototype.putShow = function (val) {
		this.show = val;
	};
	asc_ValAxisSettings.prototype.getLabel = function () {
		return this.label;
	};
	asc_ValAxisSettings.prototype.putLabel = function (v) {
		this.label = v;
	};
	asc_ValAxisSettings.prototype.getGridlines = function () {
		return this.gridlines;
	};
	asc_ValAxisSettings.prototype.putGridlines = function (v) {
		this.gridlines = v;
	};
	asc_ValAxisSettings.prototype.getNumFmt = function () {
		return this.numFmt;
	};
	asc_ValAxisSettings.prototype.putNumFmt = function (v) {
		this.numFmt = v;
	};
	asc_ValAxisSettings.prototype.isRadarAxis = function () {
		return this.isRadar;
	};
	asc_ValAxisSettings.prototype.putIsRadarAxis = function (v) {
		this.isRadar = v;
	};
	asc_ValAxisSettings.prototype.getMajorUnit = function () {
		return this.majorUnit;
	};
	asc_ValAxisSettings.prototype.getMinorUnit = function () {
		return this.minorUnit;
	};
	asc_ValAxisSettings.prototype.putMajorUnit = function (v) {
		this.majorUnit = v;
	};
	asc_ValAxisSettings.prototype.putMinorUnit = function (v) {
		this.minorUnit = v;
	};
	asc_ValAxisSettings.prototype.getMajorUnitRule = function () {
		return this.majorUnitRule;
	};
	asc_ValAxisSettings.prototype.getMinorUnitRule = function () {
		return this.minorUnitRule;
	};
	asc_ValAxisSettings.prototype.putMajorUnitRule = function (v) {
		this.majorUnitRule = v;
	};
	asc_ValAxisSettings.prototype.putMinorUnitRule = function (v) {
		this.minorUnitRule = v;
	};

	/** @constructor */
	function asc_CatAxisSettings() {
		this.intervalBetweenTick = null;
		this.intervalBetweenLabelsRule = null;
		this.intervalBetweenLabels = null;
		this.invertCatOrder = null;
		this.labelsAxisDistance = null;
		this.majorTickMark = null;
		this.minorTickMark = null;
		this.tickLabelsPos = null;
		this.crossesRule = null;
		this.crosses = null;
		this.labelsPosition = null;
		this.axisType = c_oAscAxisType.cat;
		this.crossMinVal = null;
		this.crossMaxVal = null;
		this.show = true;
		this.label = null;
		this.gridlines = null;
		this.numFmt = null;
		this.auto = false;
		this.isRadar = false;
	}

	asc_CatAxisSettings.prototype.isEqual = function (oPr) {
		if (!oPr) {
			return false;
		}
		if (this.intervalBetweenTick !== oPr.intervalBetweenTick) {
			return false;
		}
		if (this.intervalBetweenLabelsRule !== oPr.intervalBetweenLabelsRule) {
			return false;
		}
		if (this.intervalBetweenLabels !== oPr.intervalBetweenLabels) {
			return false;
		}
		if (this.invertCatOrder !== oPr.invertCatOrder) {
			return false;
		}
		if (this.labelsAxisDistance !== oPr.labelsAxisDistance) {
			return false;
		}
		if (this.majorTickMark !== oPr.majorTickMark) {
			return false;
		}
		if (this.minorTickMark !== oPr.minorTickMark) {
			return false;
		}
		if (this.tickLabelsPos !== oPr.tickLabelsPos) {
			return false;
		}
		if (this.crossesRule !== oPr.crossesRule) {
			return false;
		}
		if (this.crosses !== oPr.crosses) {
			return false;
		}
		if (this.labelsPosition !== oPr.labelsPosition) {
			return false;
		}
		if (this.axisType !== oPr.axisType) {
			return false;
		}
		if (this.crossMinVal !== oPr.crossMinVal) {
			return false;
		}
		if (this.crossMaxVal !== oPr.crossMaxVal) {
			return false;
		}
		if (this.show !== oPr.show) {
			return false;
		}
		if (this.label !== oPr.label) {
			return false;
		}
		if (this.gridlines !== oPr.gridlines) {
			return false;
		}
		if (this.auto !== oPr.auto) {
			return false;
		}
		let bEqualNumFmt = false;
		if (this.numFmt) {
			bEqualNumFmt = this.numFmt.isEqual(oPr.numFmt);
		}
		else {
			bEqualNumFmt = (this.numFmt === oPr.numFmt);
		}
		if (!bEqualNumFmt) {
			return false;
		}
		return true;
	};
	asc_CatAxisSettings.prototype.putIntervalBetweenTick = function (v) {
		this.intervalBetweenTick = v;
	};
	asc_CatAxisSettings.prototype.putIntervalBetweenLabelsRule = function (v) {
		this.intervalBetweenLabelsRule = v;
	};
	asc_CatAxisSettings.prototype.putIntervalBetweenLabels = function (v) {
		this.intervalBetweenLabels = v;
	};
	asc_CatAxisSettings.prototype.putInvertCatOrder = function (v) {
		this.invertCatOrder = v;
	};
	asc_CatAxisSettings.prototype.putLabelsAxisDistance = function (v) {
		this.labelsAxisDistance = v;
	};
	asc_CatAxisSettings.prototype.putMajorTickMark = function (v) {
		this.majorTickMark = v;
	};
	asc_CatAxisSettings.prototype.putMinorTickMark = function (v) {
		this.minorTickMark = v;
	};
	asc_CatAxisSettings.prototype.putTickLabelsPos = function (v) {
		this.tickLabelsPos = v;
	};
	asc_CatAxisSettings.prototype.putCrossesRule = function (v) {
		this.crossesRule = v;
	};
	asc_CatAxisSettings.prototype.putCrosses = function (v) {
		this.crosses = v;
	};
	asc_CatAxisSettings.prototype.putAxisType = function (v) {
		this.axisType = v;
	};
	asc_CatAxisSettings.prototype.putLabelsPosition = function (v) {
		this.labelsPosition = v;
	};
	asc_CatAxisSettings.prototype.getIntervalBetweenTick = function () {
		return this.intervalBetweenTick;
	};
	asc_CatAxisSettings.prototype.getIntervalBetweenLabelsRule = function () {
		return this.intervalBetweenLabelsRule;
	};
	asc_CatAxisSettings.prototype.getIntervalBetweenLabels = function () {
		return this.intervalBetweenLabels;
	};
	asc_CatAxisSettings.prototype.getInvertCatOrder = function () {
		return this.invertCatOrder;
	};
	asc_CatAxisSettings.prototype.getLabelsAxisDistance = function () {
		return this.labelsAxisDistance;
	};
	asc_CatAxisSettings.prototype.getMajorTickMark = function () {
		return this.majorTickMark;
	};
	asc_CatAxisSettings.prototype.getMinorTickMark = function () {
		return this.minorTickMark;
	};
	asc_CatAxisSettings.prototype.getTickLabelsPos = function () {
		return this.tickLabelsPos;
	};
	asc_CatAxisSettings.prototype.getCrossesRule = function () {
		return this.crossesRule;
	};
	asc_CatAxisSettings.prototype.getCrosses = function () {
		return this.crosses;
	};
	asc_CatAxisSettings.prototype.getAxisType = function () {
		return this.axisType;
	};
	asc_CatAxisSettings.prototype.getLabelsPosition = function () {
		return this.labelsPosition;
	};
	asc_CatAxisSettings.prototype.getCrossMinVal = function () {
		return this.crossMinVal;
	};
	asc_CatAxisSettings.prototype.getCrossMaxVal = function () {
		return this.crossMaxVal;
	};
	asc_CatAxisSettings.prototype.putCrossMinVal = function (val) {
		this.crossMinVal = val;
	};
	asc_CatAxisSettings.prototype.putCrossMaxVal = function (val) {
		this.crossMaxVal = val;
	};
	asc_CatAxisSettings.prototype.setDefault = function () {
		this.putIntervalBetweenLabelsRule(Asc.c_oAscBetweenLabelsRule.auto);
		this.putLabelsPosition(Asc.c_oAscLabelsPosition.betweenDivisions);
		this.putTickLabelsPos(Asc.c_oAscTickLabelsPos.TICK_LABEL_POSITION_NEXT_TO);
		this.putLabelsAxisDistance(100);
		this.putMajorTickMark(c_oAscTickMark.TICK_MARK_OUT);
		this.putMinorTickMark(c_oAscTickMark.TICK_MARK_NONE);
		this.putIntervalBetweenTick(1);
		this.putCrossesRule(Asc.c_oAscCrossesRule.auto);
		this.putShow(true);
		this.putAuto(true);
	};
	asc_CatAxisSettings.prototype.getShow = function () {
		return this.show;
	};
	asc_CatAxisSettings.prototype.putShow = function (val) {
		this.show = val;
	};
	asc_CatAxisSettings.prototype.getLabel = function () {
		return this.label;
	};
	asc_CatAxisSettings.prototype.putLabel = function (v) {
		this.label = v;
	};
	asc_CatAxisSettings.prototype.getGridlines = function () {
		return this.gridlines;
	};
	asc_CatAxisSettings.prototype.putGridlines = function (v) {
		this.gridlines = v;
	};
	asc_CatAxisSettings.prototype.getNumFmt = function () {
		return this.numFmt;
	};
	asc_CatAxisSettings.prototype.putNumFmt = function (v) {
		this.numFmt = v;
	};
	asc_CatAxisSettings.prototype.getAuto = function () {
		return this.auto;
	};
	asc_CatAxisSettings.prototype.putAuto = function (v) {
		this.auto = v;
	};
	asc_CatAxisSettings.prototype.isRadarAxis = function () {
		return this.isRadar;
	};
	asc_CatAxisSettings.prototype.putIsRadarAxis = function (v) {
		this.isRadar = v;
	};

	/** @constructor */
	function asc_ChartSettings() {
		this.style = null;
		this.title = null;
		this.rowCols = null;
		this.legendPos = null;
		this.dataLabelsPos = null;
		this.type = null;
		this.showSerName = null;
		this.showCatName = null;
		this.showVal = null;
		this.separator = null;
		this.inColumns = null;

		this.sRange = null;


		this.showMarker = null;
		this.bLine = null;
		this.smooth = null;
		this.chartSpace = null;

		this.bStartEdit = false;
		this.horizontalAxes = [];
		this.verticalAxes = [];
		this.depthAxes = [];

		this.view3D = null;

		this.displayTrendlinesEquation = false;
	}

	//TODO:remove this---------------------
	asc_ChartSettings.prototype.putHorAxisProps = function (v) {
		if (!AscCommon.isRealObject(v)) {
			this.horizontalAxes.length = 0;
		}
		else {
			this.horizontalAxes[0] = v;
		}
	};
	asc_ChartSettings.prototype.getHorAxisProps = function () {
		return this.horizontalAxes[0] || null;
	};
	asc_ChartSettings.prototype.putVertAxisProps = function (v) {
		if (!AscCommon.isRealObject(v)) {
			this.verticalAxes.length = 0;
		}
		else {
			this.verticalAxes[0] = v;
		}
	};
	asc_ChartSettings.prototype.getVertAxisProps = function () {
		return this.verticalAxes[0] || null;
	};
	asc_ChartSettings.prototype.putShowHorAxis = function (v) {
		let oAx = this.horizontalAxes[0];
		if (oAx) {
			oAx.putShow(v);
		}
	};
	asc_ChartSettings.prototype.getShowHorAxis = function () {
		let oAx = this.horizontalAxes[0];
		if (oAx) {
			return oAx.getShow();
		}
		return false;
	};
	asc_ChartSettings.prototype.putShowVerAxis = function (v) {
		let oAx = this.verticalAxes[0];
		if (oAx) {
			oAx.putShow(v);
		}
	};
	asc_ChartSettings.prototype.getShowVerAxis = function () {
		let oAx = this.verticalAxes[0];
		if (oAx) {
			return oAx.getShow();
		}
		return false;
	};
	asc_ChartSettings.prototype.putHorAxisLabel = function (v) {
		let oAx = this.horizontalAxes[0];
		if (oAx) {
			oAx.putLabel(v);
		}
	};
	asc_ChartSettings.prototype.getHorAxisLabel = function () {
		let oAx = this.horizontalAxes[0];
		if (oAx) {
			return oAx.getLabel();
		}
		return null;
	};
	asc_ChartSettings.prototype.putVertAxisLabel = function () {
		let oAx = this.verticalAxes[0];
		if (oAx) {
			return oAx.getLabel();
		}
		return null;
	};
	asc_ChartSettings.prototype.getVertAxisLabel = function () {
		let oAx = this.verticalAxes[0];
		if (oAx) {
			return oAx.getLabel();
		}
		return null;
	};
	asc_ChartSettings.prototype.putHorGridLines = function (v) {
		let oAx = this.verticalAxes[0];
		if (oAx) {
			oAx.putGridlines(v);
		}
	};
	asc_ChartSettings.prototype.getHorGridLines = function () {
		let oAx = this.verticalAxes[0];
		if (oAx) {
			return oAx.getGridlines();
		}
		return null;
	};
	asc_ChartSettings.prototype.putVertGridLines = function (v) {
		let oAx = this.horizontalAxes[0];
		if (oAx) {
			oAx.putGridlines(v);
		}
	};
	asc_ChartSettings.prototype.getVertGridLines = function () {
		let oAx = this.horizontalAxes[0];
		if (oAx) {
			return oAx.getGridlines();
		}
		return null;
	};
	//------------------------------
	asc_ChartSettings.prototype.getHorAxesProps = function () {
		return this.horizontalAxes;
	};
	asc_ChartSettings.prototype.getVertAxesProps = function () {
		return this.verticalAxes;
	};
	asc_ChartSettings.prototype.getDepthAxesProps = function () {
		return this.depthAxes;
	};
	asc_ChartSettings.prototype.getView3d = function () {
		if (this.chartSpace) {
			return this.chartSpace.getView3d();
		}
		return this.view3D ? this.view3D.createDuplicate() : null;
	};
	asc_ChartSettings.prototype.putView3d = function (v) {
		this.view3D = v;
	};
	asc_ChartSettings.prototype.setView3d = function (v) {
		this.putView3d(v);
		if (this.chartSpace) {
			if (v) {
				this.chartSpace.changeView3d(v.createDuplicate());
			}
			else {
				this.chartSpace.changeView3d(null);
			}
			this.updateChart();
		}
	};

	asc_ChartSettings.prototype.addHorAxesProps = function (v) {
		this.horizontalAxes.push(v);
	};
	asc_ChartSettings.prototype.addVertAxesProps = function (v) {
		this.verticalAxes.push(v);
	};
	asc_ChartSettings.prototype.addDepthAxesProps = function (v) {
		this.depthAxes.push(v);
	};
	asc_ChartSettings.prototype.removeAllAxesProps = function () {
		this.horizontalAxes.length = 0;
		this.verticalAxes.length = 0;
		this.depthAxes.length = 0;
	};
	asc_ChartSettings.prototype.equalBool = function (a, b) {
		return ((!!a) === (!!b));
	};
	asc_ChartSettings.prototype.isEqual = function (oPr) {
		if (!oPr) {
			return false;
		}
		if (this.style !== oPr.style) {
			return false;
		}
		if (this.title !== oPr.title) {
			return false;
		}
		if (this.rowCols !== oPr.rowCols) {
			return false;
		}
		if (this.legendPos !== oPr.legendPos) {
			return false;
		}
		if (this.dataLabelsPos !== oPr.dataLabelsPos) {
			return false;
		}
		if (this.type !== oPr.type) {
			return false;
		}
		if (!this.equalBool(this.showSerName, oPr.showSerName)) {
			return false;
		}
		if (!this.equalBool(this.showCatName, oPr.showCatName)) {
			return false;
		}
		if (!this.equalBool(this.showVal, oPr.showVal)) {
			return false;
		}

		if (this.separator !== oPr.separator && !(this.separator === ' ' && oPr.separator == null || oPr.separator === ' ' && this.separator == null)) {
			return false;
		}
		if (this.sRange !== oPr.sRange) {
			return false;
		}
		if (!this.equalBool(this.inColumns, oPr.inColumns)) {
			return false;
		}

		if (!this.equalBool(this.showMarker, oPr.showMarker)) {
			return false;
		}
		if (!this.equalBool(this.bLine, oPr.bLine)) {
			return false;
		}
		if (!this.equalBool(this.smooth, oPr.smooth)) {
			return false;
		}
		if (this.verticalAxes.length !== oPr.verticalAxes.length) {
			return false;
		}
		if (this.horizontalAxes.length !== oPr.horizontalAxes.length) {
			return false;
		}
		let nAx;
		for (nAx = 0; nAx < this.verticalAxes.length; ++nAx) {
			if (!this.verticalAxes[nAx].isEqual(oPr.verticalAxes[nAx])) {
				return false;
			}
		}
		for (nAx = 0; nAx < this.horizontalAxes.length; ++nAx) {
			if (!this.horizontalAxes[nAx].isEqual(oPr.horizontalAxes[nAx])) {
				return false;
			}
		}
		if (this.view3D && !oPr.view3D || !this.view3D && oPr.view3D) {
			return false;
		}
		if (this.view3D && oPr.view3D && !this.view3D.isEqual(oPr.view3D)) {
			return false;
		}
		if(this.displayTrendlinesEquation !== oPr.displayTrendlinesEquation) {
			return false;
		}
		return true;
	};
	asc_ChartSettings.prototype.isEmpty = function () {
		return this.isEqual(new asc_ChartSettings());
	};
	asc_ChartSettings.prototype.putShowMarker = function (v) {
		this.showMarker = v;
	};
	asc_ChartSettings.prototype.getShowMarker = function () {
		return this.showMarker;
	};
	asc_ChartSettings.prototype.putLine = function (v) {
		this.bLine = v;
	};
	asc_ChartSettings.prototype.getLine = function () {
		return this.bLine;
	};
	asc_ChartSettings.prototype.putRanges = function (aRanges) {
		if (Array.isArray(aRanges) && aRanges.length > 0) {
			let sRange = "=";
			for (let nRange = 0; nRange < aRanges.length; ++nRange) {
				if (nRange > 0) {
					sRange += ",";
				}
				sRange += aRanges[nRange];
			}
			this.sRange = sRange;
		}
		else {
			this.sRange = null;
		}
	};
	asc_ChartSettings.prototype.getRanges = function () {
		let sRange = this.sRange;
		if (typeof sRange === "string" && sRange > 0) {
			if (sRange.charAt(0) === '=') {
				sRange = sRange.slice(1);
			}
			return sRange.split(",");
		}
		else {
			return [];
		}
	};
	asc_ChartSettings.prototype.putSmooth = function (v) {
		this.smooth = v;
	};
	asc_ChartSettings.prototype.getSmooth = function () {
		return this.smooth;
	};
	asc_ChartSettings.prototype.putStyle = function (index) {
		if (!AscFormat.isRealNumber(index)) {
			this.style = null;
			return;
		}
		this.style = parseInt(index, 10);
		if (this.bStartEdit && this.chartSpace) {
			if (AscFormat.isRealNumber(this.style)) {
				let aStyle = AscCommon.g_oChartStyles[this.type] && AscCommon.g_oChartStyles[this.type][this.style - 1];
				if (Array.isArray(aStyle)) {
					this.chartSpace.applyChartStyleByIds(aStyle);
					this.updateChart();
				}
			}
		}
	};
	asc_ChartSettings.prototype.getStyle = function () {
		return this.style;
	};
	asc_ChartSettings.prototype.putRange = function (range) {
		this.sRange = range;
	};
	asc_ChartSettings.prototype.setRange = function (sRange) {
		if (this.chartSpace) {
			const oDataRefs = new AscFormat.CChartDataRefs(null);
			const nCheckResult = oDataRefs.checkDataRange(sRange, this.getInRows(), this.getType());
			if (nCheckResult !== Asc.c_oAscError.ID.No) {
				this.sendError(nCheckResult);
				return;
			}
			this.chartSpace.setRange(sRange);
			this.updateChart();
		}
	};
	asc_ChartSettings.prototype.isValidRange = function (sRange) {
		if (this.getRange() !== sRange) {
			const oDataRefs = new AscFormat.CChartDataRefs(null);
			const nCheckResult = oDataRefs.checkDataRange(sRange, this.getInRows(), this.getType());
			if (nCheckResult === Asc.c_oAscError.ID.MaxDataPointsError) {
				return nCheckResult;
			}
		}
		return AscFormat.isValidChartRange(sRange);
	};
	asc_ChartSettings.prototype.getRange = function () {
		if (this.chartSpace) {
			return this.chartSpace.getCommonRange();
		}
		return this.sRange;
	};
	asc_ChartSettings.prototype.putInColumns = function (inColumns) {
		this.inColumns = inColumns;
	};
	asc_ChartSettings.prototype.getInColumns = function () {
		return this.inColumns;
	};
	asc_ChartSettings.prototype.getInRows = function () {
		if (this.inColumns === true || this.inColumns === false) {
			return !this.inColumns;
		}
		return null;
	};
	asc_ChartSettings.prototype.putTitle = function (v) {
		this.title = v;
	};
	asc_ChartSettings.prototype.getTitle = function () {
		return this.title;
	};
	asc_ChartSettings.prototype.putRowCols = function (v) {
		this.rowCols = v;
	};
	asc_ChartSettings.prototype.getRowCols = function () {
		return this.rowCols;
	};
	asc_ChartSettings.prototype.putLegendPos = function (v) {
		this.legendPos = v;
	};
	asc_ChartSettings.prototype.putDataLabelsPos = function (v) {
		this.dataLabelsPos = v;
	};
	asc_ChartSettings.prototype.getLegendPos = function () {
		return this.legendPos;
	};
	asc_ChartSettings.prototype.getDataLabelsPos = function () {
		return this.dataLabelsPos;
	};
	asc_ChartSettings.prototype.getType = function () {
		if (this.chartSpace) {
			return this.chartSpace.getChartType();
		}
		return this.type;
	};
	asc_ChartSettings.prototype.checkParams = function () {
		if (this.type === null || this.type === Asc.c_oAscChartTypeSettings.comboAreaBar || this.type === Asc.c_oAscChartTypeSettings.comboBarLine || this.type === Asc.c_oAscChartTypeSettings.comboBarLineSecondary || this.type === Asc.c_oAscChartTypeSettings.comboCustom) {
			return;
		}
		if (AscFormat.getIsMarkerByType(this.type)) {
			this.showMarker = true;
		}
		else {
			this.showMarker = false;
		}
		if (AscFormat.getIsSmoothByType(this.type)) {
			this.smooth = true;
		}
		else {
			this.smooth = false;
		}
		if (AscFormat.getIsLineByType(this.type)) {
			this.bLine = true;
		}
		else {
			this.bLine = false;
		}
	};
	asc_ChartSettings.prototype.putType = function (v) {
		this.type = v;
		this.checkParams();
	};
	asc_ChartSettings.prototype.putShowSerName = function (v) {
		this.showSerName = v;
	};
	asc_ChartSettings.prototype.putShowCatName = function (v) {
		this.showCatName = v;
	};
	asc_ChartSettings.prototype.putShowVal = function (v) {
		this.showVal = v;
	};
	asc_ChartSettings.prototype.getShowSerName = function () {
		return this.showSerName;
	};
	asc_ChartSettings.prototype.getShowCatName = function () {
		return this.showCatName;
	};
	asc_ChartSettings.prototype.getShowVal = function () {
		return this.showVal;
	};
	asc_ChartSettings.prototype.putSeparator = function (v) {
		this.separator = v;
	};
	asc_ChartSettings.prototype.getSeparator = function () {
		return this.separator;
	};
	asc_ChartSettings.prototype.sendErrorOnChangeType = function (nType) {
		this.sendError(nType);
	};
	asc_ChartSettings.prototype.sendError = function (nType) {
		let oApi = Asc.editor || editor;
		if (oApi) {
			oApi.sendEvent("asc_onError", nType, Asc.c_oAscError.Level.NoCritical);
			if (oApi.UpdateInterfaceState) {
				oApi.UpdateInterfaceState();
			}
		}
	};
	asc_ChartSettings.prototype.changeType = function (type) {
		if (this.chartSpace) {
			if (type === Asc.c_oAscChartTypeSettings.stock) {
				if (!this.chartSpace.canChangeToStockChart()) {
					this.sendErrorOnChangeType(Asc.c_oAscError.ID.StockChartError);
					return false;
				}
			}
			if (type === Asc.c_oAscChartTypeSettings.comboCustom || type === Asc.c_oAscChartTypeSettings.comboAreaBar || type === Asc.c_oAscChartTypeSettings.comboBarLine || type === Asc.c_oAscChartTypeSettings.comboBarLineSecondary) {
				if (!this.chartSpace.canChangeToComboChart()) {
					this.sendErrorOnChangeType(Asc.c_oAscError.ID.ComboSeriesError);
					return false;
				}
			}
			this.putType(type);
			if (this.chartSpace) {
				let oController = this.chartSpace.getDrawingObjectsController();
				if (oController) {
					let oThis = this;
					let oChartSpace = this.chartSpace;
					oController.checkSelectedObjectsAndCallback(function () {
						oChartSpace.changeChartType(type);
						oThis.updateChart();
						let oApi = Asc.editor || editor;
						if (oApi) {
							if (oApi.UpdateInterfaceState) {
								oApi.UpdateInterfaceState();
							}
						}
					}, [], false, 0, []);
				}
			}
		}
		else {
			this.putType(type);
		}
		return true;
	};
	asc_ChartSettings.prototype.getSeries = function () {
		if (this.chartSpace) {
			return this.chartSpace.getAllSeries();
		}
		return [];
	};
	asc_ChartSettings.prototype.getCatValues = function () {
		if (this.chartSpace) {
			return this.chartSpace.getCatValues();
		}
		return [];
	};
	asc_ChartSettings.prototype.getCatFormula = function () {
		if (this.chartSpace) {
			return this.chartSpace.getCatFormula();
		}
		return "";
	};
	asc_ChartSettings.prototype.setCatFormula = function (sFormula) {
		if (this.chartSpace) {
			return this.chartSpace.setCatFormula(sFormula);
		}
		this.updateChart();
	};
	asc_ChartSettings.prototype.isValidCatFormula = function (sFormula) {
		if (sFormula === "" || sFormula === null) {
			return Asc.c_oAscError.ID.No;
		}
		return AscFormat.ExecuteNoHistory(function () {
			let oCat = new AscFormat.CCat();
			return oCat.setValues(sFormula).getError();
		}, this, []);
	};
	asc_ChartSettings.prototype.switchRowCol = function () {
		let nError = Asc.c_oAscError.ID.No;
		if (this.chartSpace) {
			nError = this.chartSpace.switchRowCol();
		}
		this.updateChart();
		return nError;
	};
	asc_ChartSettings.prototype.addSeries = function () {
		let oRet = null;
		if (this.chartSpace) {
			oRet = this.chartSpace.addNewSeries();
		}
		this.updateChart();
		return oRet;
	};
	asc_ChartSettings.prototype.addScatterSeries = function () {
		let oRet = null;
		if (this.chartSpace) {
			oRet = this.chartSpace.addNewSeries();
		}
		this.updateChart();
		return oRet;
	};
	asc_ChartSettings.prototype.startEdit = function () {
		this.bStartEdit = true;
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
	};
	asc_ChartSettings.prototype.updateInterface = function () {
		let oApi = Asc.editor || editor;
		if (oApi) {
			if (oApi.UpdateInterfaceState) {
				oApi.UpdateInterfaceState();
			}
			else {
				let oWbView = oApi.wb;
				if (oWbView) {
					let oWSView = oWbView.getWorksheet();
					if (oWSView) {
						let oRender = oWSView.objectRender;
						if (oRender) {
							oRender.sendGraphicObjectProps();
						}
					}
				}
			}
		}
	};
	asc_ChartSettings.prototype.endEdit = function () {
		if (AscCommon.History.Is_LastPointEmpty()) {
			this.cancelEdit();
			return;
		}
		this.bStartEdit = false;
		AscCommon.History.EndTransaction();
		this.updateChart();
		this.updateInterface();
	};
	asc_ChartSettings.prototype.cancelEdit = function () {
		this.bStartEdit = false;
		const bLastPointEmpty = AscCommon.History.Is_LastPointEmpty();
		AscCommon.History.EndTransaction();
		if (!bLastPointEmpty) {
			AscCommon.History.Undo();
			AscCommon.History.Clear_Redo();
		}
		AscCommon.History._sendCanUndoRedo();
		this.updateChart();
		this.updateInterface();
	};
	asc_ChartSettings.prototype.startEditData = function () {
		AscCommon.History.SavePointIndex();
	};
	asc_ChartSettings.prototype.cancelEditData = function () {
		AscCommon.History.UndoToPointIndex();
		this.updateChart();
	};
	asc_ChartSettings.prototype.endEditData = function () {
		AscCommon.History.ClearPointIndex();
		this.updateChart();
	};
	asc_ChartSettings.prototype.updateChart = function () {
		if (this.chartSpace) {
			this.chartSpace.onDataUpdate();
		}
	};
	asc_ChartSettings.prototype.getDisplayTrendlinesEquation = function() {
		return this.displayTrendlinesEquation;
	};
	asc_ChartSettings.prototype.putDisplayTrendlinesEquation = function(v) {
		this.displayTrendlinesEquation = v;
	};

	/** @constructor */
	function asc_CRect(x, y, width, height) {
		// private members
		this._x = x;
		this._y = y;
		this._width = width;
		this._height = height;
	}

	asc_CRect.prototype.asc_getX = function () {
		return this._x;
	};
	asc_CRect.prototype.asc_getY = function () {
		return this._y;
	};
	asc_CRect.prototype.asc_getWidth = function () {
		return this._width;
	};
	asc_CRect.prototype.asc_getHeight = function () {
		return this._height;
	};


	const STANDART_COLORS_MAP = {};
	STANDART_COLORS_MAP[0x000000] = "Black";
	STANDART_COLORS_MAP[0xFFFFFF] = "White";
	STANDART_COLORS_MAP[0xFF0000] = "Red";
	STANDART_COLORS_MAP[0x00FF00] = "Green";
	STANDART_COLORS_MAP[0x0000FF] = "Blue";
	STANDART_COLORS_MAP[0xFFFF00] = "Yellow";
	STANDART_COLORS_MAP[0xFF00FF] = "Purple";
	STANDART_COLORS_MAP[0x00FFFF] = "Aqua";
	STANDART_COLORS_MAP[0x800000] = "Dark Red";
	STANDART_COLORS_MAP[0x008000] = "Dark Green";
	STANDART_COLORS_MAP[0x000080] = "Dark Blue";
	STANDART_COLORS_MAP[0x808000] = "Dark Yellow";
	STANDART_COLORS_MAP[0x800080] = "Dark Purple";
	STANDART_COLORS_MAP[0x008080] = "Dark Teal";
	STANDART_COLORS_MAP[0xC0C0C0] = "Light Gray";
	STANDART_COLORS_MAP[0x808080] = "Gray";
	STANDART_COLORS_MAP[0x9999FF] = "Light Blue";
	STANDART_COLORS_MAP[0x993366] = "Pink";
	STANDART_COLORS_MAP[0xFFFFCC] = "Light Yellow";
	STANDART_COLORS_MAP[0xCCFFFF] = "Sky Blue";
	STANDART_COLORS_MAP[0x660066] = "Dark Purple";
	STANDART_COLORS_MAP[0xFF8080] = "Rose";
	STANDART_COLORS_MAP[0x0066CC] = "Blue";
	STANDART_COLORS_MAP[0xCCCCFF] = "Light Blue";
	STANDART_COLORS_MAP[0x00CCFF] = "Turquosie";
	STANDART_COLORS_MAP[0xCCFFCC] = "Light Green";
	STANDART_COLORS_MAP[0xFFFF99] = "Light Yellow";
	STANDART_COLORS_MAP[0x99CCFF] = "Light Blue";
	STANDART_COLORS_MAP[0xFF99CC] = "Pink";
	STANDART_COLORS_MAP[0xCC99FF] = "Lavender";
	STANDART_COLORS_MAP[0xFFCC99] = "Light Orange";
	STANDART_COLORS_MAP[0x3366FF] = "Blue";
	STANDART_COLORS_MAP[0x33CCCC] = "Teal";
	STANDART_COLORS_MAP[0x99CC00] = "Green";
	STANDART_COLORS_MAP[0xFFCC00] = "Gold";
	STANDART_COLORS_MAP[0xFF9900] = "Orange";
	STANDART_COLORS_MAP[0xFF6600] = "Orange";
	STANDART_COLORS_MAP[0x666699] = "Indigo";
	STANDART_COLORS_MAP[0x969696] = "Gray";
	STANDART_COLORS_MAP[0x003366] = "Dark Blue";
	STANDART_COLORS_MAP[0x339966] = "Green";
	STANDART_COLORS_MAP[0x003300] = "Dark Green";
	STANDART_COLORS_MAP[0x333300] = "Dark Yellow";
	STANDART_COLORS_MAP[0x993300] = "Brown";
	STANDART_COLORS_MAP[0x333399] = "Indigo";
	STANDART_COLORS_MAP[0x333333] = "Dark Gray";


	/**
	 * Класс CColor для работы с цветами
	 * -----------------------------------------------------------------------------
	 *
	 * @constructor
	 * @memberOf window
	 */

	function CColor(r, g, b, a) {
		this.r = (undefined == r) ? 0 : r;
		this.g = (undefined == g) ? 0 : g;
		this.b = (undefined == b) ? 0 : b;
		this.a = (undefined == a) ? 1 : a;
	}

	CColor.prototype.constructor = CColor;
	CColor.prototype.getR = function () {
		return this.r
	};
	CColor.prototype.get_r = function () {
		return this.r
	};
	CColor.prototype.put_r = function (v) {
		this.r = v;
		this.hex = undefined;
	};
	CColor.prototype.getG = function () {
		return this.g
	};
	CColor.prototype.get_g = function () {
		return this.g;
	};
	CColor.prototype.put_g = function (v) {
		this.g = v;
		this.hex = undefined;
	};
	CColor.prototype.getB = function () {
		return this.b
	};
	CColor.prototype.get_b = function () {
		return this.b;
	};
	CColor.prototype.put_b = function (v) {
		this.b = v;
		this.hex = undefined;
	};
	CColor.prototype.getA = function () {
		return this.a;
	};
	CColor.prototype.get_hex = function () {
		if (!this.hex) {
			let r = this.r.toString(16);
			let g = this.g.toString(16);
			let b = this.b.toString(16);
			this.hex = (r.length == 1 ? "0" + r : r) + (g.length == 1 ? "0" + g : g) + (b.length == 1 ? "0" + b : b);
		}
		return this.hex;
	};

	CColor.prototype.Compare = function (Color) {
		return (this.r === Color.r && this.g === Color.g && this.b === Color.b && this.a === Color.a);
	};
	CColor.prototype.Copy = function () {
		return new CColor(this.r, this.g, this.b, this.a);
	};

	CColor.prototype.getVal = function () {
		return (((this.r << 16) & 0xFF0000) + ((this.g << 8) & 0xFF00) + this.b);
	};

	CColor.prototype.getColorName = function () {
		return (new asc_CColor(this.r, this.g, this.b)).asc_getName();
	};

	CColor.prototype.getAscColor = function () {
		return (new asc_CColor(this.r, this.g, this.b));
	};
	CColor.prototype.fromAscColor = function (oAscColor) {
		return new CColor(oAscColor.r, oAscColor.g, oAscColor.b);
	};
	CColor.prototype.isEqual = function (r, g, b, a) {
		return this.r === r && this.g === g && this.b === b && this.a === a;
	};

	/** @constructor */
	function asc_CColor() {
		this.type = c_oAscColor.COLOR_TYPE_SRGB;
		this.value = null;
		this.r = 0;
		this.g = 0;
		this.b = 0;
		this.a = 255;

		this.Auto = false;

		this.Mods = [];
		this.ColorSchemeId = -1;
		this.EffectValue = 0;
		if (1 === arguments.length) {
			this.r = arguments[0].r;
			this.g = arguments[0].g;
			this.b = arguments[0].b;
		}
		else {
			if (3 <= arguments.length) {
				this.r = arguments[0];
				this.g = arguments[1];
				this.b = arguments[2];
			}
			if (4 === arguments.length) {
				this.a = arguments[3];
			}
		}
	}

	asc_CColor.prototype.constructor = asc_CColor;
	asc_CColor.prototype.asc_getR = function () {
		return this.r
	};
	asc_CColor.prototype.asc_putR = function (v) {
		this.r = v;
		this.hex = undefined;
	};
	asc_CColor.prototype.asc_getG = function () {
		return this.g;
	};
	asc_CColor.prototype.asc_putG = function (v) {
		this.g = v;
		this.hex = undefined;
	};
	asc_CColor.prototype.asc_getB = function () {
		return this.b;
	};
	asc_CColor.prototype.asc_putB = function (v) {
		this.b = v;
		this.hex = undefined;
	};
	asc_CColor.prototype.asc_getA = function () {
		return this.a;
	};
	asc_CColor.prototype.asc_putA = function (v) {
		this.a = v;
		this.hex = undefined;
	};
	asc_CColor.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CColor.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CColor.prototype.asc_getValue = function () {
		return this.value;
	};
	asc_CColor.prototype.asc_putValue = function (v) {
		this.value = v;
	};
	asc_CColor.prototype.asc_getHex = function () {
		if (!this.hex) {
			let a = this.a.toString(16);
			let r = this.r.toString(16);
			let g = this.g.toString(16);
			let b = this.b.toString(16);
			this.hex = (a.length == 1 ? "0" + a : a) + (r.length == 1 ? "0" + r : r) + (g.length == 1 ? "0" + g : g) + (b.length == 1 ? "0" + b : b);
		}
		return this.hex;
	};
	asc_CColor.prototype.asc_getColor = function () {
		return new CColor(this.r, this.g, this.b);
	};
	asc_CColor.prototype.asc_putAuto = function (v) {
		this.Auto = v;
	};
	asc_CColor.prototype.asc_getAuto = function () {
		return this.Auto;
	};
	asc_CColor.prototype.getColorDiff = function (nC1, nC2) {
		let nC1R = (nC1 >> 16) & 0xFF;
		let nC1G = (nC1 >> 8) & 0xFF;
		let nC1B = nC1 & 0xFF;
		let nC2R = (nC2 >> 16) & 0xFF;
		let nC2G = (nC2 >> 8) & 0xFF;
		let nC2B = nC2 & 0xFF;
		let lab1 = this.RGB2LAB(nC1R, nC1G, nC1B);
		let lab2 = this.RGB2LAB(nC2R, nC2G, nC2B);

		const d2r = AscCommon.deg2rad;

		const L1 = lab1[0];
		const a1 = lab1[1];
		const b1 = lab1[2];

		const L2 = lab2[0];
		const a2 = lab2[1];
		const b2 = lab2[2];

		const k_L = 1.0, k_C = 1.0, k_H = 1.0;
		const deg360InRad = d2r(360.0);
		const deg180InRad = d2r(180.0);
		const pow25To7 = 6103515625.0; /* Math.pow(25, 7) */

		let C1 = Math.sqrt((a1 * a1) + (b1 * b1));
		let C2 = Math.sqrt((a2 * a2) + (b2 * b2));
		let barC = (C1 + C2) / 2.0;
		let G = 0.5 * (1 - Math.sqrt(Math.pow(barC, 7) / (Math.pow(barC, 7) + pow25To7)));
		let a1Prime = (1.0 + G) * a1;
		let a2Prime = (1.0 + G) * a2;
		let CPrime1 = Math.sqrt((a1Prime * a1Prime) + (b1 * b1));
		let CPrime2 = Math.sqrt((a2Prime * a2Prime) + (b2 * b2));
		let hPrime1;
		const fAE = function (a, b) {
			return Math.abs(a - b) < 1e-15;
		};
		if (fAE(b1, 0.0) && fAE(a1Prime, 0.0)) hPrime1 = 0.0; else {
			hPrime1 = Math.atan2(b1, a1Prime);
			if (hPrime1 < 0) hPrime1 += deg360InRad;
		}
		let hPrime2;
		if (fAE(b2, 0.0) && fAE(a2Prime, 0.0)) hPrime2 = 0.0; else {
			hPrime2 = Math.atan2(b2, a2Prime);
			if (hPrime2 < 0) hPrime2 += deg360InRad;
		}

		let deltaLPrime = L2 - L1;
		let deltaCPrime = CPrime2 - CPrime1;
		let deltahPrime;
		let CPrimeProduct = CPrime1 * CPrime2;
		if (fAE(CPrimeProduct, 0.0)) deltahPrime = 0; else {
			deltahPrime = hPrime2 - hPrime1;
			if (deltahPrime < -deg180InRad) deltahPrime += deg360InRad; else if (deltahPrime > deg180InRad) deltahPrime -= deg360InRad;
		}
		let deltaHPrime = 2.0 * Math.sqrt(CPrimeProduct) * Math.sin(deltahPrime / 2.0);

		let barLPrime = (L1 + L2) / 2.0;
		let barCPrime = (CPrime1 + CPrime2) / 2.0;
		let barhPrime, hPrimeSum = hPrime1 + hPrime2;
		if (fAE(CPrime1 * CPrime2, 0.0)) {
			barhPrime = hPrimeSum;
		}
		else {
			if (Math.abs(hPrime1 - hPrime2) <= deg180InRad) barhPrime = hPrimeSum / 2.0; else {
				if (hPrimeSum < deg360InRad) barhPrime = (hPrimeSum + deg360InRad) / 2.0; else barhPrime = (hPrimeSum - deg360InRad) / 2.0;
			}
		}
		let T = 1.0 - (0.17 * Math.cos(barhPrime - d2r(30.0))) + (0.24 * Math.cos(2.0 * barhPrime)) + (0.32 * Math.cos((3.0 * barhPrime) + d2r(6.0))) - (0.20 * Math.cos((4.0 * barhPrime) - d2r(63.0)));
		let deltaTheta = d2r(30.0) * Math.exp(-Math.pow((barhPrime - d2r(275.0)) / d2r(25.0), 2.0));
		let R_C = 2.0 * Math.sqrt(Math.pow(barCPrime, 7.0) / (Math.pow(barCPrime, 7.0) + pow25To7));
		let S_L = 1 + ((0.015 * Math.pow(barLPrime - 50.0, 2.0)) / Math.sqrt(20 + Math.pow(barLPrime - 50.0, 2.0)));
		let S_C = 1 + (0.045 * barCPrime);
		let S_H = 1 + (0.015 * barCPrime * T);
		let R_T = (-Math.sin(2.0 * deltaTheta)) * R_C;

		let deltaE = Math.sqrt(Math.pow(deltaLPrime / (k_L * S_L), 2.0) + Math.pow(deltaCPrime / (k_C * S_C), 2.0) + Math.pow(deltaHPrime / (k_H * S_H), 2.0) + (R_T * (deltaCPrime / (k_C * S_C)) * (deltaHPrime / (k_H * S_H))));

		return deltaE;
	};
	asc_CColor.prototype.RGB2LAB = function (R, G, B) {
		let r, g, b, X, Y, Z, fx, fy, fz, xr, yr, zr;
		let Ls, as, bs;
		let eps = 216.0 / 24389.0;
		let k = 24389.0 / 27.0;

		let Xr = 0.964221;  // reference white D50
		let Yr = 1.0;
		let Zr = 0.825211;

		// RGB to XYZ
		r = R / 255; //R 0..1
		g = G / 255; //G 0..1
		b = B / 255; //B 0..1

		// assuming sRGB (D65)
		if (r <= 0.04045) r = r / 12; else r = Math.pow((r + 0.055) / 1.055, 2.4);

		if (g <= 0.04045) g = g / 12; else g = Math.pow((g + 0.055) / 1.055, 2.4);

		if (b <= 0.04045) b = b / 12; else b = Math.pow((b + 0.055) / 1.055, 2.4);


		X = 0.436052025 * r + 0.385081593 * g + 0.143087414 * b;
		Y = 0.222491598 * r + 0.71688606 * g + 0.060621486 * b;
		Z = 0.013929122 * r + 0.097097002 * g + 0.71418547 * b;

		// XYZ to Lab
		xr = X / Xr;
		yr = Y / Yr;
		zr = Z / Zr;

		if (xr > eps) fx = Math.pow(xr, 1 / 3.); else fx = ((k * xr + 16.) / 116.);

		if (yr > eps) fy = Math.pow(yr, 1 / 3.); else fy = ((k * yr + 16.) / 116.);

		if (zr > eps) fz = Math.pow(zr, 1 / 3.); else fz = ((k * zr + 16.) / 116);

		Ls = (116 * fy) - 16;
		as = 500 * (fx - fy);
		bs = 200 * (fy - fz);

		let lab = [];
		lab[0] = (2.55 * Ls + .5) >> 0;
		lab[1] = (as + .5) >> 0;
		lab[2] = (bs + .5) >> 0;
		return lab;
	};
	asc_CColor.prototype.asc_getName = function () {
		const nColorVal = this.getVal();
		if (STANDART_COLORS_MAP.hasOwnProperty(nColorVal)) {
			return STANDART_COLORS_MAP[nColorVal];
		}
		let dMinDistance = 1000000;
		let sMinName = "Black";
		for (let nCurColor in STANDART_COLORS_MAP) {
			if (STANDART_COLORS_MAP.hasOwnProperty(nCurColor)) {
				let dDist = this.getColorDiff(nColorVal, nCurColor);
				if (dDist < dMinDistance) {
					dMinDistance = dDist;
					sMinName = STANDART_COLORS_MAP[nCurColor];
				}
			}
		}
		return sMinName;
	};
	asc_CColor.prototype.getVal = function () {
		return (((this.r << 16) & 0xFF0000) + ((this.g << 8) & 0xFF00) + this.b);
	};
	asc_CColor.prototype.asc_putEffectValue = function (v) {
		let dVal = Math.abs(v);
		dVal = ((dVal * 100 + 0.5) >> 0) / 100;
		if (v < 0) {
			dVal = -dVal;
		}
		this.EffectValue = dVal;
	};
	asc_CColor.prototype.asc_getEffectValue = function () {
		return this.EffectValue;
	};
	asc_CColor.prototype.print = function () {
		console.log("Color");
		console.log("r: " + this.r);
		console.log("g: " + this.g);
		console.log("b: " + this.b);
		console.log("effect val: " + this.asc_getEffectValue());
		console.log("name: " + this.asc_getName());
		console.log("name in scheme: " + this.asc_getNameInColorScheme());
		console.log("---------------");
	};
	asc_CColor.prototype.asc_getNameInColorScheme = function () {
		if (this.ColorSchemeId === -1) {
			return null;
		}
		switch (this.ColorSchemeId) {
			// bg1,tx1,bg2,tx2,accent1 - accent6
			case 6: {
				return "background 1";
			}
			case 15: {
				return "text 1";
			}
			case 7: {
				return "background 2";
			}
			case 16: {
				return "text 2";
			}
			case 0: {
				return "accent 1";
			}
			case 1: {
				return "accent 2";
			}
			case 2: {
				return "accent 3";
			}
			case 3: {
				return "accent 4";
			}
			case 4: {
				return "accent 5";
			}
			case 5: {
				return "accent 6";
			}
		}
		return null;
	};
	asc_CColor.prototype.setColorSchemeId = function (v) {
		this.ColorSchemeId = v;
		if (!AscFormat.isRealNumber(this.ColorSchemeId)) {
			this.ColorSchemeId = -1;
		}
	};

	/** @constructor */
	function asc_CTextBorder(obj) {
		if (obj) {
			if (obj.Color instanceof asc_CColor) {
				this.Color = obj.Color;
			}
			else {
				this.Color = (undefined != obj.Color && null != obj.Color) ? CreateAscColorCustom(obj.Color.r, obj.Color.g, obj.Color.b) : null;
			}
			this.Size = (undefined != obj.Size) ? obj.Size : null;
			this.Value = (undefined != obj.Value) ? obj.Value : null;
			this.Space = (undefined != obj.Space) ? obj.Space : null;
		}
		else {
			this.Color = CreateAscColorCustom(0, 0, 0);
			this.Size = 0.5 * window["AscCommonWord"].g_dKoef_pt_to_mm;
			this.Value = window["AscCommonWord"].border_Single;
			this.Space = 0;
		}
	}

	asc_CTextBorder.prototype.asc_getColor = function () {
		return this.Color;
	};
	asc_CTextBorder.prototype.asc_putColor = function (v) {
		this.Color = v;
	};
	asc_CTextBorder.prototype.asc_getSize = function () {
		return this.Size;
	};
	asc_CTextBorder.prototype.asc_putSize = function (v) {
		this.Size = v;
	};
	asc_CTextBorder.prototype.asc_getValue = function () {
		return this.Value;
	};
	asc_CTextBorder.prototype.asc_putValue = function (v) {
		this.Value = v;
	};
	asc_CTextBorder.prototype.asc_getSpace = function () {
		return this.Space;
	};
	asc_CTextBorder.prototype.asc_putSpace = function (v) {
		this.Space = v;
	};
	asc_CTextBorder.prototype.asc_getForSelectedCells = function () {
		return this.ForSelectedCells;
	};
	asc_CTextBorder.prototype.asc_putForSelectedCells = function (v) {
		this.ForSelectedCells = v;
	};

	/** @constructor */
	function asc_CParagraphBorders(obj) {

		if (obj) {
			this.Left = (undefined != obj.Left && null != obj.Left) ? new asc_CTextBorder(obj.Left) : null;
			this.Top = (undefined != obj.Top && null != obj.Top) ? new asc_CTextBorder(obj.Top) : null;
			this.Right = (undefined != obj.Right && null != obj.Right) ? new asc_CTextBorder(obj.Right) : null;
			this.Bottom = (undefined != obj.Bottom && null != obj.Bottom) ? new asc_CTextBorder(obj.Bottom) : null;
			this.Between = (undefined != obj.Between && null != obj.Between) ? new asc_CTextBorder(obj.Between) : null;
		}
		else {
			this.Left = null;
			this.Top = null;
			this.Right = null;
			this.Bottom = null;
			this.Between = null;
		}
	}

	asc_CParagraphBorders.prototype.asc_getLeft = function () {
		return this.Left;
	};
	asc_CParagraphBorders.prototype.asc_putLeft = function (v) {
		this.Left = (v) ? new asc_CTextBorder(v) : null;
	};
	asc_CParagraphBorders.prototype.asc_getTop = function () {
		return this.Top;
	};
	asc_CParagraphBorders.prototype.asc_putTop = function (v) {
		this.Top = (v) ? new asc_CTextBorder(v) : null;
	};
	asc_CParagraphBorders.prototype.asc_getRight = function () {
		return this.Right;
	};
	asc_CParagraphBorders.prototype.asc_putRight = function (v) {
		this.Right = (v) ? new asc_CTextBorder(v) : null;
	};
	asc_CParagraphBorders.prototype.asc_getBottom = function () {
		return this.Bottom;
	};
	asc_CParagraphBorders.prototype.asc_putBottom = function (v) {
		this.Bottom = (v) ? new asc_CTextBorder(v) : null;
	};
	asc_CParagraphBorders.prototype.asc_getBetween = function () {
		return this.Between;
	};
	asc_CParagraphBorders.prototype.asc_putBetween = function (v) {
		this.Between = (v) ? new asc_CTextBorder(v) : null;
	};

	/** @constructor */
	function asc_CCustomListType(obj) {
		this.type = null;

		this.imageId = null;
		this.token = null;

		this.char = null;
		this.specialFont = null;

		this.numberingType = null;
		if (obj) {
			this.fillFromObject(obj);
		}
	}

	asc_CCustomListType.prototype.fillFromObject = function (obj) {
		if (obj.type) this.type = obj.type;
		if (obj.imageId) this.imageId = obj.imageId;
		if (obj.token) this.token = obj.token;
		if (obj.char) this.char = obj.char;
		if (obj.specialFont) this.specialFont = obj.specialFont;
		if (obj.numberingType) this.numberingType = obj.numberingType;
	};
	asc_CCustomListType.prototype.setType = function (pr) {
		this.type = pr;
	};
	asc_CCustomListType.prototype.setImageId = function (pr) {
		this.imageId = pr;
	};
	asc_CCustomListType.prototype.setToken = function (pr) {
		this.token = pr;
	};
	asc_CCustomListType.prototype.setChar = function (pr) {
		this.char = pr;
	};
	asc_CCustomListType.prototype.setSpecialFont = function (pr) {
		this.specialFont = pr;
	};
	asc_CCustomListType.prototype.setNumberingType = function (pr) {
		this.numberingType = pr;
	};
	asc_CCustomListType.prototype.getType = function () {
		return this.type;
	};
	asc_CCustomListType.prototype.getImageId = function () {
		return this.imageId;
	};
	asc_CCustomListType.prototype.getToken = function () {
		return this.token;
	};
	asc_CCustomListType.prototype.getChar = function () {
		return this.char;
	};
	asc_CCustomListType.prototype.getSpecialFont = function () {
		return this.specialFont;
	};
	asc_CCustomListType.prototype.getNumberingType = function () {
		return this.numberingType;
	};

	/** @constructor */
	function asc_CListType(obj) {

		if (obj) {
			this.Type = (undefined == obj.Type) ? null : obj.Type;
			this.SubType = (undefined == obj.Type) ? null : obj.SubType;
			this.Custom = (undefined == obj.Type) ? null : new asc_CCustomListType(obj.Custom);
		}
		else {
			this.Type = null;
			this.SubType = null;
			this.Custom = null;
		}
	}

	asc_CListType.prototype.asc_getListType = function () {
		return this.Type;
	};
	asc_CListType.prototype.asc_getListSubType = function () {
		return this.SubType;
	};

	asc_CListType.prototype.asc_getListCustom = function () {
		return this.Custom;
	};

	/** @constructor */
	function asc_CTextFontFamily(obj) {

		if (obj) {
			this.Name = (undefined != obj.Name) ? obj.Name : null; 		// "Times New Roman"
			this.Index = (undefined != obj.Index) ? obj.Index : null;	// -1
		}
		else {
			this.Name = "Times New Roman";
			this.Index = -1;
		}
	}

	asc_CTextFontFamily.prototype.asc_getName = function () {
		let _name = AscFonts.g_fontApplication ? AscFonts.g_fontApplication.NameToInterface[this.Name] : null;
		return _name ? _name : this.Name;
	};
	asc_CTextFontFamily.prototype.asc_getIndex = function () {
		return this.Index;
	};
	asc_CTextFontFamily.prototype.asc_putName = function (v) {
		this.Name = v;
	};
	asc_CTextFontFamily.prototype.asc_putIndex = function (v) {
		this.Index = v;
	};

	/** @constructor */
	function asc_CParagraphTab(Pos, Value, Leader) {
		this.Pos = Pos;
		this.Value = Value;
		this.Leader = Leader;
	}

	asc_CParagraphTab.prototype.asc_getValue = function () {
		return this.Value;
	};
	asc_CParagraphTab.prototype.asc_putValue = function (v) {
		this.Value = v;
	};
	asc_CParagraphTab.prototype.asc_getPos = function () {
		return this.Pos;
	};
	asc_CParagraphTab.prototype.asc_putPos = function (v) {
		this.Pos = v;
	};
	asc_CParagraphTab.prototype.asc_getLeader = function () {
		if (Asc.c_oAscTabLeader.Heavy === this.Leader) return Asc.c_oAscTabLeader.Underscore;

		return this.Leader;
	};
	asc_CParagraphTab.prototype.asc_putLeader = function (v) {
		this.Leader = v;
	};

	/** @constructor */
	function asc_CParagraphTabs(obj) {
		this.Tabs = [];

		if (undefined != obj) {
			let Count = obj.Tabs.length;
			for (let Index = 0; Index < Count; Index++) {
				this.Tabs.push(new asc_CParagraphTab(obj.Tabs[Index].Pos, obj.Tabs[Index].Value, obj.Tabs[Index].Leader));
			}
		}
	}

	asc_CParagraphTabs.prototype.asc_getCount = function () {
		return this.Tabs.length;
	};
	asc_CParagraphTabs.prototype.asc_getTab = function (Index) {
		return this.Tabs[Index];
	};
	asc_CParagraphTabs.prototype.asc_addTab = function (Tab) {
		this.Tabs.push(Tab)
	};
	asc_CParagraphTabs.prototype.asc_clear = function () {
		this.Tabs.length = 0;
	};

	/** @constructor */
	function asc_CParagraphShd(obj) {

		if (obj) {
			this.Value = (undefined != obj.Value) ? obj.Value : null;

			// TODO: В UI пока поддерживается ровно два типа заливки Nil, Clear
			if (null !== this.Value && this.Value !== Asc.c_oAscShd.Nil) this.Value = Asc.c_oAscShd.Clear;

			if (obj.GetSimpleColor) {

				if (Asc.c_oAscShd.Clear === obj.Value && obj.Unifill && obj.Unifill.fill && obj.Unifill.fill.type === c_oAscFill.FILL_TYPE_SOLID && obj.Unifill.fill.color) {
					this.Color = CreateAscColor(obj.Unifill.fill.color);
				}
				else {
					let oColor = obj.GetSimpleColor();
					if (oColor.Auto) this.Color = null; else this.Color = CreateAscColorCustom(oColor.r, oColor.g, oColor.b, oColor.Auto);
				}
			}
			else {
				if (obj.Unifill && obj.Unifill.fill && obj.Unifill.fill.type === c_oAscFill.FILL_TYPE_SOLID && obj.Unifill.fill.color) {
					this.Color = CreateAscColor(obj.Unifill.fill.color);
				}
				else {
					this.Color = (undefined != obj.Color && null != obj.Color) ? CreateAscColorCustom(obj.Color.r, obj.Color.g, obj.Color.b) : null;
				}
			}
		}
		else {

			// TODO: Пока мы не работает отдельно с Color и Fill, поэтому пишем и тот и другой
			this.Value = Asc.c_oAscShdNil;
			this.Color = CreateAscColorCustom(255, 255, 255);
			this.Fill = CreateAscColorCustom(255, 255, 255);
		}
	}

	asc_CParagraphShd.prototype.asc_getValue = function () {
		return this.Value;
	};
	asc_CParagraphShd.prototype.asc_putValue = function (v) {
		this.Value = v;
	};
	asc_CParagraphShd.prototype.asc_getColor = function () {
		return this.Color;
	};
	asc_CParagraphShd.prototype.asc_putColor = function (v) {
		this.Color = (v) ? v : null;
		this.Fill = (v) ? v : null;
	};

	/** @constructor */
	function asc_CParagraphFrame(obj) {
		if (obj) {
			this.FromDropCapMenu = false;

			this.DropCap = obj.DropCap;
			this.H = obj.H;
			this.HAnchor = obj.HAnchor;
			this.HRule = obj.HRule;
			this.HSpace = obj.HSpace;
			this.Lines = obj.Lines;
			this.VAnchor = obj.VAnchor;
			this.VSpace = obj.VSpace;
			this.W = obj.W;
			this.Wrap = obj.Wrap;
			this.X = obj.X;
			this.XAlign = obj.XAlign;
			this.Y = obj.Y;
			this.YAlign = obj.YAlign;
			this.Brd = (undefined != obj.Brd && null != obj.Brd) ? new asc_CParagraphBorders(obj.Brd) : null;
			this.Shd = (undefined != obj.Shd && null != obj.Shd) ? new asc_CParagraphShd(obj.Shd) : null;
			this.FontFamily = (undefined != obj.FontFamily && null != obj.FontFamily) ? new asc_CTextFontFamily(obj.FontFamily) : null;
		}
		else {
			this.FromDropCapMenu = false;

			this.DropCap = undefined;
			this.H = undefined;
			this.HAnchor = undefined;
			this.HRule = undefined;
			this.HSpace = undefined;
			this.Lines = undefined;
			this.VAnchor = undefined;
			this.VSpace = undefined;
			this.W = undefined;
			this.Wrap = undefined;
			this.X = undefined;
			this.XAlign = undefined;
			this.Y = undefined;
			this.YAlign = undefined;
			this.Shd = null;
			this.Brd = null;
			this.FontFamily = null;
		}
	}

	asc_CParagraphFrame.prototype.asc_getDropCap = function () {
		return this.DropCap;
	};
	asc_CParagraphFrame.prototype.asc_putDropCap = function (v) {
		this.DropCap = v;
	};
	asc_CParagraphFrame.prototype.asc_getH = function () {
		return this.H;
	};
	asc_CParagraphFrame.prototype.asc_putH = function (v) {
		this.H = v;
	};
	asc_CParagraphFrame.prototype.asc_getHAnchor = function () {
		return this.HAnchor;
	};
	asc_CParagraphFrame.prototype.asc_putHAnchor = function (v) {
		this.HAnchor = v;
	};
	asc_CParagraphFrame.prototype.asc_getHRule = function () {
		return this.HRule;
	};
	asc_CParagraphFrame.prototype.asc_putHRule = function (v) {
		this.HRule = v;
	};
	asc_CParagraphFrame.prototype.asc_getHSpace = function () {
		return this.HSpace;
	};
	asc_CParagraphFrame.prototype.asc_putHSpace = function (v) {
		this.HSpace = v;
	};
	asc_CParagraphFrame.prototype.asc_getLines = function () {
		return this.Lines;
	};
	asc_CParagraphFrame.prototype.asc_putLines = function (v) {
		this.Lines = v;
	};
	asc_CParagraphFrame.prototype.asc_getVAnchor = function () {
		return this.VAnchor;
	};
	asc_CParagraphFrame.prototype.asc_putVAnchor = function (v) {
		this.VAnchor = v;
	};
	asc_CParagraphFrame.prototype.asc_getVSpace = function () {
		return this.VSpace;
	};
	asc_CParagraphFrame.prototype.asc_putVSpace = function (v) {
		this.VSpace = v;
	};
	asc_CParagraphFrame.prototype.asc_getW = function () {
		return this.W;
	};
	asc_CParagraphFrame.prototype.asc_putW = function (v) {
		this.W = v;
	};
	asc_CParagraphFrame.prototype.asc_getWrap = function () {
		return this.Wrap;
	};
	asc_CParagraphFrame.prototype.asc_putWrap = function (v) {
		this.Wrap = v;
	};
	asc_CParagraphFrame.prototype.asc_getX = function () {
		return this.X;
	};
	asc_CParagraphFrame.prototype.asc_putX = function (v) {
		this.X = v;
	};
	asc_CParagraphFrame.prototype.asc_getXAlign = function () {
		return this.XAlign;
	};
	asc_CParagraphFrame.prototype.asc_putXAlign = function (v) {
		this.XAlign = v;
	};
	asc_CParagraphFrame.prototype.asc_getY = function () {
		return this.Y;
	};
	asc_CParagraphFrame.prototype.asc_putY = function (v) {
		this.Y = v;
	};
	asc_CParagraphFrame.prototype.asc_getYAlign = function () {
		return this.YAlign;
	};
	asc_CParagraphFrame.prototype.asc_putYAlign = function (v) {
		this.YAlign = v;
	};
	asc_CParagraphFrame.prototype.asc_getBorders = function () {
		return this.Brd;
	};
	asc_CParagraphFrame.prototype.asc_putBorders = function (v) {
		this.Brd = v;
	};
	asc_CParagraphFrame.prototype.asc_getShade = function () {
		return this.Shd;
	};
	asc_CParagraphFrame.prototype.asc_putShade = function (v) {
		this.Shd = v;
	};
	asc_CParagraphFrame.prototype.asc_getFontFamily = function () {
		return this.FontFamily;
	};
	asc_CParagraphFrame.prototype.asc_putFontFamily = function (v) {
		this.FontFamily = v;
	};
	asc_CParagraphFrame.prototype.asc_putFromDropCapMenu = function (v) {
		this.FromDropCapMenu = v;
	};

	/** @constructor */
	function asc_CParagraphSpacing(obj) {

		if (obj) {
			this.Line = (undefined != obj.Line) ? obj.Line : null; // Расстояние между строками внутри абзаца
			this.LineRule = (undefined != obj.LineRule) ? obj.LineRule : null; // Тип расстрояния между строками
			this.Before = (undefined != obj.Before) ? obj.Before : null; // Дополнительное расстояние до абзаца
			this.After = (undefined != obj.After) ? obj.After : null; // Дополнительное расстояние после абзаца
		}
		else {
			this.Line = undefined; // Расстояние между строками внутри абзаца
			this.LineRule = undefined; // Тип расстрояния между строками
			this.Before = undefined; // Дополнительное расстояние до абзаца
			this.After = undefined; // Дополнительное расстояние после абзаца
		}
	}

	asc_CParagraphSpacing.prototype.asc_getLine = function () {
		return this.Line;
	};
	asc_CParagraphSpacing.prototype.asc_getLineRule = function () {
		return this.LineRule;
	};
	asc_CParagraphSpacing.prototype.asc_getBefore = function () {
		return this.Before;
	};
	asc_CParagraphSpacing.prototype.asc_getAfter = function () {
		return this.After;
	};
	asc_CParagraphSpacing.prototype.asc_putLine = function (v) {
		this.Line = v;
	};
	asc_CParagraphSpacing.prototype.asc_putLineRule = function (v) {
		this.LineRule = v;
	};
	asc_CParagraphSpacing.prototype.asc_putBefore = function (v) {
		this.Before = v;
	};
	asc_CParagraphSpacing.prototype.asc_putAfter = function (v) {
		this.After = v;
	};

	/** @constructor */
	function asc_CParagraphInd(obj) {
		if (obj) {
			this.Left = (undefined != obj.Left) ? obj.Left : null; // Левый отступ
			this.Right = (undefined != obj.Right) ? obj.Right : null; // Правый отступ
			this.FirstLine = (undefined != obj.FirstLine) ? obj.FirstLine : null; // Первая строка
		}
		else {
			this.Left = undefined; // Левый отступ
			this.Right = undefined; // Правый отступ
			this.FirstLine = undefined; // Первая строка
		}
	}

	asc_CParagraphInd.prototype.asc_getLeft = function () {
		return this.Left;
	};
	asc_CParagraphInd.prototype.asc_putLeft = function (v) {
		this.Left = v;
	};
	asc_CParagraphInd.prototype.asc_getRight = function () {
		return this.Right;
	};
	asc_CParagraphInd.prototype.asc_putRight = function (v) {
		this.Right = v;
	};
	asc_CParagraphInd.prototype.asc_getFirstLine = function () {
		return this.FirstLine;
	};
	asc_CParagraphInd.prototype.asc_putFirstLine = function (v) {
		this.FirstLine = v;
	};

	/** @constructor */
	function asc_CParagraphProperty(obj) {

		if (obj) {
			this.Bidi = undefined !== obj.Bidi ? obj.Bidi : undefined;
			this.ContextualSpacing = (undefined != obj.ContextualSpacing) ? obj.ContextualSpacing : null;
			this.Ind = (undefined != obj.Ind && null != obj.Ind) ? new asc_CParagraphInd(obj.Ind) : null;
			this.KeepLines = (undefined != obj.KeepLines) ? obj.KeepLines : null;
			this.KeepNext = (undefined != obj.KeepNext) ? obj.KeepNext : undefined;
			this.WidowControl = (undefined != obj.WidowControl ? obj.WidowControl : undefined);
			this.PageBreakBefore = (undefined != obj.PageBreakBefore) ? obj.PageBreakBefore : null;
			this.Spacing = (undefined != obj.Spacing && null != obj.Spacing) ? new asc_CParagraphSpacing(obj.Spacing) : null;
			this.Brd = (undefined != obj.Brd && null != obj.Brd) ? new asc_CParagraphBorders(obj.Brd) : null;
			this.Shd = (undefined != obj.Shd && null != obj.Shd) ? new asc_CParagraphShd(obj.Shd) : null;
			this.Tabs = (undefined != obj.Tabs) ? new asc_CParagraphTabs(obj.Tabs) : undefined;
			this.DefaultTab = obj.DefaultTab != null ? obj.DefaultTab : window["AscCommonWord"].Default_Tab_Stop;
			this.Locked = (undefined != obj.Locked && null != obj.Locked) ? obj.Locked : false;
			this.CanAddTable = (undefined != obj.CanAddTable) ? obj.CanAddTable : true;

			this.FramePr = (undefined != obj.FramePr) ? new asc_CParagraphFrame(obj.FramePr) : undefined;
			this.CanAddDropCap = (undefined != obj.CanAddDropCap) ? obj.CanAddDropCap : false;
			this.CanAddImage = (undefined != obj.CanAddImage) ? obj.CanAddImage : false;

			this.Subscript = (undefined != obj.Subscript) ? obj.Subscript : undefined;
			this.Superscript = (undefined != obj.Superscript) ? obj.Superscript : undefined;
			this.SmallCaps = (undefined != obj.SmallCaps) ? obj.SmallCaps : undefined;
			this.AllCaps = (undefined != obj.AllCaps) ? obj.AllCaps : undefined;
			this.Strikeout = (undefined != obj.Strikeout) ? obj.Strikeout : undefined;
			this.DStrikeout = (undefined != obj.DStrikeout) ? obj.DStrikeout : undefined;
			this.TextSpacing = (undefined != obj.TextSpacing) ? obj.TextSpacing : undefined;
			this.Position = (undefined != obj.Position) ? obj.Position : undefined;
			this.Jc = (undefined != obj.Jc) ? obj.Jc : undefined;
			this.ListType = (undefined != obj.ListType) ? obj.ListType : undefined;
			this.OutlineLvl = (undefined != obj.OutlineLvl) ? obj.OutlineLvl : undefined;
			this.OutlineLvlStyle = (undefined != obj.OutlineLvlStyle) ? obj.OutlineLvlStyle : false;
			this.SuppressLineNumbers = undefined !== obj.SuppressLineNumbers ? obj.SuppressLineNumbers : false;
			this.Bullet = obj.Bullet;
			let oBullet = obj.Bullet;
			if (oBullet) {
				oBullet.FirstTextPr = obj.FirstTextPr;
			}
			this.Ligatures = undefined !== obj.Ligatures ? obj.Ligatures : undefined;

			this.CanDeleteBlockCC = undefined !== obj.CanDeleteBlockCC ? obj.CanDeleteBlockCC : true;
			this.CanEditBlockCC = undefined !== obj.CanEditBlockCC ? obj.CanEditBlockCC : true;
			this.CanDeleteInlineCC = undefined !== obj.CanDeleteInlineCC ? obj.CanDeleteInlineCC : true;
			this.CanEditInlineCC = undefined !== obj.CanEditInlineCC ? obj.CanEditInlineCC : true;
		}
		else {
			//ContextualSpacing : false,            // Удалять ли интервал между параграфами одинакового стиля
			//
			//    Ind :
			//    {
			//        Left      : 0,                    // Левый отступ
			//        Right     : 0,                    // Правый отступ
			//        FirstLine : 0                     // Первая строка
			//    },
			//
			//    Jc : align_Left,                      // Прилегание параграфа
			//
			//    KeepLines : false,                    // переносить параграф на новую страницу,
			//                                          // если на текущей он целиком не убирается
			//    KeepNext  : false,                    // переносить параграф вместе со следующим параграфом
			//
			//    PageBreakBefore : false,              // начинать параграф с новой страницы

			this.Bidi = undefined;
			this.ContextualSpacing = undefined;
			this.Ind = new asc_CParagraphInd();
			this.KeepLines = undefined;
			this.KeepNext = undefined;
			this.WidowControl = undefined;
			this.PageBreakBefore = undefined;
			this.Spacing = new asc_CParagraphSpacing();
			this.Brd = undefined;
			this.Shd = undefined;
			this.Locked = false;
			this.CanAddTable = true;
			this.Tabs = undefined;

			this.Subscript = undefined;
			this.Superscript = undefined;
			this.SmallCaps = undefined;
			this.AllCaps = undefined;
			this.Strikeout = undefined;
			this.DStrikeout = undefined;
			this.TextSpacing = undefined;
			this.Position = undefined;
			this.Jc = undefined;
			this.ListType = undefined;
			this.OutlineLvl = undefined;
			this.OutlineLvlStyle = false;
			this.SuppressLineNumbers = false;
			this.Bullet = undefined;
			this.Ligatures = undefined;

			this.CanDeleteBlockCC = true;
			this.CanEditBlockCC = true;
			this.CanDeleteInlineCC = true;
			this.CanEditInlineCC = true;
		}
	}
	
	asc_CParagraphProperty.prototype.asc_getRtlDirection = function() {
		return this.Bidi;
	};
	asc_CParagraphProperty.prototype.asc_putRtlDirection = function(v) {
		this.Bidi = v;
	};
	asc_CParagraphProperty.prototype.asc_getContextualSpacing = function () {
		return this.ContextualSpacing;
	};
	asc_CParagraphProperty.prototype.asc_putContextualSpacing = function (v) {
		this.ContextualSpacing = v;
	};
	asc_CParagraphProperty.prototype.asc_getInd = function () {
		return this.Ind;
	};
	asc_CParagraphProperty.prototype.asc_putInd = function (v) {
		this.Ind = v;
	};
	asc_CParagraphProperty.prototype.asc_getJc = function () {
		return this.Jc;
	};
	asc_CParagraphProperty.prototype.asc_putJc = function (v) {
		this.Jc = v;
	};
	asc_CParagraphProperty.prototype.asc_getKeepLines = function () {
		return this.KeepLines;
	};
	asc_CParagraphProperty.prototype.asc_putKeepLines = function (v) {
		this.KeepLines = v;
	};
	asc_CParagraphProperty.prototype.asc_getKeepNext = function () {
		return this.KeepNext;
	};
	asc_CParagraphProperty.prototype.asc_putKeepNext = function (v) {
		this.KeepNext = v;
	};
	asc_CParagraphProperty.prototype.asc_getPageBreakBefore = function () {
		return this.PageBreakBefore;
	};
	asc_CParagraphProperty.prototype.asc_putPageBreakBefore = function (v) {
		this.PageBreakBefore = v;
	};
	asc_CParagraphProperty.prototype.asc_getWidowControl = function () {
		return this.WidowControl;
	};
	asc_CParagraphProperty.prototype.asc_putWidowControl = function (v) {
		this.WidowControl = v;
	};
	asc_CParagraphProperty.prototype.asc_getSpacing = function () {
		return this.Spacing;
	};
	asc_CParagraphProperty.prototype.asc_putSpacing = function (v) {
		this.Spacing = v;
	};
	asc_CParagraphProperty.prototype.asc_getBorders = function () {
		return this.Brd;
	};
	asc_CParagraphProperty.prototype.asc_putBorders = function (v) {
		this.Brd = v;
	};
	asc_CParagraphProperty.prototype.asc_getShade = function () {
		return this.Shd;
	};
	asc_CParagraphProperty.prototype.asc_putShade = function (v) {
		this.Shd = v;
	};
	asc_CParagraphProperty.prototype.asc_getLocked = function () {
		return this.Locked;
	};
	asc_CParagraphProperty.prototype.asc_getCanAddTable = function () {
		return this.CanAddTable;
	};
	asc_CParagraphProperty.prototype.asc_getSubscript = function () {
		return this.Subscript;
	};
	asc_CParagraphProperty.prototype.asc_putSubscript = function (v) {
		this.Subscript = v;
	};
	asc_CParagraphProperty.prototype.asc_getSuperscript = function () {
		return this.Superscript;
	};
	asc_CParagraphProperty.prototype.asc_putSuperscript = function (v) {
		this.Superscript = v;
	};
	asc_CParagraphProperty.prototype.asc_getSmallCaps = function () {
		return this.SmallCaps;
	};
	asc_CParagraphProperty.prototype.asc_putSmallCaps = function (v) {
		this.SmallCaps = v;
	};
	asc_CParagraphProperty.prototype.asc_getAllCaps = function () {
		return this.AllCaps;
	};
	asc_CParagraphProperty.prototype.asc_putAllCaps = function (v) {
		this.AllCaps = v;
	};
	asc_CParagraphProperty.prototype.asc_getStrikeout = function () {
		return this.Strikeout;
	};
	asc_CParagraphProperty.prototype.asc_putStrikeout = function (v) {
		this.Strikeout = v;
	};
	asc_CParagraphProperty.prototype.asc_getDStrikeout = function () {
		return this.DStrikeout;
	};
	asc_CParagraphProperty.prototype.asc_putDStrikeout = function (v) {
		this.DStrikeout = v;
	};
	asc_CParagraphProperty.prototype.asc_getTextSpacing = function () {
		return this.TextSpacing;
	};
	asc_CParagraphProperty.prototype.asc_putTextSpacing = function (v) {
		this.TextSpacing = v;
	};
	asc_CParagraphProperty.prototype.asc_getPosition = function () {
		return this.Position;
	};
	asc_CParagraphProperty.prototype.asc_putPosition = function (v) {
		this.Position = v;
	};
	asc_CParagraphProperty.prototype.asc_getTabs = function () {
		return this.Tabs;
	};
	asc_CParagraphProperty.prototype.asc_putTabs = function (v) {
		this.Tabs = v;
	};
	asc_CParagraphProperty.prototype.asc_getDefaultTab = function () {
		return this.DefaultTab;
	};
	asc_CParagraphProperty.prototype.asc_putDefaultTab = function (v) {
		this.DefaultTab = v;
	};
	asc_CParagraphProperty.prototype.asc_getFramePr = function () {
		return this.FramePr;
	};
	asc_CParagraphProperty.prototype.asc_putFramePr = function (v) {
		this.FramePr = v;
	};
	asc_CParagraphProperty.prototype.asc_getCanAddDropCap = function () {
		return this.CanAddDropCap;
	};
	asc_CParagraphProperty.prototype.asc_getCanAddImage = function () {
		return this.CanAddImage;
	};
	asc_CParagraphProperty.prototype.asc_getOutlineLvl = function () {
		return this.OutlineLvl;
	};
	asc_CParagraphProperty.prototype.asc_putOutLineLvl = function (nLvl) {
		this.OutlineLvl = nLvl;
	};
	asc_CParagraphProperty.prototype.asc_getOutlineLvlStyle = function () {
		return this.OutlineLvlStyle;
	};
	asc_CParagraphProperty.prototype.asc_getSuppressLineNumbers = function () {
		return this.SuppressLineNumbers;
	};
	asc_CParagraphProperty.prototype.asc_putSuppressLineNumbers = function (isSuppress) {
		this.SuppressLineNumbers = isSuppress;
	};
	asc_CParagraphProperty.prototype.asc_putBullet = function (val) {
		this.Bullet = val;
	};
	asc_CParagraphProperty.prototype.asc_getBullet = function () {
		return this.Bullet;
	};
	asc_CParagraphProperty.prototype.asc_putBulletSize = function (size) {
		if (!this.Bullet) {
			this.Bullet = new Asc.asc_CBullet();
		}
		this.Bullet.asc_putSize(size);
	};
	asc_CParagraphProperty.prototype.asc_getBulletSize = function () {
		if (!this.Bullet) {
			return undefined;
		}
		return this.Bullet.asc_getSize();
	};
	asc_CParagraphProperty.prototype.asc_putBulletColor = function (color) {
		if (!this.Bullet) {
			this.Bullet = new Asc.asc_CBullet();
		}
		this.Bullet.asc_putColor(color);
	};
	asc_CParagraphProperty.prototype.asc_getBulletColor = function () {
		if (!this.Bullet) {
			return undefined;
		}
		return this.Bullet.asc_getColor();
	};
	asc_CParagraphProperty.prototype.asc_putNumStartAt = function (NumStartAt) {
		if (!this.Bullet) {
			this.Bullet = new Asc.asc_CBullet();
		}
		this.Bullet.asc_putNumStartAt(NumStartAt);
	};
	asc_CParagraphProperty.prototype.asc_getNumStartAt = function () {
		if (!this.Bullet) {
			return undefined;
		}
		return this.Bullet.asc_getNumStartAt();
	};
	asc_CParagraphProperty.prototype.asc_getBulletFont = function () {
		if (!this.Bullet) {
			return undefined;
		}
		return this.Bullet.asc_getFont();
	};
	asc_CParagraphProperty.prototype.asc_putBulletFont = function (v) {
		if (!this.Bullet) {
			this.Bullet = new Asc.asc_CBullet();
		}
		this.Bullet.asc_putFont(v);
	};
	asc_CParagraphProperty.prototype.asc_getBulletSymbol = function () {
		if (!this.Bullet) {
			return undefined;
		}
		return this.Bullet.asc_getSymbol();
	};
	asc_CParagraphProperty.prototype.asc_putBulletSymbol = function (v) {
		if (!this.Bullet) {
			this.Bullet = new Asc.asc_CBullet();
		}
		this.Bullet.asc_putSymbol(v);
	};
	asc_CParagraphProperty.prototype.asc_canDeleteBlockContentControl = function () {
		return this.CanDeleteBlockCC;
	};
	asc_CParagraphProperty.prototype.asc_canEditBlockContentControl = function () {
		return this.CanEditBlockCC;
	};
	asc_CParagraphProperty.prototype.asc_canDeleteInlineContentControl = function () {
		return this.CanDeleteInlineCC;
	};
	asc_CParagraphProperty.prototype.asc_canEditInlineContentControl = function () {
		return this.CanEditInlineCC;
	};
	asc_CParagraphProperty.prototype.asc_getLigatures = function () {
		return this.Ligatures;
	};
	asc_CParagraphProperty.prototype.asc_putLigatures = function (v) {
		this.Ligatures = v;
	};

	/** @constructor */
	function asc_CTexture() {
		this.Id = 0;
		this.Image = "";
	}

	asc_CTexture.prototype.asc_getId = function () {
		return this.Id;
	};
	asc_CTexture.prototype.asc_getImage = function () {
		return this.Image;
	};

	/** @constructor */
	function asc_CImageSize(width, height, isCorrect) {
		this.Width = (undefined == width) ? 0.0 : width;
		this.Height = (undefined == height) ? 0.0 : height;
		this.IsCorrect = isCorrect;
	}

	asc_CImageSize.prototype.asc_getImageWidth = function () {
		return this.Width;
	};
	asc_CImageSize.prototype.asc_getImageHeight = function () {
		return this.Height;
	};
	asc_CImageSize.prototype.asc_getIsCorrect = function () {
		return this.IsCorrect;
	};

	/** @constructor */
	function asc_CPaddings(obj) {

		if (obj) {
			this.Left = (undefined == obj.Left) ? null : obj.Left;
			this.Top = (undefined == obj.Top) ? null : obj.Top;
			this.Bottom = (undefined == obj.Bottom) ? null : obj.Bottom;
			this.Right = (undefined == obj.Right) ? null : obj.Right;
		}
		else {
			this.Left = null;
			this.Top = null;
			this.Bottom = null;
			this.Right = null;
		}
	}

	asc_CPaddings.prototype.asc_getLeft = function () {
		return this.Left;
	};
	asc_CPaddings.prototype.asc_putLeft = function (v) {
		this.Left = v;
	};
	asc_CPaddings.prototype.asc_getTop = function () {
		return this.Top;
	};
	asc_CPaddings.prototype.asc_putTop = function (v) {
		this.Top = v;
	};
	asc_CPaddings.prototype.asc_getBottom = function () {
		return this.Bottom;
	};
	asc_CPaddings.prototype.asc_putBottom = function (v) {
		this.Bottom = v;
	};
	asc_CPaddings.prototype.asc_getRight = function () {
		return this.Right;
	};
	asc_CPaddings.prototype.asc_putRight = function (v) {
		this.Right = v;
	};

	/** @constructor */
	function asc_CShapeProperty() {
		this.type = null; // custom
		this.fill = null;
		this.stroke = null;
		this.paddings = null;
		this.canFill = true;
		this.canChangeArrows = false;
		this.canEditText = false; // used in pdf editor
		this.bFromChart = false;
		this.bFromGroup = false;
		this.bFromImage = false;
		this.bFromSmartArt = false;
		this.bFromSmartArtInternal = false;
		this.Locked = false;
		this.w = null;
		this.h = null;
		this.vert = null;
		this.verticalTextAlign = null;
		this.textArtProperties = null;
		this.lockAspect = null;
		this.title = null;
		this.description = null;
		this.name = null;

		this.columnNumber = null;
		this.columnSpace = null;
		this.textFitType = null;
		this.vertOverflowType = null;
		this.signatureId = null;

		this.rot = null;
		this.rotAdd = null;
		this.flipH = null;
		this.flipV = null;
		this.flipHInvert = null;
		this.flipVInvert = null;
		this.Position = undefined;
		this.shadow = undefined;
		this.anchor = null;

		this.protectionLockText = null;
		this.protectionLocked = null;
		this.protectionPrint = null;
		this.isMotionPath = false;
	}

	asc_CShapeProperty.prototype.constructor = asc_CShapeProperty;
	asc_CShapeProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CShapeProperty.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CShapeProperty.prototype.asc_getFill = function () {
		return this.fill;
	};
	asc_CShapeProperty.prototype.asc_putFill = function (v) {
		this.fill = v;
	};
	asc_CShapeProperty.prototype.asc_getStroke = function () {
		return this.stroke;
	};
	asc_CShapeProperty.prototype.asc_putStroke = function (v) {
		this.stroke = v;
	};
	asc_CShapeProperty.prototype.asc_getPaddings = function () {
		return this.paddings;
	};
	asc_CShapeProperty.prototype.asc_putPaddings = function (v) {
		this.paddings = v;
	};
	asc_CShapeProperty.prototype.asc_getCanFill = function () {
		return this.canFill;
	};
	asc_CShapeProperty.prototype.asc_putCanFill = function (v) {
		this.canFill = v;
	};
	asc_CShapeProperty.prototype.asc_getCanChangeArrows = function () {
		return this.canChangeArrows;
	};
	asc_CShapeProperty.prototype.asc_setCanChangeArrows = function (v) {
		this.canChangeArrows = v;
	};
	asc_CShapeProperty.prototype.asc_getFromChart = function () {
		return this.bFromChart;
	};
	asc_CShapeProperty.prototype.asc_setFromChart = function (v) {
		this.bFromChart = v;
	};
	asc_CShapeProperty.prototype.asc_getFromSmartArt = function () {
		return this.bFromSmartArt;
	};
	asc_CShapeProperty.prototype.asc_setFromSmartArt = function (v) {
		this.bFromSmartArt = v;
	};
	asc_CShapeProperty.prototype.asc_getFromSmartArtInternal = function () {
		return this.bFromSmartArtInternal;
	};
	asc_CShapeProperty.prototype.asc_setFromSmartArtInternal = function (v) {
		this.bFromSmartArtInternal = v;
	};
	asc_CShapeProperty.prototype.asc_getFromGroup = function () {
		return this.bFromGroup;
	};
	asc_CShapeProperty.prototype.asc_setFromGroup = function (v) {
		this.bFromGroup = v;
	};
	asc_CShapeProperty.prototype.asc_getLocked = function () {
		return this.Locked;
	};
	asc_CShapeProperty.prototype.asc_setLocked = function (v) {
		this.Locked = v;
	};
	asc_CShapeProperty.prototype.asc_getWidth = function () {
		return this.w;
	};
	asc_CShapeProperty.prototype.asc_putWidth = function (v) {
		this.w = v;
	};
	asc_CShapeProperty.prototype.asc_getHeight = function () {
		return this.h;
	};
	asc_CShapeProperty.prototype.asc_putHeight = function (v) {
		this.h = v;
	};
	asc_CShapeProperty.prototype.asc_getVerticalTextAlign = function () {
		return this.verticalTextAlign;
	};
	asc_CShapeProperty.prototype.asc_putVerticalTextAlign = function (v) {
		this.verticalTextAlign = v;
	};
	asc_CShapeProperty.prototype.asc_getVert = function () {
		return this.vert;
	};
	asc_CShapeProperty.prototype.asc_putVert = function (v) {
		this.vert = v;
	};
	asc_CShapeProperty.prototype.asc_getTextArtProperties = function () {
		return this.textArtProperties;
	};
	asc_CShapeProperty.prototype.asc_putTextArtProperties = function (v) {
		this.textArtProperties = v;
	};
	asc_CShapeProperty.prototype.asc_getLockAspect = function () {
		return this.lockAspect
	};
	asc_CShapeProperty.prototype.asc_putLockAspect = function (v) {
		this.lockAspect = v;
	};
	asc_CShapeProperty.prototype.asc_getTitle = function () {
		return this.title;
	};
	asc_CShapeProperty.prototype.asc_putTitle = function (v) {
		this.title = v;
	};
	asc_CShapeProperty.prototype.asc_getDescription = function () {
		return this.description;
	};
	asc_CShapeProperty.prototype.asc_putDescription = function (v) {
		this.description = v;
	};
	asc_CShapeProperty.prototype.asc_getName = function () {
		return this.name;
	};
	asc_CShapeProperty.prototype.asc_putName = function (v) {
		this.name = v;
	};
	asc_CShapeProperty.prototype.asc_getColumnNumber = function () {
		return this.columnNumber;
	};
	asc_CShapeProperty.prototype.asc_putColumnNumber = function (v) {
		this.columnNumber = v;
	};
	asc_CShapeProperty.prototype.asc_getColumnSpace = function () {
		return this.columnSpace;
	};
	asc_CShapeProperty.prototype.asc_getTextFitType = function () {
		return this.textFitType;
	};
	asc_CShapeProperty.prototype.asc_getVertOverflowType = function () {
		return this.vertOverflowType;
	};
	asc_CShapeProperty.prototype.asc_putColumnSpace = function (v) {
		this.columnSpace = v;
	};
	asc_CShapeProperty.prototype.asc_putTextFitType = function (v) {
		this.textFitType = v;
	};
	asc_CShapeProperty.prototype.asc_putVertOverflowType = function (v) {
		this.vertOverflowType = v;
	};
	asc_CShapeProperty.prototype.asc_getSignatureId = function () {
		return this.signatureId;
	};
	asc_CShapeProperty.prototype.asc_putSignatureId = function (v) {
		this.signatureId = v;
	};
	asc_CShapeProperty.prototype.asc_getFromImage = function () {
		return this.bFromImage;
	};
	asc_CShapeProperty.prototype.asc_putFromImage = function (v) {
		this.bFromImage = v;
	};
	asc_CShapeProperty.prototype.asc_getRot = function () {
		return this.rot;
	};
	asc_CShapeProperty.prototype.asc_putRot = function (v) {
		this.rot = v;
	};
	asc_CShapeProperty.prototype.asc_getRotAdd = function () {
		return this.rotAdd;
	};
	asc_CShapeProperty.prototype.asc_putRotAdd = function (v) {
		this.rotAdd = v;
	};
	asc_CShapeProperty.prototype.asc_getFlipH = function () {
		return this.flipH;
	};
	asc_CShapeProperty.prototype.asc_putFlipH = function (v) {
		this.flipH = v;
	};
	asc_CShapeProperty.prototype.asc_getFlipV = function () {
		return this.flipV;
	};
	asc_CShapeProperty.prototype.asc_putFlipV = function (v) {
		this.flipV = v;
	};
	asc_CShapeProperty.prototype.asc_getFlipHInvert = function () {
		return this.flipHInvert;
	};
	asc_CShapeProperty.prototype.asc_putFlipHInvert = function (v) {
		this.flipHInvert = v;
	};
	asc_CShapeProperty.prototype.asc_getFlipVInvert = function () {
		return this.flipVInvert;
	};
	asc_CShapeProperty.prototype.asc_putFlipVInvert = function (v) {
		this.flipVInvert = v;
	};
	asc_CShapeProperty.prototype.asc_getShadow = function () {
		return this.shadow;
	};
	asc_CShapeProperty.prototype.asc_putShadow = function (v) {
		this.shadow = v;
	};
	asc_CShapeProperty.prototype.asc_getAnchor = function () {
		return this.anchor;
	};
	asc_CShapeProperty.prototype.asc_putAnchor = function (v) {
		this.anchor = v;
	};
	asc_CShapeProperty.prototype.asc_getProtectionLockText = function () {
		return this.protectionLockText;
	};
	asc_CShapeProperty.prototype.asc_putProtectionLockText = function (v) {
		this.protectionLockText = v;
	};
	asc_CShapeProperty.prototype.asc_getProtectionLocked = function () {
		return this.protectionLocked;
	};
	asc_CShapeProperty.prototype.asc_putProtectionLocked = function (v) {
		this.protectionLocked = v;
	};
	asc_CShapeProperty.prototype.asc_getProtectionPrint = function () {
		return this.protectionPrint;
	};
	asc_CShapeProperty.prototype.asc_putProtectionPrint = function (v) {
		this.protectionPrint = v;
	};
	asc_CShapeProperty.prototype.asc_getPosition = function () {
		return this.Position;
	};
	asc_CShapeProperty.prototype.asc_putPosition = function (v) {
		this.Position = v;
	};
	asc_CShapeProperty.prototype.asc_getIsMotionPath = function () {
		return this.isMotionPath;
	};
	asc_CShapeProperty.prototype.asc_getCanEditText = function () {
		return this.canEditText;
	};
	asc_CShapeProperty.prototype.asc_setCanEditText = function (v) {
		this.canEditText = v;
	};

	/** @constructor */
	function asc_CAnnotProperty() {
		this.type				= null; // custom
		this.fill				= null;
		this.stroke				= null;
		this.canFill			= true;
		this.canChangeArrows	= true;
		this.Locked				= false;
		this.subject			= undefined;
		this.canEditText		= false;

		this.Position = undefined;
	}

	asc_CAnnotProperty.prototype.constructor = asc_CAnnotProperty;
	asc_CAnnotProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CAnnotProperty.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CAnnotProperty.prototype.asc_getFill = function () {
		return this.fill;
	};
	asc_CAnnotProperty.prototype.asc_putFill = function (v) {
		this.fill = v;
	};
	asc_CAnnotProperty.prototype.asc_getStroke = function () {
		return this.stroke;
	};
	asc_CAnnotProperty.prototype.asc_putStroke = function (v) {
		this.stroke = v;
	};
	asc_CAnnotProperty.prototype.asc_getCanFill = function () {
		return this.canFill;
	};
	asc_CAnnotProperty.prototype.asc_putCanFill = function (v) {
		this.canFill = v;
	};
	asc_CAnnotProperty.prototype.asc_getCanChangeArrows = function () {
		return this.canChangeArrows;
	};
	asc_CAnnotProperty.prototype.asc_setCanChangeArrows = function (v) {
		this.canChangeArrows = v;
	};
	asc_CAnnotProperty.prototype.asc_getSubject = function () {
		return this.subject;
	};
	asc_CAnnotProperty.prototype.asc_setSubject = function (v) {
		this.subject = v;
	};
	asc_CAnnotProperty.prototype.asc_getCanEditText = function () {
		return this.canEditText;
	};
	asc_CAnnotProperty.prototype.asc_setCanEditText = function (v) {
		this.canEditText = v;
	};

	//////////////////////////////////////////////////////////////////
	///// Common field
	/////////////////////////////////////////////////////////////////
	function asc_CBaseFieldProperty() {
		// common
		this.type				= undefined;
		
		this.name				= undefined;
		this.required			= undefined;
		this.readOnly			= undefined;
		this.rot				= undefined;
		this.display			= undefined;
		this.fill				= null;
		this.stroke				= null;
		this.strokeWidth		= undefined;
		this.strokeStyle		= undefined;
		this.locked				= false;

		this.fieldProps	= null;
	}
	asc_CBaseFieldProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CBaseFieldProperty.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getName = function () {
		return this.name;
	};
	asc_CBaseFieldProperty.prototype.asc_putName = function (v) {
		this.name = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getRequired = function () {
		return this.required;
	};
	asc_CBaseFieldProperty.prototype.asc_putRequired = function (v) {
		this.required = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getReadOnly = function () {
		return this.readOnly;
	};
	asc_CBaseFieldProperty.prototype.asc_putReadOnly = function (v) {
		this.readOnly = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getRot = function () {
		return this.rot;
	};
	asc_CBaseFieldProperty.prototype.asc_putRot = function (v) {
		this.rot = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getDisplay = function () {
		return this.display;
	};
	asc_CBaseFieldProperty.prototype.asc_putDisplay = function (v) {
		this.display = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getFill = function () {
		return this.fill;
	};
	asc_CBaseFieldProperty.prototype.asc_putFill = function (v) {
		this.fill = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getStroke = function () {
		return this.stroke;
	};
	asc_CBaseFieldProperty.prototype.asc_putStroke = function (v) {
		this.stroke = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getStrokeWidth = function () {
		return this.strokeWidth;
	};
	asc_CBaseFieldProperty.prototype.asc_putStrokeWidth = function (v) {
		this.strokeWidth = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getStrokeStyle = function () {
		return this.strokeStyle;
	};
	asc_CBaseFieldProperty.prototype.asc_putStrokeStyle = function (v) {
		this.strokeStyle = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getPropLocked = function () {
		return this.locked;
	};
	asc_CBaseFieldProperty.prototype.asc_putPropLocked = function (v) {
		this.locked = v;
	};
	asc_CBaseFieldProperty.prototype.get_Locked = function () {
		return this.coEditLocked;
	};
	asc_CBaseFieldProperty.prototype.put_Locked = function (v) {
		this.coEditLocked = v;
	};
	asc_CBaseFieldProperty.prototype.asc_getFieldProps = function () {
		return this.fieldProps;
	};
	asc_CBaseFieldProperty.prototype.asc_putFieldProps = function (v) {
		this.fieldProps = v;
	};
	//////////////////////////////////////////////////////////////////
	///// Text field
	//////////////////////////////////////////////////////////////////
	function asc_CTextFieldProperty() {
		// format
		this.format				= null;
		this.validate			= null;

		// text
		this.defaultValue		= undefined;
		this.multiline			= undefined;
		this.scrollLongText		= undefined;
		this.charLimit			= undefined;
		this.comb				= undefined;
		this.placeholder		= undefined;
		this.autoFit			= undefined;
	}
	asc_CTextFieldProperty.prototype.asc_getDefaultValue = function () {
		return this.defaultValue;
	};
	asc_CTextFieldProperty.prototype.asc_putDefaultValue = function (v) {
		this.defaultValue = v;
	};
	asc_CTextFieldProperty.prototype.asc_getMultiline = function () {
		return this.multiline;
	};
	asc_CTextFieldProperty.prototype.asc_putMultiline = function (v) {
		this.multiline = v;
	};
	asc_CTextFieldProperty.prototype.asc_getScrollLongText = function () {
		return this.scrollLongText;
	};
	asc_CTextFieldProperty.prototype.asc_putScrollLongText = function (v) {
		this.scrollLongText = v;
	};
	asc_CTextFieldProperty.prototype.asc_getCharLimit = function () {
		return this.charLimit;
	};
	asc_CTextFieldProperty.prototype.asc_putCharLimit = function (v) {
		this.charLimit = v;
	};
	asc_CTextFieldProperty.prototype.asc_getComb = function () {
		return this.comb;
	};
	asc_CTextFieldProperty.prototype.asc_putComb = function (v) {
		this.comb = v;
	};
	asc_CTextFieldProperty.prototype.asc_getPlaceholder = function () {
		return this.placeholder;
	};
	asc_CTextFieldProperty.prototype.asc_putPlaceholder = function (v) {
		this.placeholder = v;
	};
	asc_CTextFieldProperty.prototype.asc_getAutoFit = function () {
		return this.autoFit;
	};
	asc_CTextFieldProperty.prototype.asc_putAutoFit = function (v) {
		this.autoFit = v;
	};
	asc_CTextFieldProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CTextFieldProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	asc_CTextFieldProperty.prototype.asc_getValidate = function () {
		return this.validate;
	};
	asc_CTextFieldProperty.prototype.asc_putValidate = function (v) {
		this.validate = v;
	};
	
	//////////////////////////////////////////////////////////////////
	///// Combobox field
	//////////////////////////////////////////////////////////////////
	function asc_CComboboxFieldProperty() {
		// format
		this.format				= null;
		this.validate			= null;

		this.options			= null;
		this.commitOnSelChange	= undefined;
		this.editable			= undefined;
		this.placeholder		= undefined;
		this.autoFit			= undefined;
	}
	asc_CComboboxFieldProperty.prototype.asc_getOptions = function () {
		return this.options;
	};
	asc_CComboboxFieldProperty.prototype.asc_putOptions = function (v) {
		this.options = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getCommitOnSelChange = function () {
		return this.commitOnSelChange;
	};
	asc_CComboboxFieldProperty.prototype.asc_putCommitOnSelChange = function (v) {
		this.commitOnSelChange = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getEditable = function () {
		return this.editable;
	};
	asc_CComboboxFieldProperty.prototype.asc_putEditable = function (v) {
		this.editable = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getPlaceholder = function () {
		return this.placeholder;
	};
	asc_CComboboxFieldProperty.prototype.asc_putPlaceholder = function (v) {
		this.placeholder = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getAutoFit = function () {
		return this.autoFit;
	};
	asc_CComboboxFieldProperty.prototype.asc_putAutoFit = function (v) {
		this.autoFit = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CComboboxFieldProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	asc_CComboboxFieldProperty.prototype.asc_getValidate = function () {
		return this.validate;
	};
	asc_CComboboxFieldProperty.prototype.asc_putValidate = function (v) {
		this.validate = v;
	};

	//////////////////////////////////////////////////////////////////
	///// Listbox field
	//////////////////////////////////////////////////////////////////
	function asc_CListboxFieldProperty() {
		this.commitOnSelChange	= undefined;
		this.multipleSelection	= undefined;
	}
	asc_CListboxFieldProperty.prototype.asc_getOptions = function () {
		return this.options;
	};
	asc_CListboxFieldProperty.prototype.asc_putOptions = function (v) {
		this.options = v;
	};
	asc_CListboxFieldProperty.prototype.asc_getCommitOnSelChange = function () {
		return this.commitOnSelChange;
	};
	asc_CListboxFieldProperty.prototype.asc_putCommitOnSelChange = function (v) {
		this.commitOnSelChange = v;
	};
	asc_CListboxFieldProperty.prototype.asc_getMultipleSelection = function () {
		return this.multipleSelection;
	};
	asc_CListboxFieldProperty.prototype.asc_putMultipleSelection = function (v) {
		this.multipleSelection = v;
	};
	//////////////////////////////////////////////////////////////////
	///// Checkbox
	//////////////////////////////////////////////////////////////////
	function asc_CCheckboxFieldProperty() {
		this.checkboxStyle	= undefined;
		this.exportValue	= undefined;
		this.defaultChecked	= undefined;
		this.toggleToOff	= undefined;
	}
	asc_CCheckboxFieldProperty.prototype.asc_getCheckboxStyle = function () {
		return this.checkboxStyle;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putCheckboxStyle = function (v) {
		this.checkboxStyle = v;
	};
	asc_CCheckboxFieldProperty.prototype.asc_getExportValue = function () {
		return this.exportValue;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putExportValue = function (v) {
		this.exportValue = v;
	};
	asc_CCheckboxFieldProperty.prototype.asc_getDefaultChecked = function () {
		return this.defaultChecked;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putDefaultChecked = function (v) {
		this.defaultChecked = v;
	};
	asc_CCheckboxFieldProperty.prototype.asc_getToggleToOff = function () {
		return this.toggleToOff;
	};
	asc_CCheckboxFieldProperty.prototype.asc_putToggleToOff = function (v) {
		this.toggleToOff = v;
	};
	
	//////////////////////////////////////////////////////////////////
	///// Radiobutton
	//////////////////////////////////////////////////////////////////
	function asc_CRadiobuttonFieldProperty() {
		asc_CCheckboxFieldProperty.call(this);
		this.radiosInUnison	= undefined;
	}
	asc_CRadiobuttonFieldProperty.prototype = Object.create(asc_CCheckboxFieldProperty.prototype);
	asc_CRadiobuttonFieldProperty.prototype.constructor = asc_CRadiobuttonFieldProperty;

	asc_CRadiobuttonFieldProperty.prototype.asc_getRadiosInUnison = function () {
		return this.radiosInUnison;
	};
	asc_CRadiobuttonFieldProperty.prototype.asc_putRadiosInUnison = function (v) {
		this.radiosInUnison = v;
	};
	//////////////////////////////////////////////////////////////////
	///// Pushbutton
	//////////////////////////////////////////////////////////////////
	function asc_CButtonFieldProperty(buttonField) {
		this.parentField	= buttonField;

		this.highlight		= undefined;
		this.layout			= undefined;
		this.scaleWhen		= undefined;
		this.scaleHow		= undefined;
		this.fitBounds		= undefined;
		this.iconPos		= null;
		this.behavior		= undefined;
		this.currentState	= undefined;
		this.normalCaption	= undefined;
		this.hoverCaption	= undefined;
		this.downCaption	= undefined;
		this.normalImage	= undefined;
		this.hoverImage		= undefined;
		this.downImage		= undefined;

		this.DivId			= undefined;
	}
	asc_CButtonFieldProperty.prototype.getParentField = function () {
		return this.parentField;
	};
	asc_CButtonFieldProperty.prototype.asc_getHighlight = function () {
		return this.highlight;
	};
	asc_CButtonFieldProperty.prototype.asc_putHighlight = function (v) {
		this.highlight = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getLayout = function () {
		return this.layout;
	};
	asc_CButtonFieldProperty.prototype.asc_putLayout = function (v) {
		this.layout = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getScaleWhen = function () {
		return this.scaleWhen;
	};
	asc_CButtonFieldProperty.prototype.asc_putScaleWhen = function (v) {
		this.scaleWhen = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getScaleHow = function () {
		return this.scaleHow;
	};
	asc_CButtonFieldProperty.prototype.asc_putScaleHow = function (v) {
		this.scaleHow = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getFitBounds = function () {
		return this.fitBounds;
	};
	asc_CButtonFieldProperty.prototype.asc_putFitBounds = function (v) {
		this.fitBounds = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getIconPos = function () {
		return this.iconPos;
	};
	asc_CButtonFieldProperty.prototype.asc_putIconPos = function (v) {
		this.iconPos = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getBehavior = function () {
		return this.behavior;
	};
	asc_CButtonFieldProperty.prototype.asc_putBehavior = function (v) {
		this.behavior = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getNormalCaption = function () {
		return this.normalCaption;
	};
	asc_CButtonFieldProperty.prototype.asc_putNormalCaption = function (v) {
		this.normalCaption = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getNormalImage = function () {
		return this.normalImage;
	};
	asc_CButtonFieldProperty.prototype.asc_putNormalImage = function (v) {
		this.normalImage = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getHoverCaption = function () {
		return this.hoverCaption;
	};
	asc_CButtonFieldProperty.prototype.asc_putHoverCaption = function (v) {
		this.hoverCaption = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getHoverImage = function () {
		return this.hoverImage;
	};
	asc_CButtonFieldProperty.prototype.asc_putHoverImage = function (v) {
		this.hoverImage = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getDownCaption = function () {
		return this.downCaption;
	};
	asc_CButtonFieldProperty.prototype.asc_putDownCaption = function (v) {
		this.downCaption = v;
	};
	asc_CButtonFieldProperty.prototype.asc_getDownImage = function () {
		return this.downImage;
	};
	asc_CButtonFieldProperty.prototype.asc_putDownImage = function (v) {
		this.downImage = v;
	};
	asc_CButtonFieldProperty.prototype.asc_putCurrentState = function(v) {
		this.currentState = v;

		// set to field state to add image (will clear after set image)
		let oField = this.getParentField();
		oField.asc_curImageState = v;

		this.drawTexture(v);
	};
	asc_CButtonFieldProperty.prototype.asc_getCurrentState = function(v) {
		return this.currentState;
	};
	asc_CButtonFieldProperty.prototype.put_DivId = function (v) {
		this.DivId = v;
		this.drawTexture(this.currentState);
	};
	asc_CButtonFieldProperty.prototype.drawTexture = function (nState) {
		let sImageRasterId;
		switch (nState) {
			case AscPDF.APPEARANCE_TYPES.normal:
				sImageRasterId = this.normalImage;
				break;
			case AscPDF.APPEARANCE_TYPES.mouseDown:
				sImageRasterId = this.downImage;
				break;
			case AscPDF.APPEARANCE_TYPES.rollover:
				sImageRasterId = this.hoverImage;
				break;
		}

		var oDiv = document.getElementById(this.DivId);
		if(!oDiv){
			return;
		}

		var aChildren = oDiv.children;
		var oCanvas = null;
		for(var i = 0; i < aChildren.length; ++i){
			if(aChildren[i].nodeName && aChildren[i].nodeName.toUpperCase() === 'CANVAS'){
				oCanvas = aChildren[i];
				break;
			}
		}
		var nWidth = oDiv.clientWidth;
		var nHeight = oDiv.clientHeight;
		if(null === oCanvas){
			oCanvas = document.createElement('canvas');
			oCanvas.width = parseInt(nWidth);
			oCanvas.height = parseInt(nHeight);
			oDiv.appendChild(oCanvas);
		}
		var oContext = oCanvas.getContext('2d');
		oContext.clearRect(0, 0, oCanvas.width, oCanvas.height);
		if (!sImageRasterId) {
			return;
		}
		
		var _img = Asc.editor.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(sImageRasterId)];
		if (_img != undefined && _img.Image != null && _img.Status != AscFonts.ImageLoadStatus.Loading)
		{
			var _x = 0;
			var _y = 0;
			var _w = Math.max(_img.Image.width, 1);
			var _h = Math.max(_img.Image.height, 1);

			var dAspect1 = nWidth / nHeight;
			var dAspect2 = _w / _h;

			_w = nWidth;
			_h = nHeight;
			if (dAspect1 >= dAspect2)
			{
				_w = dAspect2 * nHeight;
				_x = (nWidth - _w) / 2;
			}
			else
			{
				_h = _w / dAspect2;
				_y = (nHeight - _h) / 2;
			}
			oContext.drawImage(_img.Image, _x, _y, _w, _h);
		}
		else if (!_img || !_img.Image)
		{
			oContext.lineWidth = 1;

			oContext.beginPath();
			oContext.moveTo(0, 0);
			oContext.lineTo(nWidth, nHeight);
			oContext.moveTo(nWidth, 0);
			oContext.lineTo(0, nHeight);
			oContext.strokeStyle = "#FF0000";
			oContext.stroke();

			oContext.beginPath();
			oContext.moveTo(0, 0);
			oContext.lineTo(nWidth, 0);
			oContext.lineTo(nWidth, nHeight);
			oContext.lineTo(0, nHeight);
			oContext.closePath();

			oContext.strokeStyle = "#000000";
			oContext.stroke();
			oContext.beginPath();
		}
	};
	asc_CButtonFieldProperty.prototype.put_ImageUrl = function (sUrl, nState) {
		if (!this.DivId){
			return;
		}
		let Api = Asc.editor;

		// set to field state to add image (will clear after set image)
		let oField = this.getParentField();

		Api._addImageUrl([sUrl], oField);
	};
	asc_CButtonFieldProperty.prototype.showFileDialog = function (nState) {
		if (!this.DivId){
			return;
		}
		let Api = Asc.editor;

		// set to field state to add image (will clear after set image)
		let oField = this.getParentField();
		let oDoc = oField.GetDocument();
		let oActionsQueue = oDoc.GetActionsQueue();

		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]()) {
            window["AscDesktopEditor"]["OpenFilenameDialog"]("images", false, function(_file) {
                var file = _file;
                if (Array.isArray(file))
                    file = file[0];
        
                var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](file);
                editor._addImageUrl([AscCommon.g_oDocumentUrls.getImageUrl(_url)], oField);
            });
        }
        else {
            AscCommon.ShowImageFileDialog(Api.documentId, Api.documentUserId, undefined, Api.documentShardKey, Api.documentWopiSrc, Api.documentUserSessionId, function(error, files) {
                if (error.canceled == true) {
                    oActionsQueue.Continue();
                }
                else {
                    Api._uploadCallback(error, files, oField);
                }

				AscCommon.global_mouseEvent.UnLockMouse();

            }, function(error) {
                if (c_oAscError.ID.No !== error) {
                    Api.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
                }

                Api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
                AscCommon.global_mouseEvent.UnLockMouse();
            });
        }
	};
	//////////////////////////////////////////////////////////////////
	///// Number format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldNumberFormatProperty() {
		this.type				= AscPDF.FormatType.NUMBER;
		this.decimals			= undefined;
		this.sepStyle			= undefined;
		this.negStyle			= undefined;
		this.currency			= undefined;
		this.currencyPrepend	= undefined;
	};

	asc_CFieldNumberFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getDecimals = function () {
		return this.decimals;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putDecimals = function (v) {
		this.decimals = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getSepStyle = function () {
		return this.sepStyle;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putSepStyle = function (v) {
		this.sepStyle = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getNegStyle = function () {
		return this.negStyle;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putNegStyle = function (v) {
		this.negStyle = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getCurrency = function () {
		return this.currency;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putCurrency = function (v) {
		this.currency = v;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_getCurrencyPrepend = function () {
		return this.currencyPrepend;
	};
	asc_CFieldNumberFormatProperty.prototype.asc_putCurrencyPrepend = function (v) {
		this.currencyPrepend = v;
	};
	//////////////////////////////////////////////////////////////////
	///// Percentage format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldPercentageFormatProperty() {
		this.type		= AscPDF.FormatType.PERCENTAGE;
		this.decimals	= undefined;
		this.sepStyle	= undefined;
	};

	asc_CFieldPercentageFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_getDecimals = function () {
		return this.decimals;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_putDecimals = function (v) {
		this.decimals = v;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_getSepStyle = function () {
		return this.sepStyle;
	};
	asc_CFieldPercentageFormatProperty.prototype.asc_putSepStyle = function (v) {
		this.sepStyle = v;
	};

	//////////////////////////////////////////////////////////////////
	///// Date format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldDateFormatProperty() {
		this.type		= AscPDF.FormatType.DATE;
		this.format		= undefined;
	};

	asc_CFieldDateFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldDateFormatProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CFieldDateFormatProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	
	//////////////////////////////////////////////////////////////////
	///// Time format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldTimeFormatProperty() {
		this.type		= AscPDF.FormatType.TIME;
		this.format		= undefined;
	};

	asc_CFieldTimeFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldTimeFormatProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CFieldTimeFormatProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};

	//////////////////////////////////////////////////////////////////
	///// Special format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldSpecialFormatProperty() {
		this.type		= AscPDF.FormatType.SPECIAL;
		this.format		= undefined;
		this.mask		= undefined;
	};

	asc_CFieldSpecialFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_getFormat = function () {
		return this.format;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_putFormat = function (v) {
		this.format = v;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_getMask = function () {
		return this.mask;
	};
	asc_CFieldSpecialFormatProperty.prototype.asc_putMask = function (v) {
		this.mask = v;
	};

	//////////////////////////////////////////////////////////////////
	///// Regular (our custom) format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldRegularFormatProperty() {
		this.type		= AscPDF.FormatType.REGULAR;
		this.regExp		= undefined;
	};

	asc_CFieldRegularFormatProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldRegularFormatProperty.prototype.asc_getRegExp = function () {
		return this.regExp;
	};
	asc_CFieldRegularFormatProperty.prototype.asc_putRegExp = function (v) {
		this.regExp = v;
	};

	//////////////////////////////////////////////////////////////////
	///// Validate format
	//////////////////////////////////////////////////////////////////
	function asc_CFieldValidateProperty() {
		this.type			= undefined;
		this.greaterThen	= undefined;
		this.lessThen		= undefined;
	};

	asc_CFieldValidateProperty.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CFieldValidateProperty.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CFieldValidateProperty.prototype.asc_getBeGreaterThen = function () {
		return this.bGreaterThen;
	};
	asc_CFieldValidateProperty.prototype.asc_putBeGreaterThen = function (v) {
		this.bGreaterThen = v;
	};
	asc_CFieldValidateProperty.prototype.asc_getGreaterThen = function () {
		return this.greaterThen;
	};
	asc_CFieldValidateProperty.prototype.asc_putGreaterThen = function (v) {
		this.greaterThen = v;
	};
	asc_CFieldValidateProperty.prototype.asc_getBeLessThen = function () {
		return this.bLessThen;
	};
	asc_CFieldValidateProperty.prototype.asc_putBeLessThen = function (v) {
		this.bLessThen = v;
	};
	asc_CFieldValidateProperty.prototype.asc_getLessThen = function () {
		return this.lessThen;
	};
	asc_CFieldValidateProperty.prototype.asc_putLessThen = function (v) {
		this.lessThen = v;
	};

	/** @constructor */
	function asc_CPdfPageProperty() {
		this.deleteLock	= false;
		this.rotateLock	= false;
		this.editLock	= false;
	}

	asc_CPdfPageProperty.prototype.constructor = asc_CPdfPageProperty;
	asc_CPdfPageProperty.prototype.asc_getDeleteLock = function () {
		return this.deleteLock;
	};
	asc_CPdfPageProperty.prototype.asc_putDeleteLock = function (v) {
		this.deleteLock = v;
	};
	asc_CPdfPageProperty.prototype.asc_getRotateLock = function () {
		return this.rotateLock;
	};
	asc_CPdfPageProperty.prototype.asc_putRotateLock = function (v) {
		this.rotateLock = v;
	};
	asc_CPdfPageProperty.prototype.asc_getEditLock = function () {
		return this.editLock;
	};
	asc_CPdfPageProperty.prototype.asc_putEditLock = function (v) {
		this.editLock = v;
	};

	/** @constructor */
	function asc_TextArtProperties(obj) {
		if (obj) {
			this.Fill = obj.Fill;//asc_Fill
			this.Line = obj.Line;//asc_Stroke
			this.Form = obj.Form;//srting
			this.Style = obj.Style;//
		}
		else {
			this.Fill = undefined;
			this.Line = undefined;
			this.Form = undefined;
			this.Style = undefined;
		}
	}

	asc_TextArtProperties.prototype.asc_putFill = function (oAscFill) {
		this.Fill = oAscFill;
	};
	asc_TextArtProperties.prototype.asc_getFill = function () {
		return this.Fill;
	};
	asc_TextArtProperties.prototype.asc_putLine = function (oAscStroke) {
		this.Line = oAscStroke;
	};
	asc_TextArtProperties.prototype.asc_getLine = function () {
		return this.Line;
	};
	asc_TextArtProperties.prototype.asc_putForm = function (sForm) {
		this.Form = sForm;
	};
	asc_TextArtProperties.prototype.asc_getForm = function () {
		return this.Form;
	};
	asc_TextArtProperties.prototype.asc_putStyle = function (Style) {
		this.Style = Style;
	};
	asc_TextArtProperties.prototype.asc_getStyle = function () {
		return this.Style;
	};

	function CImagePositionH(obj) {
		if (obj) {
			this.RelativeFrom = (undefined === obj.RelativeFrom) ? undefined : obj.RelativeFrom;
			this.UseAlign = (undefined === obj.UseAlign) ? undefined : obj.UseAlign;
			this.Align = (undefined === obj.Align) ? undefined : obj.Align;
			this.Value = (undefined === obj.Value) ? undefined : obj.Value;
			this.Percent = (undefined === obj.Percent) ? undefined : obj.Percent;
		}
		else {
			this.RelativeFrom = undefined;
			this.UseAlign = undefined;
			this.Align = undefined;
			this.Value = undefined;
			this.Percent = undefined;
		}
	}

	CImagePositionH.prototype.get_RelativeFrom = function () {
		return this.RelativeFrom;
	};
	CImagePositionH.prototype.put_RelativeFrom = function (v) {
		this.RelativeFrom = v;
	};
	CImagePositionH.prototype.get_UseAlign = function () {
		return this.UseAlign;
	};
	CImagePositionH.prototype.put_UseAlign = function (v) {
		this.UseAlign = v;
	};
	CImagePositionH.prototype.get_Align = function () {
		return this.Align;
	};
	CImagePositionH.prototype.put_Align = function (v) {
		this.Align = v;
	};
	CImagePositionH.prototype.get_Value = function () {
		return this.Value;
	};
	CImagePositionH.prototype.put_Value = function (v) {
		this.Value = v;
	};
	CImagePositionH.prototype.get_Percent = function () {
		return this.Percent
	};
	CImagePositionH.prototype.put_Percent = function (v) {
		this.Percent = v;
	};

	function CImagePositionV(obj) {
		if (obj) {
			this.RelativeFrom = (undefined === obj.RelativeFrom) ? undefined : obj.RelativeFrom;
			this.UseAlign = (undefined === obj.UseAlign) ? undefined : obj.UseAlign;
			this.Align = (undefined === obj.Align) ? undefined : obj.Align;
			this.Value = (undefined === obj.Value) ? undefined : obj.Value;
			this.Percent = (undefined === obj.Percent) ? undefined : obj.Percent;
		}
		else {
			this.RelativeFrom = undefined;
			this.UseAlign = undefined;
			this.Align = undefined;
			this.Value = undefined;
			this.Percent = undefined;
		}
	}

	CImagePositionV.prototype.get_RelativeFrom = function () {
		return this.RelativeFrom;
	};
	CImagePositionV.prototype.put_RelativeFrom = function (v) {
		this.RelativeFrom = v;
	};
	CImagePositionV.prototype.get_UseAlign = function () {
		return this.UseAlign;
	};
	CImagePositionV.prototype.put_UseAlign = function (v) {
		this.UseAlign = v;
	};
	CImagePositionV.prototype.get_Align = function () {
		return this.Align;
	};
	CImagePositionV.prototype.put_Align = function (v) {
		this.Align = v;
	};
	CImagePositionV.prototype.get_Value = function () {
		return this.Value;
	};
	CImagePositionV.prototype.put_Value = function (v) {
		this.Value = v;
	};
	CImagePositionV.prototype.get_Percent = function () {
		return this.Percent
	};
	CImagePositionV.prototype.put_Percent = function (v) {
		this.Percent = v;
	};

	function CPosition(obj) {
		if (obj) {
			this.X = (undefined == obj.X) ? null : obj.X;
			this.Y = (undefined == obj.Y) ? null : obj.Y;
		}
		else {
			this.X = null;
			this.Y = null;
		}
	}

	CPosition.prototype.get_X = function () {
		return this.X;
	};
	CPosition.prototype.put_X = function (v) {
		this.X = v;
	};
	CPosition.prototype.get_Y = function () {
		return this.Y;
	};
	CPosition.prototype.put_Y = function (v) {
		this.Y = v;
	};

	/** @constructor */
	function asc_CImgProperty(obj) {

		if (obj) {
			this.CanBeFlow = (undefined != obj.CanBeFlow) ? obj.CanBeFlow : true;

			this.Width = (undefined != obj.Width) ? obj.Width : undefined;
			this.Height = (undefined != obj.Height) ? obj.Height : undefined;
			this.WrappingStyle = (undefined != obj.WrappingStyle) ? obj.WrappingStyle : undefined;
			this.Paddings = (undefined != obj.Paddings) ? new asc_CPaddings(obj.Paddings) : undefined;
			this.Position = (undefined != obj.Position) ? new CPosition(obj.Position) : undefined;
			this.AllowOverlap = (undefined != obj.AllowOverlap) ? obj.AllowOverlap : undefined;
			this.PositionH = (undefined != obj.PositionH) ? new CImagePositionH(obj.PositionH) : undefined;
			this.PositionV = (undefined != obj.PositionV) ? new CImagePositionV(obj.PositionV) : undefined;

			this.SizeRelH = (undefined != obj.SizeRelH) ? new CImagePositionH(obj.SizeRelH) : undefined;
			this.SizeRelV = (undefined != obj.SizeRelV) ? new CImagePositionV(obj.SizeRelV) : undefined;

			this.Internal_Position = (undefined != obj.Internal_Position) ? obj.Internal_Position : null;

			this.ImageUrl = (undefined != obj.ImageUrl) ? obj.ImageUrl : null;
			this.Token = obj.Token;
			this.Locked = (undefined != obj.Locked) ? obj.Locked : false;
			this.lockAspect = (undefined != obj.lockAspect) ? obj.lockAspect : false;


			this.ChartProperties = (undefined != obj.ChartProperties) ? obj.ChartProperties : null;
			this.ShapeProperties = (undefined != obj.ShapeProperties) ? obj.ShapeProperties : null;
			this.SlicerProperties = (undefined != obj.SlicerProperties) ? obj.SlicerProperties : null;

			this.ChangeLevel = (undefined != obj.ChangeLevel) ? obj.ChangeLevel : null;
			this.Group = (obj.Group != undefined) ? obj.Group : null;

			this.fromGroup = obj.fromGroup != undefined ? obj.fromGroup : null;
			this.severalCharts = obj.severalCharts != undefined ? obj.severalCharts : false;
			this.severalChartTypes = obj.severalChartTypes != undefined ? obj.severalChartTypes : undefined;
			this.severalChartStyles = obj.severalChartStyles != undefined ? obj.severalChartStyles : undefined;
			this.verticalTextAlign = obj.verticalTextAlign != undefined ? obj.verticalTextAlign : undefined;
			this.vert = obj.vert != undefined ? obj.vert : undefined;

			//oleObjects
			this.pluginGuid = obj.pluginGuid !== undefined ? obj.pluginGuid : undefined;
			this.pluginData = obj.pluginData !== undefined ? obj.pluginData : undefined;

			this.title = obj.title != undefined ? obj.title : undefined;
			this.description = obj.description != undefined ? obj.description : undefined;
			this.name = obj.name != undefined ? obj.name : undefined;

			this.columnNumber = obj.columnNumber != undefined ? obj.columnNumber : undefined;
			this.columnSpace = obj.columnSpace != undefined ? obj.columnSpace : undefined;
			this.textFitType = obj.textFitType != undefined ? obj.textFitType : undefined;
			this.vertOverflowType = obj.vertOverflowType != undefined ? obj.vertOverflowType : undefined;
			this.shadow = obj.shadow != undefined ? obj.shadow : undefined;

			this.rot = obj.rot != undefined ? obj.rot : undefined;
			this.flipH = obj.flipH != undefined ? obj.flipH : undefined;
			this.flipV = obj.flipV != undefined ? obj.flipV : undefined;
			this.resetCrop = obj.resetCrop != undefined ? obj.resetCrop : undefined;
			this.anchor = obj.anchor != undefined ? obj.anchor : undefined;

			this.protectionLockText = obj.protectionLockText;
			this.protectionLocked = obj.protectionLocked;
			this.protectionPrint = obj.protectionPrint;

			this.bSetOriginalSize = obj.bSetOriginalSize;
			this.transparent = obj.transparent;
			this.isCrop      = obj.isCrop;
			this.cropHeightCoefficient = obj.cropHeightCoefficient;
			this.cropWidthCoefficient = obj.cropWidthCoefficient;
		}
		else {
			this.CanBeFlow = true;
			this.Width = undefined;
			this.Height = undefined;
			this.WrappingStyle = undefined;
			this.Paddings = undefined;
			this.Position = undefined;
			this.PositionH = undefined;
			this.PositionV = undefined;

			this.SizeRelH = undefined;
			this.SizeRelV = undefined;

			this.Internal_Position = null;
			this.ImageUrl = null;
			this.Token = undefined;
			this.Locked = false;

			this.ChartProperties = null;
			this.ShapeProperties = null;
			this.SlicerProperties = null;

			this.ChangeLevel = null;
			this.Group = null;
			this.fromGroup = null;
			this.severalCharts = false;
			this.severalChartTypes = undefined;
			this.severalChartStyles = undefined;
			this.verticalTextAlign = undefined;
			this.vert = undefined;

			//oleObjects
			this.pluginGuid = undefined;
			this.pluginData = undefined;

			this.title = undefined;
			this.description = undefined;
			this.name = undefined;

			this.columnNumber = undefined;
			this.columnSpace = undefined;
			this.textFitType = undefined;
			this.vertOverflowType = undefined;


			this.rot = undefined;
			this.rotAdd = undefined;
			this.flipH = undefined;
			this.flipV = undefined;
			this.resetCrop = undefined;
			this.anchor = undefined;

			this.protectionLockText = undefined;
			this.protectionLocked = undefined;
			this.protectionPrint = undefined;
			this.bSetOriginalSize = undefined;

			this.transparent = undefined;
			this.isCrop      = undefined;
			this.cropHeightCoefficient = 1;
			this.cropWidthCoefficient = 1;
		}
	}

	asc_CImgProperty.prototype.constructor = asc_CImgProperty;
	asc_CImgProperty.prototype.asc_getChangeLevel = function () {
		return this.ChangeLevel;
	};
	asc_CImgProperty.prototype.asc_putChangeLevel = function (v) {
		this.ChangeLevel = v;
	};

	asc_CImgProperty.prototype.asc_getCanBeFlow = function () {
		return this.CanBeFlow;
	};
	asc_CImgProperty.prototype.asc_getWidth = function () {
		return this.Width;
	};
	asc_CImgProperty.prototype.asc_putWidth = function (v) {
		this.Width = v;
	};
	asc_CImgProperty.prototype.asc_getHeight = function () {
		return this.Height;
	};
	asc_CImgProperty.prototype.asc_putHeight = function (v) {
		this.Height = v;
	};
	asc_CImgProperty.prototype.asc_getWrappingStyle = function () {
		return this.WrappingStyle;
	};
	asc_CImgProperty.prototype.asc_putWrappingStyle = function (v) {
		this.WrappingStyle = v;
	};

	// Возвращается объект класса Asc.asc_CPaddings
	asc_CImgProperty.prototype.asc_getPaddings = function () {
		return this.Paddings;
	}; // Аргумент объект класса Asc.asc_CPaddings
	asc_CImgProperty.prototype.asc_putPaddings = function (v) {
		this.Paddings = v;
	};
	asc_CImgProperty.prototype.asc_getAllowOverlap = function () {
		return this.AllowOverlap;
	};
	asc_CImgProperty.prototype.asc_putAllowOverlap = function (v) {
		this.AllowOverlap = v;
	}; // Возвращается объект класса CPosition
	asc_CImgProperty.prototype.asc_getPosition = function () {
		return this.Position;
	}; // Аргумент объект класса CPosition
	asc_CImgProperty.prototype.asc_putPosition = function (v) {
		this.Position = v;
	};
	asc_CImgProperty.prototype.asc_getPositionH = function () {
		return this.PositionH;
	};
	asc_CImgProperty.prototype.asc_putPositionH = function (v) {
		this.PositionH = v;
	};
	asc_CImgProperty.prototype.asc_getPositionV = function () {
		return this.PositionV;
	};
	asc_CImgProperty.prototype.asc_putPositionV = function (v) {
		this.PositionV = v;
	};

	asc_CImgProperty.prototype.asc_getSizeRelH = function () {
		return this.SizeRelH;
	};

	asc_CImgProperty.prototype.asc_putSizeRelH = function (v) {
		this.SizeRelH = v;
	};

	asc_CImgProperty.prototype.asc_getSizeRelV = function () {
		return this.SizeRelV;
	};

	asc_CImgProperty.prototype.asc_putSizeRelV = function (v) {
		this.SizeRelV = v;
	};

	asc_CImgProperty.prototype.asc_getValue_X = function (RelativeFrom) {
		if (null != this.Internal_Position) {
			return this.Internal_Position.Calculate_X_Value(RelativeFrom);
		}
		return 0;
	};
	asc_CImgProperty.prototype.asc_getValue_Y = function (RelativeFrom) {
		if (null != this.Internal_Position) {
			return this.Internal_Position.Calculate_Y_Value(RelativeFrom);
		}
		return 0;
	};

	asc_CImgProperty.prototype.asc_getImageUrl = function () {
		return this.ImageUrl;
	};
	asc_CImgProperty.prototype.asc_putImageUrl = function (v, sToken) {
		this.ImageUrl = v;
		this.Token = sToken;
	};
	asc_CImgProperty.prototype.asc_getGroup = function () {
		return this.Group;
	};
	asc_CImgProperty.prototype.asc_putGroup = function (v) {
		this.Group = v;
	};
	asc_CImgProperty.prototype.asc_getFromGroup = function () {
		return this.fromGroup;
	};
	asc_CImgProperty.prototype.asc_putFromGroup = function (v) {
		this.fromGroup = v;
	};

	asc_CImgProperty.prototype.asc_getisChartProps = function () {
		return this.isChartProps;
	};
	asc_CImgProperty.prototype.asc_putisChartPross = function (v) {
		this.isChartProps = v;
	};

	asc_CImgProperty.prototype.asc_getSeveralCharts = function () {
		return this.severalCharts;
	};
	asc_CImgProperty.prototype.asc_putSeveralCharts = function (v) {
		this.severalCharts = v;
	};
	asc_CImgProperty.prototype.asc_getSeveralChartTypes = function () {
		return this.severalChartTypes;
	};
	asc_CImgProperty.prototype.asc_putSeveralChartTypes = function (v) {
		this.severalChartTypes = v;
	};

	asc_CImgProperty.prototype.asc_getSeveralChartStyles = function () {
		return this.severalChartStyles;
	};
	asc_CImgProperty.prototype.asc_putSeveralChartStyles = function (v) {
		this.severalChartStyles = v;
	};

	asc_CImgProperty.prototype.asc_getVerticalTextAlign = function () {
		return this.verticalTextAlign;
	};
	asc_CImgProperty.prototype.asc_putVerticalTextAlign = function (v) {
		this.verticalTextAlign = v;
	};
	asc_CImgProperty.prototype.asc_getVert = function () {
		return this.vert;
	};
	asc_CImgProperty.prototype.asc_putVert = function (v) {
		this.vert = v;
	};

	asc_CImgProperty.prototype.asc_getLocked = function () {
		return this.Locked;
	};
	asc_CImgProperty.prototype.asc_getLockAspect = function () {
		return this.lockAspect;
	};
	asc_CImgProperty.prototype.asc_putLockAspect = function (v) {
		this.lockAspect = v;
	};
	asc_CImgProperty.prototype.asc_getChartProperties = function () {
		return this.ChartProperties;
	};
	asc_CImgProperty.prototype.asc_putChartProperties = function (v) {
		this.ChartProperties = v;
	};
	asc_CImgProperty.prototype.asc_getShapeProperties = function () {
		return this.ShapeProperties;
	};
	asc_CImgProperty.prototype.asc_putShapeProperties = function (v) {
		this.ShapeProperties = v;
	};
	asc_CImgProperty.prototype.asc_getSlicerProperties = function () {
		return this.SlicerProperties;
	};
	asc_CImgProperty.prototype.asc_putSlicerProperties = function (v) {
		this.SlicerProperties = v;
	};

	asc_CImgProperty.prototype.asc_getOriginSize = function (api) {
		if (this.ImageUrl === null) {
			return new asc_CImageSize(50, 50, false);
		}

		let origW = 0;
		let origH = 0;

		if(window["IS_NATIVE_EDITOR"]) {
			let fGetOriginalImageSize = window["native"]["GetOriginalImageSize"] ||
				window["native"]["DD_GetOriginalImageSize"];
			let sizes = fGetOriginalImageSize.call(window["native"], this.ImageUrl);
			let w = sizes[0];
			let h = sizes[1];
			let isN = AscFormat.isRealNumber;
			if (isN(w) && isN(h)) {
				origW = w;
				origH = h;
			}
		}
		else {
			let _image = api.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(this.ImageUrl)];
			if (_image != undefined && _image.Image != null && _image.Status == window['AscFonts'].ImageLoadStatus.Complete) {
				origW = _image.Image.width;
				origH = _image.Image.height;
			}
			else if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["GetImageOriginalSize"]) {
				let _size = window["AscDesktopEditor"]["GetImageOriginalSize"](this.ImageUrl);
				if (_size["W"] != 0 && _size["H"] != 0) {
					origW = _size["W"];
					origH = _size["H"];
				}
			}
		}
		if (origW != 0 && origH != 0) {
			let __w = Math.max((origW * AscCommon.g_dKoef_pix_to_mm), 1);
			let __h = Math.max((origH * AscCommon.g_dKoef_pix_to_mm), 1);

			return new asc_CImageSize(__w, __h, true);
		}
		return new asc_CImageSize(50, 50, false);
	};

	asc_CImgProperty.prototype.asc_getCropOriginSize = function(api) {
		const oSizes = this.asc_getOriginSize(api);
		oSizes.Width *= this.cropWidthCoefficient;
		oSizes.Height *= this.cropHeightCoefficient;
		return oSizes;
	};

	//oleObjects
	asc_CImgProperty.prototype.asc_getPluginGuid = function () {
		return this.pluginGuid;
	};

	asc_CImgProperty.prototype.asc_putPluginGuid = function (v) {
		this.pluginGuid = v;
	};

	asc_CImgProperty.prototype.asc_getPluginData = function () {
		return this.pluginData;
	};

	asc_CImgProperty.prototype.asc_putPluginData = function (v) {
		this.pluginData = v;
	};

	asc_CImgProperty.prototype.asc_getTitle = function () {
		return this.title;
	};

	asc_CImgProperty.prototype.asc_putTitle = function (v) {
		this.title = v;
	};

	asc_CImgProperty.prototype.asc_getDescription = function () {
		return this.description;
	};

	asc_CImgProperty.prototype.asc_putDescription = function (v) {
		this.description = v;
	};
	asc_CImgProperty.prototype.asc_getName = function () {
		return this.name;
	};

	asc_CImgProperty.prototype.asc_putName = function (v) {
		this.name = v;
	};

	asc_CImgProperty.prototype.asc_getColumnNumber = function () {
		return this.columnNumber;
	};

	asc_CImgProperty.prototype.asc_putColumnNumber = function (v) {
		this.columnNumber = v;
	};

	asc_CImgProperty.prototype.asc_getColumnSpace = function () {
		return this.columnSpace;
	};

	asc_CImgProperty.prototype.asc_getTextFitType = function () {
		return this.textFitType;
	};
	asc_CImgProperty.prototype.asc_getVertOverflowType = function () {
		return this.vertOverflowType;
	};

	asc_CImgProperty.prototype.asc_putColumnSpace = function (v) {
		this.columnSpace = v;
	};
	asc_CImgProperty.prototype.asc_putTextFitType = function (v) {
		this.textFitType = v;
	};
	asc_CImgProperty.prototype.asc_putVertOverflowType = function (v) {
		this.vertOverflowType = v;
	};

	asc_CImgProperty.prototype.asc_getSignatureId = function () {
		if (this.ShapeProperties) return this.ShapeProperties.asc_getSignatureId();
		return undefined;
	};

	asc_CImgProperty.prototype.asc_getRot = function () {
		return this.rot;
	};

	asc_CImgProperty.prototype.asc_putRot = function (v) {
		this.rot = v;
	};
	asc_CImgProperty.prototype.asc_getRotAdd = function () {
		return this.rotAdd;
	};

	asc_CImgProperty.prototype.asc_putRotAdd = function (v) {
		this.rotAdd = v;
	};

	asc_CImgProperty.prototype.asc_getFlipH = function () {
		return this.flipH;
	};

	asc_CImgProperty.prototype.asc_putFlipH = function (v) {
		this.flipH = v;
	};
	asc_CImgProperty.prototype.asc_getFlipHInvert = function () {
		return this.flipHInvert;
	};

	asc_CImgProperty.prototype.asc_putFlipHInvert = function (v) {
		this.flipHInvert = v;
	};

	asc_CImgProperty.prototype.asc_getFlipV = function () {
		return this.flipV;
	};

	asc_CImgProperty.prototype.asc_putFlipV = function (v) {
		this.flipV = v;
	};
	asc_CImgProperty.prototype.asc_getFlipVInvert = function () {
		return this.flipVInvert;
	};

	asc_CImgProperty.prototype.asc_putFlipVInvert = function (v) {
		this.flipVInvert = v;
	};
	asc_CImgProperty.prototype.asc_putResetCrop = function (v) {
		this.resetCrop = v;
	};
	asc_CImgProperty.prototype.asc_getShadow = function () {
		return this.shadow;
	};

	asc_CImgProperty.prototype.asc_putShadow = function (v) {
		this.shadow = v;
	};
	asc_CImgProperty.prototype.asc_getAnchor = function () {
		return this.anchor;
	};

	asc_CImgProperty.prototype.asc_putAnchor = function (v) {
		this.anchor = v;
	};
	asc_CImgProperty.prototype.asc_getProtectionLockText = function () {
		return this.protectionLockText;
	};
	asc_CImgProperty.prototype.asc_putProtectionLockText = function (v) {
		this.protectionLockText = v;
	};
	asc_CImgProperty.prototype.asc_getProtectionLocked = function () {
		return this.protectionLocked;
	};

	asc_CImgProperty.prototype.asc_putProtectionLocked = function (v) {
		this.protectionLocked = v;
	};
	asc_CImgProperty.prototype.asc_getProtectionPrint = function () {
		return this.protectionPrint;
	};
	asc_CImgProperty.prototype.asc_putProtectionPrint = function (v) {
		this.protectionPrint = v;
	};
	asc_CImgProperty.prototype.asc_getTransparent = function () {
		return this.transparent;
	};
	asc_CImgProperty.prototype.asc_putTransparent = function (v) {
		this.transparent = v;
	};

	asc_CImgProperty.prototype.asc_getIsCrop = function () {
		return this.isCrop;
	};
	asc_CImgProperty.prototype.asc_getCropHeightCoefficient = function () {
		return this.cropHeightCoefficient;
	};
	asc_CImgProperty.prototype.asc_putCropHeightCoefficient = function (v) {
		this.cropHeightCoefficient = v;
	};
	asc_CImgProperty.prototype.asc_getCropWidthCoefficient = function () {
		return this.cropWidthCoefficient;
	};
	asc_CImgProperty.prototype.asc_putCropWidthCoefficient = function (v) {
		this.cropWidthCoefficient = v;
	};

	/** @constructor */
	function asc_CSelectedObject(type, val) {
		this.Type = (undefined != type) ? type : null;
		this.Value = (undefined != val) ? val : null;
	}

	asc_CSelectedObject.prototype.asc_getObjectType = function () {
		return this.Type;
	};
	asc_CSelectedObject.prototype.asc_getObjectValue = function () {
		return this.Value;
	};



	/** @constructor */
	function asc_CShapeFill() {
		this.type = null;
		this.fill = null;
		this.transparent = null;
	}

	asc_CShapeFill.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CShapeFill.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CShapeFill.prototype.asc_getFill = function () {
		return this.fill;
	};
	asc_CShapeFill.prototype.asc_putFill = function (v) {
		this.fill = v;
	};
	asc_CShapeFill.prototype.asc_getTransparent = function () {
		return this.transparent;
	};
	asc_CShapeFill.prototype.asc_putTransparent = function (v) {
		this.transparent = v;
	};
	asc_CShapeFill.prototype.asc_CheckForseSet = function () {
		if (null != this.transparent) {
			return true;
		}
		if (null != this.fill && this.fill.Positions != null) {
			return true;
		}
		return false;
	};

	/** @constructor */
	function asc_CFillBlip() {
		this.type = c_oAscFillBlipType.STRETCH;
		this.url = "";
		this.token = undefined;
		this.texture_id = null;
	}

	asc_CFillBlip.prototype.asc_getType = function () {
		return this.type
	};
	asc_CFillBlip.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CFillBlip.prototype.asc_getUrl = function () {
		return this.url;
	};
	asc_CFillBlip.prototype.asc_putUrl = function (v, sToken) {
		this.url = v;
		this.token = sToken;
	};
	asc_CFillBlip.prototype.asc_getTextureId = function () {
		return this.texture_id;
	};
	asc_CFillBlip.prototype.asc_putTextureId = function (v) {
		this.texture_id = v;
	};

	/** @constructor */
	function asc_CFillHatch() {
		this.PatternType = undefined;
		this.fgClr = undefined;
		this.bgClr = undefined;
	}

	asc_CFillHatch.prototype.asc_getPatternType = function () {
		return this.PatternType;
	};
	asc_CFillHatch.prototype.asc_putPatternType = function (v) {
		this.PatternType = v;
	};
	asc_CFillHatch.prototype.asc_getColorFg = function () {
		return this.fgClr;
	};
	asc_CFillHatch.prototype.asc_putColorFg = function (v) {
		this.fgClr = v;
	};
	asc_CFillHatch.prototype.asc_getColorBg = function () {
		return this.bgClr;
	};
	asc_CFillHatch.prototype.asc_putColorBg = function (v) {
		this.bgClr = v;
	};

	/** @constructor */
	function asc_CFillGrad() {
		this.Colors = undefined;
		this.Positions = undefined;
		this.GradType = 0;

		this.LinearAngle = undefined;
		this.LinearScale = true;

		this.PathType = 0;
	}

	asc_CFillGrad.prototype.asc_getColors = function () {
		return this.Colors;
	};
	asc_CFillGrad.prototype.asc_putColors = function (v) {
		this.Colors = v;
	};
	asc_CFillGrad.prototype.asc_getPositions = function () {
		return this.Positions;
	};
	asc_CFillGrad.prototype.asc_putPositions = function (v) {
		this.Positions = v;
	};
	asc_CFillGrad.prototype.asc_getGradType = function () {
		return this.GradType;
	};
	asc_CFillGrad.prototype.asc_putGradType = function (v) {
		this.GradType = v;
	};
	asc_CFillGrad.prototype.asc_getLinearAngle = function () {
		return this.LinearAngle;
	};
	asc_CFillGrad.prototype.asc_putLinearAngle = function (v) {
		this.LinearAngle = v;
	};
	asc_CFillGrad.prototype.asc_getLinearScale = function () {
		return this.LinearScale;
	};
	asc_CFillGrad.prototype.asc_putLinearScale = function (v) {
		this.LinearScale = v;
	};
	asc_CFillGrad.prototype.asc_getPathType = function () {
		return this.PathType;
	};
	asc_CFillGrad.prototype.asc_putPathType = function (v) {
		this.PathType = v;
	};

	/** @constructor */
	function asc_CFillSolid() {
		this.color = new asc_CColor();
	}

	asc_CFillSolid.prototype.asc_getColor = function () {
		return this.color
	};
	asc_CFillSolid.prototype.asc_putColor = function (v) {
		this.color = v;
	};


	/** @constructor */
	function asc_CStroke() {
		this.type = null;
		this.width = null;
		this.color = null;
		this.prstDash = null;

		this.LineJoin = null;
		this.LineCap = null;

		this.LineBeginStyle = null;
		this.LineBeginSize = null;

		this.LineEndStyle = null;
		this.LineEndSize = null;

		this.canChangeArrows = false;
		this.transparent = null;
	}

	asc_CStroke.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CStroke.prototype.asc_putType = function (v) {
		this.type = v;
	};
	asc_CStroke.prototype.asc_getWidth = function () {
		return this.width;
	};
	asc_CStroke.prototype.asc_putWidth = function (v) {
		this.width = v;
	};
	asc_CStroke.prototype.asc_getColor = function () {
		return this.color;
	};
	asc_CStroke.prototype.asc_putColor = function (v) {
		this.color = v;
	};
	asc_CStroke.prototype.asc_getLinejoin = function () {
		return this.LineJoin;
	};
	asc_CStroke.prototype.asc_putLinejoin = function (v) {
		this.LineJoin = v;
	};
	asc_CStroke.prototype.asc_getLinecap = function () {
		return this.LineCap;
	};
	asc_CStroke.prototype.asc_putLinecap = function (v) {
		this.LineCap = v;
	};
	asc_CStroke.prototype.asc_getLinebeginstyle = function () {
		return this.LineBeginStyle;
	};
	asc_CStroke.prototype.asc_putLinebeginstyle = function (v) {
		this.LineBeginStyle = v;
	};
	asc_CStroke.prototype.asc_getLinebeginsize = function () {
		return this.LineBeginSize;
	};
	asc_CStroke.prototype.asc_putLinebeginsize = function (v) {
		this.LineBeginSize = v;
	};
	asc_CStroke.prototype.asc_getLineendstyle = function () {
		return this.LineEndStyle;
	};
	asc_CStroke.prototype.asc_putLineendstyle = function (v) {
		this.LineEndStyle = v;
	};
	asc_CStroke.prototype.asc_getLineendsize = function () {
		return this.LineEndSize;
	};
	asc_CStroke.prototype.asc_putLineendsize = function (v) {
		this.LineEndSize = v;
	};
	asc_CStroke.prototype.asc_getCanChangeArrows = function () {
		return this.canChangeArrows;
	};
	asc_CStroke.prototype.asc_getTransparent = function () {
		return this.transparent;
	};
	asc_CStroke.prototype.asc_putTransparent = function (v) {
		this.transparent = v;
	};
	asc_CStroke.prototype.asc_putPrstDash = function (v) {
		this.prstDash = v;
	};
	asc_CStroke.prototype.asc_getPrstDash = function () {
		return this.prstDash;
	};

	// цвет. может быть трех типов:
	// c_oAscColor.COLOR_TYPE_SRGB		: value - не учитывается
	// c_oAscColor.COLOR_TYPE_PRST		: value - имя стандартного цвета (map_prst_color)
	// c_oAscColor.COLOR_TYPE_SCHEME	: value - тип цвета в схеме
	// c_oAscColor.COLOR_TYPE_SYS		: конвертируется в srgb
	function CAscColorScheme() {
		this.colors = [];
		this.name = "";
		this.scheme = null;
		this.summ = 0;
	}

	CAscColorScheme.prototype.get_colors = function () {
		return this.colors;
	};
	CAscColorScheme.prototype.get_name = function () {
		return AscCommon.translateManager.getValue(this.name);
	};
	CAscColorScheme.prototype.get_dk1 = function () {
		return this.colors[0];
	};
	CAscColorScheme.prototype.get_lt1 = function () {
		return this.colors[1];
	};
	CAscColorScheme.prototype.get_dk2 = function () {
		return this.colors[2];
	};
	CAscColorScheme.prototype.get_lt2 = function () {
		return this.colors[3];
	};
	CAscColorScheme.prototype.get_accent1 = function () {
		return this.colors[4];
	};
	CAscColorScheme.prototype.get_accent2 = function () {
		return this.colors[5];
	};
	CAscColorScheme.prototype.get_accent3 = function () {
		return this.colors[6];
	};
	CAscColorScheme.prototype.get_accent4 = function () {
		return this.colors[7];
	};
	CAscColorScheme.prototype.get_accent5 = function () {
		return this.colors[8];
	};
	CAscColorScheme.prototype.get_accent6 = function () {
		return this.colors[9];
	};
	CAscColorScheme.prototype.get_hlink = function () {
		return this.colors[10];
	};
	CAscColorScheme.prototype.get_folHlink = function () {
		return this.colors[11];
	};
	CAscColorScheme.prototype.putColor = function (color) {
		this.colors.push(color);
		this.summ += color.getVal();
	};
	CAscColorScheme.prototype.isEqual = function (oColorScheme) {
		if (this.summ === oColorScheme.summ) {
			for (let i = 0; i < this.colors.length; ++i) {
				let oColor1 = this.colors[i];
				let oColor2 = oColorScheme.colors[i];
				if (!(!oColor1 && !oColor2 || oColor2 && oColor2 && oColor1.Compare(oColor2))) {
					return false;
				}
			}
			return this.name === oColorScheme.name;
		}
		return false;
	};


	//-----------------------------------------------------------------
	// События движения мыши
	//-----------------------------------------------------------------
	function CMouseMoveData(obj) {
		if (obj) {
			this.Type = (undefined != obj.Type) ? obj.Type : c_oAscMouseMoveDataTypes.Common;
			this.X_abs = (undefined != obj.X_abs) ? obj.X_abs : 0;
			this.Y_abs = (undefined != obj.Y_abs) ? obj.Y_abs : 0;
			this.EyedropperColor = (undefined != obj.EyedropperColor) ? obj.EyedropperColor : undefined;
			this.PlaceholderType = obj.PlaceholderType;
			switch (this.Type) {
				case c_oAscMouseMoveDataTypes.Hyperlink : {
					this.Hyperlink = (undefined != obj.PageNum) ? obj.PageNum : 0;
					break;
				}

				case c_oAscMouseMoveDataTypes.LockedObject : {
					this.UserId = (undefined != obj.UserId) ? obj.UserId : "";
					this.HaveChanges = (undefined != obj.HaveChanges) ? obj.HaveChanges : false;
					this.LockedObjectType = (undefined != obj.LockedObjectType) ? obj.LockedObjectType : Asc.c_oAscMouseMoveLockedObjectType.Common;
					break;
				}
				case c_oAscMouseMoveDataTypes.Footnote: {
					this.Text = "";
					this.Number = 1;
					break;
				}
				case c_oAscMouseMoveDataTypes.Review: {
					this.ReviewChange = obj && obj.ReviewChange ? obj.ReviewChange : null;
					break;
				}
			}
		}
		else {
			this.Type = c_oAscMouseMoveDataTypes.Common;
			this.X_abs = 0;
			this.Y_abs = 0;
			this.EyedropperColor = undefined;
		}
	}

	CMouseMoveData.prototype.get_Type = function () {
		return this.Type;
	};
	CMouseMoveData.prototype.get_X = function () {
		return this.X_abs;
	};
	CMouseMoveData.prototype.get_Y = function () {
		return this.Y_abs;
	};
	CMouseMoveData.prototype.get_Hyperlink = function () {
		return this.Hyperlink;
	};
	CMouseMoveData.prototype.get_UserId = function () {
		return this.UserId;
	};
	CMouseMoveData.prototype.get_HaveChanges = function () {
		return this.HaveChanges;
	};
	CMouseMoveData.prototype.get_LockedObjectType = function () {
		return this.LockedObjectType;
	};
	CMouseMoveData.prototype.get_FootnoteText = function () {
		return this.Text;
	};
	CMouseMoveData.prototype.get_FootnoteNumber = function () {
		return this.Number;
	};
	CMouseMoveData.prototype.get_FormHelpText = function () {
		return this.Text;
	};
	CMouseMoveData.prototype.get_ReviewChange = function () {
		return this.ReviewChange;
	};
	CMouseMoveData.prototype.get_EyedropperColor = function () {
		return this.EyedropperColor;
	};
	CMouseMoveData.prototype.get_PlaceholderType = function () {
		return this.PlaceholderType;
	};
	CMouseMoveData.prototype.get_EffectText = function () {
		return this.EffectText;
	};
	CMouseMoveData.prototype.get_EffectDescription = function () {
		return this.EffectDescription;
	};


	/**
	 * Класс для работы с интерфейсом для гиперссылок
	 * @param obj
	 * @constructor
	 */
	function CHyperlinkProperty(obj) {
		if (obj) {
			this.Text = (undefined != obj.Text) ? obj.Text : null;
			this.Value = (undefined != obj.Value) ? obj.Value : "";
			this.ToolTip = (undefined != obj.ToolTip) ? obj.ToolTip : "";
			this.Class = (undefined !== obj.Class) ? obj.Class : null;
			this.Anchor = (undefined !== obj.Anchor) ? obj.Anchor : null;
			this.Heading = (obj.Heading ? obj.Heading : null);
		}
		else {
			this.Text = null;
			this.Value = "";
			this.ToolTip = "";
			this.Class = null;
			this.Anchor = null;
			this.Heading = null;
		}
	}

	CHyperlinkProperty.prototype.get_Value = function () {
		return this.Value;
	};
	CHyperlinkProperty.prototype.put_Value = function (v) {
		this.Value = v;
	};
	CHyperlinkProperty.prototype.get_ToolTip = function () {
		return this.ToolTip;
	};
	CHyperlinkProperty.prototype.put_ToolTip = function (v) {
		this.ToolTip = v ? v.slice(0, Asc.c_oAscMaxTooltipLength) : v;
	};
	CHyperlinkProperty.prototype.get_Text = function () {
		return this.Text;
	};
	CHyperlinkProperty.prototype.put_Text = function (v) {
		this.Text = v;
	};
	CHyperlinkProperty.prototype.put_InternalHyperlink = function (oClass) {
		this.Class = oClass;
	};
	CHyperlinkProperty.prototype.get_InternalHyperlink = function () {
		return this.Class;
	};
	CHyperlinkProperty.prototype.is_TopOfDocument = function () {
		return (this.Anchor === "_top");
	};
	CHyperlinkProperty.prototype.put_TopOfDocument = function () {
		this.Anchor = "_top";
	};
	CHyperlinkProperty.prototype.get_Bookmark = function () {
		return this.Anchor;
	};
	CHyperlinkProperty.prototype.put_Bookmark = function (sBookmark) {
		this.Anchor = sBookmark;
	};
	CHyperlinkProperty.prototype.is_Heading = function () {
		return (this.Heading instanceof AscWord.Paragraph ? true : false)
	};
	CHyperlinkProperty.prototype.put_Heading = function (oParagraph) {
		this.Heading = oParagraph;
	};
	CHyperlinkProperty.prototype.get_Heading = function () {
		return this.Heading;
	};

	window['Asc']['CHyperlinkProperty'] = window['Asc'].CHyperlinkProperty = CHyperlinkProperty;
	CHyperlinkProperty.prototype['get_Value'] = CHyperlinkProperty.prototype.get_Value;
	CHyperlinkProperty.prototype['put_Value'] = CHyperlinkProperty.prototype.put_Value;
	CHyperlinkProperty.prototype['get_ToolTip'] = CHyperlinkProperty.prototype.get_ToolTip;
	CHyperlinkProperty.prototype['put_ToolTip'] = CHyperlinkProperty.prototype.put_ToolTip;
	CHyperlinkProperty.prototype['get_Text'] = CHyperlinkProperty.prototype.get_Text;
	CHyperlinkProperty.prototype['put_Text'] = CHyperlinkProperty.prototype.put_Text;
	CHyperlinkProperty.prototype['get_InternalHyperlink'] = CHyperlinkProperty.prototype.get_InternalHyperlink;
	CHyperlinkProperty.prototype['put_InternalHyperlink'] = CHyperlinkProperty.prototype.put_InternalHyperlink;
	CHyperlinkProperty.prototype['is_TopOfDocument'] = CHyperlinkProperty.prototype.is_TopOfDocument;
	CHyperlinkProperty.prototype['put_TopOfDocument'] = CHyperlinkProperty.prototype.put_TopOfDocument;
	CHyperlinkProperty.prototype['get_Bookmark'] = CHyperlinkProperty.prototype.get_Bookmark;
	CHyperlinkProperty.prototype['put_Bookmark'] = CHyperlinkProperty.prototype.put_Bookmark;
	CHyperlinkProperty.prototype['is_Heading'] = CHyperlinkProperty.prototype.is_Heading;
	CHyperlinkProperty.prototype['put_Heading'] = CHyperlinkProperty.prototype.put_Heading;
	CHyperlinkProperty.prototype['get_Heading'] = CHyperlinkProperty.prototype.get_Heading;


	/**
	 * @property {string|null} Id
	 * @property {string|null} FullName
	 * @property {string|null} FirstName
	 * @property {string|null} LastName
	 * @property {boolean|null} IsAnonymousUser
	 * @constructor
	 *  */
	function asc_CUserInfo() {
		this.Id = null;
		this.FullName = null;
		this.FirstName = null;
		this.LastName = null;
		this.IsAnonymousUser = false;
	}
	asc_CUserInfo.prototype.clone = function () {
		let res = new asc_CUserInfo();
		res.Id = this.Id;
		res.FullName = this.FullName;
		res.FirstName = this.FirstName;
		res.LastName = this.LastName;
		res.IsAnonymousUser = this.IsAnonymousUser;
		return res;
	};
	asc_CUserInfo.prototype.asc_putId = asc_CUserInfo.prototype.put_Id = function (v) {
		this.Id = v;
	};
	asc_CUserInfo.prototype.asc_getId = asc_CUserInfo.prototype.get_Id = function () {
		return this.Id;
	};
	asc_CUserInfo.prototype.asc_putFullName = asc_CUserInfo.prototype.put_FullName = function (v) {
		this.FullName = v;
	};
	asc_CUserInfo.prototype.asc_getFullName = asc_CUserInfo.prototype.get_FullName = function () {
		return this.FullName;
	};
	asc_CUserInfo.prototype.asc_putFirstName = asc_CUserInfo.prototype.put_FirstName = function (v) {
		this.FirstName = v;
	};
	asc_CUserInfo.prototype.asc_getFirstName = asc_CUserInfo.prototype.get_FirstName = function () {
		return this.FirstName;
	};
	asc_CUserInfo.prototype.asc_putLastName = asc_CUserInfo.prototype.put_LastName = function (v) {
		this.LastName = v;
	};
	asc_CUserInfo.prototype.asc_getLastName = asc_CUserInfo.prototype.get_LastName = function () {
		return this.LastName;
	};
	asc_CUserInfo.prototype.asc_getIsAnonymousUser = asc_CUserInfo.prototype.get_IsAnonymousUser = function () {
		return this.IsAnonymousUser;
	};
	asc_CUserInfo.prototype.asc_putIsAnonymousUser = asc_CUserInfo.prototype.put_IsAnonymousUser = function (v) {
		this.IsAnonymousUser = v;
	};

	/**
	 * @property {string|null} Id
	 * @property {string|null} Url
	 * @property {string|null} Title
	 * @property {string|null} Format
	 * @property {string|null} VKey
	 * @property {string|null} Token
	 * @property {asc_CUserInfo|null} UserInfo
	 * @property {object|null} Options
	 * @property {string|null} CallbackUrl
	 * @property {object|null} TemplateReplacement
	 * @property {string|null} Mode
	 * @property {object|null} Permissions
	 * @property {string|null} Lang
	 * @property {boolean|null} OfflineApp
	 * @property {boolean|undefined} Encrypted
	 * @property {object|undefined} EncryptedInfo
	 * @property {boolean|null} IsEnabledPlugins
	 * @property {boolean|null} IsEnabledMacroses
	 * @property {boolean|null} IsWebOpening
	 * @property {boolean|null} SupportsOnSaveDocument
	 * @property {object|null} Wopi
	 * @property {string|null} shardkey
	 * @property {object|null} ReferenceData
	 * @constructor
	 * */
	function asc_CDocInfo() {
		this.Id = null;
		this.Url = null;
		this.Title = null;
		this.Format = null;
		this.VKey = null;
		this.Token = null;
		this.UserInfo = null;
		this.Options = null;
		this.CallbackUrl = null;
		this.TemplateReplacement = null;
		this.Mode = null;
		this.Permissions = null;
		this.Lang = null;
		this.OfflineApp = false;
		this.Encrypted;
		this.EncryptedInfo;
		this.IsEnabledPlugins = true;
		this.IsEnabledMacroses = true;
		this.IsWebOpening = false;
		this.SupportsOnSaveDocument = false;
		this.Wopi = null;
		this.shardkey = null;

		//for external reference
		this.ReferenceData = null;

		this.showVerticalScroll = null;
		this.showHorizontalScroll = null;
	}

	prot = asc_CDocInfo.prototype;
	prot.clone = function () {
		let res = new asc_CDocInfo();
		res.Id = this.Id;
		res.Url = this.Url;
		res.Title = this.Title;
		res.Format = this.Format;
		res.VKey = this.VKey;
		res.Token = this.Token;
		res.UserInfo = this.UserInfo ? this.UserInfo.clone() : null;
		res.Options = this.Options ? JSON.parse(JSON.stringify(this.Options)) : null;
		res.CallbackUrl = this.CallbackUrl;
		res.TemplateReplacement = this.TemplateReplacement ? JSON.parse(JSON.stringify(this.TemplateReplacement)) : null;
		res.Mode = this.Mode;
		res.Permissions = this.Permissions ? JSON.parse(JSON.stringify(this.Permissions)) : null;
		res.Lang = this.Lang;
		res.OfflineApp = this.OfflineApp;
		res.Encrypted = this.Encrypted;
		res.EncryptedInfo = this.EncryptedInfo ? JSON.parse(JSON.stringify(this.EncryptedInfo)) : undefined;
		res.IsEnabledPlugins = this.IsEnabledPlugins;
		res.IsEnabledMacroses = this.IsEnabledMacroses;
		res.IsWebOpening = this.IsWebOpening;
		res.SupportsOnSaveDocument = this.SupportsOnSaveDocument;
		res.Wopi = this.Wopi ? JSON.parse(JSON.stringify(this.Wopi)) : null;
		res.shardkey = this.shardkey;
		res.ReferenceData = this.ReferenceData ? JSON.parse(JSON.stringify(this.ReferenceData)) : null;
		return res;
	};
	prot.extendWithWopiParams = function(data) {
		let docInfo = this.clone();
		//like in web-apps/apps/api/wopi/editor-wopi.ejs and onRefreshFile(web-apps/apps/documenteditor/main/app/controller/Main.js)
		let key = data["key"];
		let userAuth = data["userAuth"];
		let fileInfo = data["fileInfo"];
		let token = data["token"];

		docInfo.put_Id(key);
		if (fileInfo["FileUrl"]) {
			docInfo.put_Url(fileInfo["FileUrl"]);
		} else if (fileInfo["TemplateSource"]) {
			docInfo.put_Url(fileInfo["TemplateSource"]);
		} else if (userAuth) {
			docInfo.put_Url(userAuth["wopiSrc"] + "/contents?access_token=" + userAuth["access_token"]);
		}
		docInfo.put_Title(fileInfo["BreadcrumbDocName"] || fileInfo["BaseFileName"]);
		docInfo.put_CallbackUrl(JSON.stringify(userAuth));
		docInfo.put_Token(token);
		//todo does userInfo can change? (IsAnonymousUser)

		let fileType = fileInfo["BaseFileName"] ? fileInfo["BaseFileName"].substr(fileInfo["BaseFileName"].lastIndexOf('.') + 1) : "";
		fileType = fileInfo["FileExtension"] ? fileInfo["FileExtension"].substr(1) : fileType;
		fileType = fileType.toLowerCase();
		docInfo.put_Format(fileType);
		docInfo.put_Mode(userAuth["mode"]);
		//todo does permissions can change? (formsubmit, dchat)
		docInfo.put_CoEditingMode(userAuth["mode"] !== "view" ? "fast" : "strict");

		docInfo.put_Wopi({
			"FileNameMaxLength": fileInfo["FileNameMaxLength"] && fileInfo["FileNameMaxLength"] > 0 ? fileInfo["FileNameMaxLength"] : 250,
			"WOPISrc": userAuth["wopiSrc"],
			"UserSessionId": userAuth["userSessionId"],
			"Version": fileInfo["Version"],
			"LastModifiedTime": fileInfo["LastModifiedTime"]
		});
		return docInfo;
	};
	prot.isFormatWithForms = function () {
		return this.Format === "oform" || this.Format === "docxf" || this.Format === "pdf";
	};
	prot.get_Id = prot.asc_getId = function () {
		return this.Id
	};
	prot.put_Id = prot.asc_putId = function (v) {
		this.Id = v;
	};
	prot.get_Url = prot.asc_getUrl = function () {
		return this.Url;
	};
	prot.put_Url = prot.asc_putUrl = function (v) {
		this.Url = v;
	};
	prot.get_DirectUrl = prot.asc_getDirectUrl = function () {
		return this.DirectUrl;
	};
	prot.put_DirectUrl = prot.asc_putDirectUrl = function (v) {
		this.DirectUrl = v;
	};
	prot.get_Title = prot.asc_getTitle = function () {
		return this.Title;
	};
	prot.put_Title = prot.asc_putTitle = function (v) {
		this.Title = v;
	};
	prot.get_Format = prot.asc_getFormat = function () {
		return this.Format;
	};
	prot.put_Format = prot.asc_putFormat = function (v) {
		this.Format = v;
	};
	prot.get_VKey = prot.asc_getVKey = function () {
		return this.VKey;
	};
	prot.put_VKey = prot.asc_putVKey = function (v) {
		this.VKey = v;
	};
	prot.get_Token = prot.asc_getToken = function () {
		return this.Token;
	};
	prot.put_Token = prot.asc_putToken = function (v) {
		this.Token = v;
	};
	prot.get_OfflineApp = function () {
		return this.OfflineApp;
	};
	prot.put_OfflineApp = function (v) {
		this.OfflineApp = v;
	};
	prot.get_UserId = prot.asc_getUserId = function () {
		return (this.UserInfo ? this.UserInfo.get_Id() : null);
	};
	prot.get_UserName = prot.asc_getUserName = function () {
		return (this.UserInfo ? this.UserInfo.get_FullName() : null);
	};
	prot.get_FirstName = prot.asc_getFirstName = function () {
		return (this.UserInfo ? this.UserInfo.get_FirstName() : null);
	};
	prot.get_LastName = prot.asc_getLastName = function () {
		return (this.UserInfo ? this.UserInfo.get_LastName() : null);
	};
	prot.get_IsAnonymousUser = prot.get_IsAnonymousUser = function () {
		return (this.UserInfo ? this.UserInfo.get_IsAnonymousUser() : null);
	};
	prot.get_Options = prot.asc_getOptions = function () {
		return this.Options;
	};
	prot.put_Options = prot.asc_putOptions = function (v) {
		this.Options = v;
	};
	prot.get_CallbackUrl = prot.asc_getCallbackUrl = function () {
		return this.CallbackUrl;
	};
	prot.put_CallbackUrl = prot.asc_putCallbackUrl = function (v) {
		this.CallbackUrl = v;
	};
	prot.get_TemplateReplacement = prot.asc_getTemplateReplacement = function () {
		return this.TemplateReplacement;
	};
	prot.put_TemplateReplacement = prot.asc_putTemplateReplacement = function (v) {
		this.TemplateReplacement = v;
	};
	prot.get_UserInfo = prot.asc_getUserInfo = function () {
		return this.UserInfo;
	};
	prot.put_UserInfo = prot.asc_putUserInfo = function (v) {
		this.UserInfo = v;
	};
	prot.get_Mode = prot.asc_getMode = function () {
		return this.Mode;
	};
	prot.put_Mode = prot.asc_putMode = function (v) {
		this.Mode = v;
	};
	prot.get_Permissions = prot.asc_getPermissions = function () {
		return this.Permissions;
	};
	prot.put_Permissions = prot.asc_putPermissions = function (v) {
		this.Permissions = v;
	};
	prot.get_Lang = prot.asc_getLang = function () {
		return this.Lang;
	};
	prot.put_Lang = prot.asc_putLang = function (v) {
		this.Lang = v;
	};
	prot.get_Encrypted = prot.asc_getEncrypted = function () {
		return this.Encrypted;
	};
	prot.put_Encrypted = prot.asc_putEncrypted = function (v) {
		this.Encrypted = v;
	};
	prot.get_EncryptedInfo = prot.asc_getEncryptedInfo = function () {
		return this.EncryptedInfo;
	};
	prot.put_EncryptedInfo = prot.asc_putEncryptedInfo = function (v) {
		this.EncryptedInfo = v;
	};
	prot.get_IsEnabledPlugins = prot.asc_getIsEnabledPlugins = function () {
		return this.IsEnabledPlugins;
	};
	prot.put_IsEnabledPlugins = prot.asc_putIsEnabledPlugins = function (v) {
		this.IsEnabledPlugins = v;
	};
	prot.get_IsEnabledMacroses = prot.asc_getIsEnabledMacroses = function () {
		return this.IsEnabledMacroses;
	};
	prot.put_IsEnabledMacroses = prot.asc_putIsEnabledMacroses = function (v) {
		this.IsEnabledMacroses = v;
	};
	prot.get_CoEditingMode = prot.asc_getCoEditingMode = function () {
		return this.coEditingMode;
	};
	prot.put_CoEditingMode = prot.asc_putCoEditingMode = function (v) {
		this.coEditingMode = v;
	};
	prot.put_ReferenceData = prot.asc_putReferenceData = function (v) {
		this.ReferenceData = v;
	};
	prot.put_IsWebOpening = prot.asc_putIsWebOpening = function (v) {
		this.IsWebOpening = v;
	};
	prot.get_IsWebOpening = prot.asc_getIsWebOpening = function () {
		return this.IsWebOpening;
	};
	prot.put_SupportsOnSaveDocument = prot.asc_putSupportsOnSaveDocument = function (v) {
		this.SupportsOnSaveDocument = v;
	};
	prot.get_SupportsOnSaveDocument = prot.asc_getSupportsOnSaveDocument = function () {
		return this.SupportsOnSaveDocument;
	};
	prot.put_Wopi = prot.asc_putWopi = function (v) {
		this.Wopi = v;
	};
	prot.get_Wopi = prot.asc_getWopi = function () {
		return this.Wopi;
	};
	prot.put_Shardkey = prot.asc_putShardkey = function (v) {
		this.shardkey = v;
	};
	prot.get_Shardkey = prot.asc_getShardkey = function () {
		return this.shardkey;
	};
	prot.put_ShowVerticalScroll = prot.asc_putShowVerticalScroll = function (v) {
		this.showVerticalScroll = v;
	};
	prot.get_ShowVerticalScroll = prot.asc_getShowVerticalScroll = function () {
		return this.showVerticalScroll;
	};
	prot.put_ShowHorizontalScroll = prot.asc_putShowHorizontalScroll = function (v) {
		this.showHorizontalScroll = v;
	};
	prot.get_ShowHorizontalScroll = prot.asc_getShowHorizontalScroll = function () {
		return this.showHorizontalScroll;
	};

	function COpenProgress() {
		this.Type = Asc.c_oAscAsyncAction.Open;

		this.FontsCount = 0;
		this.CurrentFont = 0;

		this.ImagesCount = 0;
		this.CurrentImage = 0;
	}

	COpenProgress.prototype.asc_getType = function () {
		return this.Type
	};
	COpenProgress.prototype.asc_getFontsCount = function () {
		return this.FontsCount
	};
	COpenProgress.prototype.asc_getCurrentFont = function () {
		return this.CurrentFont
	};
	COpenProgress.prototype.asc_getImagesCount = function () {
		return this.ImagesCount
	};
	COpenProgress.prototype.asc_getCurrentImage = function () {
		return this.CurrentImage
	};

	function CErrorData() {
		this.Value = 0;
	}

	CErrorData.prototype.put_Value = function (v) {
		this.Value = v;
	};
	CErrorData.prototype.get_Value = function () {
		return this.Value;
	};

	function CAscMathType() {
		this.Id = 0;

		this.X = 0;
		this.Y = 0;
	}

	CAscMathType.prototype.get_Id = function () {
		return this.Id;
	};
	CAscMathType.prototype.get_X = function () {
		return this.X;
	};
	CAscMathType.prototype.get_Y = function () {
		return this.Y;
	};

	function CAscMathCategory() {
		this.Id = 0;
		this.Data = [];

		this.W = 0;
		this.H = 0;
	}

	CAscMathCategory.prototype.get_Id = function () {
		return this.Id;
	};
	CAscMathCategory.prototype.get_Data = function () {
		return this.Data;
	};
	CAscMathCategory.prototype.get_W = function () {
		return this.W;
	};
	CAscMathCategory.prototype.get_H = function () {
		return this.H;
	};
	CAscMathCategory.prototype.private_Sort = function () {
		this.Data.sort(function (a, b) {
			return a.Id - b.Id;
		});
	};

	function CStyleImage(name, type, image, uiPriority) {
		this.name = name;
		this.displayName = null;
		this.type = type;
		this.image = image;
		this.uiPriority = uiPriority;
	}

	CStyleImage.prototype.asc_getId = CStyleImage.prototype.asc_getName = CStyleImage.prototype.get_Name = function () {
		return this.name;
	};
	CStyleImage.prototype.asc_getDisplayName = function () {
		return this.displayName;
	};
	CStyleImage.prototype.asc_getType = CStyleImage.prototype.get_Type = function () {
		return this.type;
	};
	CStyleImage.prototype.asc_getImage = function () {
		return this.image;
	};

	/** @constructor */
	function asc_CSpellCheckProperty(Word, Checked, Variants, ParaId, Element) {
		this.Word = Word;
		this.Checked = Checked;
		this.Variants = Variants;

		this.ParaId = ParaId;
		this.Element = Element;
	}

	asc_CSpellCheckProperty.prototype.get_Word = function () {
		return this.Word;
	};
	asc_CSpellCheckProperty.prototype.get_Checked = function () {
		return this.Checked;
	};
	asc_CSpellCheckProperty.prototype.get_Variants = function () {
		return this.Variants;
	};

	function CWatermarkOnDraw(htmlContent, api) {
		// example content:
		/*
		{
			"type" : "rect",
			"width" : 100, // mm
			"height" : 100, // mm
			"rotate" : -45, // degrees
			"margins" : [ 10, 10, 10, 10 ], // text margins
			"fill" : [255, 0, 0], // [] => none // "image_url"
			"stroke-width" : 1, // mm
			"stroke" : [0, 0, 255], // [] => none
			"align" : 1, // vertical text align (4 - top, 1 - center, 0 - bottom)

			"paragraphs" : [
				{
					"align" : 4, // horizontal text align [1 - left, 2 - center, 0 - right, 3 - justify]
					"fill" : [255, 0, 0], // paragraph highlight. [] => none
					"linespacing" : 0,

					"runs" : [
						{
							"text" : "some text",
							"fill" : [255, 255, 255], // text highlight. [] => none,
							"font-family" : "Arial",
							"font-size" : 24, // pt
							"bold" : true,
							"italic" : false,
							"strikeout" : "false",
							"underline" : "false"
						},
						{
							"text" : "<%br%>"
						}
					]
				}
			]
		}
		*/

		this.api = api;
		this.isFontsLoaded = false;

		this.inputContentSrc = htmlContent;
		if (typeof this.inputContentSrc === "object") this.inputContentSrc = JSON.stringify(this.inputContentSrc);

		this.replaceMap = {};

		this.image = null;
		this.imageBase64 = undefined;
		this.width = 0;
		this.height = 0;

		this.imageBackgroundUrl = "";
		this.imageBackground = null;

		this.transparent = 0.3;
		this.zoom = 1;
		this.calculatezoom = -1;

		this.isNativeGlobalAlpha = false;

		this.contentObjects = null;

		this.CheckParams = function () {
			this.replaceMap["%user_name%"] = this.api.User.userName;

			let content = this.inputContentSrc;
			for (let key in this.replaceMap) {
				if (!this.replaceMap.hasOwnProperty(key)) continue;
				content = content.replace(new RegExp(key, 'g'), this.replaceMap[key]);
			}
			this.contentObjects = {};
			try {
				let _objTmp = JSON.parse(content);
				this.contentObjects = _objTmp;
			} catch (err) {
			}

			this.transparent = (undefined == this.contentObjects['transparent']) ? 0.3 : this.contentObjects['transparent'];
		};

		this.Generate = function () {
			if (!this.isFontsLoaded) return;

			if (this.zoom == this.calculatezoom) return;

			this.calculatezoom = this.zoom;
			this.privateGenerateShape(this.contentObjects);
			//console.log( this.image.toDataURL("image/png"));
		};

		this.getCorrectedInputContentSrc = function() {
			let content = this.inputContentSrc;
			for (let key in this.replaceMap) {
				if (!this.replaceMap.hasOwnProperty(key)) continue;
				content = content.replace(new RegExp(key, 'g'), this.replaceMap[key]);
			}
			return content;
		};

		this.Draw = function (context, dw_or_dx, dh_or_dy, dw, dh) {
			if (!this.image || !this.isFontsLoaded) return;

			let x = 0;
			let y = 0;

			if (undefined == dw) {
				x = (dw_or_dx - this.width) >> 1;
				y = (dh_or_dy - this.height) >> 1;
			}
			else {
				x = (dw_or_dx + ((dw - this.width) / 2)) >> 0;
				y = (dh_or_dy + ((dh - this.height) / 2)) >> 0;
			}
			let oldGlobalAlpha = context.globalAlpha;
			context.globalAlpha = this.transparent;
			context.drawImage(this.image, x, y);
			context.globalAlpha = oldGlobalAlpha;
		};

		this.StartRenderer = function () {
			if (window["NATIVE_EDITOR_ENJINE"]) return;

			let canvasTransparent = document.createElement("canvas");
			canvasTransparent.width = this.image.width;
			canvasTransparent.height = this.image.height;
			let ctx = canvasTransparent.getContext("2d");
			ctx.globalAlpha = this.transparent;
			ctx.drawImage(this.image, 0, 0);
			try {
				this.imageBase64 = canvasTransparent.toDataURL("image/png");
			}
			catch (e) {
				this.imageBase64 = undefined;
				this.api.sendEvent("asc_onError", Asc.c_oAscError.ID.CannotSaveWatermark, Asc.c_oAscError.Level.NoCritical);
			}
			canvasTransparent = null;
		};
		this.EndRenderer = function () {
			if (window["NATIVE_EDITOR_ENJINE"]) return;

			delete this.imageBase64;
			this.imageBase64 = undefined;
		};
		this.DrawOnRenderer = function (renderer, w, h) {
			if(!this.imageBase64) {
				return;
			}
			let wMM = this.width * AscCommon.g_dKoef_pix_to_mm / this.zoom;
			let hMM = this.height * AscCommon.g_dKoef_pix_to_mm / this.zoom;
			let x = (w - wMM) / 2;
			let y = (h - hMM) / 2;

			if (window["NATIVE_EDITOR_ENJINE"]) {
				renderer.put_brushTexture(this.imageBase64, 0);
				renderer.put_BrushTextureAlpha((255 * this.transparent) >> 0);
				renderer._e();
				renderer.rect(x, y, wMM, hMM);
				renderer.df();
				renderer._e();
			}
			else {
				renderer.UseOriginImageUrl = true;
				renderer.drawImage(this.imageBase64, x, y, wMM, hMM);
				renderer.UseOriginImageUrl = false;
			}
		};

		this.privateGenerateShape = function (obj) {

			AscFormat.ExecuteNoHistory(function (obj) {

				let oShape = new AscFormat.CShape();
				let bWord = false;
				let oApi = Asc['editor'] || editor;
				if (!oApi) {
					return null;
				}
				switch (oApi.getEditorId()) {
					case AscCommon.c_oEditorId.Word: {
						bWord = true && !oApi.isPdfEditor();
						oShape.setWordShape(bWord);
						break;
					}
					case AscCommon.c_oEditorId.Presentation: {
						oShape.setWordShape(false);
						oShape.setParent(oApi.WordControl.m_oLogicDocument.GetCurrentSlide());
						break;
					}
					case AscCommon.c_oEditorId.Spreadsheet: {
						oShape.setWordShape(false);
						oShape.setWorksheet(oApi.wb.getWorksheet().model);
						break;
					}
					case AscCommon.c_oEditorId.Visio: {
						oShape.setWordShape(false);
						oShape.setParent(oApi.WordControl.m_oLogicDocument);
						break;
					}
				}

				let _oldTrackRevision = false;
				if (oApi.getEditorId() === AscCommon.c_oEditorId.Word && oApi.WordControl && oApi.WordControl.m_oLogicDocument && !oApi.isPdfEditor())
					_oldTrackRevision = oApi.WordControl.m_oLogicDocument.GetLocalTrackRevisions();

				if (false !== _oldTrackRevision)
					oApi.WordControl.m_oLogicDocument.SetLocalTrackRevisions(false);

				let bRemoveDocument = false;
				if (oApi.WordControl && !oApi.WordControl.m_oLogicDocument) {
					// TODO: Зачем это здесь вообще?
					bRemoveDocument = true;
					oApi.WordControl.m_oLogicDocument = new AscWord.CDocument(null, false);
					oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oApi.WordControl.m_oLogicDocument;
				}
				oShape.setBDeleted(false);
				oShape.spPr = new AscFormat.CSpPr();
				oShape.spPr.setParent(oShape);
				oShape.spPr.setXfrm(new AscFormat.CXfrm());
				oShape.spPr.xfrm.setParent(oShape.spPr);
				oShape.spPr.xfrm.setOffX(0);
				oShape.spPr.xfrm.setOffY(0);
				oShape.spPr.xfrm.setExtX(obj['width']);
				oShape.spPr.xfrm.setExtY(obj['height']);
				oShape.spPr.xfrm.setRot(AscFormat.normalizeRotate(obj['rotate'] ? (obj['rotate'] * Math.PI / 180) : 0));
				oShape.spPr.setGeometry(AscFormat.CreateGeometry(obj['type'] || "rect"));
				if (obj['fill'] && obj['fill'].length === 3) {
					oShape.spPr.setFill(AscFormat.CreateSolidFillRGB(obj['fill'][0], obj['fill'][1], obj['fill'][2]));
				}
				else if (this.imageBackground) {
					oApi.ImageLoader.map_image_index[this.imageBackgroundUrl] = {
						Image: this.imageBackground, Status: AscFonts.ImageLoadStatus.Complete
					};
					oShape.spPr.setFill(AscFormat.builder_CreateBlipFill(this.imageBackgroundUrl, "stretch"));
				}
				if (AscFormat.isRealNumber(obj['stroke-width']) || Array.isArray(obj['stroke']) && obj['stroke'].length === 3) {
					let oUnifill;
					if (Array.isArray(obj['stroke']) && obj['stroke'].length === 3) {
						oUnifill = AscFormat.CreateSolidFillRGB(obj['stroke'][0], obj['stroke'][1], obj['stroke'][2]);
					}
					else {
						oUnifill = AscFormat.CreateSolidFillRGB(0, 0, 0);
					}
					oShape.spPr.setLn(AscFormat.CreatePenFromParams(oUnifill, undefined, undefined, undefined, undefined, AscFormat.isRealNumber(obj['stroke-width']) ? obj['stroke-width'] : 12700.0 / 36000.0));
				}

				if (bWord) {
					oShape.createTextBoxContent();
				}
				else {
					oShape.createTextBody();
				}
				let align = obj['align'];
				if (undefined != align) {
					oShape.setVerticalAlign(align);
				}
				else {

					oShape.setVerticalAlign(1);//ctr
				}
				oShape.setVertOverflowType(AscFormat.nVOTOverflow);
				if (Array.isArray(obj['margins']) && obj['margins'].length === 4) {
					oShape.setPaddings({
						Left: obj['margins'][0],
						Top: obj['margins'][1],
						Right: obj['margins'][2],
						Bottom: obj['margins'][3]
					});
				}
				let oContent = oShape.getDocContent();
				let aParagraphsS = obj['paragraphs'] || [];
				if (aParagraphsS.length > 0) {
					oContent.Content.length = 0;
				}
				for (let i = 0; i < aParagraphsS.length; ++i) {
					let oCurParS = aParagraphsS[i];
					let oNewParagraph = new AscWord.Paragraph(oContent, !bWord);
					if (AscFormat.isRealNumber(oCurParS['align'])) {
						oNewParagraph.Set_Align(oCurParS['align'])
					}
					if (Array.isArray(oCurParS['fill']) && oCurParS['fill'].length === 3) {
						let oShd = new AscCommonWord.CDocumentShd();
						oShd.Value = Asc.c_oAscShdClear;
						oShd.Color.r = oCurParS['fill'][0];
						oShd.Color.g = oCurParS['fill'][1];
						oShd.Color.b = oCurParS['fill'][2];
						oShd.Fill = new AscCommonWord.CDocumentColor();
						oShd.Fill.r = oCurParS['fill'][0];
						oShd.Fill.g = oCurParS['fill'][1];
						oShd.Fill.b = oCurParS['fill'][2];
						oNewParagraph.Set_Shd(oShd, true);
					}
					if (AscFormat.isRealNumber(oCurParS['linespacing'])) {
						oNewParagraph.Set_Spacing({
							Line: oCurParS['linespacing'], Before: 0, After: 0, LineRule: Asc.linerule_Auto
						}, true);
					}
					let aRunsS = oCurParS['runs'];
					for (let j = 0; j < aRunsS.length; ++j) {
						let oRunS = aRunsS[j];
						let oRun = new AscCommonWord.ParaRun(oNewParagraph, false);
						if (Array.isArray(oRunS['fill']) && oRunS['fill'].length === 3) {
							oRun.Set_Unifill(AscFormat.CreateSolidFillRGB(oRunS['fill'][0], oRunS['fill'][1], oRunS['fill'][2]));
						}
						let fontFamilyName = oRunS['font-family'] ? oRunS['font-family'] : "Arial";
						let fontSize = (oRunS['font-size'] != null) ? oRunS['font-size'] : 50;

						oRun.SetRFontsAscii({Name: fontFamilyName, Index: -1});
						oRun.SetRFontsCS({Name: fontFamilyName, Index: -1});
						oRun.SetRFontsEastAsia({Name: fontFamilyName, Index: -1});
						oRun.SetRFontsHAnsi({Name: fontFamilyName, Index: -1});

						oRun.SetFontSize(fontSize);

						oRun.SetBold(oRunS['bold'] === true);
						oRun.SetItalic(oRunS['italic'] === true);
						oRun.SetStrikeout(oRunS['strikeout'] === true);
						oRun.SetUnderline(oRunS['underline'] === true);

						let sCustomText = oRunS['text'];
						if (sCustomText === "<%br%>") {
							oRun.AddToContent(0, new AscWord.CRunBreak(AscWord.break_Line), false);
						}
						else {
							oRun.AddText(sCustomText);
						}

						oNewParagraph.Internal_Content_Add(j, oRun, false);
					}
					oContent.Internal_Content_Add(oContent.Content.length, oNewParagraph);
				}

				let bLoad = AscCommon.g_oIdCounter.m_bLoad;
				AscCommon.g_oIdCounter.Set_Load(false);
				oShape.recalculate();
				if (oShape.bWordShape) {
					oShape.recalculateText();
				}

				AscCommon.g_oIdCounter.Set_Load(bLoad);
				let oldShowParaMarks;
				if (window.editor) {
					oldShowParaMarks = oApi.ShowParaMarks;
					oApi.ShowParaMarks = false;
				}

				AscCommon.IsShapeToImageConverter = true;
				let _bounds_cheker = new AscFormat.CSlideBoundsChecker();

				let w_mm = 210;
				let h_mm = 297;
				let w_px = AscCommon.AscBrowser.convertToRetinaValue(w_mm * AscCommon.g_dKoef_mm_to_pix * this.zoom, true);
				let h_px = AscCommon.AscBrowser.convertToRetinaValue(h_mm * AscCommon.g_dKoef_mm_to_pix * this.zoom, true);

				_bounds_cheker.init(w_px, h_px, w_mm, h_mm);
				_bounds_cheker.transform(1, 0, 0, 1, 0, 0);

				_bounds_cheker.AutoCheckLineWidth = true;
				_bounds_cheker.CheckLineWidth(oShape);
				oShape.draw(_bounds_cheker, 0);
				_bounds_cheker.CorrectBounds2();

				let _need_pix_width = _bounds_cheker.Bounds.max_x - _bounds_cheker.Bounds.min_x + 1;
				let _need_pix_height = _bounds_cheker.Bounds.max_y - _bounds_cheker.Bounds.min_y + 1;

				if (_need_pix_width <= 0 || _need_pix_height <= 0) return;

				if (!this.image) this.image = document.createElement("canvas");

				this.image.width = _need_pix_width;
				this.image.height = _need_pix_height;
				this.width = _need_pix_width;
				this.height = _need_pix_height;

				let _ctx = this.image.getContext('2d');

				let g = null;

				if (window["NATIVE_EDITOR_ENJINE"]) {
					g = new AscCommon.CNativeGraphics();
					g.width = _need_pix_width;
					g.height = _need_pix_height;
					g.create(window["native"], _need_pix_width, _need_pix_height, _need_pix_width / AscCommon.g_dKoef_mm_to_pix, _need_pix_height / AscCommon.g_dKoef_mm_to_pix);
					g.CoordTransformOffset(-_bounds_cheker.Bounds.min_x, -_bounds_cheker.Bounds.min_y);
					g.transform(1, 0, 0, 1, 0, 0);

					if (this.isNativeGlobalAlpha)
						g.CreateLayer(this.transparent);
				}
				else {
					g = new AscCommon.CGraphics();
					g.init(_ctx, w_px, h_px, w_mm, h_mm);
					g.m_oFontManager = AscCommon.g_fontManager;

					g.m_oCoordTransform.tx = -_bounds_cheker.Bounds.min_x;
					g.m_oCoordTransform.ty = -_bounds_cheker.Bounds.min_y;
					g.transform(1, 0, 0, 1, 0, 0);
				}

				oShape.draw(g, 0);

				if (window["NATIVE_EDITOR_ENJINE"])
				{
					if (this.isNativeGlobalAlpha)
						g.BlendLayer();
					this.imageBase64 = g.toDataURL("image/png");
				}

				AscCommon.IsShapeToImageConverter = false;

				if (bRemoveDocument) {
					oApi.WordControl.m_oLogicDocument = null;
					oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = null;
				}
				if (window.editor) {
					oApi.ShowParaMarks = oldShowParaMarks;
				}

				if (false !== _oldTrackRevision)
					oApi.WordControl.m_oLogicDocument.SetLocalTrackRevisions(_oldTrackRevision);

				if (this.imageBackground)
					delete oApi.ImageLoader.map_image_index[this.imageBackgroundUrl];

			}, this, [obj]);

		};

		this.onReady = function () {
			this.isFontsLoaded = true;
			let oApi = this.api;

			switch (oApi.editorId) {
				case AscCommon.c_oEditorId.Word: {
					if (oApi.WordControl) {
						if (oApi.watermarkDraw) {
							oApi.watermarkDraw.zoom = oApi.WordControl.m_nZoomValue / 100;
							oApi.watermarkDraw.Generate();
						}

						oApi.WordControl.OnRePaintAttack();
					}

					break;
				}
				case AscCommon.c_oEditorId.Presentation:
				case AscCommon.c_oEditorId.Visio: {
					if (oApi.WordControl) {
						if (oApi.watermarkDraw) {
							oApi.watermarkDraw.zoom = oApi.WordControl.m_nZoomValue / 100;
							oApi.watermarkDraw.Generate();
						}

						oApi.WordControl.OnRePaintAttack();
					}
					break;
				}
				case AscCommon.c_oEditorId.Spreadsheet: {
					let ws = oApi.wb && oApi.wb.getWorksheet();
					if (ws && ws.objectRender && ws.objectRender) {
						ws.objectRender.OnUpdateOverlay();
					}
					break;
				}
			}
		};

		this.checkOnReady = function () {
			this.CheckParams();

			let fonts = [];
			let pars = this.contentObjects['paragraphs'] || [];
			let i, j;
			for (i = 0; i < pars.length; i++) {
				let runs = pars[i]['runs'];
				for (j = 0; j < runs.length; j++) {
					if (undefined === runs[j]["font-family"]) runs[j]["font-family"] = "Arial";
					fonts.push(runs[j]["font-family"]);
				}
			}

			for (i = 0; i < fonts.length; i++) {
				fonts[i] = new AscFonts.CFont(AscFonts.g_fontApplication.GetFontInfoName(fonts[i]));
			}

			if ("string" === typeof this.contentObjects["fill"]) this.imageBackgroundUrl = this.contentObjects["fill"];

			if (false === AscCommon.g_font_loader.CheckFontsNeedLoading(fonts)) {
				this.loadBackgroundImage();
				return;
			}

			this.api.asyncMethodCallback = function () {
				let oApi = Asc['editor'] || editor;
				oApi.watermarkDraw.loadBackgroundImage();
			};

			AscCommon.g_font_loader.LoadDocumentFonts2(fonts);
		};

		this.loadBackgroundImage = function () {
			if ("" === this.imageBackgroundUrl) {
				this.onReady();
				return;
			}

			this.imageBackground = new Image();
			this.imageBackground.onload = function () {
				Asc["editor"].watermarkDraw.onReady();
			};
			this.imageBackground.onerror = function () {
				Asc["editor"].watermarkDraw.imageBackground = null;
				Asc["editor"].watermarkDraw.onReady();
			};
			this.imageBackground.src = this.imageBackgroundUrl;
		};

		this.generateNative = function () {
			this.isFontsLoaded = true;
			this.zoom = 1;
			this.calculatezoom = 0;
			this.CheckParams();

			if (this.contentObjects && "string" === typeof this.contentObjects["fill"])
			{
				this.imageBackgroundUrl = this.contentObjects["fill"];
				this.imageBackground = {};
			}

			this.Generate();
		};
	}

	// ----------------------------- plugins ------------------------------- //
	let PluginType = {
		System: 0,      // Системный, неотключаемый плагин.
		Background: 1,  // Фоновый плагин. Тоже самое, что и системный, но отключаемый.
		Window: 2,      // Окно
		Panel: 3,       // Панель
		Invisible : 4,  // Невидимый
		PanelRight: 5   // Панель справа
	};

	PluginType["System"] = PluginType.System;
	PluginType["Background"] = PluginType.Background;
	PluginType["Window"] = PluginType.Window;
	PluginType["Panel"] = PluginType.Panel;
	PluginType["PanelRight"] = PluginType.PanelRight;
	PluginType["Unvisible"] = PluginType.Unvisible;

	PluginType["getType"] = PluginType.getType = function(type) {
		if (undefined === type)
			return undefined;

		if (typeof type !== "string")
			return type;

		switch (type) {
			case "system" : return this.System;
			case "background" : return this.Background;
			case "window" : return this.Window;
			case "panel" : return this.Panel;
			case "panelRight" : return this.PanelRight;
			case "invisible" : return this.Invisible;
			default: break;
		}

		return this.Background;
	};

	function CPluginVariation() {
		this.description = "";
		this.url = "";
		this.help = "";
		this.baseUrl = "";
		this.index = 0;     // сверху не выставляем. оттуда в каком порядке пришли - в таком порядке и работают

		this.descriptionLocale = undefined;
		this.icons = ["1x", "2x"];
		this.isViewer = false;
		this.isDisplayedInViewer = undefined;
		this.EditorsSupport = ["word", "cell", "slide"];

		this.type = PluginType.Background;

		this.isCustomWindow = false;	// используется только если this.type === PluginType.Window
		this.isModal = true;     // используется только если this.type === PluginType.Window

		this.isCanDocked = false;

		this.initDataType = EPluginDataType.none;
		this.initData = "";

		this.isUpdateOleOnResize = false;

		this.buttons = [{"text": "Ok", "primary": true}, {"text": "Cancel", "primary": false}];

		this.size = undefined;
		this.initOnSelectionChanged = undefined;

		this.store = undefined;

		this.events = [];
		this.eventsMap = {};
	}

	CPluginVariation.prototype["get_Description"] = function () {
		return this.description;
	};
	CPluginVariation.prototype["get_Url"] = function () {
		return this.url;
	};
	CPluginVariation.prototype["get_Help"] = function () {
		return this.help;
	};

	CPluginVariation.prototype["get_Icons"] = function () {
		return this.icons;
	};

	CPluginVariation.prototype["get_Type"] = function () {
		return this.type;
	};

	CPluginVariation.prototype["get_Visual"] = function () {
		return (this.type === PluginType.Window || this.type === PluginType.Panel || this.type === PluginType.PanelRight) ? true : false;
	};

	CPluginVariation.prototype["get_Viewer"] = function () {
		return this.isViewer;
	};
	CPluginVariation.prototype["get_EditorsSupport"] = function () {
		return this.EditorsSupport;
	};

	CPluginVariation.prototype["get_Modal"] = function () {
		return this.isModal;
	};
	CPluginVariation.prototype["get_InsideMode"] = function () {
		return (this.type === PluginType.Panel || this.type === PluginType.PanelRight) ? true : false;
	};
	CPluginVariation.prototype["get_CustomWindow"] = function () {
		return this.isCustomWindow;
	};

	CPluginVariation.prototype["get_IsCanDocked"] = function () {
		return this.isCanDocked;
	};

	CPluginVariation.prototype["get_Buttons"] = function () {
		return this.buttons;
	};
	CPluginVariation.prototype["get_Size"] = function () {
		return this.size;
	};
	CPluginVariation.prototype["get_Events"] = function () {
		return this.events;
	};
	CPluginVariation.prototype["set_Events"] = function (value) {
		if (!value) return;

		this.events = value.slice(0, value.length);
		this.eventsMap = {};
		for (let i = 0; i < this.events.length; i++) this.eventsMap[this.events[i]] = true;
	};

	CPluginVariation.prototype["serialize"] = function () {
		let _object = {};
		_object["description"] = this.description;
		_object["url"] = this.url;
		_object["help"] = this.help;
		_object["index"] = this.index;

		_object["descriptionLocale"] = this.descriptionLocale;

		_object["icons"] = this.icons;
		_object["icons2"] = this.icons2;

		_object["isViewer"] = this.isViewer;
		_object["isDisplayedInViewer"] = this.isDisplayedInViewer;
		_object["EditorsSupport"] = this.EditorsSupport;

		_object["type"] = this.type;

		_object["isCustomWindow"] = this.isCustomWindow;
		_object["isModal"] = this.isModal;

		_object["isCanDocked"] = this.isCanDocked;

		_object["initDataType"] = this.initDataType;
		_object["initData"] = this.initData;

		_object["isUpdateOleOnResize"] = this.isUpdateOleOnResize;

		_object["buttons"] = this.buttons;

		_object["size"] = this.size;
		_object["initOnSelectionChanged"] = this.initOnSelectionChanged;

		_object["store"] = this.store;

		if (this.events)
			_object["events"] = this.events.slice(0, this.events.length);

		return _object;
	};
	CPluginVariation.prototype["deserialize"] = function (_object) {
		this.description = (_object["description"] != null) ? _object["description"] : this.description;
		this.url = (_object["url"] != null) ? _object["url"] : this.url;
		this.help = (_object["help"] != null) ? _object["help"] : this.help;
		this.index = (_object["index"] != null) ? _object["index"] : this.index;

		this.descriptionLocale = (_object["descriptionLocale"] != null) ? _object["descriptionLocale"] : this.descriptionLocale;
		this.icons = (_object["icons"] != null) ? _object["icons"] : this.icons;
		this.icons2 = (_object["icons2"] != null) ? _object["icons2"] : this.icons2;
		this.isViewer = (_object["isViewer"] != null) ? _object["isViewer"] : this.isViewer;
		this.isDisplayedInViewer = (_object["isDisplayedInViewer"] != null) ? _object["isDisplayedInViewer"] : this.isDisplayedInViewer;
		this.EditorsSupport = (_object["EditorsSupport"] != null) ? _object["EditorsSupport"] : this.EditorsSupport;

		// default: background
		this.type = PluginType.Background;

		let _type = PluginType.getType(_object["type"]);
		if (undefined !== _type) {
			this.type = _type;
		}
		else {
			// old version: not support background plugins
			if (true === _object["isSystem"])
				this.type = PluginType.System;
			else if (true === _object["isVisual"])
				this.type = (true === _object["isInsideMode"]) ? PluginType.Panel : PluginType.Window;
			else
				this.type = PluginType.Invisible;
		}

		this.isCustomWindow = (_object["isCustomWindow"] != null) ? _object["isCustomWindow"] : this.isCustomWindow;
		this.isModal = (_object["isModal"] != null) ? _object["isModal"] : this.isModal;

		this.isCanDocked = (_object["isCanDocked"] != null) ? _object["isCanDocked"] : this.isCanDocked;

		this.initDataType = (_object["initDataType"] != null) ? _object["initDataType"] : this.initDataType;
		this.initData = (_object["initData"] != null) ? _object["initData"] : this.initData;

		this.isUpdateOleOnResize = (_object["isUpdateOleOnResize"] != null) ? _object["isUpdateOleOnResize"] : this.isUpdateOleOnResize;

		this.buttons = (_object["buttons"] != null) ? _object["buttons"] : this.buttons;

		this.store = (_object["store"] != null) ? _object["store"] : this.store;

		if (_object["events"] != null) this["set_Events"](_object["events"]);

		this.size = (_object["size"] != null) ? _object["size"] : this.size;
		this.initOnSelectionChanged = (_object["initOnSelectionChanged"] != null) ? _object["initOnSelectionChanged"] : this.initOnSelectionChanged;
	};

	function CPlugin() {
		this.name = "";
		this.nameLocale = {};
		this.guid = "";
		this.baseUrl = "";
		this.minVersion = "";
		this.version = "";
		this.isConnector = false;
		this.loader;

		this.variations = [];
	}

	CPlugin.prototype["get_Name"] = function (locale) {
		if (locale && this.nameLocale && this.nameLocale[locale]) return this.nameLocale[locale];
		return this.name;
	};
	CPlugin.prototype["set_Name"] = function (value) {
		this.name = value;
	};
	CPlugin.prototype["get_NameLocale"] = function () {
		return this.nameLocale;
	};
	CPlugin.prototype["set_NameLocale"] = function (value) {
		this.nameLocale = value;
	};
	CPlugin.prototype["get_Guid"] = function () {
		return this.guid;
	};
	CPlugin.prototype["set_Guid"] = function (value) {
		this.guid = value;
	};
	CPlugin.prototype["get_BaseUrl"] = function () {
		return this.baseUrl;
	};
	CPlugin.prototype["set_BaseUrl"] = function (value) {
		this.baseUrl = value;
	};
	CPlugin.prototype["get_MinVersion"] = function () {
		return this.minVersion;
	};
	CPlugin.prototype["set_MinVersion"] = function (value) {
		this.minVersion = value;
	};
	CPlugin.prototype["get_Version"] = function () {
		return this.version;
	};
	CPlugin.prototype["set_Version"] = function (value) {
		this.version = value;
	};

	CPlugin.prototype["get_Variations"] = function () {
		return this.variations;
	};
	CPlugin.prototype["set_Variations"] = function (value) {
		this.variations = value;
	};

	CPlugin.prototype["get_Loader"] = function () {
		return this.loader;
	};
	CPlugin.prototype["set_Loader"] = function (value) {
		this.loader = value;
	};

	CPlugin.prototype["serialize"] = function () {
		let _object = {};
		_object["name"] = this.name;
		_object["nameLocale"] = this.nameLocale;
		_object["guid"] = this.guid;
		_object["version"] = this.version;
		_object["baseUrl"] = this.baseUrl;
		_object["minVersion"] = this.minVersion;
		_object["isConnector"] = this.isConnector;
		_object["loader"] = this.loader;

		if (this.group) {
			_object["group"] = {};
			_object["group"]["name"] = this.group.name;
			_object["group"]["rank"] = this.group.rank;
		}

		_object["variations"] = [];
		for (let i = 0; i < this.variations.length; i++) {
			_object["variations"].push(this.variations[i].serialize());
		}
		return _object;
	};
	CPlugin.prototype["deserialize"] = function (_object) {
		this.name = (_object["name"] != null) ? _object["name"] : this.name;
		this.nameLocale = (_object["nameLocale"] != null) ? _object["nameLocale"] : this.nameLocale;
		this.guid = (_object["guid"] != null) ? _object["guid"] : this.guid;
		this.version = (_object["version"] != null) ? _object["version"] : this.version;
		this.baseUrl = (_object["baseUrl"] != null) ? _object["baseUrl"] : this.baseUrl;
		this.minVersion = (_object["minVersion"] != null) ? _object["minVersion"] : this.minVersion;
		this.isConnector = (_object["isConnector"] != null) ? _object["isConnector"] : this.isConnector;
		this.loader = (_object["loader"] != null) ? _object["loader"] : this.loader;

		if (true) {
			// удалим этот if, как передем на просто прокидку объекта в интерфейсе
			if (_object["groupName"] || _object["groupRank"]) this.group = {};

			if (_object["groupName"]) this.group.name = _object["groupName"];
			if (_object["groupRank"]) this.group.rank = _object["groupRank"];
		}

		if (_object["group"]) {
			this.group = {};
			this.group.name = (_object["group"]["name"] != null) ? _object["group"]["name"] : "";
			this.group.rank = (_object["group"]["rank"] != null) ? _object["group"]["rank"] : 0;
		}

		this.variations = [];
		for (let i = 0; i < _object["variations"].length; i++) {
			let _variation = new CPluginVariation();
			_variation["deserialize"](_object["variations"][i]);
			this.variations.push(_variation);
		}
	};

	// no export
	CPlugin.prototype.isType = function (type) {
		if (this.variations && this.variations[0] && this.variations[0].type === type) return true;
		return false;
	};
	CPlugin.prototype.isSystem = function () {
		return this.isType(PluginType.System);
	};
	CPlugin.prototype.isBackground = function () {
		return this.isType(PluginType.Background);
	};
	CPlugin.prototype.getIntVersion = function () {
		if (!this.version) return 0;
		let arrayVersion = this.version.split(".");

		while (arrayVersion.length < 3) arrayVersion.push("0");

		try {
			let intVer = parseInt(arrayVersion[0]) * 10000 + parseInt(arrayVersion[1]) * 100 + parseInt(arrayVersion[2]);
			return intVer;
		} catch (e) {
		}
		return 0;
	};

	/**
	 * @constructor
	 */
	function CDocInfoProp(obj) {
		if (obj) {
			this.PageCount = obj.PageCount;
			this.WordsCount = obj.WordsCount;
			this.ParagraphCount = obj.ParagraphCount;
			this.SymbolsCount = obj.SymbolsCount;
			this.SymbolsWSCount = obj.SymbolsWSCount;
		}
		else {
			this.PageCount = -1;
			this.WordsCount = -1;
			this.ParagraphCount = -1;
			this.SymbolsCount = -1;
			this.SymbolsWSCount = -1;
		}
	}

	CDocInfoProp.prototype.get_PageCount = function () {
		return this.PageCount;
	};
	CDocInfoProp.prototype.put_PageCount = function (v) {
		this.PageCount = v;
	};
	CDocInfoProp.prototype.get_WordsCount = function () {
		return this.WordsCount;
	};
	CDocInfoProp.prototype.put_WordsCount = function (v) {
		this.WordsCount = v;
	};
	CDocInfoProp.prototype.get_ParagraphCount = function () {
		return this.ParagraphCount;
	};
	CDocInfoProp.prototype.put_ParagraphCount = function (v) {
		this.ParagraphCount = v;
	};
	CDocInfoProp.prototype.get_SymbolsCount = function () {
		return this.SymbolsCount;
	};
	CDocInfoProp.prototype.put_SymbolsCount = function (v) {
		this.SymbolsCount = v;
	};
	CDocInfoProp.prototype.get_SymbolsWSCount = function () {
		return this.SymbolsWSCount;
	};
	CDocInfoProp.prototype.put_SymbolsWSCount = function (v) {
		this.SymbolsWSCount = v;
	};
	
	/**
	 * @constructor
	 */
	function RangePermProp(obj) {
		if (obj) {
			this.editText      = undefined !== obj.editText ? obj.editText : true;
			this.editParagraph = undefined !== obj.editParagraph ? obj.editParagraph : true;
			this.insertObject  = undefined !== obj.insertObject ? obj.insertObject : true;
		} else {
			this.editText      = true;
			this.editParagraph = true;
			this.insertObject  = true;
		}
	}
	RangePermProp.prototype.get_canEditText = function() {
		return this.editText;
	};
	RangePermProp.prototype.get_canEditPara = function() {
		return this.editParagraph;
	};
	RangePermProp.prototype.get_canInsObject = function() {
		return this.insertObject;
	};


	function CButtonData(oData) {
		this["obj"] = oData["obj"];
		this["type"] = oData["type"];
		this["button"] = oData["button"];
		this["isForm"] = oData["isForm"];
		this["pr"] = oData["pr"];
	}
	CButtonData.prototype.get_Obj = function() {
		return this["obj"];
	};
	CButtonData.prototype.get_ObjId = function() {
		if(!this["obj"]) return null;
		return this["obj"].Id;
	};
	CButtonData.prototype.get_Type = function() {
		return this["type"];
	};
	CButtonData.prototype.get_Button = function() {
		return this["button"];
	};
	CButtonData.prototype.get_IsForm = function() {
		return this["isForm"];
	};
	CButtonData.prototype.get_Properties = function() {
		return this["pr"];
	};

	/*
	 * Export
	 * -----------------------------------------------------------------------------
	 */
	window['AscCommon'] = window['AscCommon'] || {};
	window['Asc'] = window['Asc'] || {};

	window['Asc']['c_oAscArrUserColors'] = window['Asc'].c_oAscArrUserColors = c_oAscArrUserColors;

	window["AscCommon"].CreateAscColorCustom = CreateAscColorCustom;
	window["AscCommon"].CreateAscColor = CreateAscColor;
	window["AscCommon"].CreateGUID = CreateGUID;
	window["AscCommon"].CreateUUID = CreateUUID;
	window["AscCommon"].CreateUInt32 = CreateUInt32;
	window["AscCommon"].CreateDurableId = CreateDurableId;
	window["AscCommon"].FixDurableId = FixDurableId;
	window["AscCommon"].ExtendPrototype = ExtendPrototype;

	window['Asc']['c_oLicenseResult'] = window['Asc'].c_oLicenseResult = c_oLicenseResult;
	prot = c_oLicenseResult;
	prot['Error'] = prot.Error;
	prot['Expired'] = prot.Expired;
	prot['Success'] = prot.Success;
	prot['UnknownUser'] = prot.UnknownUser;
	prot['Connections'] = prot.Connections;
	prot['ExpiredTrial'] = prot.ExpiredTrial;
	prot['SuccessLimit'] = prot.SuccessLimit;
	prot['UsersCount'] = prot.UsersCount;
	prot['ConnectionsOS'] = prot.ConnectionsOS;
	prot['UsersCountOS'] = prot.UsersCountOS;
	prot['ExpiredLimited'] = prot.ExpiredLimited;
	prot['ConnectionsLiveOS'] = prot.ConnectionsLiveOS;
	prot['ConnectionsLive'] = prot.ConnectionsLive;
	prot['UsersViewCount'] = prot.UsersViewCount;
	prot['UsersViewCountOS'] = prot.UsersViewCountOS;
	prot['NotBefore'] = prot.NotBefore;

	window['Asc']['c_oRights'] = window['Asc'].c_oRights = c_oRights;
	prot = c_oRights;
	prot['None'] = prot.None;
	prot['Edit'] = prot.Edit;
	prot['Review'] = prot.Review;
	prot['Comment'] = prot.Comment;
	prot['View'] = prot.View;

	window['Asc']['c_oLicenseMode'] = window['Asc'].c_oLicenseMode = c_oLicenseMode;
	prot = c_oLicenseMode;
	prot['None'] = prot.None;
	prot['Trial'] = prot.Trial;
	prot['Developer'] = prot.Developer;
	prot['Limited'] = prot.Limited;

	window["Asc"]["EPluginDataType"] = window["Asc"].EPluginDataType = EPluginDataType;
	prot = EPluginDataType;
	prot['none'] = prot.none;
	prot['text'] = prot.text;
	prot['ole'] = prot.ole;
	prot['html'] = prot.html;

	window["AscCommon"]["asc_CSignatureLine"] = window["AscCommon"].asc_CSignatureLine = asc_CSignatureLine;
	prot = asc_CSignatureLine.prototype;
	prot["asc_getId"] = prot.asc_getId;
	prot["asc_setId"] = prot.asc_setId;
	prot["asc_getGuid"] = prot.asc_getGuid;
	prot["asc_setGuid"] = prot.asc_setGuid;
	prot["asc_getSigner1"] = prot.asc_getSigner1;
	prot["asc_setSigner1"] = prot.asc_setSigner1;
	prot["asc_getSigner2"] = prot.asc_getSigner2;
	prot["asc_setSigner2"] = prot.asc_setSigner2;
	prot["asc_getEmail"] = prot.asc_getEmail;
	prot["asc_setEmail"] = prot.asc_setEmail;
	prot["asc_getInstructions"] = prot.asc_getInstructions;
	prot["asc_setInstructions"] = prot.asc_setInstructions;
	prot["asc_getShowDate"] = prot.asc_getShowDate;
	prot["asc_setShowDate"] = prot.asc_setShowDate;
	prot["asc_getValid"] = prot.asc_getValid;
	prot["asc_setValid"] = prot.asc_setValid;
	prot["asc_getDate"] = prot.asc_getDate;
	prot["asc_setDate"] = prot.asc_setDate;
	prot["asc_getVisible"] = prot.asc_getVisible;
	prot["asc_setVisible"] = prot.asc_setVisible;
	prot["asc_getRequested"] = prot.asc_getRequested;
	prot["asc_setRequested"] = prot.asc_setRequested;
	prot["asc_getIsForm"] = prot.asc_getIsForm;
	prot["asc_setIsForm"] = prot.asc_setIsForm;

	window["AscCommon"].asc_CAscEditorPermissions = asc_CAscEditorPermissions;
	prot = asc_CAscEditorPermissions.prototype;
	prot["asc_getLicenseType"] = prot.asc_getLicenseType;
	prot["asc_getCanCoAuthoring"] = prot.asc_getCanCoAuthoring;
	prot["asc_getCanReaderMode"] = prot.asc_getCanReaderMode;
	prot["asc_getCanBranding"] = prot.asc_getCanBranding;
	prot["asc_getCustomization"] = prot.asc_getCustomization;
	prot["asc_getIsAutosaveEnable"] = prot.asc_getIsAutosaveEnable;
	prot["asc_getAutosaveMinInterval"] = prot.asc_getAutosaveMinInterval;
	prot["asc_getIsAnalyticsEnable"] = prot.asc_getIsAnalyticsEnable;
	prot["asc_getIsLight"] = prot.asc_getIsLight;
	prot["asc_getLicenseMode"] = prot.asc_getLicenseMode;
	prot["asc_getRights"] = prot.asc_getRights;
	prot["asc_getBuildVersion"] = prot.asc_getBuildVersion;
	prot["asc_getBuildNumber"] = prot.asc_getBuildNumber;
	prot["asc_getLiveViewerSupport"] = prot.asc_getLiveViewerSupport;
	prot["asc_getIsBeta"] = prot.asc_getIsBeta;

	window["AscCommon"].asc_CAxNumFmt = asc_CAxNumFmt;
	prot = asc_CAxNumFmt.prototype;
	prot["getFormatCode"] = prot.getFormatCode;
	prot["putFormatCode"] = prot.putFormatCode;
	prot["getFormatCellsInfo"] = prot.getFormatCellsInfo;
	prot["getSourceLinked"] = prot.getSourceLinked;
	prot["putSourceLinked"] = prot.putSourceLinked;

	window["AscCommon"].asc_ValAxisSettings = asc_ValAxisSettings;
	prot = asc_ValAxisSettings.prototype;
	prot["putMinValRule"] = prot.putMinValRule;
	prot["putMinVal"] = prot.putMinVal;
	prot["putMaxValRule"] = prot.putMaxValRule;
	prot["putMaxVal"] = prot.putMaxVal;
	prot["putInvertValOrder"] = prot.putInvertValOrder;
	prot["putLogScale"] = prot.putLogScale;
	prot["putLogBase"] = prot.putLogBase;
	prot["putUnits"] = prot.putUnits;
	prot["putShowUnitsOnChart"] = prot.putShowUnitsOnChart;
	prot["putMajorTickMark"] = prot.putMajorTickMark;
	prot["putMinorTickMark"] = prot.putMinorTickMark;
	prot["putTickLabelsPos"] = prot.putTickLabelsPos;
	prot["putCrossesRule"] = prot.putCrossesRule;
	prot["putCrosses"] = prot.putCrosses;
	prot["putDispUnitsRule"] = prot.putDispUnitsRule;
	prot["getDispUnitsRule"] = prot.getDispUnitsRule;
	prot["putAxisType"] = prot.putAxisType;
	prot["getAxisType"] = prot.getAxisType;
	prot["getMinValRule"] = prot.getMinValRule;
	prot["getMinVal"] = prot.getMinVal;
	prot["getMaxValRule"] = prot.getMaxValRule;
	prot["getMaxVal"] = prot.getMaxVal;
	prot["getInvertValOrder"] = prot.getInvertValOrder;
	prot["getLogScale"] = prot.getLogScale;
	prot["getLogBase"] = prot.getLogBase;
	prot["getUnits"] = prot.getUnits;
	prot["getShowUnitsOnChart"] = prot.getShowUnitsOnChart;
	prot["getMajorTickMark"] = prot.getMajorTickMark;
	prot["getMinorTickMark"] = prot.getMinorTickMark;
	prot["getTickLabelsPos"] = prot.getTickLabelsPos;
	prot["getCrossesRule"] = prot.getCrossesRule;
	prot["getCrosses"] = prot.getCrosses;
	prot["setDefault"] = prot.setDefault;
	prot["getShow"] = prot.getShow;
	prot["putShow"] = prot.putShow;
	prot["putLabel"] = prot.putLabel;
	prot["getLabel"] = prot.getLabel;
	prot["putGridlines"] = prot.putGridlines;
	prot["getGridlines"] = prot.getGridlines;
	prot["putNumFmt"] = prot.putNumFmt;
	prot["getNumFmt"] = prot.getNumFmt;
	prot["isRadarAxis"] = prot.isRadarAxis;
	prot["getMajorUnit"] = prot.getMajorUnit;
	prot["getMinorUnit"] = prot.getMinorUnit;
	prot["putMajorUnit"] = prot.putMajorUnit;
	prot["putMinorUnit"] = prot.putMinorUnit;
	prot["getMajorUnitRule"] = prot.getMajorUnitRule;
	prot["getMinorUnitRule"] = prot.getMinorUnitRule;
	prot["putMajorUnitRule"] = prot.putMajorUnitRule;
	prot["putMinorUnitRule"] = prot.putMinorUnitRule;

	window["AscCommon"].asc_CatAxisSettings = asc_CatAxisSettings;
	prot = asc_CatAxisSettings.prototype;
	prot["putIntervalBetweenTick"] = prot.putIntervalBetweenTick;
	prot["putIntervalBetweenLabelsRule"] = prot.putIntervalBetweenLabelsRule;
	prot["putIntervalBetweenLabels"] = prot.putIntervalBetweenLabels;
	prot["putInvertCatOrder"] = prot.putInvertCatOrder;
	prot["putLabelsAxisDistance"] = prot.putLabelsAxisDistance;
	prot["putMajorTickMark"] = prot.putMajorTickMark;
	prot["putMinorTickMark"] = prot.putMinorTickMark;
	prot["putTickLabelsPos"] = prot.putTickLabelsPos;
	prot["putCrossesRule"] = prot.putCrossesRule;
	prot["putCrosses"] = prot.putCrosses;
	prot["putAxisType"] = prot.putAxisType;
	prot["putLabelsPosition"] = prot.putLabelsPosition;
	prot["putCrossMaxVal"] = prot.putCrossMaxVal;
	prot["putCrossMinVal"] = prot.putCrossMinVal;
	prot["getIntervalBetweenTick"] = prot.getIntervalBetweenTick;
	prot["getIntervalBetweenLabelsRule"] = prot.getIntervalBetweenLabelsRule;
	prot["getIntervalBetweenLabels"] = prot.getIntervalBetweenLabels;
	prot["getInvertCatOrder"] = prot.getInvertCatOrder;
	prot["getLabelsAxisDistance"] = prot.getLabelsAxisDistance;
	prot["getMajorTickMark"] = prot.getMajorTickMark;
	prot["getMinorTickMark"] = prot.getMinorTickMark;
	prot["getTickLabelsPos"] = prot.getTickLabelsPos;
	prot["getCrossesRule"] = prot.getCrossesRule;
	prot["getCrosses"] = prot.getCrosses;
	prot["getAxisType"] = prot.getAxisType;
	prot["getLabelsPosition"] = prot.getLabelsPosition;
	prot["getCrossMaxVal"] = prot.getCrossMaxVal;
	prot["getCrossMinVal"] = prot.getCrossMinVal;
	prot["setDefault"] = prot.setDefault;
	prot["getShow"] = prot.getShow;
	prot["putShow"] = prot.putShow;
	prot["getLabel"] = prot.getLabel;
	prot["putLabel"] = prot.putLabel;
	prot["putGridlines"] = prot.putGridlines;
	prot["getGridlines"] = prot.getGridlines;
	prot["putNumFmt"] = prot.putNumFmt;
	prot["getNumFmt"] = prot.getNumFmt;
	prot["getAuto"] = prot.getAuto;
	prot["putAuto"] = prot.putAuto;
	prot["isRadarAxis"] = prot.isRadarAxis;

	window["Asc"]["asc_ChartSettings"] = window["Asc"].asc_ChartSettings = asc_ChartSettings;
	prot = asc_ChartSettings.prototype;
	prot["putStyle"] = prot.putStyle;
	prot["putTitle"] = prot.putTitle;
	prot["putRowCols"] = prot.putRowCols;
	prot["putHorAxisLabel"] = prot.putHorAxisLabel;
	prot["putVertAxisLabel"] = prot.putVertAxisLabel;
	prot["putLegendPos"] = prot.putLegendPos;
	prot["putDataLabelsPos"] = prot.putDataLabelsPos;
	prot["putCatAx"] = prot.putCatAx;
	prot["putValAx"] = prot.putValAx;
	prot["getStyle"] = prot.getStyle;
	prot["getTitle"] = prot.getTitle;
	prot["getRowCols"] = prot.getRowCols;
	prot["getHorAxisLabel"] = prot.getHorAxisLabel;
	prot["getVertAxisLabel"] = prot.getVertAxisLabel;
	prot["getLegendPos"] = prot.getLegendPos;
	prot["getDataLabelsPos"] = prot.getDataLabelsPos;
	prot["getHorGridLines"] = prot.getHorGridLines;
	prot["putHorGridLines"] = prot.putHorGridLines;
	prot["getVertGridLines"] = prot.getVertGridLines;
	prot["putVertGridLines"] = prot.putVertGridLines;
	prot["getType"] = prot.getType;
	prot["putType"] = prot.putType;
	prot["putShowSerName"] = prot.putShowSerName;
	prot["getShowSerName"] = prot.getShowSerName;
	prot["putShowCatName"] = prot.putShowCatName;
	prot["getShowCatName"] = prot.getShowCatName;
	prot["putShowVal"] = prot.putShowVal;
	prot["getShowVal"] = prot.getShowVal;
	prot["putSeparator"] = prot.putSeparator;
	prot["getSeparator"] = prot.getSeparator;
	prot["putHorAxisProps"] = prot.putHorAxisProps;
	prot["getHorAxisProps"] = prot.getHorAxisProps;
	prot["putVertAxisProps"] = prot.putVertAxisProps;
	prot["getVertAxisProps"] = prot.getVertAxisProps;
	prot["putRange"] = prot.putRange;
	prot["getRange"] = prot.getRange;
	prot["putRanges"] = prot.putRanges;
	prot["getRanges"] = prot.getRanges;
	prot["putInColumns"] = prot.putInColumns;
	prot["getInColumns"] = prot.getInColumns;
	prot["getInRows"] = prot.getInRows;
	prot["putShowMarker"] = prot.putShowMarker;
	prot["getShowMarker"] = prot.getShowMarker;
	prot["putLine"] = prot.putLine;
	prot["getLine"] = prot.getLine;
	prot["putSmooth"] = prot.putSmooth;
	prot["getSmooth"] = prot.getSmooth;
	prot["changeType"] = prot.changeType;
	prot["putShowHorAxis"] = prot.putShowHorAxis;
	prot["getShowHorAxis"] = prot.getShowHorAxis;
	prot["putShowVerAxis"] = prot.putShowVerAxis;
	prot["getShowVerAxis"] = prot.getShowVerAxis;
	prot["getSeries"] = prot.getSeries;
	prot["getCatValues"] = prot.getCatValues;
	prot["switchRowCol"] = prot.switchRowCol;
	prot["addSeries"] = prot.addSeries;
	prot["addScatterSeries"] = prot.addScatterSeries;
	prot["getCatFormula"] = prot.getCatFormula;
	prot["setCatFormula"] = prot.setCatFormula;
	prot["isValidCatFormula"] = prot.isValidCatFormula;
	prot["setRange"] = prot.setRange;
	prot["isValidRange"] = prot.isValidRange;
	prot["startEdit"] = prot.startEdit;
	prot["endEdit"] = prot.endEdit;
	prot["cancelEdit"] = prot.cancelEdit;
	prot["startEditData"] = prot.startEditData;
	prot["cancelEditData"] = prot.cancelEditData;
	prot["endEditData"] = prot.endEditData;
	prot["getHorAxesProps"] = prot.getHorAxesProps;
	prot["getVertAxesProps"] = prot.getVertAxesProps;
	prot["getDepthAxesProps"] = prot.getDepthAxesProps;
	prot["getView3d"] = prot.getView3d;
	prot["putView3d"] = prot.putView3d;
	prot["setView3d"] = prot.setView3d;
	prot["getDisplayTrendlinesEquation"] = prot.getDisplayTrendlinesEquation;
	prot["putDisplayTrendlinesEquation"] = prot.putDisplayTrendlinesEquation;

	window["AscCommon"].asc_CRect = asc_CRect;
	prot = asc_CRect.prototype;
	prot["asc_getX"] = prot.asc_getX;
	prot["asc_getY"] = prot.asc_getY;
	prot["asc_getWidth"] = prot.asc_getWidth;
	prot["asc_getHeight"] = prot.asc_getHeight;

	window["AscCommon"].CColor = CColor;
	prot = CColor.prototype;
	prot["getR"] = prot.getR;
	prot["get_r"] = prot.get_r;
	prot["put_r"] = prot.put_r;
	prot["getG"] = prot.getG;
	prot["get_g"] = prot.get_g;
	prot["put_g"] = prot.put_g;
	prot["getB"] = prot.getB;
	prot["get_b"] = prot.get_b;
	prot["put_b"] = prot.put_b;
	prot["getA"] = prot.getA;
	prot["get_hex"] = prot.get_hex;

	window["Asc"]["asc_CColor"] = window["Asc"].asc_CColor = asc_CColor;
	prot = asc_CColor.prototype;
	prot["get_r"] = prot["asc_getR"] = prot.get_r = prot.asc_getR;
	prot["put_r"] = prot["asc_putR"] = prot.put_r = prot.asc_putR;
	prot["get_g"] = prot["asc_getG"] = prot.get_g = prot.asc_getG;
	prot["put_g"] = prot["asc_putG"] = prot.put_g = prot.asc_putG;
	prot["get_b"] = prot["asc_getB"] = prot.get_b = prot.asc_getB;
	prot["put_b"] = prot["asc_putB"] = prot.put_b = prot.asc_putB;
	prot["get_a"] = prot["asc_getA"] = prot.get_a = prot.asc_getA;
	prot["put_a"] = prot["asc_putA"] = prot.put_a = prot.asc_putA;
	prot["get_auto"] = prot["asc_getAuto"] = prot.get_auto = prot.asc_getAuto;
	prot["put_auto"] = prot["asc_putAuto"] = prot.put_auto = prot.asc_putAuto;
	prot["get_type"] = prot["asc_getType"] = prot.get_type = prot.asc_getType;
	prot["put_type"] = prot["asc_putType"] = prot.put_type = prot.asc_putType;
	prot["get_value"] = prot["asc_getValue"] = prot.get_value = prot.asc_getValue;
	prot["put_value"] = prot["asc_putValue"] = prot.put_value = prot.asc_putValue;
	prot["get_hex"] = prot["asc_getHex"] = prot.get_hex = prot.asc_getHex;
	prot["get_color"] = prot["asc_getColor"] = prot.get_color = prot.asc_getColor;
	prot["get_name"] = prot["asc_getName"] = prot.get_name = prot.asc_getName;
	prot["get_effectValue"] = prot["asc_getEffectValue"] = prot.get_effectValue = prot.asc_getEffectValue;
	prot["put_effectValue"] = prot["asc_putEffectValue"] = prot.put_effectValue = prot.asc_putEffectValue;
	prot["get_nameInColorScheme"] = prot["asc_getNameInColorScheme"] = prot.get_nameInColorScheme = prot.asc_getNameInColorScheme;


	window["Asc"]["asc_CTextBorder"] = window["Asc"].asc_CTextBorder = asc_CTextBorder;
	prot = asc_CTextBorder.prototype;
	prot["get_Color"] = prot["asc_getColor"] = prot.asc_getColor;
	prot["put_Color"] = prot["asc_putColor"] = prot.asc_putColor;
	prot["get_Size"] = prot["asc_getSize"] = prot.asc_getSize;
	prot["put_Size"] = prot["asc_putSize"] = prot.asc_putSize;
	prot["get_Value"] = prot["asc_getValue"] = prot.asc_getValue;
	prot["put_Value"] = prot["asc_putValue"] = prot.asc_putValue;
	prot["get_Space"] = prot["asc_getSpace"] = prot.asc_getSpace;
	prot["put_Space"] = prot["asc_putSpace"] = prot.asc_putSpace;
	prot["get_ForSelectedCells"] = prot["asc_getForSelectedCells"] = prot.asc_getForSelectedCells;
	prot["put_ForSelectedCells"] = prot["asc_putForSelectedCells"] = prot.asc_putForSelectedCells;

	window["Asc"]["asc_CParagraphBorders"] = window["Asc"].asc_CParagraphBorders = asc_CParagraphBorders;
	prot = asc_CParagraphBorders.prototype;
	prot["get_Left"] = prot["asc_getLeft"] = prot.asc_getLeft;
	prot["put_Left"] = prot["asc_putLeft"] = prot.asc_putLeft;
	prot["get_Top"] = prot["asc_getTop"] = prot.asc_getTop;
	prot["put_Top"] = prot["asc_putTop"] = prot.asc_putTop;
	prot["get_Right"] = prot["asc_getRight"] = prot.asc_getRight;
	prot["put_Right"] = prot["asc_putRight"] = prot.asc_putRight;
	prot["get_Bottom"] = prot["asc_getBottom"] = prot.asc_getBottom;
	prot["put_Bottom"] = prot["asc_putBottom"] = prot.asc_putBottom;
	prot["get_Between"] = prot["asc_getBetween"] = prot.asc_getBetween;
	prot["put_Between"] = prot["asc_putBetween"] = prot.asc_putBetween;

	window["AscCommon"].asc_CCustomListType = window["Asc"]["asc_CCustomListType"] = window["Asc"].asc_CCustomListType = asc_CCustomListType;
	prot = asc_CCustomListType.prototype;
	prot["setType"] = prot["asc_setType"] = prot.setType;
	prot["setImageId"] = prot["asc_setImageId"] = prot.setImageId;
	prot["setToken"] = prot["asc_setToken"] = prot.setToken;
	prot["setChar"] = prot["asc_setChar"] = prot.setChar;
	prot["setSpecialFont"] = prot["asc_setSpecialFont"] = prot.setSpecialFont;
	prot["setNumberingType"] = prot["asc_setNumberingType"] = prot.setNumberingType;
	prot["getType"] = prot["asc_getType"] = prot.getType;
	prot["getImageId"] = prot["asc_getImageId"] = prot.getImageId;
	prot["getToken"] = prot["asc_getToken"] = prot.getToken;
	prot["getChar"] = prot["asc_getChar"] = prot.getChar;
	prot["getSpecialFont"] = prot["asc_getSpecialFont"] = prot.getSpecialFont;
	prot["getNumberingType"] = prot["asc_getNumberingType"] = prot.getNumberingType;

	window["AscCommon"].asc_CListType = asc_CListType;
	prot = asc_CListType.prototype;
	prot["get_ListType"] = prot["asc_getListType"] = prot.asc_getListType;
	prot["get_ListSubType"] = prot["asc_getListSubType"] = prot.asc_getListSubType;
	prot["get_ListCustom"] = prot["asc_getListCustom"] = prot.asc_getListCustom;

	window["AscCommon"].asc_CTextFontFamily = asc_CTextFontFamily;
	window["AscCommon"]["asc_CTextFontFamily"] = asc_CTextFontFamily;
	prot = asc_CTextFontFamily.prototype;
	prot["get_Name"] = prot["asc_getName"] = prot.get_Name = prot.asc_getName;
	prot["get_Index"] = prot["asc_getIndex"] = prot.get_Index = prot.asc_getIndex;
	prot["put_Name"] = prot["asc_putName"] = prot.put_Name = prot.asc_putName;
	prot["put_Index"] = prot["asc_putIndex"] = prot.put_Index = prot.asc_putIndex;

	window["Asc"]["asc_CParagraphTab"] = window["Asc"].asc_CParagraphTab = asc_CParagraphTab;
	prot = asc_CParagraphTab.prototype;
	prot["get_Value"] = prot["asc_getValue"] = prot.asc_getValue;
	prot["put_Value"] = prot["asc_putValue"] = prot.asc_putValue;
	prot["get_Pos"] = prot["asc_getPos"] = prot.asc_getPos;
	prot["put_Pos"] = prot["asc_putPos"] = prot.asc_putPos;
	prot["get_Leader"] = prot["asc_getLeader"] = prot.asc_getLeader;
	prot["put_Leader"] = prot["asc_putLeader"] = prot.asc_putLeader;

	window["Asc"]["asc_CParagraphTabs"] = window["Asc"].asc_CParagraphTabs = asc_CParagraphTabs;
	prot = asc_CParagraphTabs.prototype;
	prot["get_Count"] = prot["asc_getCount"] = prot.asc_getCount;
	prot["get_Tab"] = prot["asc_getTab"] = prot.asc_getTab;
	prot["add_Tab"] = prot["asc_addTab"] = prot.asc_addTab;
	prot["clear"] = prot.clear = prot["asc_clear"] = prot.asc_clear;

	window["Asc"]["asc_CParagraphShd"] = window["Asc"].asc_CParagraphShd = asc_CParagraphShd;
	prot = asc_CParagraphShd.prototype;
	prot["get_Value"] = prot["asc_getValue"] = prot.asc_getValue;
	prot["put_Value"] = prot["asc_putValue"] = prot.asc_putValue;
	prot["get_Color"] = prot["asc_getColor"] = prot.asc_getColor;
	prot["put_Color"] = prot["asc_putColor"] = prot.asc_putColor;

	window["Asc"]["asc_CParagraphFrame"] = window["Asc"].asc_CParagraphFrame = asc_CParagraphFrame;
	prot = asc_CParagraphFrame.prototype;
	prot["asc_getDropCap"] = prot["get_DropCap"] = prot.asc_getDropCap;
	prot["asc_putDropCap"] = prot["put_DropCap"] = prot.asc_putDropCap;
	prot["asc_getH"] = prot["get_H"] = prot.asc_getH;
	prot["asc_putH"] = prot["put_H"] = prot.asc_putH;
	prot["asc_getHAnchor"] = prot["get_HAnchor"] = prot.asc_getHAnchor;
	prot["asc_putHAnchor"] = prot["put_HAnchor"] = prot.asc_putHAnchor;
	prot["asc_getHRule"] = prot["get_HRule"] = prot.asc_getHRule;
	prot["asc_putHRule"] = prot["put_HRule"] = prot.asc_putHRule;
	prot["asc_getHSpace"] = prot["get_HSpace"] = prot.asc_getHSpace;
	prot["asc_putHSpace"] = prot["put_HSpace"] = prot.asc_putHSpace;
	prot["asc_getLines"] = prot["get_Lines"] = prot.asc_getLines;
	prot["asc_putLines"] = prot["put_Lines"] = prot.asc_putLines;
	prot["asc_getVAnchor"] = prot["get_VAnchor"] = prot.asc_getVAnchor;
	prot["asc_putVAnchor"] = prot["put_VAnchor"] = prot.asc_putVAnchor;
	prot["asc_getVSpace"] = prot["get_VSpace"] = prot.asc_getVSpace;
	prot["asc_putVSpace"] = prot["put_VSpace"] = prot.asc_putVSpace;
	prot["asc_getW"] = prot["get_W"] = prot.asc_getW;
	prot["asc_putW"] = prot["put_W"] = prot.asc_putW;
	prot["asc_getWrap"] = prot["get_Wrap"] = prot.asc_getWrap;
	prot["asc_putWrap"] = prot["put_Wrap"] = prot.asc_putWrap;
	prot["asc_getX"] = prot["get_X"] = prot.asc_getX;
	prot["asc_putX"] = prot["put_X"] = prot.asc_putX;
	prot["asc_getXAlign"] = prot["get_XAlign"] = prot.asc_getXAlign;
	prot["asc_putXAlign"] = prot["put_XAlign"] = prot.asc_putXAlign;
	prot["asc_getY"] = prot["get_Y"] = prot.asc_getY;
	prot["asc_putY"] = prot["put_Y"] = prot.asc_putY;
	prot["asc_getYAlign"] = prot["get_YAlign"] = prot.asc_getYAlign;
	prot["asc_putYAlign"] = prot["put_YAlign"] = prot.asc_putYAlign;
	prot["asc_getBorders"] = prot["get_Borders"] = prot.asc_getBorders;
	prot["asc_putBorders"] = prot["put_Borders"] = prot.asc_putBorders;
	prot["asc_getShade"] = prot["get_Shade"] = prot.asc_getShade;
	prot["asc_putShade"] = prot["put_Shade"] = prot.asc_putShade;
	prot["asc_getFontFamily"] = prot["get_FontFamily"] = prot.asc_getFontFamily;
	prot["asc_putFontFamily"] = prot["put_FontFamily"] = prot.asc_putFontFamily;
	prot["asc_putFromDropCapMenu"] = prot["put_FromDropCapMenu"] = prot.asc_putFromDropCapMenu;

	window["AscCommon"].asc_CParagraphSpacing = asc_CParagraphSpacing;
	prot = asc_CParagraphSpacing.prototype;
	prot["get_Line"] = prot["asc_getLine"] = prot.asc_getLine;
	prot["put_Line"] = prot["asc_putLine"] = prot.asc_putLine;
	prot["get_LineRule"] = prot["asc_getLineRule"] = prot.asc_getLineRule;
	prot["put_LineRule"] = prot["asc_putLineRule"] = prot.asc_putLineRule;
	prot["get_Before"] = prot["asc_getBefore"] = prot.asc_getBefore;
	prot["put_Before"] = prot["asc_putBefore"] = prot.asc_putBefore;
	prot["get_After"] = prot["asc_getAfter"] = prot.asc_getAfter;
	prot["put_After"] = prot["asc_putAfter"] = prot.asc_putAfter;

	window["Asc"]["asc_CParagraphInd"] = window["Asc"].asc_CParagraphInd = asc_CParagraphInd;
	prot = asc_CParagraphInd.prototype;
	prot["get_Left"] = prot["asc_getLeft"] = prot.asc_getLeft;
	prot["put_Left"] = prot["asc_putLeft"] = prot.asc_putLeft;
	prot["get_Right"] = prot["asc_getRight"] = prot.asc_getRight;
	prot["put_Right"] = prot["asc_putRight"] = prot.asc_putRight;
	prot["get_FirstLine"] = prot["asc_getFirstLine"] = prot.asc_getFirstLine;
	prot["put_FirstLine"] = prot["asc_putFirstLine"] = prot.asc_putFirstLine;

	window["Asc"]["asc_CParagraphProperty"] = window["Asc"].asc_CParagraphProperty = asc_CParagraphProperty;
	prot = asc_CParagraphProperty.prototype;
	prot["get_RtlDirection"] = prot["asc_getRtlDirection"] = prot.asc_getRtlDirection;
	prot["put_RtlDirection"] = prot["asc_putRtlDirection"] = prot.asc_putRtlDirection;
	prot["get_ContextualSpacing"] = prot["asc_getContextualSpacing"] = prot.asc_getContextualSpacing;
	prot["put_ContextualSpacing"] = prot["asc_putContextualSpacing"] = prot.asc_putContextualSpacing;
	prot["get_Ind"] = prot["asc_getInd"] = prot.asc_getInd;
	prot["put_Ind"] = prot["asc_putInd"] = prot.asc_putInd;
	prot["get_Jc"] = prot["asc_getJc"] = prot.asc_getJc;
	prot["put_Jc"] = prot["asc_putJc"] = prot.asc_putJc;
	prot["get_KeepLines"] = prot["asc_getKeepLines"] = prot.asc_getKeepLines;
	prot["put_KeepLines"] = prot["asc_putKeepLines"] = prot.asc_putKeepLines;
	prot["get_KeepNext"] = prot["asc_getKeepNext"] = prot.asc_getKeepNext;
	prot["put_KeepNext"] = prot["asc_putKeepNext"] = prot.asc_putKeepNext;
	prot["get_PageBreakBefore"] = prot["asc_getPageBreakBefore"] = prot.asc_getPageBreakBefore;
	prot["put_PageBreakBefore"] = prot["asc_putPageBreakBefore"] = prot.asc_putPageBreakBefore;
	prot["get_WidowControl"] = prot["asc_getWidowControl"] = prot.asc_getWidowControl;
	prot["put_WidowControl"] = prot["asc_putWidowControl"] = prot.asc_putWidowControl;
	prot["get_Spacing"] = prot["asc_getSpacing"] = prot.asc_getSpacing;
	prot["put_Spacing"] = prot["asc_putSpacing"] = prot.asc_putSpacing;
	prot["get_Borders"] = prot["asc_getBorders"] = prot.asc_getBorders;
	prot["put_Borders"] = prot["asc_putBorders"] = prot.asc_putBorders;
	prot["get_Shade"] = prot["asc_getShade"] = prot.asc_getShade;
	prot["put_Shade"] = prot["asc_putShade"] = prot.asc_putShade;
	prot["get_Locked"] = prot["asc_getLocked"] = prot.asc_getLocked;
	prot["get_CanAddTable"] = prot["asc_getCanAddTable"] = prot.asc_getCanAddTable;
	prot["get_Subscript"] = prot["asc_getSubscript"] = prot.asc_getSubscript;
	prot["put_Subscript"] = prot["asc_putSubscript"] = prot.asc_putSubscript;
	prot["get_Superscript"] = prot["asc_getSuperscript"] = prot.asc_getSuperscript;
	prot["put_Superscript"] = prot["asc_putSuperscript"] = prot.asc_putSuperscript;
	prot["get_SmallCaps"] = prot["asc_getSmallCaps"] = prot.asc_getSmallCaps;
	prot["put_SmallCaps"] = prot["asc_putSmallCaps"] = prot.asc_putSmallCaps;
	prot["get_AllCaps"] = prot["asc_getAllCaps"] = prot.asc_getAllCaps;
	prot["put_AllCaps"] = prot["asc_putAllCaps"] = prot.asc_putAllCaps;
	prot["get_Strikeout"] = prot["asc_getStrikeout"] = prot.asc_getStrikeout;
	prot["put_Strikeout"] = prot["asc_putStrikeout"] = prot.asc_putStrikeout;
	prot["get_DStrikeout"] = prot["asc_getDStrikeout"] = prot.asc_getDStrikeout;
	prot["put_DStrikeout"] = prot["asc_putDStrikeout"] = prot.asc_putDStrikeout;
	prot["get_TextSpacing"] = prot["asc_getTextSpacing"] = prot.asc_getTextSpacing;
	prot["put_TextSpacing"] = prot["asc_putTextSpacing"] = prot.asc_putTextSpacing;
	prot["get_Position"] = prot["asc_getPosition"] = prot.asc_getPosition;
	prot["put_Position"] = prot["asc_putPosition"] = prot.asc_putPosition;
	prot["get_Tabs"] = prot["asc_getTabs"] = prot.asc_getTabs;
	prot["put_Tabs"] = prot["asc_putTabs"] = prot.asc_putTabs;
	prot["get_DefaultTab"] = prot["asc_getDefaultTab"] = prot.asc_getDefaultTab;
	prot["put_DefaultTab"] = prot["asc_putDefaultTab"] = prot.asc_putDefaultTab;
	prot["get_FramePr"] = prot["asc_getFramePr"] = prot.asc_getFramePr;
	prot["put_FramePr"] = prot["asc_putFramePr"] = prot.asc_putFramePr;
	prot["get_CanAddDropCap"] = prot["asc_getCanAddDropCap"] = prot.asc_getCanAddDropCap;
	prot["get_CanAddImage"] = prot["asc_getCanAddImage"] = prot.asc_getCanAddImage;
	prot["get_OutlineLvl"] = prot["asc_getOutlineLvl"] = prot.asc_getOutlineLvl;
	prot["put_OutlineLvl"] = prot["asc_putOutLineLvl"] = prot.asc_putOutLineLvl;
	prot["get_OutlineLvlStyle"] = prot["asc_getOutlineLvlStyle"] = prot.asc_getOutlineLvlStyle;
	prot["get_SuppressLineNumbers"] = prot["asc_getSuppressLineNumbers"] = prot.asc_getSuppressLineNumbers;
	prot["put_SuppressLineNumbers"] = prot["asc_putSuppressLineNumbers"] = prot.asc_putSuppressLineNumbers;
	prot["put_Bullet"] = prot["asc_putBullet"] = prot.asc_putBullet;
	prot["get_Bullet"] = prot["asc_getBullet"] = prot.asc_getBullet;
	prot["put_BulletSize"] = prot["asc_putBulletSize"] = prot.asc_putBulletSize;
	prot["get_BulletSize"] = prot["asc_getBulletSize"] = prot.asc_getBulletSize;
	prot["put_BulletColor"] = prot["asc_putBulletColor"] = prot.asc_putBulletColor;
	prot["get_BulletColor"] = prot["asc_getBulletColor"] = prot.asc_getBulletColor;
	prot["put_NumStartAt"] = prot["asc_putNumStartAt"] = prot.asc_putNumStartAt;
	prot["get_NumStartAt"] = prot["asc_getNumStartAt"] = prot.asc_getNumStartAt;
	prot["get_BulletFont"] = prot["asc_getBulletFont"] = prot.asc_getBulletFont;
	prot["put_BulletFont"] = prot["asc_putBulletFont"] = prot.asc_putBulletFont;
	prot["get_BulletSymbol"] = prot["asc_getBulletSymbol"] = prot.asc_getBulletSymbol;
	prot["put_BulletSymbol"] = prot["asc_putBulletSymbol"] = prot.asc_putBulletSymbol;
	prot["can_DeleteBlockContentControl"] = prot["asc_canDeleteBlockContentControl"] = prot.asc_canDeleteBlockContentControl;
	prot["can_EditBlockContentControl"] = prot["asc_canEditBlockContentControl"] = prot.asc_canEditBlockContentControl;
	prot["can_DeleteInlineContentControl"] = prot["asc_canDeleteInlineContentControl"] = prot.asc_canDeleteInlineContentControl;
	prot["can_EditInlineContentControl"] = prot["asc_canEditInlineContentControl"] = prot.asc_canEditInlineContentControl;
	prot["get_Ligatures"] = prot["asc_getLigatures"] = prot.asc_getLigatures;
	prot["put_Ligatures"] = prot["asc_putLigatures"] = prot.asc_putLigatures;

	window["AscCommon"].asc_CTexture = asc_CTexture;
	prot = asc_CTexture.prototype;
	prot["get_id"] = prot["asc_getId"] = prot.asc_getId;
	prot["get_image"] = prot["asc_getImage"] = prot.asc_getImage;

	window["AscCommon"].asc_CImageSize = asc_CImageSize;
	prot = asc_CImageSize.prototype;
	prot["get_ImageWidth"] = prot["asc_getImageWidth"] = prot.asc_getImageWidth;
	prot["get_ImageHeight"] = prot["asc_getImageHeight"] = prot.asc_getImageHeight;
	prot["get_IsCorrect"] = prot["asc_getIsCorrect"] = prot.asc_getIsCorrect;

	window["Asc"]["asc_CPaddings"] = window["Asc"].asc_CPaddings = asc_CPaddings;
	prot = asc_CPaddings.prototype;
	prot["get_Left"] = prot["asc_getLeft"] = prot.asc_getLeft;
	prot["put_Left"] = prot["asc_putLeft"] = prot.asc_putLeft;
	prot["get_Top"] = prot["asc_getTop"] = prot.asc_getTop;
	prot["put_Top"] = prot["asc_putTop"] = prot.asc_putTop;
	prot["get_Bottom"] = prot["asc_getBottom"] = prot.asc_getBottom;
	prot["put_Bottom"] = prot["asc_putBottom"] = prot.asc_putBottom;
	prot["get_Right"] = prot["asc_getRight"] = prot.asc_getRight;
	prot["put_Right"] = prot["asc_putRight"] = prot.asc_putRight;

	window["Asc"]["asc_CShapeProperty"] = window["Asc"].asc_CShapeProperty = asc_CShapeProperty;
	prot = asc_CShapeProperty.prototype;
	prot["get_type"] = prot["asc_getType"] = prot.asc_getType;
	prot["put_type"] = prot["asc_putType"] = prot.asc_putType;
	prot["get_fill"] = prot["asc_getFill"] = prot.asc_getFill;
	prot["put_fill"] = prot["asc_putFill"] = prot.asc_putFill;
	prot["get_stroke"] = prot["asc_getStroke"] = prot.asc_getStroke;
	prot["put_stroke"] = prot["asc_putStroke"] = prot.asc_putStroke;
	prot["get_paddings"] = prot["asc_getPaddings"] = prot.asc_getPaddings;
	prot["put_paddings"] = prot["asc_putPaddings"] = prot.asc_putPaddings;
	prot["get_CanFill"] = prot["asc_getCanFill"] = prot.asc_getCanFill;
	prot["put_CanFill"] = prot["asc_putCanFill"] = prot.asc_putCanFill;
	prot["get_CanChangeArrows"] = prot["asc_getCanChangeArrows"] = prot.asc_getCanChangeArrows;
	prot["set_CanChangeArrows"] = prot["asc_setCanChangeArrows"] = prot.asc_setCanChangeArrows;
	prot["get_FromChart"] = prot["asc_getFromChart"] = prot.asc_getFromChart;
	prot["set_FromChart"] = prot["asc_setFromChart"] = prot.asc_setFromChart;
	prot["set_FromSmartArt"] = prot["asc_setFromSmartArt"] = prot.asc_setFromSmartArt;
	prot["get_FromSmartArt"] = prot["asc_getFromSmartArt"] = prot.asc_getFromSmartArt;
	prot["set_FromSmartArtInternal"] = prot["asc_setFromSmartArtInternal"] = prot.asc_setFromSmartArtInternal;
	prot["get_FromSmartArtInternal"] = prot["asc_getFromSmartArtInternal"] = prot.asc_getFromSmartArtInternal;
	prot["get_FromGroup"] = prot["asc_getFromGroup"] = prot.asc_getFromGroup;
	prot["set_FromGroup"] = prot["asc_setFromGroup"] = prot.asc_setFromGroup;
	prot["get_Locked"] = prot["asc_getLocked"] = prot.asc_getLocked;
	prot["set_Locked"] = prot["asc_setLocked"] = prot.asc_setLocked;
	prot["get_Width"] = prot["asc_getWidth"] = prot.asc_getWidth;
	prot["put_Width"] = prot["asc_putWidth"] = prot.asc_putWidth;
	prot["get_Height"] = prot["asc_getHeight"] = prot.asc_getHeight;
	prot["put_Height"] = prot["asc_putHeight"] = prot.asc_putHeight;
	prot["get_VerticalTextAlign"] = prot["asc_getVerticalTextAlign"] = prot.asc_getVerticalTextAlign;
	prot["put_VerticalTextAlign"] = prot["asc_putVerticalTextAlign"] = prot.asc_putVerticalTextAlign;
	prot["get_Vert"] = prot["asc_getVert"] = prot.asc_getVert;
	prot["put_Vert"] = prot["asc_putVert"] = prot.asc_putVert;
	prot["get_TextArtProperties"] = prot["asc_getTextArtProperties"] = prot.asc_getTextArtProperties;
	prot["put_TextArtProperties"] = prot["asc_putTextArtProperties"] = prot.asc_putTextArtProperties;
	prot["get_LockAspect"] = prot["asc_getLockAspect"] = prot.asc_getLockAspect;
	prot["put_LockAspect"] = prot["asc_putLockAspect"] = prot.asc_putLockAspect;
	prot["get_Title"] = prot["asc_getTitle"] = prot.asc_getTitle;
	prot["put_Title"] = prot["asc_putTitle"] = prot.asc_putTitle;
	prot["get_Description"] = prot["asc_getDescription"] = prot.asc_getDescription;
	prot["put_Description"] = prot["asc_putDescription"] = prot.asc_putDescription;
	prot["get_Name"] = prot["asc_getName"] = prot.asc_getName;
	prot["put_Name"] = prot["asc_putName"] = prot.asc_putName;
	prot["get_ColumnNumber"] = prot["asc_getColumnNumber"] = prot.asc_getColumnNumber;
	prot["put_ColumnNumber"] = prot["asc_putColumnNumber"] = prot.asc_putColumnNumber;
	prot["get_ColumnSpace"] = prot["asc_getColumnSpace"] = prot.asc_getColumnSpace;
	prot["get_TextFitType"] = prot["asc_getTextFitType"] = prot.asc_getTextFitType;
	prot["get_VertOverflowType"] = prot["asc_getVertOverflowType"] = prot.asc_getVertOverflowType;
	prot["put_ColumnSpace"] = prot["asc_putColumnSpace"] = prot.asc_putColumnSpace;
	prot["put_TextFitType"] = prot["asc_putTextFitType"] = prot.asc_putTextFitType;
	prot["put_VertOverflowType"] = prot["asc_putVertOverflowType"] = prot.asc_putVertOverflowType;
	prot["get_SignatureId"] = prot["asc_getSignatureId"] = prot.asc_getSignatureId;
	prot["put_SignatureId"] = prot["asc_putSignatureId"] = prot.asc_putSignatureId;
	prot["get_FromImage"] = prot["asc_getFromImage"] = prot.asc_getFromImage;
	prot["put_FromImage"] = prot["asc_putFromImage"] = prot.asc_putFromImage;
	prot["get_Rot"] = prot["asc_getRot"] = prot.asc_getRot;
	prot["put_Rot"] = prot["asc_putRot"] = prot.asc_putRot;
	prot["get_RotAdd"] = prot["asc_getRotAdd"] = prot.asc_getRotAdd;
	prot["put_RotAdd"] = prot["asc_putRotAdd"] = prot.asc_putRotAdd;
	prot["get_FlipH"] = prot["asc_getFlipH"] = prot.asc_getFlipH;
	prot["put_FlipH"] = prot["asc_putFlipH"] = prot.asc_putFlipH;
	prot["get_FlipV"] = prot["asc_getFlipV"] = prot.asc_getFlipV;
	prot["put_FlipV"] = prot["asc_putFlipV"] = prot.asc_putFlipV;
	prot["get_FlipHInvert"] = prot["asc_getFlipHInvert"] = prot.asc_getFlipHInvert;
	prot["put_FlipHInvert"] = prot["asc_putFlipHInvert"] = prot.asc_putFlipHInvert;
	prot["get_FlipVInvert"] = prot["asc_getFlipVInvert"] = prot.asc_getFlipVInvert;
	prot["put_FlipVInvert"] = prot["asc_putFlipVInvert"] = prot.asc_putFlipVInvert;
	prot["put_Shadow"] = prot.put_Shadow = prot["put_shadow"] = prot.put_shadow = prot["asc_putShadow"] = prot.asc_putShadow;
	prot["get_Shadow"] = prot.get_Shadow = prot["get_shadow"] = prot.get_shadow = prot["asc_getShadow"] = prot.asc_getShadow;
	prot["put_Anchor"] = prot.put_Anchor = prot["asc_putAnchor"] = prot.asc_putAnchor;
	prot["get_Anchor"] = prot.get_Anchor = prot["asc_getAnchor"] = prot.asc_getAnchor;
	prot["get_ProtectionLockText"] = prot["asc_getProtectionLockText"] = prot.asc_getProtectionLockText;
	prot["put_ProtectionLockText"] = prot["asc_putProtectionLockText"] = prot.asc_putProtectionLockText;
	prot["get_ProtectionLocked"] = prot["asc_getProtectionLocked"] = prot.asc_getProtectionLocked;
	prot["put_ProtectionLocked"] = prot["asc_putProtectionLocked"] = prot.asc_putProtectionLocked;
	prot["get_ProtectionPrint"] = prot["asc_getProtectionPrint"] = prot.asc_getProtectionPrint;
	prot["put_ProtectionPrint"] = prot["asc_putProtectionPrint"] = prot.asc_putProtectionPrint;
	prot["get_Position"] = prot["asc_getPosition"] = prot.asc_getPosition;
	prot["put_Position"] = prot["asc_putPosition"] = prot.asc_putPosition;
	prot["get_IsMotionPath"] = prot["asc_getIsMotionPath"] = prot.asc_getIsMotionPath;
	prot["asc_getCanEditText"]		= prot.asc_getCanEditText;
	prot["asc_setCanEditText"]		= prot.asc_setCanEditText;

	window["Asc"]["asc_CAnnotProperty"] = window["Asc"].asc_CAnnotProperty = asc_CAnnotProperty;
	prot = asc_CAnnotProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_putType"]				= prot.asc_putType;
	prot["asc_getFill"]				= prot.asc_getFill;
	prot["asc_putFill"]				= prot.asc_putFill;
	prot["asc_getStroke"]			= prot.asc_getStroke;
	prot["asc_putStroke"]			= prot.asc_putStroke;
	prot["asc_getCanFill"]			= prot.asc_getCanFill;
	prot["asc_putCanFill"]			= prot.asc_putCanFill;
	prot["asc_getCanChangeArrows"]	= prot.asc_getCanChangeArrows;
	prot["asc_setCanChangeArrows"]	= prot.asc_setCanChangeArrows;
	prot["asc_getSubject"]			= prot.asc_getSubject;
	prot["asc_setSubject"]			= prot.asc_setSubject;
	prot["asc_getCanEditText"]		= prot.asc_getCanEditText;
	prot["asc_setCanEditText"]		= prot.asc_setCanEditText;

	window["Asc"]["asc_CBaseFieldProperty"] = window["Asc"].asc_CBaseFieldProperty = asc_CBaseFieldProperty;
	prot = asc_CBaseFieldProperty.prototype;
	prot["asc_getType"]			= prot.asc_getType;
	prot["asc_putType"]			= prot.asc_putType;
	prot["asc_getName"]			= prot.asc_getName;
	prot["asc_putName"]			= prot.asc_putName;
	prot["asc_getRequired"]		= prot.asc_getRequired;
	prot["asc_putRequired"]		= prot.asc_putRequired;
	prot["asc_getReadOnly"]		= prot.asc_getReadOnly;
	prot["asc_putReadOnly"]		= prot.asc_putReadOnly;
	prot["asc_getRot"]			= prot.asc_getRot;
	prot["asc_putRot"]			= prot.asc_putRot;
	prot["asc_getDisplay"]		= prot.asc_getDisplay;
	prot["asc_putDisplay"]		= prot.asc_putDisplay;
	prot["asc_getFill"]			= prot.asc_getFill;
	prot["asc_putFill"]			= prot.asc_putFill;
	prot["asc_getStroke"]		= prot.asc_getStroke;
	prot["asc_putStroke"]		= prot.asc_putStroke;
	prot["asc_getStrokeWidth"]	= prot.asc_getStrokeWidth;
	prot["asc_putStrokeWidth"]	= prot.asc_putStrokeWidth;
	prot["asc_getStrokeStyle"]	= prot.asc_getStrokeStyle;
	prot["asc_putStrokeStyle"]	= prot.asc_putStrokeStyle;
	prot["asc_getPropLocked"]	= prot.asc_getPropLocked;
	prot["asc_putPropLocked"]	= prot.asc_putPropLocked;
	prot["get_Locked"]			= prot.get_Locked;
	prot["put_Locked"]			= prot.put_Locked;
	prot["asc_getFieldProps"]	= prot.asc_getFieldProps;
	prot["asc_putFieldProps"]	= prot.asc_putFieldProps;

	window["Asc"]["asc_CTextFieldProperty"] = window["Asc"].asc_CTextFieldProperty = asc_CTextFieldProperty;
	prot = asc_CTextFieldProperty.prototype;
	prot["asc_getDefaultValue"]			= prot.asc_getDefaultValue;
	prot["asc_putDefaultValue"]			= prot.asc_putDefaultValue;
	prot["asc_getMultiline"]			= prot.asc_getMultiline;
	prot["asc_putMultiline"]			= prot.asc_putMultiline;
	prot["asc_getScrollLongText"]		= prot.asc_getScrollLongText;
	prot["asc_putScrollLongText"]		= prot.asc_putScrollLongText;
	prot["asc_getCharLimit"]			= prot.asc_getCharLimit;
	prot["asc_putCharLimit"]			= prot.asc_putCharLimit;
	prot["asc_getComb"]					= prot.asc_getComb;
	prot["asc_putComb"]					= prot.asc_putComb;
	prot["asc_getPlaceholder"]			= prot.asc_getPlaceholder;
	prot["asc_putPlaceholder"]			= prot.asc_putPlaceholder;
	prot["asc_getAutoFit"]				= prot.asc_getAutoFit;
	prot["asc_putAutoFit"]				= prot.asc_putAutoFit;
	prot["asc_getFormat"]				= prot.asc_getFormat;
	prot["asc_putFormat"]				= prot.asc_putFormat;
	prot["asc_getValidate"]				= prot.asc_getValidate;
	prot["asc_putValidate"]				= prot.asc_putValidate;

	window["Asc"]["asc_CComboboxFieldProperty"] = window["Asc"].asc_CComboboxFieldProperty = asc_CComboboxFieldProperty;
	prot = asc_CComboboxFieldProperty.prototype;
	prot["asc_getOptions"]				= prot.asc_getOptions;
	prot["asc_putOptions"]				= prot.asc_putOptions;
	prot["asc_getCommitOnSelChange"]	= prot.asc_getCommitOnSelChange;
	prot["asc_putCommitOnSelChange"]	= prot.asc_putCommitOnSelChange;
	prot["asc_getEditable"]				= prot.asc_getEditable;
	prot["asc_putEditable"]				= prot.asc_putEditable;
	prot["asc_getPlaceholder"]			= prot.asc_getPlaceholder;
	prot["asc_putPlaceholder"]			= prot.asc_putPlaceholder;
	prot["asc_getAutoFit"]				= prot.asc_getAutoFit;
	prot["asc_putAutoFit"]				= prot.asc_putAutoFit;
	prot["asc_getFormat"]				= prot.asc_getFormat;
	prot["asc_putFormat"]				= prot.asc_putFormat;
	prot["asc_getValidate"]				= prot.asc_getValidate;
	prot["asc_putValidate"]				= prot.asc_putValidate;

	window["Asc"]["asc_CListboxFieldProperty"] = window["Asc"].asc_CListboxFieldProperty = asc_CListboxFieldProperty;
	prot = asc_CListboxFieldProperty.prototype;
	prot["asc_getOptions"]				= prot.asc_getOptions;
	prot["asc_putOptions"]				= prot.asc_putOptions;
	prot["asc_getCommitOnSelChange"]	= prot.asc_getCommitOnSelChange;
	prot["asc_putCommitOnSelChange"]	= prot.asc_putCommitOnSelChange;
	prot["asc_getMultipleSelection"]	= prot.asc_getMultipleSelection;
	prot["asc_putMultipleSelection"]	= prot.asc_putMultipleSelection;

	window["Asc"]["asc_CCheckboxFieldProperty"] = window["Asc"].asc_CCheckboxFieldProperty = asc_CCheckboxFieldProperty;
	prot = asc_CCheckboxFieldProperty.prototype;
	prot["asc_getCheckboxStyle"]		= prot.asc_getCheckboxStyle;
	prot["asc_putCheckboxStyle"]		= prot.asc_putCheckboxStyle;
	prot["asc_getExportValue"]		= prot.asc_getExportValue;
	prot["asc_putExportValue"]		= prot.asc_putExportValue;
	prot["asc_getDefaultChecked"]	= prot.asc_getDefaultChecked;
	prot["asc_putDefaultChecked"]	= prot.asc_putDefaultChecked;
	prot["asc_getToggleToOff"]		= prot.asc_getToggleToOff;
	prot["asc_putToggleToOff"]		= prot.asc_putToggleToOff;

	window["Asc"]["asc_CRadiobuttonFieldProperty"] = window["Asc"].asc_CRadiobuttonFieldProperty = asc_CRadiobuttonFieldProperty;
	prot = asc_CRadiobuttonFieldProperty.prototype;
	prot["asc_getRadiosInUnison"]	= prot.asc_getRadiosInUnison;
	prot["asc_putRadiosInUnison"]	= prot.asc_putRadiosInUnison;

	window["Asc"]["asc_CButtonFieldProperty"] = window["Asc"].asc_CButtonFieldProperty = asc_CButtonFieldProperty;
	prot = asc_CButtonFieldProperty.prototype;
	prot["asc_getHighlight"]	= prot.asc_getHighlight;
	prot["asc_putHighlight"]	= prot.asc_putHighlight;
	prot["asc_getLayout"]		= prot.asc_getLayout;
	prot["asc_putLayout"]		= prot.asc_putLayout;
	prot["asc_getScaleWhen"]	= prot.asc_getScaleWhen;
	prot["asc_putScaleWhen"]	= prot.asc_putScaleWhen;
	prot["asc_getScaleHow"]		= prot.asc_getScaleHow;
	prot["asc_putScaleHow"]		= prot.asc_putScaleHow;
	prot["asc_getFitBounds"]	= prot.asc_getFitBounds;
	prot["asc_putFitBounds"]	= prot.asc_putFitBounds;
	prot["asc_getIconPos"]		= prot.asc_getIconPos;
	prot["asc_putIconPos"]		= prot.asc_putIconPos;
	prot["asc_getBehavior"]		= prot.asc_getBehavior;
	prot["asc_putBehavior"]		= prot.asc_putBehavior;
	prot["asc_getNormalCaption"]= prot.asc_getNormalCaption;
	prot["asc_putNormalCaption"]= prot.asc_putNormalCaption;
	prot["asc_getNormalImage"]	= prot.asc_getNormalImage;
	prot["asc_putNormalImage"]	= prot.asc_putNormalImage;
	prot["asc_getHoverCaption"]	= prot.asc_getHoverCaption;
	prot["asc_putHoverCaption"]	= prot.asc_putHoverCaption;
	prot["asc_getHoverImage"]	= prot.asc_getHoverImage;
	prot["asc_putHoverImage"]	= prot.asc_putHoverImage;
	prot["asc_getDownCaption"]	= prot.asc_getDownCaption;
	prot["asc_putDownCaption"]	= prot.asc_putDownCaption;
	prot["asc_getDownImage"]	= prot.asc_getDownImage;
	prot["asc_putDownImage"]	= prot.asc_putDownImage;
	prot["asc_putCurrentState"]	= prot.asc_putCurrentState;
	prot["asc_getCurrentState"]	= prot.asc_getCurrentState;
	prot["put_DivId"]			= prot.put_DivId;
	prot["drawTexture"]			= prot.drawTexture;
	prot["put_ImageUrl"]		= prot.put_ImageUrl;
	prot["showFileDialog"]		= prot.showFileDialog;

	window["Asc"]["asc_CFieldNumberFormatProperty"] = window["Asc"].asc_CFieldNumberFormatProperty = asc_CFieldNumberFormatProperty;
	prot = asc_CFieldNumberFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getDecimals"]			= prot.asc_getDecimals;
	prot["asc_putDecimals"]			= prot.asc_putDecimals;
	prot["asc_getSepStyle"]			= prot.asc_getSepStyle;
	prot["asc_putSepStyle"]			= prot.asc_putSepStyle;
	prot["asc_getNegStyle"]			= prot.asc_getNegStyle;
	prot["asc_putNegStyle"]			= prot.asc_putNegStyle;
	prot["asc_getCurrency"]			= prot.asc_getCurrency;
	prot["asc_putCurrency"]			= prot.asc_putCurrency;
	prot["asc_getCurrencyPrepend"]	= prot.asc_getCurrencyPrepend;
	prot["asc_putCurrencyPrepend"]	= prot.asc_putCurrencyPrepend;

	
	window["Asc"]["asc_CFieldPercentageFormatProperty"] = window["Asc"].asc_CFieldPercentageFormatProperty = asc_CFieldPercentageFormatProperty;
	prot = asc_CFieldPercentageFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getDecimals"]			= prot.asc_getDecimals;
	prot["asc_putDecimals"]			= prot.asc_putDecimals;
	prot["asc_getSepStyle"]			= prot.asc_getSepStyle;
	prot["asc_putSepStyle"]			= prot.asc_putSepStyle;

	window["Asc"]["asc_CFieldDateFormatProperty"] = window["Asc"].asc_CFieldDateFormatProperty = asc_CFieldDateFormatProperty;
	prot = asc_CFieldDateFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getFormat"]			= prot.asc_getFormat;
	prot["asc_putFormat"]			= prot.asc_putFormat;
	
	window["Asc"]["asc_CFieldTimeFormatProperty"] = window["Asc"].asc_CFieldTimeFormatProperty = asc_CFieldTimeFormatProperty;
	prot = asc_CFieldTimeFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getFormat"]			= prot.asc_getFormat;
	prot["asc_putFormat"]			= prot.asc_putFormat;
	
	window["Asc"]["asc_CFieldSpecialFormatProperty"] = window["Asc"].asc_CFieldSpecialFormatProperty = asc_CFieldSpecialFormatProperty;
	prot = asc_CFieldSpecialFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getFormat"]			= prot.asc_getFormat;
	prot["asc_putFormat"]			= prot.asc_putFormat;
	prot["asc_getMask"]				= prot.asc_getMask;
	prot["asc_putMask"]				= prot.asc_putMask;
	
	window["Asc"]["asc_CFieldRegularFormatProperty"] = window["Asc"].asc_CFieldRegularFormatProperty = asc_CFieldRegularFormatProperty;
	prot = asc_CFieldRegularFormatProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_getRegExp"]			= prot.asc_getRegExp;
	prot["asc_putRegExp"]			= prot.asc_putRegExp;
	
	window["Asc"]["asc_CFieldValidateProperty"] = window["Asc"].asc_CFieldValidateProperty = asc_CFieldValidateProperty;
	prot = asc_CFieldValidateProperty.prototype;
	prot["asc_getType"]				= prot.asc_getType;
	prot["asc_putType"]				= prot.asc_putType;
	prot["asc_getBeGreaterThen"]	= prot.asc_getBeGreaterThen;
	prot["asc_putBeGreaterThen"]	= prot.asc_putBeGreaterThen;
	prot["asc_getGreaterThen"]		= prot.asc_getGreaterThen;
	prot["asc_putGreaterThen"]		= prot.asc_putGreaterThen;
	prot["asc_getBeLessThen"]		= prot.asc_getBeLessThen;
	prot["asc_putBeLessThen"]		= prot.asc_putBeLessThen;
	prot["asc_getLessThen"]			= prot.asc_getLessThen;
	prot["asc_putLessThen"]			= prot.asc_putLessThen;

	window["Asc"]["asc_CPdfPageProperty"] = window["Asc"].asc_CPdfPageProperty = asc_CPdfPageProperty;
	prot = asc_CPdfPageProperty.prototype;
	prot["asc_getDeleteLock"]	= prot.asc_getDeleteLock;
	prot["asc_putDeleteLock"]	= prot.asc_putDeleteLock;
	prot["asc_getRotateLock"]	= prot.asc_getRotateLock;
	prot["asc_putRotateLock"]	= prot.asc_putRotateLock;
	prot["asc_getEditLock"]		= prot.asc_getEditLock;
	prot["asc_putEditLock"]		= prot.asc_putEditLock;

	window["Asc"]["asc_TextArtProperties"] = window["Asc"].asc_TextArtProperties = asc_TextArtProperties;
	prot = asc_TextArtProperties.prototype;
	prot["asc_putFill"] = prot.asc_putFill;
	prot["asc_getFill"] = prot.asc_getFill;
	prot["asc_putLine"] = prot.asc_putLine;
	prot["asc_getLine"] = prot.asc_getLine;
	prot["asc_putForm"] = prot.asc_putForm;
	prot["asc_getForm"] = prot.asc_getForm;
	prot["asc_putStyle"] = prot.asc_putStyle;
	prot["asc_getStyle"] = prot.asc_getStyle;

	window['Asc']['CImagePositionH'] = window["Asc"].CImagePositionH = CImagePositionH;
	prot = CImagePositionH.prototype;
	prot['get_RelativeFrom'] = prot.get_RelativeFrom;
	prot['put_RelativeFrom'] = prot.put_RelativeFrom;
	prot['get_UseAlign'] = prot.get_UseAlign;
	prot['put_UseAlign'] = prot.put_UseAlign;
	prot['get_Align'] = prot.get_Align;
	prot['put_Align'] = prot.put_Align;
	prot['get_Value'] = prot.get_Value;
	prot['put_Value'] = prot.put_Value;
	prot['get_Percent'] = prot.get_Percent;
	prot['put_Percent'] = prot.put_Percent;

	window['Asc']['CImagePositionV'] = window["Asc"].CImagePositionV = CImagePositionV;
	prot = CImagePositionV.prototype;
	prot['get_RelativeFrom'] = prot.get_RelativeFrom;
	prot['put_RelativeFrom'] = prot.put_RelativeFrom;
	prot['get_UseAlign'] = prot.get_UseAlign;
	prot['put_UseAlign'] = prot.put_UseAlign;
	prot['get_Align'] = prot.get_Align;
	prot['put_Align'] = prot.put_Align;
	prot['get_Value'] = prot.get_Value;
	prot['put_Value'] = prot.put_Value;
	prot['get_Percent'] = prot.get_Percent;
	prot['put_Percent'] = prot.put_Percent;

	window['Asc']['CPosition'] = window["Asc"].CPosition = CPosition;
	prot = CPosition.prototype;
	prot['get_X'] = prot.get_X;
	prot['put_X'] = prot.put_X;
	prot['get_Y'] = prot.get_Y;
	prot['put_Y'] = prot.put_Y;

	window["Asc"]["asc_CImgProperty"] = window["Asc"].asc_CImgProperty = asc_CImgProperty;
	prot = asc_CImgProperty.prototype;
	prot["get_ChangeLevel"] = prot["asc_getChangeLevel"] = prot.asc_getChangeLevel;
	prot["put_ChangeLevel"] = prot["asc_putChangeLevel"] = prot.asc_putChangeLevel;
	prot["get_CanBeFlow"] = prot["asc_getCanBeFlow"] = prot.asc_getCanBeFlow;
	prot["get_Width"] = prot["asc_getWidth"] = prot.asc_getWidth;
	prot["put_Width"] = prot["asc_putWidth"] = prot.asc_putWidth;
	prot["get_Height"] = prot["asc_getHeight"] = prot.asc_getHeight;
	prot["put_Height"] = prot["asc_putHeight"] = prot.asc_putHeight;
	prot["get_WrappingStyle"] = prot["asc_getWrappingStyle"] = prot.asc_getWrappingStyle;
	prot["put_WrappingStyle"] = prot["asc_putWrappingStyle"] = prot.asc_putWrappingStyle;
	prot["get_Paddings"] = prot["asc_getPaddings"] = prot.asc_getPaddings;
	prot["put_Paddings"] = prot["asc_putPaddings"] = prot.asc_putPaddings;
	prot["get_AllowOverlap"] = prot["asc_getAllowOverlap"] = prot.asc_getAllowOverlap;
	prot["put_AllowOverlap"] = prot["asc_putAllowOverlap"] = prot.asc_putAllowOverlap;
	prot["get_Position"] = prot["asc_getPosition"] = prot.asc_getPosition;
	prot["put_Position"] = prot["asc_putPosition"] = prot.asc_putPosition;
	prot["get_PositionH"] = prot["asc_getPositionH"] = prot.asc_getPositionH;
	prot["put_PositionH"] = prot["asc_putPositionH"] = prot.asc_putPositionH;
	prot["get_PositionV"] = prot["asc_getPositionV"] = prot.asc_getPositionV;
	prot["put_PositionV"] = prot["asc_putPositionV"] = prot.asc_putPositionV;
	prot["get_SizeRelH"] = prot["asc_getSizeRelH"] = prot.asc_getSizeRelH;
	prot["put_SizeRelH"] = prot["asc_putSizeRelH"] = prot.asc_putSizeRelH;
	prot["get_SizeRelV"] = prot["asc_getSizeRelV"] = prot.asc_getSizeRelV;
	prot["put_SizeRelV"] = prot["asc_putSizeRelV"] = prot.asc_putSizeRelV;
	prot["get_Value_X"] = prot["asc_getValue_X"] = prot.asc_getValue_X;
	prot["get_Value_Y"] = prot["asc_getValue_Y"] = prot.asc_getValue_Y;
	prot["get_ImageUrl"] = prot["asc_getImageUrl"] = prot.asc_getImageUrl;
	prot["put_ImageUrl"] = prot["asc_putImageUrl"] = prot.asc_putImageUrl;
	prot["get_Group"] = prot["asc_getGroup"] = prot.asc_getGroup;
	prot["put_Group"] = prot["asc_putGroup"] = prot.asc_putGroup;
	prot["get_FromGroup"] = prot["asc_getFromGroup"] = prot.asc_getFromGroup;
	prot["put_FromGroup"] = prot["asc_putFromGroup"] = prot.asc_putFromGroup;
	prot["get_isChartProps"] = prot["asc_getisChartProps"] = prot.asc_getisChartProps;
	prot["put_isChartPross"] = prot["asc_putisChartPross"] = prot.asc_putisChartPross;
	prot["get_SeveralCharts"] = prot["asc_getSeveralCharts"] = prot.asc_getSeveralCharts;
	prot["put_SeveralCharts"] = prot["asc_putSeveralCharts"] = prot.asc_putSeveralCharts;
	prot["get_SeveralChartTypes"] = prot["asc_getSeveralChartTypes"] = prot.asc_getSeveralChartTypes;
	prot["put_SeveralChartTypes"] = prot["asc_putSeveralChartTypes"] = prot.asc_putSeveralChartTypes;
	prot["get_SeveralChartStyles"] = prot["asc_getSeveralChartStyles"] = prot.asc_getSeveralChartStyles;
	prot["put_SeveralChartStyles"] = prot["asc_putSeveralChartStyles"] = prot.asc_putSeveralChartStyles;
	prot["get_VerticalTextAlign"] = prot["asc_getVerticalTextAlign"] = prot.asc_getVerticalTextAlign;
	prot["put_VerticalTextAlign"] = prot["asc_putVerticalTextAlign"] = prot.asc_putVerticalTextAlign;
	prot["get_Vert"] = prot["asc_getVert"] = prot.asc_getVert;
	prot["put_Vert"] = prot["asc_putVert"] = prot.asc_putVert;
	prot["get_Locked"] = prot["asc_getLocked"] = prot.asc_getLocked;
	prot["getLockAspect"] = prot["asc_getLockAspect"] = prot.asc_getLockAspect;
	prot["putLockAspect"] = prot["asc_putLockAspect"] = prot.asc_putLockAspect;
	prot["get_ChartProperties"] = prot["asc_getChartProperties"] = prot.asc_getChartProperties;
	prot["put_ChartProperties"] = prot["asc_putChartProperties"] = prot.asc_putChartProperties;
	prot["get_ShapeProperties"] = prot["asc_getShapeProperties"] = prot.asc_getShapeProperties;
	prot["put_ShapeProperties"] = prot["asc_putShapeProperties"] = prot.asc_putShapeProperties;
	prot["put_SlicerProperties"] = prot["asc_putSlicerProperties"] = prot.asc_putSlicerProperties;
	prot["get_SlicerProperties"] = prot["asc_getSlicerProperties"] = prot.asc_getSlicerProperties;
	prot["get_OriginSize"] = prot["asc_getOriginSize"] = prot.asc_getOriginSize;
	prot["get_CropOriginSize"] = prot["asc_getCropOriginSize"] = prot.asc_getCropOriginSize;
	prot["get_PluginGuid"] = prot["asc_getPluginGuid"] = prot.asc_getPluginGuid;
	prot["put_PluginGuid"] = prot["asc_putPluginGuid"] = prot.asc_putPluginGuid;
	prot["get_PluginData"] = prot["asc_getPluginData"] = prot.asc_getPluginData;
	prot["put_PluginData"] = prot["asc_putPluginData"] = prot.asc_putPluginData;
	prot["get_Rot"] = prot["asc_getRot"] = prot.asc_getRot;
	prot["put_Rot"] = prot["asc_putRot"] = prot.asc_putRot;
	prot["get_RotAdd"] = prot["asc_getRotAdd"] = prot.asc_getRotAdd;
	prot["put_RotAdd"] = prot["asc_putRotAdd"] = prot.asc_putRotAdd;
	prot["get_FlipH"] = prot["asc_getFlipH"] = prot.asc_getFlipH;
	prot["put_FlipH"] = prot["asc_putFlipH"] = prot.asc_putFlipH;
	prot["get_FlipV"] = prot["asc_getFlipV"] = prot.asc_getFlipV;
	prot["put_FlipV"] = prot["asc_putFlipV"] = prot.asc_putFlipV;
	prot["get_FlipHInvert"] = prot["asc_getFlipHInvert"] = prot.asc_getFlipHInvert;
	prot["put_FlipHInvert"] = prot["asc_putFlipHInvert"] = prot.asc_putFlipHInvert;
	prot["get_FlipVInvert"] = prot["asc_getFlipVInvert"] = prot.asc_getFlipVInvert;
	prot["put_FlipVInvert"] = prot["asc_putFlipVInvert"] = prot.asc_putFlipVInvert;
	prot["put_ResetCrop"] = prot["asc_putResetCrop"] = prot.asc_putResetCrop;

	prot["get_Title"] = prot["asc_getTitle"] = prot.asc_getTitle;
	prot["put_Title"] = prot["asc_putTitle"] = prot.asc_putTitle;
	prot["get_Description"] = prot["asc_getDescription"] = prot.asc_getDescription;
	prot["put_Description"] = prot["asc_putDescription"] = prot.asc_putDescription;
	prot["get_Name"] = prot["asc_getName"] = prot.asc_getName;
	prot["put_Name"] = prot["asc_putName"] = prot.asc_putName;
	prot["get_ColumnNumber"] = prot["asc_getColumnNumber"] = prot.asc_getColumnNumber;
	prot["put_ColumnNumber"] = prot["asc_putColumnNumber"] = prot.asc_putColumnNumber;
	prot["get_ColumnSpace"] = prot["asc_getColumnSpace"] = prot.asc_getColumnSpace;
	prot["get_TextFitType"] = prot["asc_getTextFitType"] = prot.asc_getTextFitType;
	prot["get_VertOverflowType"] = prot["asc_getVertOverflowType"] = prot.asc_getVertOverflowType;
	prot["put_ColumnSpace"] = prot["asc_putColumnSpace"] = prot.asc_putColumnSpace;
	prot["put_TextFitType"] = prot["asc_putTextFitType"] = prot.asc_putTextFitType;
	prot["put_VertOverflowType"] = prot["asc_putVertOverflowType"] = prot.asc_putVertOverflowType;
	prot["asc_getSignatureId"] = prot["asc_getSignatureId"] = prot.asc_getSignatureId;

	prot["put_Shadow"] = prot.put_Shadow = prot["put_shadow"] = prot.put_shadow = prot["asc_putShadow"] = prot.asc_putShadow;
	prot["get_Shadow"] = prot.get_Shadow = prot["get_shadow"] = prot.get_shadow = prot["asc_getShadow"] = prot.asc_getShadow;

	prot["put_Anchor"] = prot.put_Anchor = prot["asc_putAnchor"] = prot.asc_putAnchor;
	prot["get_Anchor"] = prot.get_Anchor = prot["asc_getAnchor"] = prot.asc_getAnchor;
	prot["get_ProtectionLockText"] = prot["asc_getProtectionLockText"] = prot.asc_getProtectionLockText;
	prot["put_ProtectionLockText"] = prot["asc_putProtectionLockText"] = prot.asc_putProtectionLockText;
	prot["get_ProtectionLocked"] = prot["asc_getProtectionLocked"] = prot.asc_getProtectionLocked;
	prot["put_ProtectionLocked"] = prot["asc_putProtectionLocked"] = prot.asc_putProtectionLocked;
	prot["get_ProtectionPrint"] = prot["asc_getProtectionPrint"] = prot.asc_getProtectionPrint;
	prot["put_ProtectionPrint"] = prot["asc_putProtectionPrint"] = prot.asc_putProtectionPrint;
	prot["get_Transparent"] = prot["asc_getTransparent"] = prot.asc_getTransparent;
	prot["put_Transparent"] = prot["asc_putTransparent"] = prot.asc_putTransparent;
	prot["get_CropHeightCoefficient"] = prot["asc_getCropHeightCoefficient"] = prot.asc_getCropHeightCoefficient;
	prot["put_CropHeightCoefficient"] = prot["asc_putCropHeightCoefficient"] = prot.asc_putCropHeightCoefficient;
	prot["get_CropWidthCoefficient"] = prot["asc_getCropWidthCoefficient"] = prot.asc_getCropWidthCoefficient;
	prot["put_CropWidthCoefficient"] = prot["asc_putCropWidthCoefficient"] = prot.asc_putCropWidthCoefficient;
	prot["get_IsCrop"] = prot["asc_getIsCrop"] = prot.asc_getIsCrop;

	window["AscCommon"].asc_CSelectedObject = asc_CSelectedObject;
	prot = asc_CSelectedObject.prototype;
	prot["get_ObjectType"] = prot["asc_getObjectType"] = prot.asc_getObjectType;
	prot["get_ObjectValue"] = prot["asc_getObjectValue"] = prot.asc_getObjectValue;

	window["Asc"]["asc_CShapeFill"] = window["Asc"].asc_CShapeFill = asc_CShapeFill;
	prot = asc_CShapeFill.prototype;
	prot["get_type"] = prot["asc_getType"] = prot.asc_getType;
	prot["put_type"] = prot["asc_putType"] = prot.asc_putType;
	prot["get_fill"] = prot["asc_getFill"] = prot.asc_getFill;
	prot["put_fill"] = prot["asc_putFill"] = prot.asc_putFill;
	prot["get_transparent"] = prot["asc_getTransparent"] = prot.asc_getTransparent;
	prot["put_transparent"] = prot["asc_putTransparent"] = prot.asc_putTransparent;
	prot["asc_CheckForseSet"] = prot["asc_CheckForseSet"] = prot.asc_CheckForseSet;

	window["Asc"]["asc_CFillBlip"] = window["Asc"].asc_CFillBlip = asc_CFillBlip;
	prot = asc_CFillBlip.prototype;
	prot["get_type"] = prot["asc_getType"] = prot.asc_getType;
	prot["put_type"] = prot["asc_putType"] = prot.asc_putType;
	prot["get_url"] = prot["asc_getUrl"] = prot.asc_getUrl;
	prot["put_url"] = prot["asc_putUrl"] = prot.asc_putUrl;
	prot["get_texture_id"] = prot["asc_getTextureId"] = prot.asc_getTextureId;
	prot["put_texture_id"] = prot["asc_putTextureId"] = prot.asc_putTextureId;

	window["Asc"]["asc_CFillHatch"] = window["Asc"].asc_CFillHatch = asc_CFillHatch;
	prot = asc_CFillHatch.prototype;
	prot["get_pattern_type"] = prot["asc_getPatternType"] = prot.asc_getPatternType;
	prot["put_pattern_type"] = prot["asc_putPatternType"] = prot.asc_putPatternType;
	prot["get_color_fg"] = prot["asc_getColorFg"] = prot.asc_getColorFg;
	prot["put_color_fg"] = prot["asc_putColorFg"] = prot.asc_putColorFg;
	prot["get_color_bg"] = prot["asc_getColorBg"] = prot.asc_getColorBg;
	prot["put_color_bg"] = prot["asc_putColorBg"] = prot.asc_putColorBg;

	window["Asc"]["asc_CFillGrad"] = window["Asc"].asc_CFillGrad = asc_CFillGrad;
	prot = asc_CFillGrad.prototype;
	prot["get_colors"] = prot["asc_getColors"] = prot.asc_getColors;
	prot["put_colors"] = prot["asc_putColors"] = prot.asc_putColors;
	prot["get_positions"] = prot["asc_getPositions"] = prot.asc_getPositions;
	prot["put_positions"] = prot["asc_putPositions"] = prot.asc_putPositions;
	prot["get_grad_type"] = prot["asc_getGradType"] = prot.asc_getGradType;
	prot["put_grad_type"] = prot["asc_putGradType"] = prot.asc_putGradType;
	prot["get_linear_angle"] = prot["asc_getLinearAngle"] = prot.asc_getLinearAngle;
	prot["put_linear_angle"] = prot["asc_putLinearAngle"] = prot.asc_putLinearAngle;
	prot["get_linear_scale"] = prot["asc_getLinearScale"] = prot.asc_getLinearScale;
	prot["put_linear_scale"] = prot["asc_putLinearScale"] = prot.asc_putLinearScale;
	prot["get_path_type"] = prot["asc_getPathType"] = prot.asc_getPathType;
	prot["put_path_type"] = prot["asc_putPathType"] = prot.asc_putPathType;

	window["Asc"]["asc_CFillSolid"] = window["Asc"].asc_CFillSolid = asc_CFillSolid;
	prot = asc_CFillSolid.prototype;
	prot["get_color"] = prot["asc_getColor"] = prot.asc_getColor;
	prot["put_color"] = prot["asc_putColor"] = prot.asc_putColor;

	window["Asc"]["asc_CStroke"] = window["Asc"].asc_CStroke = asc_CStroke;
	prot = asc_CStroke.prototype;
	prot["get_type"] = prot["asc_getType"] = prot.asc_getType;
	prot["put_type"] = prot["asc_putType"] = prot.asc_putType;
	prot["get_width"] = prot["asc_getWidth"] = prot.asc_getWidth;
	prot["put_width"] = prot["asc_putWidth"] = prot.asc_putWidth;
	prot["get_color"] = prot["asc_getColor"] = prot.asc_getColor;
	prot["put_color"] = prot["asc_putColor"] = prot.asc_putColor;
	prot["get_linejoin"] = prot["asc_getLinejoin"] = prot.asc_getLinejoin;
	prot["put_linejoin"] = prot["asc_putLinejoin"] = prot.asc_putLinejoin;
	prot["get_linecap"] = prot["asc_getLinecap"] = prot.asc_getLinecap;
	prot["put_linecap"] = prot["asc_putLinecap"] = prot.asc_putLinecap;
	prot["get_linebeginstyle"] = prot["asc_getLinebeginstyle"] = prot.asc_getLinebeginstyle;
	prot["put_linebeginstyle"] = prot["asc_putLinebeginstyle"] = prot.asc_putLinebeginstyle;
	prot["get_linebeginsize"] = prot["asc_getLinebeginsize"] = prot.asc_getLinebeginsize;
	prot["put_linebeginsize"] = prot["asc_putLinebeginsize"] = prot.asc_putLinebeginsize;
	prot["get_lineendstyle"] = prot["asc_getLineendstyle"] = prot.asc_getLineendstyle;
	prot["put_lineendstyle"] = prot["asc_putLineendstyle"] = prot.asc_putLineendstyle;
	prot["get_lineendsize"] = prot["asc_getLineendsize"] = prot.asc_getLineendsize;
	prot["put_lineendsize"] = prot["asc_putLineendsize"] = prot.asc_putLineendsize;
	prot["get_canChangeArrows"] = prot["asc_getCanChangeArrows"] = prot.asc_getCanChangeArrows;
	prot["put_prstDash"] = prot["asc_putPrstDash"] = prot.asc_putPrstDash;
	prot["get_prstDash"] = prot["asc_getPrstDash"] = prot.asc_getPrstDash;
	prot["get_transparent"] = prot["asc_getTransparent"] = prot.asc_getTransparent;
	prot["put_transparent"] = prot["asc_putTransparent"] = prot.asc_putTransparent;

	window["AscCommon"].CAscColorScheme = CAscColorScheme;
	prot = CAscColorScheme.prototype;
	prot["get_colors"] = prot.get_colors;
	prot["get_name"] = prot.get_name;

	window["AscCommon"].CMouseMoveData = CMouseMoveData;
	prot = CMouseMoveData.prototype;
	prot["get_Type"] = prot.get_Type;
	prot["get_X"] = prot.get_X;
	prot["get_Y"] = prot.get_Y;
	prot["get_Hyperlink"] = prot.get_Hyperlink;
	prot["get_UserId"] = prot.get_UserId;
	prot["get_HaveChanges"] = prot.get_HaveChanges;
	prot["get_LockedObjectType"] = prot.get_LockedObjectType;
	prot["get_FootnoteText"] = prot.get_FootnoteText;
	prot["get_FootnoteNumber"] = prot.get_FootnoteNumber;
	prot["get_FormHelpText"] = prot.get_FormHelpText;
	prot["get_ReviewChange"] = prot.get_ReviewChange;
	prot["get_EyedropperColor"] = prot.get_EyedropperColor;
	prot["get_PlaceholderType"] = prot.get_PlaceholderType;
	prot["get_EffectText"] = prot.get_EffectText;
	prot["get_EffectDescription"] = prot.get_EffectDescription;

	window["Asc"]["asc_CUserInfo"] = window["Asc"].asc_CUserInfo = asc_CUserInfo;
	prot = asc_CUserInfo.prototype;
	prot["asc_putId"] = prot["put_Id"] = prot.asc_putId;
	prot["asc_getId"] = prot["get_Id"] = prot.asc_getId;
	prot["asc_putFullName"] = prot["put_FullName"] = prot.asc_putFullName;
	prot["asc_getFullName"] = prot["get_FullName"] = prot.asc_getFullName;
	prot["asc_putFirstName"] = prot["put_FirstName"] = prot.asc_putFirstName;
	prot["asc_getFirstName"] = prot["get_FirstName"] = prot.asc_getFirstName;
	prot["asc_putLastName"] = prot["put_LastName"] = prot.asc_putLastName;
	prot["asc_getLastName"] = prot["get_LastName"] = prot.asc_getLastName;
	prot["asc_putIsAnonymousUser"] = prot["put_IsAnonymousUser"] = prot.asc_putIsAnonymousUser;
	prot["asc_getIsAnonymousUser"] = prot["get_IsAnonymousUser"] = prot.asc_getIsAnonymousUser;

	window["Asc"]["asc_CDocInfo"] = window["Asc"].asc_CDocInfo = asc_CDocInfo;
	prot = asc_CDocInfo.prototype;
	prot["get_Id"] = prot["asc_getId"] = prot.asc_getId;
	prot["put_Id"] = prot["asc_putId"] = prot.asc_putId;
	prot["get_Url"] = prot["asc_getUrl"] = prot.asc_getUrl;
	prot["put_Url"] = prot["asc_putUrl"] = prot.asc_putUrl;
	prot["get_DirectUrl"] = prot["asc_getDirectUrl"] = prot.asc_getDirectUrl;
	prot["put_DirectUrl"] = prot["asc_putDirectUrl"] = prot.asc_putDirectUrl;
	prot["get_Title"] = prot["asc_getTitle"] = prot.asc_getTitle;
	prot["put_Title"] = prot["asc_putTitle"] = prot.asc_putTitle;
	prot["get_Format"] = prot["asc_getFormat"] = prot.asc_getFormat;
	prot["put_Format"] = prot["asc_putFormat"] = prot.asc_putFormat;
	prot["get_VKey"] = prot["asc_getVKey"] = prot.asc_getVKey;
	prot["put_VKey"] = prot["asc_putVKey"] = prot.asc_putVKey;
	prot["get_UserId"] = prot["asc_getUserId"] = prot.asc_getUserId;
	prot["get_UserName"] = prot["asc_getUserName"] = prot.asc_getUserName;
	prot["get_Options"] = prot["asc_getOptions"] = prot.asc_getOptions;
	prot["put_Options"] = prot["asc_putOptions"] = prot.asc_putOptions;
	prot["get_CallbackUrl"] = prot["asc_getCallbackUrl"] = prot.asc_getCallbackUrl;
	prot["put_CallbackUrl"] = prot["asc_putCallbackUrl"] = prot.asc_putCallbackUrl;
	prot["get_TemplateReplacement"] = prot["asc_getTemplateReplacement"] = prot.asc_getTemplateReplacement;
	prot["put_TemplateReplacement"] = prot["asc_putTemplateReplacement"] = prot.asc_putTemplateReplacement;
	prot["get_UserInfo"] = prot["asc_getUserInfo"] = prot.asc_getUserInfo;
	prot["put_UserInfo"] = prot["asc_putUserInfo"] = prot.asc_putUserInfo;
	prot["get_Token"] = prot["asc_getToken"] = prot.asc_getToken;
	prot["put_Token"] = prot["asc_putToken"] = prot.asc_putToken;
	prot["get_Mode"] = prot["asc_getMode"] = prot.asc_getMode;
	prot["put_Mode"] = prot["asc_putMode"] = prot.asc_putMode;
	prot["get_Permissions"] = prot["asc_getPermissions"] = prot.asc_getPermissions;
	prot["put_Permissions"] = prot["asc_putPermissions"] = prot.asc_putPermissions;
	prot["get_Lang"] = prot["asc_getLang"] = prot.asc_getLang;
	prot["put_Lang"] = prot["asc_putLang"] = prot.asc_putLang;
	prot["get_Encrypted"] = prot["asc_getEncrypted"] = prot.asc_getEncrypted;
	prot["put_Encrypted"] = prot["asc_putEncrypted"] = prot.asc_putEncrypted;
	prot["get_EncryptedInfo"] = prot["asc_getEncryptedInfo"] = prot.asc_getEncryptedInfo;
	prot["put_EncryptedInfo"] = prot["asc_putEncryptedInfo"] = prot.asc_putEncryptedInfo;
	prot["get_IsEnabledPlugins"] = prot["asc_getIsEnabledPlugins"] = prot.asc_getIsEnabledPlugins;
	prot["put_IsEnabledPlugins"] = prot["asc_putIsEnabledPlugins"] = prot.asc_putIsEnabledPlugins;
	prot["get_IsEnabledMacroses"] = prot["asc_getIsEnabledMacroses"] = prot.asc_getIsEnabledMacroses;
	prot["put_IsEnabledMacroses"] = prot["asc_putIsEnabledMacroses"] = prot.asc_putIsEnabledMacroses;
	prot["get_CoEditingMode"] = prot["asc_getCoEditingMode"] = prot.asc_getCoEditingMode;
	prot["put_CoEditingMode"] = prot["asc_putCoEditingMode"] = prot.asc_putCoEditingMode;
	prot["put_ReferenceData"] = prot["asc_putReferenceData"] = prot.asc_putReferenceData;
	prot["put_IsWebOpening"] = prot["asc_putIsWebOpening"] = prot.asc_putIsWebOpening;
	prot["get_IsWebOpening"] = prot["asc_getIsWebOpening"] = prot.asc_getIsWebOpening;
	prot["put_SupportsOnSaveDocument"] = prot["asc_putSupportsOnSaveDocument"] = prot.asc_putSupportsOnSaveDocument;
	prot["get_SupportsOnSaveDocument"] = prot["asc_getSupportsOnSaveDocument"] = prot.asc_getSupportsOnSaveDocument;
	prot["put_Wopi"] = prot["asc_putWopi"] = prot.asc_putWopi;
	prot["get_Wopi"] = prot["asc_getWopi"] = prot.asc_getWopi;
	prot["put_Shardkey"] = prot["asc_putShardkey"] = prot.asc_putShardkey;
	prot["get_Shardkey"] = prot["asc_getShardkey"] = prot.asc_getShardkey;
	prot["put_ShowVerticalScroll"] = prot["asc_putShowVerticalScroll"] = prot.asc_putShowVerticalScroll;
	prot["get_ShowVerticalScroll"] = prot["get_getShowVerticalScroll"] = prot.get_getShowVerticalScroll;
	prot["put_ShowHorizontalScroll"] = prot["asc_putShowHorizontalScroll"] = prot.asc_putShowHorizontalScroll;
	prot["get_ShowHorizontalScroll"] = prot["get_getShowHorizontalScroll"] = prot.get_getShowHorizontalScroll;

	window["AscCommon"].COpenProgress = COpenProgress;
	prot = COpenProgress.prototype;
	prot["asc_getType"] = prot.asc_getType;
	prot["asc_getFontsCount"] = prot.asc_getFontsCount;
	prot["asc_getCurrentFont"] = prot.asc_getCurrentFont;
	prot["asc_getImagesCount"] = prot.asc_getImagesCount;
	prot["asc_getCurrentImage"] = prot.asc_getCurrentImage;

	window["AscCommon"].CErrorData = CErrorData;
	prot = CErrorData.prototype;
	prot["put_Value"] = prot.put_Value;
	prot["get_Value"] = prot.get_Value;

	window["AscCommon"].CAscMathType = CAscMathType;
	prot = CAscMathType.prototype;
	prot["get_Id"] = prot.get_Id;
	prot["get_X"] = prot.get_X;
	prot["get_Y"] = prot.get_Y;

	window["AscCommon"].CAscMathCategory = CAscMathCategory;
	prot = CAscMathCategory.prototype;
	prot["get_Id"] = prot.get_Id;
	prot["get_Data"] = prot.get_Data;
	prot["get_W"] = prot.get_W;
	prot["get_H"] = prot.get_H;

	window["AscCommon"].CStyleImage = CStyleImage;
	prot = CStyleImage.prototype;
	prot["asc_getId"] = prot["asc_getName"] = prot["get_Name"] = prot.asc_getName;
	prot["asc_getDisplayName"] = prot.asc_getDisplayName;
	prot["asc_getType"] = prot["get_Type"] = prot.asc_getType;
	prot["asc_getImage"] = prot.asc_getImage;

	window["AscCommon"].asc_CSpellCheckProperty = asc_CSpellCheckProperty;
	prot = asc_CSpellCheckProperty.prototype;
	prot["get_Word"] = prot.get_Word;
	prot["get_Checked"] = prot.get_Checked;
	prot["get_Variants"] = prot.get_Variants;

	window["AscCommon"].CWatermarkOnDraw = CWatermarkOnDraw;
	window["AscCommon"].isFileBuild = isFileBuild;
	window["AscCommon"].checkCanvasInDiv = checkCanvasInDiv;
	window["AscCommon"].isValidJs = isValidJs;
    window["AscCommon"].parseJSDoc = parseJSDoc;

	window["Asc"]["PluginType"] = window["Asc"].PluginType = PluginType;
	window["Asc"]["CPluginVariation"] = window["Asc"].CPluginVariation = CPluginVariation;
	window["Asc"]["CPlugin"] = window["Asc"].CPlugin = CPlugin;

	window["AscCommon"].CDocInfoProp = CDocInfoProp;
	CDocInfoProp.prototype['get_PageCount'] = CDocInfoProp.prototype.get_PageCount;
	CDocInfoProp.prototype['put_PageCount'] = CDocInfoProp.prototype.put_PageCount;
	CDocInfoProp.prototype['get_WordsCount'] = CDocInfoProp.prototype.get_WordsCount;
	CDocInfoProp.prototype['put_WordsCount'] = CDocInfoProp.prototype.put_WordsCount;
	CDocInfoProp.prototype['get_ParagraphCount'] = CDocInfoProp.prototype.get_ParagraphCount;
	CDocInfoProp.prototype['put_ParagraphCount'] = CDocInfoProp.prototype.put_ParagraphCount;
	CDocInfoProp.prototype['get_SymbolsCount'] = CDocInfoProp.prototype.get_SymbolsCount;
	CDocInfoProp.prototype['put_SymbolsCount'] = CDocInfoProp.prototype.put_SymbolsCount;
	CDocInfoProp.prototype['get_SymbolsWSCount'] = CDocInfoProp.prototype.get_SymbolsWSCount;
	CDocInfoProp.prototype['put_SymbolsWSCount'] = CDocInfoProp.prototype.put_SymbolsWSCount;
	
	window["Asc"]["RangePermProp"] = window["Asc"].RangePermProp = RangePermProp;
	prot = RangePermProp.prototype;
	prot["get_canEditText"] = prot.get_canEditText;
	prot["get_canEditPara"] = prot.get_canEditPara;
	prot["get_canInsObject"] = prot.get_canInsObject;

	window["Asc"]["CButtonData"] = window["Asc"].CButtonData = CButtonData;
	prot = CButtonData.prototype;

	prot["get_Obj"] = prot.get_Obj;
	prot["get_Type"] = prot.get_Type;
	prot["get_Button"] = prot.get_Button;
	prot["get_IsForm"] = prot.get_IsForm;
	prot["get_Properties"] = prot.get_Properties;

	window["AscCommon"]["pix2mm"] = window["AscCommon"].pix2mm = function(pix)
	{
		return pix * AscCommon.g_dKoef_pix_to_mm;
	};
	window["AscCommon"]["mm2pix"] = window["AscCommon"].mm2pix = function(mm)
	{
		return mm * AscCommon.g_dKoef_mm_to_pix;
	};

	window["Asc"].checkReturnCommand = function(obj, recursionDepth)
	{
		let depth = (recursionDepth === undefined) ? 0 : recursionDepth;
		if (depth > 10)
			return false;

		switch (typeof obj)
		{
			case "undefined":
			case "boolean":
			case "number":
			case "string":
			case "symbol":
			case "bigint":
				return true;
			case "object":
			{
				if (!obj)
					return true;

				if (Array.isArray(obj))
				{
					for (let i = 0, len = obj.length; i < len; i++)
					{
						if (!Asc.checkReturnCommand(obj[i], depth + 1))
							return false;
					}

					return true;
				}

				if (Object.getPrototypeOf)
				{
					let prot = Object.getPrototypeOf(obj);
					if (prot && prot.__proto__ && prot.__proto__.constructor && prot.__proto__.constructor.name)
					{
						if (prot.__proto__.constructor.name === "TypedArray")
							return true;
					}
				}

				for (let prop in obj)
				{
					let isNaturalProp = true;
					if (obj.hasOwnProperty)
						isNaturalProp = obj.hasOwnProperty(prop);
					if (isNaturalProp)
					{
						if (!Asc.checkReturnCommand(obj[prop], depth + 1))
							return false;
					}
				}

				return true;
			}
			default:
				break;
		}

		return false;
	};
	
})(window);
