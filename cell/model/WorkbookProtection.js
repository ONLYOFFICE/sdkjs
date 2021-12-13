/*
 * (c) Copyright Ascensio System SIA 2010-2021
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(window, undefined){


	/*this.sheet = true;+
	 this.objects = false;
	 this.scenarios = false;
	 this.formatCells = true;
	 this.formatColumns = true;
	 this.formatRows = true;
	 this.insertColumns = true; +
	 this.insertRows = true; +
	 this.insertHyperlinks = true;+
	 this.deleteColumns = true; +
	 this.deleteRows = true; +
	 this.selectLockedCells = false;
	 this.sort = true;+
	 this.autoFilter = true; +
	 this.pivotTables = true;
	 this.selectUnlockedCells = false;*/

	var c_oSerProtectedAlgorithmNameTypes = {
		MD2: 1,
		MD4: 2,
		MD5: 3,
		RIPEMD_128: 4,
		RIPEMD_160: 5,
		SHA_1: 6,
		SHA_256: 7,
		SHA_384: 8,
		SHA_512: 9,
		WHIRLPOOL: 10
	};

	function fromModelAlgoritmName(alg) {
		switch (alg) {
			case c_oSerProtectedAlgorithmNameTypes.MD2 :
				alg = AscCommon.HashAlgs.MD2;
				break;
			case c_oSerProtectedAlgorithmNameTypes.MD4 :
				alg = AscCommon.HashAlgs.MD4;
				break;
			case c_oSerProtectedAlgorithmNameTypes.MD5 :
				alg = AscCommon.HashAlgs.MD5;
				break;
			case c_oSerProtectedAlgorithmNameTypes.RIPEMD_160 :
				alg = AscCommon.HashAlgs.RMD160;
				break;
			case c_oSerProtectedAlgorithmNameTypes.SHA_1 :
				alg = AscCommon.HashAlgs.SHA1;
				break;
			case c_oSerProtectedAlgorithmNameTypes.SHA_256 :
				alg = AscCommon.HashAlgs.SHA256;
				break;
			case c_oSerProtectedAlgorithmNameTypes.SHA_384 :
				alg = AscCommon.HashAlgs.SHA384;
				break;
			case c_oSerProtectedAlgorithmNameTypes.SHA_512 :
				alg = AscCommon.HashAlgs.SHA512;
				break;
			case c_oSerProtectedAlgorithmNameTypes.WHIRLPOOL :
				alg = AscCommon.HashAlgs.WHIRLPOOL;
				break;
			default:
				alg = AscCommon.HashAlgs.SHA256;
		}
		return alg;
	}

	function generateHashParams() {
		return {spinCount: 100000, saltValue: AscCommon.randomBytes(16).base64(), algorithmName: c_oSerProtectedAlgorithmNameTypes.SHA_512};
	}

	function getPasswordHash(password, getString) {
		var nResult = null;
		if (password.length) {
			nResult = 0;
			for (var i = password.length - 1; i >= 0; i--) {
				nResult = ((nResult >> 14) & 0x01) | ((nResult << 1) & 0x7FFF);
				nResult ^= password.charCodeAt(i);
			}

			nResult = ((nResult >> 14) & 0x01) | ((nResult << 1) & 0x7FFF);
			nResult ^= ((0x8000) | ('N'.charCodeAt(0) << 8) | ('K'.charCodeAt(0)));
			nResult ^= password.length;
		}

		return getString && nResult ? nResult.toString(16) : nResult;
	}

	function CSheetProtection(ws) {
		this.algorithmName = null;
		this.hashValue = null;
		this.saltValue = null;
		this.spinCount = null;
		this.password = null;

		this.sheet = false;
		this.objects = false;
		this.scenarios = false;
		this.formatCells = true;
		this.formatColumns = true;
		this.formatRows = true;
		this.insertColumns = true;
		this.insertRows = true;
		this.insertHyperlinks = true;
		this.deleteColumns = true;
		this.deleteRows = true;
		this.selectLockedCells = false;
		this.sort = true;
		this.autoFilter = true;
		this.pivotTables = true;
		this.selectUnlockedCells = false;

		this._ws = ws;
		this.temporaryPassword = null;

		return this;
	}

	CSheetProtection.prototype.clone = function(ws) {
		var res = new CSheetProtection(ws);

		res.algorithmName = this.algorithmName;
		res.hashValue = this.hashValue;
		res.saltValue = this.saltValue;
		res.spinCount = this.spinCount;
		res.password = this.password;

		res.sheet = this.sheet;
		res.objects = this.objects;
		res.scenarios = this.scenarios;
		res.formatCells = this.formatCells;
		res.formatColumns = this.formatColumns;

		res.formatRows = this.formatRows;
		res.insertColumns = this.insertColumns;
		res.insertRows = this.insertRows;
		res.insertHyperlinks = this.insertHyperlinks;
		res.deleteColumns = this.deleteColumns;
		res.deleteRows = this.deleteRows;
		res.selectLockedCells = this.selectLockedCells;
		res.sort = this.sort;
		res.autoFilter = this.autoFilter;
		res.pivotTables = this.pivotTables;
		res.selectUnlockedCells = this.selectUnlockedCells;

		return res;
	};

	CSheetProtection.prototype.default = function() {
		this.algorithmName = null;
		this.hashValue = null;
		this.saltValue = null;
		this.spinCount = null;
		this.password = null;

		this.sheet = false;
		this.objects = false;
		this.scenarios = false;
		this.formatCells = true;
		this.formatColumns = true;
		this.formatRows = true;
		this.insertColumns = true;
		this.insertRows = true;
		this.insertHyperlinks = true;
		this.deleteColumns = true;
		this.deleteRows = true;
		this.selectLockedCells = false;
		this.sort = true;
		this.autoFilter = true;
		this.pivotTables = true;
		this.selectUnlockedCells = false;
	};

	CSheetProtection.prototype.setDefaultInterface = function() {
		this.algorithmName = null;
		this.hashValue = null;
		this.saltValue = null;
		this.spinCount = null;
		this.password = null;

		this.sheet = false;
		this.objects = true;
		this.scenarios = true;
		this.formatCells = true;
		this.formatColumns = true;
		this.formatRows = true;
		this.insertColumns = true;
		this.insertRows = true;
		this.insertHyperlinks = true;
		this.deleteColumns = true;
		this.deleteRows = true;
		this.selectLockedCells = false;
		this.sort = true;
		this.autoFilter = true;
		this.pivotTables = true;
		this.selectUnlockedCells = false;
	};

	CSheetProtection.prototype.isDefault = function () {
		if (this.algorithmName === null && this.hashValue === null && this.saltValue === null && this.spinCount === null) {
			if (this.sheet === false && this.objects === true && this.scenarios === true && this.formatColumns === true && this.formatRows === true && this.insertColumns ===
				true && this.insertRows === true && this.deleteRows === true && this.selectLockedCells === false && this.sort === true && this.autoFilter === true &&
				this.pivotTables === true && this.selectUnlockedCells === false && this.password === null) {
				return true;
			}
		}
		return false;
	};

	CSheetProtection.prototype.set = function (val, addToHistory, ws) {
		this.algorithmName = this.checkProperty(this.algorithmName, val.algorithmName, AscCH.historyitem_Protected_SetAlgorithmName, ws, addToHistory);
		this.hashValue = this.checkProperty(this.hashValue, val.hashValue, AscCH.historyitem_Protected_SetHashValue, ws, addToHistory);
		this.saltValue = this.checkProperty(this.saltValue, val.saltValue, AscCH.historyitem_Protected_SetSaltValue, ws, addToHistory);
		this.spinCount = this.checkProperty(this.spinCount, val.spinCount, AscCH.historyitem_Protected_SetSpinCount, ws, addToHistory);
		this.password = this.checkProperty(this.password, val.password, AscCH.historyitem_Protected_SetPassword, ws, addToHistory);

		this.sheet = this.checkProperty(this.sheet, val.sheet, AscCH.historyitem_Protected_SetSheet, ws, addToHistory);
		this.objects = this.checkProperty(this.objects, val.objects, AscCH.historyitem_Protected_SetObjects, ws, addToHistory);
		this.scenarios = this.checkProperty(this.scenarios, val.scenarios, AscCH.historyitem_Protected_SetScenarios, ws, addToHistory);
		this.formatCells = this.checkProperty(this.formatCells, val.formatCells, AscCH.historyitem_Protected_SetFormatCells, ws, addToHistory);
		this.formatColumns = this.checkProperty(this.formatColumns, val.formatColumns, AscCH.historyitem_Protected_SetFormatColumns, ws, addToHistory);
		this.formatRows = this.checkProperty(this.formatRows, val.formatRows, AscCH.historyitem_Protected_SetFormatRows, ws, addToHistory);
		this.insertColumns = this.checkProperty(this.insertColumns, val.insertColumns, AscCH.historyitem_Protected_SetInsertColumns, ws, addToHistory);
		this.insertRows = this.checkProperty(this.insertRows, val.insertRows, AscCH.historyitem_Protected_SetInsertRows, ws, addToHistory);
		this.insertHyperlinks = this.checkProperty(this.insertHyperlinks, val.insertHyperlinks, AscCH.historyitem_Protected_SetInsertHyperlinks, ws, addToHistory);
		this.deleteColumns = this.checkProperty(this.deleteColumns, val.deleteColumns, AscCH.historyitem_Protected_SetDeleteColumns, ws, addToHistory);
		this.deleteRows = this.checkProperty(this.deleteRows, val.deleteRows, AscCH.historyitem_Protected_SetDeleteRows, ws, addToHistory);

		this.selectLockedCells = this.checkProperty(this.selectLockedCells, val.selectLockedCells, AscCH.historyitem_Protected_SetSelectLockedCells, ws, addToHistory);
		this.sort = this.checkProperty(this.sort, val.sort, AscCH.historyitem_Protected_SetSort, ws, addToHistory);
		this.autoFilter = this.checkProperty(this.autoFilter, val.autoFilter, AscCH.historyitem_Protected_SetAutoFilter, ws, addToHistory);
		this.pivotTables = this.checkProperty(this.pivotTables, val.pivotTables, AscCH.historyitem_Protected_SetPivotTables, ws, addToHistory);
		this.selectUnlockedCells = this.checkProperty(this.selectUnlockedCells, val.selectUnlockedCells, AscCH.historyitem_Protected_SetSelectUnlockedCells, ws, addToHistory);
	};

	CSheetProtection.prototype.checkProperty = function (propOld, propNew, type, ws, addToHistory) {
		if (propOld !== propNew) {
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoProtectedSheet, type, ws.getId(), null,
					new AscCommonExcel.UndoRedoData_ProtectedRange(null, propOld, propNew));
			}
			return propNew;
		}
		return propOld;
	};

	CSheetProtection.prototype.Write_ToBinary2 = function(w) {
		var _writeBool = function (val) {
			if (null != val) {
				w.WriteBool(true);
				w.WriteBool(val);
			} else {
				w.WriteBool(false);
			}
		};

		if (null != this.algorithmName) {
			w.WriteBool(true);
			w.WriteLong(this.algorithmName);
		} else {
			w.WriteBool(false);
		}

		/*<xsd:attribute name="hashValue" type="xsd:base64Binary" use="optional"/>
		 <xsd:attribute name="saltValue" type="xsd:base64Binary" use="optional"/>*/
		if (null != this.hashValue) {
			w.WriteBool(true);
			w.WriteString2(this.hashValue);
		} else {
			w.WriteBool(false);
		}

		if (null != this.saltValue) {
			w.WriteBool(true);
			w.WriteString2(this.saltValue);
		} else {
			w.WriteBool(false);
		}

		if (null != this.spinCount) {
			w.WriteBool(true);
			w.WriteLong(this.spinCount);
		} else {
			w.WriteBool(false);
		}

		if (null != this.password) {
			w.WriteBool(true);
			w.WriteString2(this.password);
		} else {
			w.WriteBool(false);
		}

		_writeBool(this.sheet);
		_writeBool(this.objects);
		_writeBool(this.scenarios);
		_writeBool(this.formatCells);
		_writeBool(this.formatColumns);

		_writeBool(this.formatRows);
		_writeBool(this.insertColumns);
		_writeBool(this.insertHyperlinks);
		_writeBool(this.deleteColumns);
		_writeBool(this.deleteRows);
		_writeBool(this.selectLockedCells);
		_writeBool(this.sort);
		_writeBool(this.autoFilter);
		_writeBool(this.selectUnlockedCells);
	};

	CSheetProtection.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			this.algorithmName = r.GetLong();
		}
		if (r.GetBool()) {
			this.hashValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.saltValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.spinCount = r.GetLong();
		}
		if (r.GetBool()) {
			this.password = r.GetString2();
		}

		if (r.GetBool()) {
			this.sheet = r.GetBool();
		}
		if (r.GetBool()) {
			this.objects = r.GetBool();
		}
		if (r.GetBool()) {
			this.scenarios = r.GetBool();
		}
		if (r.GetBool()) {
			this.formatCells = r.GetBool();
		}
		if (r.GetBool()) {
			this.formatColumns = r.GetBool();
		}
		if (r.GetBool()) {
			this.formatRows = r.GetBool();
		}

		if (r.GetBool()) {
			this.insertColumns = r.GetBool();
		}
		if (r.GetBool()) {
			this.insertHyperlinks = r.GetBool();
		}
		if (r.GetBool()) {
			this.deleteColumns = r.GetBool();
		}
		if (r.GetBool()) {
			this.deleteRows = r.GetBool();
		}

		if (r.GetBool()) {
			this.selectLockedCells = r.GetBool();
		}
		if (r.GetBool()) {
			this.sort = r.GetBool();
		}
		if (r.GetBool()) {
			this.autoFilter = r.GetBool();
		}
		if (r.GetBool()) {
			this.selectUnlockedCells = r.GetBool();
		}
	};

	CSheetProtection.prototype.getAlgorithmName = function () {
		return this.algorithmName;
	};
	CSheetProtection.prototype.getHashValue = function () {
		return this.hashValue;
	};
	CSheetProtection.prototype.getSaltValue = function () {
		return this.saltValue;
	};
	CSheetProtection.prototype.getSpinCount = function () {
		return this.spinCount;
	};

	CSheetProtection.prototype.getSheet = function () {
		return this.sheet;
	};
	CSheetProtection.prototype.getObjects = function () {
		return this.objects;
	};
	CSheetProtection.prototype.getScenarios = function () {
		return this.scenarios;
	};
	CSheetProtection.prototype.getFormatCells = function () {
		return this.formatCells;
	};
	CSheetProtection.prototype.getFormatColumns = function () {
		return this.formatColumns;
	};
	CSheetProtection.prototype.getFormatRows = function () {
		return this.formatRows;
	};
	CSheetProtection.prototype.getInsertColumns = function () {
		return this.insertColumns;
	};
	CSheetProtection.prototype.getInsertRows = function () {
		return this.insertRows;
	};
	CSheetProtection.prototype.getInsertHyperlinks = function () {
		return this.insertHyperlinks;
	};
	CSheetProtection.prototype.getDeleteColumns = function () {
		return this.deleteColumns;
	};
	CSheetProtection.prototype.getDeleteRows = function () {
		return this.deleteRows;
	};
	CSheetProtection.prototype.getSelectLockedCells = function () {
		return this.selectLockedCells;
	};
	CSheetProtection.prototype.getSort = function () {
		return this.sort;
	};
	CSheetProtection.prototype.getAutoFilter = function () {
		return this.autoFilter;
	};
	CSheetProtection.prototype.getPivotTables = function () {
		return this.pivotTables;
	};
	CSheetProtection.prototype.getSelectUnlockedCells = function () {
		return this.selectUnlockedCells;
	};

	CSheetProtection.prototype.setAlgorithmName = function (val) {
		this.algorithmName = val;
	};
	CSheetProtection.prototype.setHashValue = function (val) {
		this.hashValue = val;
	};
	CSheetProtection.prototype.setSaltValue = function (val) {
		this.saltValue = val;
	};
	CSheetProtection.prototype.setSpinCount = function (val) {
		this.spinCount = val;
	};

	CSheetProtection.prototype.setSheet = function (val) {
		this.sheet = val;
	};
	CSheetProtection.prototype.setPasswordXL = function (val) {
		this.password = val;
	};
	CSheetProtection.prototype.getPasswordXL = function () {
		return this.password;
	};
	CSheetProtection.prototype.isPasswordXL = function () {
		return this.password != null;
	};

	CSheetProtection.prototype.asc_setSheet = function (password, callback) {
		//просталяю временный пароль, аспинхронная проверка пароля в asc_setProtectedSheet
		this.setSheet(!this.sheet);
		if (this.sheet && password) {
			var hashParams = generateHashParams();
			this.saltValue = hashParams.saltValue;
			this.spinCount = hashParams.spinCount;
			this.algorithmName = hashParams.algorithmName;
		}
		this.temporaryPassword = password;
		if (callback) {
			callback(this);
		}
	};

	CSheetProtection.prototype.setObjects = function (val) {
		this.objects = val;
	};
	CSheetProtection.prototype.setScenarios = function (val) {
		this.scenarios = val;
	};
	CSheetProtection.prototype.setFormatCells = function (val) {
		this.formatCells = val;
	};
	CSheetProtection.prototype.setFormatColumns = function (val) {
		this.formatColumns = val;
	};
	CSheetProtection.prototype.setFormatRows = function (val) {
		this.formatRows = val;
	};
	CSheetProtection.prototype.setInsertColumns = function (val) {
		this.insertColumns = val;
	};
	CSheetProtection.prototype.setInsertRows = function (val) {
		this.insertRows = val;
	};
	CSheetProtection.prototype.setInsertHyperlinks = function (val) {
		this.insertHyperlinks = val;
	};
	CSheetProtection.prototype.setDeleteColumns = function (val) {
		this.deleteColumns = val;
	};
	CSheetProtection.prototype.setDeleteRows = function (val) {
		this.deleteRows = val;
	};
	CSheetProtection.prototype.setSelectLockedCells = function (val) {
		this.selectLockedCells = val;
	};
	CSheetProtection.prototype.setSort = function (val) {
		this.sort = val;
	};
	CSheetProtection.prototype.setAutoFilter = function (val) {
		this.autoFilter = val;
	};
	CSheetProtection.prototype.setPivotTables = function (val) {
		this.pivotTables = val;
	};
	CSheetProtection.prototype.setSelectUnlockedCells = function (val) {
		this.selectUnlockedCells = val;
	};
	CSheetProtection.prototype.asc_setPassword = function (val) {
		//генерируем хэш
		this.algorithmName = "test";
		this.hashValue = "test";
		this.saltValue = "test";
		this.spinCount = "test";
	};
	CSheetProtection.prototype.asc_isPassword = function () {
		return this.algorithmName != null || this.password != null;
	};

	function CWorkbookProtection(wb) {
		this.lockStructure = null;//false
		this.lockWindows = null;//false
		this.lockRevision = null;//false

		this.revisionsAlgorithmName = null;
		this.revisionsHashValue = null;
		this.revisionsSaltValue = null;
		this.revisionsSpinCount = null;
		this.workbookAlgorithmName = null;
		this.workbookHashValue = null;
		this.workbookSaltValue = null;
		this.workbookSpinCount = null;

		this.workbookPassword = null;

		this._wb = wb;
		this.temporaryPassword = null;

		return this;
	}

	CWorkbookProtection.prototype.clone = function(wb) {
		var res = new CWorkbookProtection(wb);

		res.lockStructure = this.lockStructure;
		res.lockWindows = this.lockWindows;
		res.lockRevision = this.lockRevision;

		res.revisionsAlgorithmName = this.revisionsAlgorithmName;
		res.revisionsHashValue = this.revisionsHashValue;
		res.revisionsSaltValue = this.revisionsSaltValue;
		res.revisionsSpinCount = this.revisionsSpinCount;

		res.workbookAlgorithmName = this.workbookAlgorithmName;
		res.workbookHashValue = this.workbookHashValue;
		res.workbookSaltValue = this.workbookSaltValue;
		res.workbookSpinCount = this.workbookSpinCount;

		res.workbookPassword = this.workbookPassword;

		return res;
	};

	CWorkbookProtection.prototype.set = function (val, addToHistory) {
		this.lockStructure = this.checkProperty(this.lockStructure, val.lockStructure, AscCH.historyitem_Protected_SetLockStructure, addToHistory);
		this.lockWindows = this.checkProperty(this.lockWindows, val.lockWindows, AscCH.historyitem_Protected_SetLockWindows, addToHistory);
		this.lockRevision = this.checkProperty(this.lockRevision, val.lockRevision, AscCH.historyitem_Protected_SetLockRevision, addToHistory);

		this.revisionsAlgorithmName = this.checkProperty(this.revisionsAlgorithmName, val.revisionsAlgorithmName, AscCH.historyitem_Protected_SetRevisionsAlgorithmName, addToHistory);
		this.revisionsHashValue = this.checkProperty(this.revisionsHashValue, val.revisionsHashValue, AscCH.historyitem_Protected_SetRevisionsHashValue, addToHistory);
		this.revisionsSaltValue = this.checkProperty(this.revisionsSaltValue, val.revisionsSaltValue, AscCH.historyitem_Protected_SetRevisionsSaltValue, addToHistory);
		this.revisionsSpinCount = this.checkProperty(this.revisionsSpinCount, val.revisionsSpinCount, AscCH.historyitem_Protected_SetRevisionsSpinCount, addToHistory);

		this.workbookAlgorithmName = this.checkProperty(this.workbookAlgorithmName, val.workbookAlgorithmName, AscCH.historyitem_Protected_SetWorkbookAlgorithmName, addToHistory);
		this.workbookHashValue = this.checkProperty(this.workbookHashValue, val.workbookHashValue, AscCH.historyitem_Protected_SetWorkbookHashValue, addToHistory);
		this.workbookSaltValue = this.checkProperty(this.workbookSaltValue, val.workbookSaltValue, AscCH.historyitem_Protected_SetWorkbookSaltValue, addToHistory);
		this.workbookSpinCount = this.checkProperty(this.workbookSpinCount, val.workbookSpinCount, AscCH.historyitem_Protected_SetWorkbookSpinCount, addToHistory);

		this.workbookPassword = this.checkProperty(this.workbookPassword, val.workbookPassword, AscCH.historyitem_Protected_SetPassword, addToHistory);
	};

	CWorkbookProtection.prototype.checkProperty = function (propOld, propNew, type, addToHistory) {
		if (propOld !== propNew) {
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoProtectedWorkbook, type, null, null,
					new AscCommonExcel.UndoRedoData_ProtectedRange(null, propOld, propNew));
			}
			return propNew;
		}
		return propOld;
	};

	CWorkbookProtection.prototype.Write_ToBinary2 = function(w) {
		if (null != this.lockStructure) {
			w.WriteBool(true);
			w.WriteBool(this.lockStructure);
		} else {
			w.WriteBool(false);
		}
		if (null != this.lockWindows) {
			w.WriteBool(true);
			w.WriteBool(this.lockWindows);
		} else {
			w.WriteBool(false);
		}
		if (null != this.lockRevision) {
			w.WriteBool(true);
			w.WriteBool(this.lockRevision);
		} else {
			w.WriteBool(false);
		}

		/*<xsd:attribute name="revisionsAlgorithmName" type="s:ST_Xstring" use="optional"/>
		 <xsd:attribute name="revisionsHashValue" type="xsd:base64Binary" use="optional"/>
		 <xsd:attribute name="revisionsSaltValue" type="xsd:base64Binary" use="optional"/>
		 <xsd:attribute name="revisionsSpinCount" type="xsd:unsignedInt" use="optional"/>

		 <xsd:attribute name="workbookAlgorithmName" type="s:ST_Xstring" use="optional"/>
		 <xsd:attribute name="workbookHashValue" type="xsd:base64Binary" use="optional"/>
		 <xsd:attribute name="workbookSaltValue" type="xsd:base64Binary" use="optional"/>
		 <xsd:attribute name="workbookSpinCount" type="xsd:unsignedInt" use="optional"/>
		 </xsd:complexType>*/
		if (null != this.revisionsAlgorithmName) {
			w.WriteBool(true);
			w.WriteLong(this.revisionsAlgorithmName);
		} else {
			w.WriteBool(false);
		}
		if (null != this.revisionsHashValue) {
			w.WriteBool(true);
			w.WriteString2(this.revisionsHashValue);
		} else {
			w.WriteBool(false);
		}
		if (null != this.revisionsSaltValue) {
			w.WriteBool(true);
			w.WriteString2(this.revisionsSaltValue);
		} else {
			w.WriteBool(false);
		}
		if (null != this.revisionsSpinCount) {
			w.WriteBool(true);
			w.WriteLong(this.revisionsSpinCount);
		} else {
			w.WriteBool(false);
		}

		if (null != this.workbookAlgorithmName) {
			w.WriteBool(true);
			w.WriteLong(this.workbookAlgorithmName);
		} else {
			w.WriteBool(false);
		}
		if (null != this.workbookHashValue) {
			w.WriteBool(true);
			w.WriteString2(this.workbookHashValue);
		} else {
			w.WriteBool(false);
		}
		if (null != this.workbookSaltValue) {
			w.WriteBool(true);
			w.WriteString2(this.workbookSaltValue);
		} else {
			w.WriteBool(false);
		}
		if (null != this.workbookSpinCount) {
			w.WriteBool(true);
			w.WriteLong(this.workbookSpinCount);
		} else {
			w.WriteBool(false);
		}
		if (null != this.workbookPassword) {
			w.WriteBool(true);
			w.WriteString2(this.workbookPassword);
		} else {
			w.WriteBool(false);
		}
	};

	CWorkbookProtection.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			this.lockStructure = r.GetBool();
		}
		if (r.GetBool()) {
			this.lockWindows = r.GetBool();
		}
		if (r.GetBool()) {
			this.lockRevision = r.GetBool();
		}

		if (r.GetBool()) {
			this.revisionsAlgorithmName = r.GetLong();
		}
		if (r.GetBool()) {
			this.revisionsHashValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.revisionsSaltValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.revisionsSpinCount = r.GetLong();
		}

		if (r.GetBool()) {
			this.workbookAlgorithmName = r.GetLong();
		}
		if (r.GetBool()) {
			this.workbookHashValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.workbookSaltValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.workbookSpinCount = r.GetLong();
		}

		if (r.GetBool()) {
			this.workbookPassword = r.GetString2();
		}
	};

	CWorkbookProtection.prototype.asc_getLockStructure = function () {
		return this.lockStructure;
	};
	CWorkbookProtection.prototype.asc_getLockWindows = function () {
		return this.lockWindows;
	};
	CWorkbookProtection.prototype.asc_getLockRevision = function () {
		return this.lockRevision;
	};
	CWorkbookProtection.prototype.asc_getRevisionsAlgorithmName = function () {
		return this.revisionsAlgorithmName;
	};
	CWorkbookProtection.prototype.asc_getRevisionsSaltValue = function () {
		return this.revisionsHashValue;
	};
	CWorkbookProtection.prototype.asc_getRevisionsSpinCount = function () {
		return this.revisionsSaltValue;
	};
	CWorkbookProtection.prototype.asc_getWorkbookAlgorithmName = function () {
		return this.revisionsSpinCount;
	};
	CWorkbookProtection.prototype.asc_getWorkbookHashValue = function () {
		return this.workbookAlgorithmName;
	};
	CWorkbookProtection.prototype.asc_getWorkbookSaltValue = function () {
		return this.workbookHashValue;
	};
	CWorkbookProtection.prototype.asc_getWorkbookSpinCount = function () {
		return this.workbookSaltValue;
	};
	CWorkbookProtection.prototype.asc_setLockStructure = function (password, callback) {
		//просталяю временный пароль, аспинхронная проверка пароля в asc_setProtectedWorkbook
		this.setLockStructure(!this.lockStructure);

		if (this.lockStructure && password) {
			var hashParams = generateHashParams();
			this.workbookSaltValue = hashParams.saltValue;
			this.workbookSpinCount = hashParams.spinCount;
			this.workbookAlgorithmName = hashParams.algorithmName;
		}
		this.temporaryPassword = password;
		if (callback) {
			callback(this);
		}
	};
	CWorkbookProtection.prototype.setLockStructure = function (val) {
		this.lockStructure = val;
	};
	CWorkbookProtection.prototype.asc_setLockWindows = function (val) {
		this.lockWindows = val;
	};
	CWorkbookProtection.prototype.asc_setLockRevision = function (val) {
		this.lockRevision = val;
	};
	CWorkbookProtection.prototype.asc_setRevisionsAlgorithmName = function (val) {
		this.revisionsAlgorithmName = val;
	};
	CWorkbookProtection.prototype.asc_setRevisionsHashValue = function (val) {
		this.revisionsHashValue = val;
	};
	CWorkbookProtection.prototype.asc_setRevisionsSaltValue = function (val) {
		this.revisionsSaltValue = val;
	};
	CWorkbookProtection.prototype.asc_setRevisionsSpinCount = function (val) {
		this.revisionsSpinCount = val;
	};
	CWorkbookProtection.prototype.asc_setWorkbookAlgorithmName = function (val) {
		this.workbookAlgorithmName = val;
	};
	CWorkbookProtection.prototype.asc_setWorkbookHashValue = function (val) {
		this.workbookHashValue = val;
	};
	CWorkbookProtection.prototype.asc_setWorkbookSaltValue = function (val) {
		this.workbookSaltValue = val;
	};
	CWorkbookProtection.prototype.asc_setWorkbookSpinCount = function (val) {
		this.workbookSpinCount = val;
	};
	CWorkbookProtection.prototype.asc_setPassword = function (val) {
		//генерируем хэш
		this.workbookAlgorithmName = "test";
		this.workbookHashValue = "test";
		this.workbookSaltValue = "test";
	};
	CWorkbookProtection.prototype.asc_isPassword = function (val) {
		return this.workbookAlgorithmName != null || this.workbookPassword != null;
	};
	CWorkbookProtection.prototype.setPasswordXL = function (val) {
		this.workbookPassword = val;
	};
	CWorkbookProtection.prototype.getPasswordXL = function () {
		return this.workbookPassword;
	};
	CWorkbookProtection.prototype.isPasswordXL = function () {
		return this.workbookPassword != null;
	};


	function CProtectedRange(ws) {
		this.sqref = null;
		this.name = null;

		this.algorithmName = null;
		this.hashValue = null;
		this.saltValue = null;
		this.spinCount = null;

		this._ws = ws;
		this.isLock = null;

		this.Id = AscCommon.g_oIdCounter.Get_NewId();

		this._isEnterPassword = null;
		this.temporaryPassword = null;

		return this;
	}

	CProtectedRange.prototype.Get_Id = function () {
		return this.Id;
	};

	CProtectedRange.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.ProtectedRangeDataInner;
	};

	CProtectedRange.prototype.clone = function(ws) {
		var res = new CProtectedRange(ws);

		res.sqref = this.sqref;
		res.name = this.name;

		res.algorithmName = this.algorithmName;
		res.hashValue = this.hashValue;
		res.saltValue = this.saltValue;
		res.spinCount = this.spinCount;

		return res;
	};

	CProtectedRange.prototype.set = function (val, addToHistory, ws) {
		
		this.cleanTemp();

		this.name = this.checkProperty(this.name, val.name, AscCH.historyitem_Protected_SetName, ws, addToHistory);
		this.algorithmName = this.checkProperty(this.algorithmName, val.algorithmName, AscCH.historyitem_Protected_SetAlgorithmName, ws, addToHistory);
		this.hashValue = this.checkProperty(this.hashValue, val.hashValue, AscCH.historyitem_Protected_SetHashValue, ws, addToHistory);
		this.saltValue = this.checkProperty(this.saltValue, val.saltValue, AscCH.historyitem_Protected_SetSaltValue, ws, addToHistory);
		this.spinCount = this.checkProperty(this.spinCount, val.spinCount, AscCH.historyitem_Protected_SetSpinCount, ws, addToHistory);


		var compareElements = function (_elem1, _elem2) {
			if (_elem1.length === _elem2.length) {
				for (var i = 0; i < _elem1.length; i++) {
					if (!_elem1[i].isEqual(_elem2[i])) {
						return false;
					}
				}
				return true;
			}
			return false;
		};

		if (this.sqref && val.sqref && !compareElements(this.sqref, val.sqref)) {
			this.setSqref(val.sqref, ws, true);
		}
	};

	CProtectedRange.prototype.cleanTemp = function () {
		this._isEnterPassword = null;
		//this.temporaryPassword = null;
	};

	CProtectedRange.prototype.checkProperty = function (propOld, propNew, type, ws, addToHistory) {
		if (propOld !== propNew) {
			if (addToHistory) {
				History.Add(AscCommonExcel.g_oUndoRedoProtectedRange, type, ws.getId(), null,
					new AscCommonExcel.UndoRedoData_ProtectedRange(this.Id, propOld, propNew));
			}
			return propNew;
		}
		return propOld;
	};

	CProtectedRange.prototype.setSqref = function (location, ws, addToHistory) {
		if (addToHistory) {
			var getUndoRedoRange = function (_ranges) {
				var needRanges = [];
				for (var i = 0; i < _ranges.length; i++) {
					needRanges.push(new AscCommonExcel.UndoRedoData_BBox(_ranges[i]));
				}
				return needRanges;
			};

			History.Add(AscCommonExcel.g_oUndoRedoProtectedRange, AscCH.historyitem_Protected_SetSqref,
				ws.getId(), null, new AscCommonExcel.UndoRedoData_ProtectedRange(this.Id, getUndoRedoRange(this.sqref), getUndoRedoRange(location)));
		}
		this.sqref = location;
	};

	CProtectedRange.prototype.setOffset = function(offset, range, ws, addToHistory) {
		var newRanges = [];
		var isChange = false;

		var _setDiff = function (_range) {
			//TODO объединть в одну функцию с dataValidation(.shift)
			var _newRanges, _offset, tempRange, intersection, otherPart, diff;

			if (range && range.getType() === Asc.c_oAscSelectionType.RangeCells) {
				if (offset.row !== 0) {
					//c_oAscInsertOptions.InsertCellsAndShiftDown
					tempRange = new Asc.Range(range.c1, range.r1, range.c2, AscCommon.gc_nMaxRow0);
					intersection = tempRange.intersection(_range);
					if (intersection) {
						diff = range.r2 - range.r1 + 1;

						_newRanges = [];
						//добавляем сдвинутую часть диапазона
						_newRanges.push(intersection);
						_offset = new AscCommon.CellBase(offset.row > 0 ? diff : -diff, 0);
						otherPart = _newRanges[0].difference(_range);
						_newRanges[0].setOffset(_offset);
						//исключаем сдвинутую часть из диапазона
						_newRanges = _newRanges.concat(otherPart);

					}
				} else if (offset.col !== 0) {
					//c_oAscInsertOptions.InsertCellsAndShiftRight
					tempRange = new Asc.Range(range.c1, range.r1, AscCommon.gc_nMaxCol0, range.r2);
					intersection = tempRange.intersection(_range);
					if (intersection) {
						diff = range.c2 - range.c1 + 1;
						_newRanges = [];
						//добавляем сдвинутую часть диапазона
						_newRanges.push(intersection);
						_offset = new AscCommon.CellBase(0, offset.col > 0 ? diff : -diff, 0);
						otherPart = _newRanges[0].difference(_range);
						_newRanges[0].setOffset(_offset);
						//исключаем сдвинутую часть из диапазона
						_newRanges = _newRanges.concat(otherPart);
					}
				}
			}

			return _newRanges;
		};

		for (var i = 0; i < this.sqref.length; i++) {
			var newRange = this.sqref[i].clone();
			if (range.isIntersectForShift(newRange, offset)) {
				if (newRange.forShift(range, offset)) {
					if (ws.autoFilters.isAddTotalRow && newRange.containsRange(this.sqref[i])) {
						newRange = this.sqref[i].clone();
					} else {
						isChange = true;
					}
				}
				newRanges.push(newRange);
			} else {
				if (ws.autoFilters.isAddTotalRow && newRange.containsRange(this.sqref[i])) {
					newRange = this.sqref[i].clone();
				} else {
					var changedRanges = _setDiff(this.sqref[i]);
					if (changedRanges) {
						newRanges = newRanges.concat(changedRanges);
						isChange = true;
					} else {
						newRanges = newRanges.concat(this.sqref[i].clone());
					}
				}
			}
		}
		if (isChange) {
			this.setSqref(newRanges, ws, addToHistory);
		}
	};

	CProtectedRange.prototype.Write_ToBinary2 = function(w) {
		if (null != this.sqref) {
			w.WriteBool(true);
			w.WriteLong(this.sqref.length);
			for (var i = 0; i < this.sqref.length; i++) {
				w.WriteLong(this.sqref[i].r1);
				w.WriteLong(this.sqref[i].c1);
				w.WriteLong(this.sqref[i].r2);
				w.WriteLong(this.sqref[i].c2);
			}
		} else {
			w.WriteBool(false);
		}
		if (null != this.name) {
			w.WriteBool(true);
			w.WriteString2(this.name);
		} else {
			w.WriteBool(false);
		}

		if (null != this.algorithmName) {
			w.WriteBool(true);
			w.WriteLong(this.algorithmName);
		} else {
			w.WriteBool(false);
		}
		if (null != this.hashValue) {
			w.WriteBool(true);
			w.WriteString2(this.hashValue);
		} else {
			w.WriteBool(false);
		}
		if (null != this.saltValue) {
			w.WriteBool(true);
			w.WriteString2(this.saltValue);
		} else {
			w.WriteBool(false);
		}
		if (null != this.spinCount) {
			w.WriteBool(true);
			w.WriteLong(this.spinCount);
		} else {
			w.WriteBool(false);
		}
	};

	CProtectedRange.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			var length = r.GetULong();
			for (var i = 0; i < length; ++i) {
				if (!this.sqref) {
					this.sqref = [];
				}
				var r1 = r.GetLong();
				var c1 = r.GetLong();
				var r2 = r.GetLong();
				var c2 = r.GetLong();
				this.sqref.push(new Asc.Range(c1, r1, c2, r2));
			}
		}
		if (r.GetBool()) {
			this.name = r.GetString2();
		}

		if (r.GetBool()) {
			this.algorithmName = r.GetLong();
		}
		if (r.GetBool()) {
			this.hashValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.saltValue = r.GetString2();
		}
		if (r.GetBool()) {
			this.spinCount = r.GetLong();
		}
	};

	CProtectedRange.prototype.contains = function (c, r) {
		//TODO  в каком виде будет хранится sqref?
		for (var i = 0; i < this.sqref.length; i++) {
			if (this.sqref[i].contains(c, r)) {
				return true;
			}
		}
		return false;
	};

	CProtectedRange.prototype.containsRange = function (range) {
		//TODO  в каком виде будет хранится sqref?
		for (var i = 0; i < this.sqref.length; i++) {
			if (this.sqref[i].containsRange(range)) {
				return true;
			}
		}
		return false;
	};
	CProtectedRange.prototype.intersection = function (range) {
		//TODO  в каком виде будет хранится sqref?
		for (var i = 0; i < this.sqref.length; i++) {
			if (this.sqref[i].intersection(range)) {
				return true;
			}
		}
		return false;
	};
	CProtectedRange.prototype.containsIntoRange = function (range) {
		//TODO  в каком виде будет хранится sqref?
		for (var i = 0; i < this.sqref.length; i++) {
			if (!range.containsRange(this.sqref[i])) {
				return false;
			}
		}
		return true;
	};
	CProtectedRange.prototype.generateNewName = function (modelProtectedRanges) {
		if (modelProtectedRanges.length) {
			var mapNames = [];
			for (var i = 0; i < modelProtectedRanges.length; i++) {
				mapNames[modelProtectedRanges[i].name] = 1;
			}
			var counter = 1;
			var newName = this.name;
			while (mapNames[newName]) {
				newName = this.name + "_" + counter;
				counter++;
			}
			this.name = newName;
		}
	};
	CProtectedRange.prototype.asc_getSqref = function () {
		var arrResult = [];

		if (this.sqref) {
			this.sqref.forEach(function (item) {
				arrResult.push(item.getAbsName());
			});
		}
		return "=" + arrResult.join(AscCommon.FormulaSeparators.functionArgumentSeparator);
	};
	CProtectedRange.prototype.asc_getName = function () {
		return this.name;
	};
	CProtectedRange.prototype.asc_getAlgorithmName = function () {
		return this.algorithmName;
	};
	CProtectedRange.prototype.asc_getHashValue = function () {
		return this.hashValue;
	};
	CProtectedRange.prototype.asc_getSaltValue = function () {
		return this.saltValue;
	};
	CProtectedRange.prototype.asc_getSpinCount = function () {
		return this.spinCount;
	};
	CProtectedRange.prototype.asc_setSqref = function (val) {
		var t = this;
		if (val) {
			if (val[0] === "=") {
				val = val.slice(1);
			}
			this.sqref = [];
			if (!val) {
				return;
			}
			val = val.split(",");
			val.forEach(function (item) {
				if (-1 !== item.indexOf("!")) {
					var is3DRef = AscCommon.parserHelp.parse3DRef(item);
					if (is3DRef) {
						item = is3DRef.range;
					}
				}
				t.sqref.push(AscCommonExcel.g_oRangeCache.getAscRange(item));
			});
		}
	};
	CProtectedRange.prototype.asc_setName = function (val) {
		this.name = val;
	};
	CProtectedRange.prototype.asc_setAlgorithmName = function (val) {
		this.algorithmName = val;
	};
	CProtectedRange.prototype.asc_setHashValue = function (val) {
		this.hashValue = val;
	};
	CProtectedRange.prototype.asc_setSaltValue = function (val) {
		this.saltValue = val;
	};
	CProtectedRange.prototype.asc_setSpinCount = function (val) {
		this.spinCount = val;
	};
	CProtectedRange.prototype.asc_setPassword = function (val) {
		if (val) {
			var hashParams = generateHashParams();
			this.saltValue = hashParams.saltValue;
			this.spinCount = hashParams.spinCount;
			this.algorithmName = hashParams.algorithmName;
		}
		//генерируем хэш
		this.temporaryPassword = val;
	};
	CProtectedRange.prototype.asc_isPassword = function () {
		return this.algorithmName != null;
	};
	CProtectedRange.prototype.asc_getIsLock = function () {
		return this.isLock;
	};
	CProtectedRange.prototype.asc_checkPassword = function (val, callback) {
		var checkHash = {password: val, salt: this.saltValue, spinCount: this.spinCount, alg: fromModelAlgoritmName(this.algorithmName)};
		AscCommon.calculateProtectHash([checkHash], function (hash) {
			callback(hash && hash[0] === this.hashValue);
		});
	};
	CProtectedRange.prototype.asc_getId = function () {
		return this.Id;
	};
	CProtectedRange.prototype.isUserEnteredPassword = function () {
		return this._isEnterPassword;
	};

	CProtectedRange.sStartLock = 'protectedRange_';



	//----------------------------------------------------------export----------------------------------------------------
	var prot;
	window['Asc'] = window['Asc'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["Asc"].CSheetProtection = CSheetProtection;
	prot = CSheetProtection.prototype;
	prot["asc_getSpinCount"] = prot.getSpinCount;
	prot["asc_getSheet"] = prot.getSheet;
	prot["asc_getObjects"] = prot.getObjects;
	prot["asc_getScenarios"] = prot.getScenarios;
	prot["asc_getFormatCells"] = prot.getFormatCells;
	prot["asc_getFormatColumns"] = prot.getFormatColumns;
	prot["asc_getFormatRows"] = prot.getFormatRows;
	prot["asc_getInsertColumns"] = prot.getInsertColumns;
	prot["asc_getInsertRows"] = prot.getInsertRows;
	prot["asc_getInsertHyperlinks"] = prot.getInsertHyperlinks;
	prot["asc_getDeleteColumns"] = prot.getDeleteColumns;
	prot["asc_getDeleteRows"] = prot.getDeleteRows;
	prot["asc_getSelectLockedCells"] = prot.getSelectLockedCells;
	prot["asc_getSort"] = prot.getSort;
	prot["asc_getAutoFilter"] = prot.getAutoFilter;
	prot["asc_getPivotTables"] = prot.getPivotTables;
	prot["asc_getSelectUnlockedCells"] = prot.getSelectUnlockedCells;
	prot["asc_setSpinCount"] = prot.setSpinCount;
	prot["asc_setSheet"] = prot.asc_setSheet;
	prot["asc_setObjects"] = prot.setObjects;
	prot["asc_setScenarios"] = prot.setScenarios;
	prot["asc_setFormatCells"] = prot.setFormatCells;
	prot["asc_setFormatColumns"] = prot.setFormatColumns;
	prot["asc_setFormatRows"] = prot.setFormatRows;
	prot["asc_setInsertColumns"] = prot.setInsertColumns;
	prot["asc_setInsertRows"] = prot.setInsertRows;
	prot["asc_setInsertHyperlinks"] = prot.setInsertHyperlinks;
	prot["asc_setDeleteColumns"] = prot.setDeleteColumns;
	prot["asc_setDeleteRows"] = prot.setDeleteRows;
	prot["asc_setSelectLockedCells"] = prot.setSelectLockedCells;
	prot["asc_setSort"] = prot.setSort;
	prot["asc_setAutoFilter"] = prot.setAutoFilter;
	prot["asc_setPivotTables"] = prot.setPivotTables;
	prot["asc_setSelectUnlockedCells"] = prot.setSelectUnlockedCells;
	prot["asc_setPassword"] = prot.asc_setPassword;
	prot["asc_isPassword"] = prot.asc_isPassword;


	window["Asc"].CWorkbookProtection = CWorkbookProtection;
	prot = CWorkbookProtection.prototype;
	prot["asc_getLockStructure"] = prot.asc_getLockStructure;
	prot["asc_getLockWindows"] = prot.asc_getLockWindows;
	prot["asc_getLockRevision"] = prot.asc_getLockRevision;
	prot["asc_getRevisionsAlgorithmName"] = prot.asc_getRevisionsAlgorithmName;
	prot["asc_getRevisionsSaltValue"] = prot.asc_getRevisionsSaltValue;
	prot["asc_getRevisionsSpinCount"] = prot.asc_getRevisionsSpinCount;
	prot["asc_getWorkbookAlgorithmName"] = prot.asc_getWorkbookAlgorithmName;
	prot["asc_getWorkbookHashValue"] = prot.asc_getWorkbookHashValue;
	prot["asc_getWorkbookSaltValue"] = prot.asc_getWorkbookSaltValue;
	prot["asc_getWorkbookSpinCount"] = prot.asc_getWorkbookSpinCount;
	prot["asc_getSelectLockedCells"] = prot.asc_getSelectLockedCells;
	prot["asc_setLockStructure"] = prot.asc_setLockStructure;
	prot["asc_setLockWindows"] = prot.asc_setLockWindows;
	prot["asc_setLockRevision"] = prot.asc_setLockRevision;
	prot["asc_setRevisionsAlgorithmName"] = prot.asc_setRevisionsAlgorithmName;
	prot["asc_setRevisionsSaltValue"] = prot.asc_setRevisionsSaltValue;
	prot["asc_setRevisionsSpinCount"] = prot.asc_setRevisionsSpinCount;
	prot["asc_setWorkbookAlgorithmName"] = prot.asc_setWorkbookAlgorithmName;
	prot["asc_setWorkbookHashValue"] = prot.asc_setWorkbookHashValue;
	prot["asc_setWorkbookSaltValue"] = prot.asc_setWorkbookSaltValue;
	prot["asc_setWorkbookSpinCount"] = prot.asc_setWorkbookSpinCount;
	prot["asc_setPassword"] = prot.asc_setPassword;
	prot["asc_isPassword"] = prot.asc_isPassword;

	window["Asc"]["CProtectedRange"] = window["Asc"].CProtectedRange = CProtectedRange;
	prot = CProtectedRange.prototype;
	prot["asc_getSqref"] = prot.asc_getSqref;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getAlgorithmName"] = prot.asc_getAlgorithmName;
	prot["asc_getHashValue"] = prot.asc_getHashValue;
	prot["asc_getSaltValue"] = prot.asc_getSaltValue;
	prot["asc_getSpinCount"] = prot.asc_getSpinCount;

	prot["asc_setSqref"] = prot.asc_setSqref;
	prot["asc_setName"] = prot.asc_setName;
	prot["asc_setAlgorithmName"] = prot.asc_setAlgorithmName;
	prot["asc_setHashValue"] = prot.asc_setHashValue;
	prot["asc_setSaltValue"] = prot.asc_setSaltValue;
	prot["asc_setSpinCount"] = prot.asc_setSpinCount;
	prot["asc_setPassword"] = prot.asc_setPassword;
	prot["asc_isPassword"] = prot.asc_isPassword;
	prot["asc_getIsLock"] = prot.asc_getIsLock;
	prot["asc_checkPassword"] = prot.asc_checkPassword;
	prot["asc_getId"] = prot.asc_getId;

	window["AscCommonExcel"].fromModelAlgoritmName = fromModelAlgoritmName;
	window["AscCommonExcel"].getPasswordHash = getPasswordHash;

})(window);
