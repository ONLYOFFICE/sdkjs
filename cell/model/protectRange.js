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

(function (window) {

	function CUserProtectedRange(ws) {
		this.ref = null;
		this.name = null;
		this.users = null;

		this.usersGroups = null;

		//for warning
		this.warningText = null;

		//everyone edit/view type
		this.type = null;//c_oSerUserProtectedRangeType -> notView/view/edit, //default -> Asc.c_oSerUserProtectedRangeType.view

		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		this._ws = ws;
		this.isLock = null;

		return this;
	}

	CUserProtectedRange.prototype.Get_Id = function () {
		return this.Id;
	};

	CUserProtectedRange.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.UserProtectedRange;
	};

	CUserProtectedRange.prototype.clone = function (ws, needCloneId) {
		var res = new CUserProtectedRange(ws);

		res.ref = this.ref ? this.ref.clone() : null;
		res.name = this.name;
		res.users = this.users;
		res.usersGroups = this.usersGroups;

		res.warningText = this.warningText;

		res.type = this.type;

		if (needCloneId) {
			res.Id = this.Id;
		}

		return res;
	};

	CUserProtectedRange.prototype.Write_ToBinary2 = function (w) {
		if (null != this.ref) {
			w.WriteBool(true);

			w.WriteLong(this.ref.r1);
			w.WriteLong(this.ref.c1);
			w.WriteLong(this.ref.r2);
			w.WriteLong(this.ref.c2);
		} else {
			w.WriteBool(false);
		}
		if (null != this.name) {
			w.WriteBool(true);
			w.WriteString2(this.name);
		} else {
			w.WriteBool(false);
		}

		if (null != this.users) {
			let count = this.users.length;
			if (count) {
				w.WriteBool(true);
				w.WriteLong(count);
				for (let i = 0; i < count; i++) {
					this.users[i].Write_ToBinary2(w);
				}
			} else {
				w.WriteBool(false);
			}
		} else {
			w.WriteBool(false);
		}

		if (null != this.usersGroups) {
			let count = this.usersGroups.length;
			if (count) {
				w.WriteBool(true);
				w.WriteLong(count);
				for (let i = 0; i < count; i++) {
					this.usersGroups[i].Write_ToBinary2(w);
				}
			} else {
				w.WriteBool(false);
			}
		} else {
			w.WriteBool(false);
		}

		if (null != this.warningText) {
			w.WriteBool(true);
			w.WriteString2(this.warningText);
		} else {
			w.WriteBool(false);
		}

		if (null != this.Id) {
			w.WriteBool(true);
			w.WriteString2(this.Id);
		} else {
			w.WriteBool(false);
		}

		if (null != this.type) {
			w.WriteBool(true);
			w.WriteLong(this.type);
		} else {
			w.WriteBool(false);
		}
	};

	CUserProtectedRange.prototype.Read_FromBinary2 = function (r) {
		if (r.GetBool()) {
			var r1 = r.GetLong();
			var c1 = r.GetLong();
			var r2 = r.GetLong();
			var c2 = r.GetLong();
			this.ref = new Asc.Range(c1, r1, c2, r2);
		}
		if (r.GetBool()) {
			this.name = r.GetString2();
		}

		if (r.GetBool()) {
			let length = r.GetULong();
			for (let i = 0; i < length; ++i) {
				if (!this.users) {
					this.users = [];
				}
				let newUserInfo = new CUserProtectedRangeUserInfo();
				newUserInfo.Read_FromBinary2(r);
				this.users.push(newUserInfo);
			}
		}

		if (r.GetBool()) {
			let length = r.GetULong();
			for (let i = 0; i < length; ++i) {
				if (!this.usersGroups) {
					this.usersGroups = [];
				}
				let newUserInfo = new CUserProtectedRangeUserInfo();
				newUserInfo.Read_FromBinary2(r);
				this.usersGroups.push(newUserInfo);
			}
		}

		if (r.GetBool()) {
			this.warningText = r.GetString2();
		}

		if (r.GetBool()) {
			this.Id = r.GetString2();
		}

		if (r.GetBool()) {
			this.type = r.GetLong();
		}

	};
	CUserProtectedRange.prototype.intersection = function (range) {
		return this.ref && this.ref.intersection(range);
	};
	CUserProtectedRange.prototype.contains2 = function (oCell) {
		return this.ref && this.ref.contains2(oCell);
	};
	CUserProtectedRange.prototype.isUserCanDoByType = function (userId, type) {
		let res = false;

		switch (type) {
			case Asc.c_oSerUserProtectedRangeType.view: {
				res = this.isUserCanView(userId);
				break;
			}
			case Asc.c_oSerUserProtectedRangeType.edit:
			default: {
				res = this.isUserCanEdit(userId);
				break;
			}
		}

		return res;
	};
	CUserProtectedRange.prototype.isUserCanEdit = function (userId) {
		if (this.users) {
			for (let i = 0; i < this.users.length; i++) {
				if (this.users[i].isCanEdit(userId)) {
					return true;
				}
			}
		}
		return false;
	};
	CUserProtectedRange.prototype.isUserCanView = function (userId) {
		if (this.asc_getType() !== Asc.c_oSerUserProtectedRangeType.notView) {
			return true;
		}
		if (this.users) {
			for (let i = 0; i < this.users.length; i++) {
				if (this.users[i].isCanView(userId)) {
					return true;
				}
			}
		}
		return false;
	};
	CUserProtectedRange.prototype.isInRange = function (bbox) {
		return bbox.containsRange(this.ref);
	};
	CUserProtectedRange.prototype.setOffset = function (offset, addToHistory) {
		var ref = this.ref.clone();
		ref.setOffset(offset);
		this.setLocation(ref, addToHistory);
	};
	CUserProtectedRange.prototype.setLocation = function (ref, addToHistory) {
		if (addToHistory) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoUserProtectedRange, AscCH.historyitem_UserProtectedRange_Ref,
				this._ws ? this._ws.getId() : null, null,
				new AscCommonExcel.UndoRedoData_UserProtectedRange(this.Get_Id(),
					new AscCommonExcel.UndoRedoData_BBox(this.ref), new AscCommonExcel.UndoRedoData_BBox(ref)));
		}
		this.ref = ref;
	};
	CUserProtectedRange.prototype.getUserById = function (userId) {
		if (this.users) {
			for (let i = 0; i < this.users.length; i++) {
				if (this.users[i].id === userId) {
					return {obj: this.users[i], index: i};
				}
			}
		}
		return null;
	};

	CUserProtectedRange.prototype.asc_getRef = function () {
		var result = null;

		if (this.ref) {
			result = this.ref.getAbsName();
			if (this._ws) {
				result = AscCommon.parserHelp.get3DRef(this._ws.getName(), result);
			}
		}

		return result ? "=" + result : null;
	};
	CUserProtectedRange.prototype.asc_getName = function () {
		return this.name;
	};
	CUserProtectedRange.prototype.asc_getUsers = function () {
		return this.users;
	};
	CUserProtectedRange.prototype.asc_getUserGroups = function () {
		return this.usersGroups;
	};
	CUserProtectedRange.prototype.asc_setRef = function (val) {
		if (val) {
			if (val[0] === "=") {
				val = val.slice(1);
			}
			this.ref = [];
			if (!val) {
				return;
			}

			let sheetName;
			if (-1 !== val.indexOf("!")) {
				var is3DRef = AscCommon.parserHelp.parse3DRef(val);
				if (is3DRef) {
					val = is3DRef.range;
					sheetName = is3DRef.sheet;
				}
			}
			let api = window["Asc"]["editor"];
			let wbModel = api.wbModel;
			if (wbModel) {
				this._ws = sheetName ? wbModel.getWorksheetByName(sheetName) : wbModel.getActiveWs();
			}
			this.ref = AscCommonExcel.g_oRangeCache.getAscRange(val);
		}
	};
	CUserProtectedRange.prototype.asc_setName = function (val) {
		this.name = val.trim && val.trim();
	};
	CUserProtectedRange.prototype.asc_setUsers = function (val) {
		this.users = val;
	};
	CUserProtectedRange.prototype.asc_setType = function (val) {
		this.type = val;
	};
	CUserProtectedRange.prototype.asc_getIsLock = function () {
		return this.isLock;
	};
	CUserProtectedRange.prototype.asc_getId = function () {
		return this.Id;
	};
	CUserProtectedRange.prototype.asc_getType = function () {
		return this.type == null ? Asc.c_oSerUserProtectedRangeType.view : this.type;
	};
	CUserProtectedRange.prototype.initPostOpen = function (ws) {
		this._ws = ws;
	};

	function CUserProtectedRangeUserInfo() {
		this.id = null;
		this.name = null;
		this.type = null;//default -> Asc.c_oSerUserProtectedRangeType.edit
	}

	CUserProtectedRangeUserInfo.prototype.isCanEdit = function (id) {
		let res = false;
		if (id === this.id) {
			res = this.asc_getType() === Asc.c_oSerUserProtectedRangeType.edit || this.asc_getType() == null;
		}
		return res;
	};

	CUserProtectedRangeUserInfo.prototype.isCanView = function (id) {
		let res = false;
		if (id === this.id) {
			res = this.asc_getType() !== Asc.c_oSerUserProtectedRangeType.notView;
		}
		return res;
	};

	CUserProtectedRangeUserInfo.prototype.Write_ToBinary2 = function (w) {
		if (null != this.id) {
			w.WriteBool(true);
			w.WriteString2(this.id);
		} else {
			w.WriteBool(false);
		}
		if (null != this.name) {
			w.WriteBool(true);
			w.WriteString2(this.name);
		} else {
			w.WriteBool(false);
		}
		if (null != this.type) {
			w.WriteBool(true);
			w.WriteLong(this.type);
		} else {
			w.WriteBool(false);
		}
	};

	CUserProtectedRangeUserInfo.prototype.Read_FromBinary2 = function (r) {
		if (r.GetBool()) {
			this.id = r.GetString2();
		}
		if (r.GetBool()) {
			this.name = r.GetString2();
		}
		if (r.GetBool()) {
			this.type = r.GetLong();
		}
	};

	CUserProtectedRangeUserInfo.prototype.getType = function () {
		return AscCommonExcel.UndoRedoDataTypes.UserProtectedRange;
	};

	CUserProtectedRangeUserInfo.prototype.clone = function (ws) {
		var res = new CUserProtectedRangeUserInfo(ws);

		res.id = this.id;
		res.name = this.name;
		res.type = this.type;

		return res;
	};

	CUserProtectedRangeUserInfo.prototype.asc_getId = function () {
		return this.id;
	};
	CUserProtectedRangeUserInfo.prototype.asc_getName = function () {
		return this.name;
	};
	CUserProtectedRangeUserInfo.prototype.asc_getType = function () {
		return this.type == null ? Asc.c_oSerUserProtectedRangeType.edit : this.type;
	};

	CUserProtectedRangeUserInfo.prototype.asc_setId = function (val) {
		this.id = val;
	};
	CUserProtectedRangeUserInfo.prototype.asc_setName = function (val) {
		this.name = val;
	};
	CUserProtectedRangeUserInfo.prototype.asc_setType = function (val) {
		this.type = val;
	};


	//----------------------------------------------------------export----------------------------------------------------
	var prot;
	window['Asc'] = window['Asc'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["Asc"]["CUserProtectedRange"] = window["Asc"].CUserProtectedRange = CUserProtectedRange;
	prot = CUserProtectedRange.prototype;
	prot["asc_getRef"] = prot.asc_getRef;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getUsers"] = prot.asc_getUsers;
	prot["asc_getUserGroups"] = prot.asc_getUserGroups;
	prot["asc_getType"] = prot.asc_getType;

	prot["asc_setRef"] = prot.asc_setRef;
	prot["asc_setName"] = prot.asc_setName;
	prot["asc_setUsers"] = prot.asc_setUsers;

	prot["asc_getId"] = prot.asc_getId;

	prot["asc_getIsLock"] = prot.asc_getIsLock;
	prot["asc_setType"] = prot.asc_setType;



	window["Asc"]["CUserProtectedRangeUserInfo"] = window["Asc"].CUserProtectedRangeUserInfo = CUserProtectedRangeUserInfo;
	prot = CUserProtectedRangeUserInfo.prototype;
	prot["asc_getId"] = prot.asc_getId;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getType"] = prot.asc_getType;
	prot["asc_setId"] = prot.asc_setId;
	prot["asc_setName"] = prot.asc_setName;
	prot["asc_setType"] = prot.asc_setType;

})(window);
