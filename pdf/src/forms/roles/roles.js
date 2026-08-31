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

(function (window) {
    /**
     * Manager of form roles for a PDF document. Holds the authoritative collection
     * of CPdfRoleInfo (each role owns the full-names of its fields).
     * @param {AscPDF.CPDFDoc} document
     * @constructor
     */

    CPdfFormRoles.DEFAULT_ROLE_NAME  = "Anyone";
    CPdfFormRoles.DEFAULT_ROLE_COLOR = "#dde4ff";

    function CPdfFormRoles(document) {
        this.document    = document;
        this.Roles       = [];
        this.CurrentRole = null;
        this.CurrentNoRole = false;
        this.Final       = false; // form finalized: structure editing forbidden
        this._EnsureDefaultRole();
    }

    CPdfFormRoles.prototype.GetDocument = function () { return this.document; };

    CPdfFormRoles.prototype._EnsureDefaultRole = function () {
        let defaultRole = null;
        for (let i = 0; i < this.Roles.length; ++i) {
            let role = this.Roles[i];
            if (role.IsDefault()) {
                if (!defaultRole)
                    defaultRole = role;
                else
                    role.SetDefault(false);
            }
        }
        if (defaultRole)
            return;

        let legacyDefault = this.GetRole(CPdfFormRoles.DEFAULT_ROLE_NAME);
        if (legacyDefault) {
            legacyDefault.SetDefault(true);
            return;
        }

        if (this.Roles.length > 0) {
            this.Roles[0].SetDefault(true);
            return;
        }

        this.Roles.push(new AscPDF.CPdfRoleInfo({
            "id":         AscCommon.CreateGUID(),
            "name":       CPdfFormRoles.DEFAULT_ROLE_NAME,
            "color":      CPdfFormRoles.DEFAULT_ROLE_COLOR,
            "isDefault":  true,
            "fieldNames": []
        }));
    };

    CPdfFormRoles.prototype.Serialize = function () {
        return {
            "final": this.Final,
            "roles": this.Roles.map(function (r) { return r.ToJSON(); })
        };
    };
    CPdfFormRoles.prototype._ParseData = function (data) {
        data = data || {};
        this.Final = !!data["final"];
        let arr = Array.isArray(data["roles"]) ? data["roles"] : [];
        this.Roles = arr.map(function (o) { return AscPDF.CPdfRoleInfo.FromJSON(o); })
            .filter(function (r) { return !!r; });
    };
    CPdfFormRoles.prototype.Load = function (data) {
        this._ParseData(data);
        this._RevalidateCurrent();
    };

    CPdfFormRoles.prototype.OnEndLoad = function () {
        this._PruneMissingFields();
        this._EnsureDefaultRole();
        this._RevalidateCurrent();
    };
    CPdfFormRoles.prototype._PruneMissingFields = function () {
        let oDoc = this.document;
        if (!oDoc || !oDoc.GetField)
            return;
        for (let i = 0; i < this.Roles.length; ++i) {
            let names = this.Roles[i].GetFieldNames().slice();
            for (let j = 0; j < names.length; ++j) {
                if (!oDoc.GetField(names[j]))
                    this.Roles[i].RemoveFieldName(names[j]);
            }
        }
    };

    // --- internal helpers ---
    CPdfFormRoles.prototype._RevalidateCurrent = function () {
        if (this.CurrentRole)
            this.CurrentRole = this.GetRoleById(this.CurrentRole.GetId());
    };
    CPdfFormRoles.prototype._Commit = function (aOld) {
        if (typeof CChangesPDFDocumentRolesInfo !== "undefined"
            && AscCommon.History && AscCommon.History.Add) {
            AscCommon.History.Add(new CChangesPDFDocumentRolesInfo(this.document, aOld, this.Serialize()));
        }
        this._RevalidateCurrent();
        this.OnUpdateRoles();
    };
	CPdfFormRoles.prototype._DoAction = function (callback, description, isIgnoreCanEditFlag) {
		let oDoc = this.GetDocument();
		if (!oDoc)
			return false;

		let api = oDoc.GetApi ? oDoc.GetApi() : null;
		if (api && api.isRestrictionSignatures && api.isRestrictionSignatures())
			return false;

		if (!isIgnoreCanEditFlag && !oDoc.IsEditFieldsMode())
			return false;

		if (oDoc.IsSelectionLocked(AscCommon.changestype_Document_Settings, null, false, !!isIgnoreCanEditFlag))
			return false;

		oDoc.StartAction(description || AscDFH.historydescription_Pdf_ChangeRoles);
		let result = callback.call(this);
		oDoc.FinalizeAction();
		return result;
	};
    CPdfFormRoles.prototype.ApplySnapshot = function (data) {
		let bOldFinal = this.Final;
        this._ParseData(data);
        this._EnsureDefaultRole();
        this._RevalidateCurrent();
        this.OnUpdateRoles();
		if (bOldFinal !== this.Final)
			this.OnChangeFinal();
    };

    // --- final / preparation ---
    CPdfFormRoles.prototype.IsFinal = function () { return this.Final; };
    CPdfFormRoles.prototype.SetFinal = function (bFinal) {
        if (this.Final === bFinal) return false;
		AscCommon.History.Add(new CChangesPDFDocumentFinal(this.document, this.Final, bFinal));
		this.Final = bFinal;
		this.OnChangeFinal();
        return true;
    };
    CPdfFormRoles.prototype.CompletePreparation = function () {
        let aOld = this.Serialize();
		let bOldFinal = this.Final;
        for (let i = 0; i < this.Roles.length; ++i) {
            let r = this.Roles[i];
            r.SetFilled(false); r.SetDate(null); r.SetUserId(null); r.SetUserName(null);
        }
        this.Final = true;
        this._Commit(aOld);
		if (bOldFinal !== this.Final)
			this.OnChangeFinal();
        return true;
    };
    CPdfFormRoles.prototype._IsRequiredFieldAvailable = function (field) {
		if (!field.IsRequired()) {
			return false;
		}

        if (field.IsReadOnly && field.IsReadOnly())
            return false;

        let widgets = field.GetAllWidgets();
        return widgets.some(function (widget) {
            return widget && (!widget.IsHidden || !widget.IsHidden());
        });
    };
    CPdfFormRoles.prototype.GetRoleUnfilledRequiredFields = function (name) {
        let role = this.GetRole(name);
        let oDoc = this.document;
        let res  = [];
        if (!role || !oDoc || !oDoc.GetField) return res;
        let names = this.GetEffectiveFieldNames(role);
        for (let i = 0; i < names.length; ++i) {
            let f = oDoc.GetField(names[i]);
            if (this._IsRequiredFieldAvailable(f)) {
				if (!f.IsFilled()) {
					res.push(names[i]);
				}
            }
        }
        return res;
    };
    CPdfFormRoles.prototype.GetAllUnfilledRequiredFields = function () {
        let res = [];
        for (let i = 0; i < this.Roles.length; ++i) {
            res = res.concat(this.GetRoleUnfilledRequiredFields(this.Roles[i].GetName()));
        }
        return res;
    };
    CPdfFormRoles.prototype.IsRequiredFilled = function (checkAll) {
        if (true !== checkAll && this.CurrentNoRole) return false;
        let name = true !== checkAll && this.CurrentRole ? this.CurrentRole.GetName() : null;
        let aMissing = name ? this.GetRoleUnfilledRequiredFields(name) : this.GetAllUnfilledRequiredFields();
        return aMissing.length === 0;
    };
    CPdfFormRoles.prototype.SetAllRolesFilled = function (bFilled, userPr) {
        let api = this.document && this.document.GetApi && this.document.GetApi();
        let sUserName = userPr && userPr["name"] !== undefined
			? userPr["name"]
			: (AscCommon.UserInfoParser ? AscCommon.UserInfoParser.getCurrentName() : null);
        let sUserId = userPr && userPr["id"] !== undefined
			? userPr["id"]
			: (api ? api.documentUserId : null);
        let aOld = this.Serialize();
        for (let i = 0; i < this.Roles.length; ++i) {
            let r = this.Roles[i];
            r.SetFilled(!!bFilled);
            if (bFilled) { r.SetDate(Date.now()); r.SetUserId(sUserId); r.SetUserName(sUserName); }
            else         { r.SetDate(null); r.SetUserId(null); r.SetUserName(null); }
        }
        this._Commit(aOld);
        return true;
    };

    CPdfFormRoles.prototype._NormalizeColor = function (color) {
        if (!color) return null;
        if (typeof color === "string") return color;
        if (typeof color.r === "number") {
            function h(n) { let s = n.toString(16); return s.length === 1 ? "0" + s : s; }
            return "#" + h(color.r) + h(color.g) + h(color.b);
        }
        return null;
    };
    CPdfFormRoles.prototype._HexFromSettings = function (roleSettings) {
        return this._NormalizeColor(roleSettings.getColor && roleSettings.getColor());
    };

    // --- queries ---
    CPdfFormRoles.prototype.GetAllRoles = function () { return this.Roles; };
    CPdfFormRoles.prototype.GetRole = function (name) {
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].GetName() === name) return this.Roles[i];
        }
        return null;
    };
    CPdfFormRoles.prototype.GetRoleById = function (id) {
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].GetId() === id) return this.Roles[i];
        }
        return null;
    };
    CPdfFormRoles.prototype.HaveRole = function (name) { return !!this.GetRole(name); };

    CPdfFormRoles.prototype.GetRoleByField = function (field) {
        let sName = field && field.GetFullName ? field.GetFullName() : null;
        if (sName == null) return null;
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].HasFieldName(sName)) return this.Roles[i];
        }
        return this.GetDefaultRole();
    };

    CPdfFormRoles.prototype.IsDefaultRole = function (role) {
        return !!role && role.IsDefault();
    };
    CPdfFormRoles.prototype.GetDefaultRole = function () {
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].IsDefault()) return this.Roles[i];
        }
        return null;
    };
    CPdfFormRoles.prototype._GetClaimedFieldNames = function () {
        let claimed = {};
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.IsDefaultRole(this.Roles[i])) continue;
            let names = this.Roles[i].GetFieldNames();
            for (let j = 0; j < names.length; ++j) claimed[names[j]] = true;
        }
        return claimed;
    };
    CPdfFormRoles.prototype.GetEffectiveFieldNames = function (role) {
        let explicit = role ? role.GetFieldNames() : [];
        if (!this.IsDefaultRole(role)) return explicit.slice();
        let oDoc = this.document;
        let seen = {};
        let res  = [];
        for (let i = 0; i < explicit.length; ++i) {
            if (!seen[explicit[i]]) { seen[explicit[i]] = true; res.push(explicit[i]); }
        }
        if (!oDoc || !oDoc.GetAllWidgets) return res;
        let claimed  = this._GetClaimedFieldNames();
        let aWidgets = oDoc.GetAllWidgets() || [];
        for (let i = 0; i < aWidgets.length; ++i) {
            let sFull = aWidgets[i].GetFullName ? aWidgets[i].GetFullName() : null;
            if (sFull && !claimed[sFull] && !seen[sFull]) { seen[sFull] = true; res.push(sFull); }
        }
        return res;
    };

    // --- authoring ---
    CPdfFormRoles.prototype.AddRole = function (name, color) {
        if (this.HaveRole(name)) return false;
        let aOld = this.Serialize();
        this.Roles.push(new AscPDF.CPdfRoleInfo({
            "id":         AscCommon.CreateGUID(),
            "name":       name,
            "color":      this._NormalizeColor(color),
            "fieldNames": []
        }));
        this._Commit(aOld);
        return true;
    };

    CPdfFormRoles.prototype.AssignFieldToRole = function (field, name) {
        let role = this.GetRole(name);
        if (!role || !field || !field.GetFullName) return false;
        let sName = field.GetFullName();
        let aOld = this.Serialize();
        // a field belongs to exactly one role
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i] !== role) this.Roles[i].RemoveFieldName(sName);
        }

        if (!this.IsDefaultRole(role))
            role.AddFieldName(sName);
        else
            role.RemoveFieldName(sName);
        this._Commit(aOld);

		field.AddToRedraw();
        return true;
    };

    CPdfFormRoles.prototype.RenameField = function (sOldName, sNewName) {
        if (sOldName === sNewName) return false;
        let role = null;
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].HasFieldName(sNewName)) { role = this.Roles[i]; break; }
        }
        if (!role) return false;
        let aOld = this.Serialize();
        role.AddFieldName(sNewName);
        let bOldStillExists = this.document && this.document.GetField && this.document.GetField(sOldName);
        if (!bOldStillExists)
            role.RemoveFieldName(sOldName);
        this._Commit(aOld);
        return true;
    };
    CPdfFormRoles.prototype.RemoveFieldFromRoles = function (sName) {
        let bHas = false;
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].HasFieldName(sName)) { bHas = true; break; }
        }
        if (!bHas) return false;
        let aOld = this.Serialize();
        for (let i = 0; i < this.Roles.length; ++i) {
            this.Roles[i].RemoveFieldName(sName);
        }
        this._Commit(aOld);
        return true;
    };

    CPdfFormRoles.prototype.EditRole = function (name, roleSettings) {
        let role = this.GetRole(name);
        if (!role) return false;

        let newName = roleSettings.getName ? roleSettings.getName() : undefined;
        if (newName && newName !== name && this.HaveRole(newName)) return false;

        let bHasColor = !!(roleSettings.getColor && roleSettings.getColor());
        let aOld = this.Serialize();
        if (newName) role.SetName(newName);
        if (bHasColor) role.SetColor(this._HexFromSettings(roleSettings));
        this._Commit(aOld);
        return true;
    };

    CPdfFormRoles.prototype.RemoveRole = function (name, delegateName) {
        let idx = -1;
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].GetName() === name) { idx = i; break; }
        }
        if (idx === -1) return false;
        let removed   = this.Roles[idx];
        let delegate  = (delegateName && delegateName !== name) ? this.GetRole(delegateName) : null;
        let aOld = this.Serialize();

        if (this.IsDefaultRole(removed)) {
            this.Roles.splice(idx, 1);

            let newDefault = delegate && delegate !== removed ? delegate : this.Roles[0];
            if (newDefault)
                newDefault.SetDefault(true);
            else
                this._EnsureDefaultRole();

            this._Commit(aOld);
            return true;
        }

        if (delegate && !this.IsDefaultRole(delegate)) {
            let names = removed.GetFieldNames();
            for (let i = 0; i < names.length; ++i) delegate.AddFieldName(names[i]);
        }
        this.Roles.splice(idx, 1);
        this._Commit(aOld);
        return true;
    };

    CPdfFormRoles.prototype.OnUpdateRoles = function () {
        let api = this.document && this.document.GetApi && this.document.GetApi();
        if (api && api.sendEvent)
            api.sendEvent("asc_onUpdateOFormRoles", this.GetAllRoles());
    };
	CPdfFormRoles.prototype.OnChangeFinal = function () {
		let api = this.document && this.document.GetApi && this.document.GetApi();
		if (api && api.sendEvent)
			api.sendEvent("asc_onOFormChangeFinal", this.IsFinal());
	};

    // --- ordering ---
    CPdfFormRoles.prototype._MoveRole = function (name, direction) {
        let idx = -1;
        for (let i = 0; i < this.Roles.length; ++i) {
            if (this.Roles[i].GetName() === name) { idx = i; break; }
        }
        if (idx === -1) return false;
        let other = idx + direction;
        if (other < 0 || other >= this.Roles.length) return false;

        let aOld = this.Serialize();
		let role = this.Roles[idx];
		this.Roles[idx] = this.Roles[other];
		this.Roles[other] = role;
        this._Commit(aOld);
        return true;
    };

	CPdfFormRoles.prototype.MoveUpRole   = function (name) { return this._MoveRole(name, -1); };
	CPdfFormRoles.prototype.MoveDownRole = function (name) { return this._MoveRole(name, +1); };

    CPdfFormRoles.prototype.CanFillRole = function (name) {
        let role = this.GetRole(name);
        if (!role || role.IsFilled()) return false;

		let roleIndex = this.Roles.indexOf(role);
		for (let i = 0; i < roleIndex; ++i) {
			if (!this.Roles[i].IsFilled()) return false;
		}
        return true;
    };

    // --- filling ---
    CPdfFormRoles.prototype.SetCurrentRole = function (name) {
        let role = this.GetRole(name);
        this.CurrentRole = role || null;
        this.CurrentNoRole = false;
        return !!role;
    };
    CPdfFormRoles.prototype.SetCurrentNoRole = function () {
        this.CurrentRole = null;
        this.CurrentNoRole = true;
    };
    CPdfFormRoles.prototype.ClearCurrentRole = function () {
        this.CurrentRole = null;
        this.CurrentNoRole = false;
    };
    CPdfFormRoles.prototype.GetCurrentRole = function () { return this.CurrentRole; };
    CPdfFormRoles.prototype.IsCurrentNoRole = function () { return this.CurrentNoRole; };

    CPdfFormRoles.prototype.SetRoleFilled = function (name, isFilled, userPr) {
        let role = this.GetRole(name);
        if (!role) return false;
        let api = this.document && this.document.GetApi && this.document.GetApi();
        let aOld = this.Serialize();
        role.SetFilled(!!isFilled);
        if (isFilled) {
            // stamp who/when
            role.SetDate(Date.now());
			role.SetUserId(userPr && userPr["id"] !== undefined
				? userPr["id"]
				: (api ? api.documentUserId : null));
			role.SetUserName(userPr && userPr["name"] !== undefined
				? userPr["name"]
				: (AscCommon.UserInfoParser ? AscCommon.UserInfoParser.getCurrentName() : null));
        }
        else {
            role.SetDate(null);
            role.SetUserId(null);
            role.SetUserName(null);
        }
        this._Commit(aOld);

        if (api && api.sendEvent)
            api.sendEvent("asc_onOFormRoleFilled", name, !!isFilled);
        return true;
    };

    CPdfFormRoles.prototype.CanFillField = function (field) {
        let role = this.GetRoleByField(field);
        if (!role) return true;
		
		let oDoc = this.document;
		if (oDoc.IsFillingFormMode()) {
			return !this.CurrentNoRole && this.CurrentRole != null && this.CurrentRole.GetId() === role.GetId();
		}
		else if (Asc.editor.IsCommentsMode()) {
			return !this.CurrentNoRole && role.IsDefault();
		}
    };

    CPdfFormRoles.prototype.IsAllRolesFilled = function () {
        for (let i = 0; i < this.Roles.length; ++i) {
            if (!this.Roles[i].IsFilled()) return false;
        }
        return this.Roles.length > 0;
    };

	// Public OForm-compatible interface.
	CPdfFormRoles.prototype.asc_getAllRoles = function () {
		return this.GetAllRoles();
	};
	CPdfFormRoles.prototype.asc_addRole = function (roleSettings, color) {
		let name = roleSettings;
		if (typeof roleSettings !== "string") {
			if (!roleSettings || !roleSettings.getName)
				return false;

			name = roleSettings.getName();
			color = roleSettings.getColor ? roleSettings.getColor() : color;
		}

		return this._DoAction(function () {
			return this.AddRole(name, color);
		});
	};
	CPdfFormRoles.prototype.asc_removeRole = function (name, delegateName) {
		return this._DoAction(function () {
			return this.RemoveRole(name, delegateName);
		});
	};
	CPdfFormRoles.prototype.asc_editRole = function (name, roleSettings) {
		return this._DoAction(function () {
			return this.EditRole(name, roleSettings);
		});
	};
	CPdfFormRoles.prototype.asc_moveUpRole = function (name) {
		return this._DoAction(function () {
			return this.MoveUpRole(name);
		});
	};
	CPdfFormRoles.prototype.asc_moveDownRole = function (name) {
		return this._DoAction(function () {
			return this.MoveDownRole(name);
		});
	};
	CPdfFormRoles.prototype.asc_haveRole = function (name) {
		return this.HaveRole(name);
	};
	CPdfFormRoles.prototype.asc_getRole = function (name) {
		let oRole = this.GetRole(name);
		return oRole ? oRole.GetSettings() : null;
	};
	CPdfFormRoles.prototype.asc_canFillRole = function (name) {
		return this.CanFillRole(name);
	};
	CPdfFormRoles.prototype.markAsFinal = function (isFinal) {
		if (this.IsFinal() === isFinal)
			return;

		return this._DoAction(function () {
			return this.SetFinal(isFinal);
		}, AscDFH.historydescription_Pdf_MarkAsFinal, true);
	};
	CPdfFormRoles.prototype.isFinal = function () {
		return this.IsFinal();
	};

    window["AscPDF"] = window["AscPDF"] || {};
    window["AscPDF"].CPdfFormRoles = CPdfFormRoles;

	CPdfFormRoles.prototype["asc_getAllRoles"]   = CPdfFormRoles.prototype.asc_getAllRoles;
	CPdfFormRoles.prototype["asc_addRole"]       = CPdfFormRoles.prototype.asc_addRole;
	CPdfFormRoles.prototype["asc_removeRole"]    = CPdfFormRoles.prototype.asc_removeRole;
	CPdfFormRoles.prototype["asc_editRole"]      = CPdfFormRoles.prototype.asc_editRole;
	CPdfFormRoles.prototype["asc_moveUpRole"]    = CPdfFormRoles.prototype.asc_moveUpRole;
	CPdfFormRoles.prototype["asc_moveDownRole"]  = CPdfFormRoles.prototype.asc_moveDownRole;
	CPdfFormRoles.prototype["asc_haveRole"]      = CPdfFormRoles.prototype.asc_haveRole;
	CPdfFormRoles.prototype["asc_getRole"]       = CPdfFormRoles.prototype.asc_getRole;
	CPdfFormRoles.prototype["asc_canFillRole"]   = CPdfFormRoles.prototype.asc_canFillRole;
	CPdfFormRoles.prototype["asc_markAsFinal"]   = CPdfFormRoles.prototype.asc_markAsFinal = CPdfFormRoles.prototype.markAsFinal;
	CPdfFormRoles.prototype["asc_isFinal"]       = CPdfFormRoles.prototype.asc_isFinal = CPdfFormRoles.prototype.isFinal;
})(window);
