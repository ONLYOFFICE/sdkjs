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

(function () {
	AscDFH.changesFactory[AscDFH.historyitem_PropLockerSetId] = AscDFH.CChangesDrawingsString;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_PropLockerSetId] = function (oClass, value) {oClass.objectId = value;};

	function PropLocker(objectId) {
		this.objectId = null;
		this.Lock = new AscCommon.CLock();
		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		AscCommon.g_oTableId.Add(this, this.Id);

		if (typeof objectId === "string") {
			this.setObjectId(objectId);
		}

	}

	PropLocker.prototype.getObjectType = function () {
		return AscDFH.historyitem_type_PropLocker;
	};
	PropLocker.prototype.setObjectId = function (id) {
		AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PropLockerSetId, this.objectId, id));
		this.objectId = id;
	};
	PropLocker.prototype.Get_Id = function () {
		return this.Id;
	};
	PropLocker.prototype.Write_ToBinary2 = function (w) {
		w.WriteLong(AscDFH.historyitem_type_PropLocker);
		w.WriteString2(this.Id);
	};

	PropLocker.prototype.Read_FromBinary2 = function (r) {
		this.Id = r.GetString2();
	};

	PropLocker.prototype.Refresh_RecalcData = function () {};

	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].PropLocker = PropLocker;
})();
