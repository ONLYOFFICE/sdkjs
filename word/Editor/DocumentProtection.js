/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

var ECryptAlgoritmName = {
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
var ECryptAlgClass = {
	Custom: 0,
	Hash: 1
};
var ECryptAlgType = {
	Custom: 0,
	TypeAny: 1
};
var ECryptProv = {
	Custom: 0,
	RsaAES: 1,
	RsaFull: 2
};

function CDocProtect() {
	this.Id = AscCommon.g_oIdCounter.Get_NewId();

	this.algorithmName = null;
	this.edit = null;
	this.enforcement = null;
	this.formatting = null;
	this.hashValue = null;
	this.saltValue = null;
	this.spinCount = null;

	this.algIdExt = null;
	this.algIdExtSource = null;
	this.cryptAlgorithmClass = null;
	this.cryptAlgorithmSid = null;
	this.cryptAlgorithmType = null;
	this.cryptProvider = null;
	this.cryptProviderType = null;
	this.cryptProviderTypeExt = null;
	this.cryptProviderTypeExtSource = null;

	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	AscCommon.g_oTableId.Add(this, this.Id);

	this.Lock = new AscCommon.CLock();
	this.lockType = AscCommon.c_oAscLockTypes.kLockTypeNone;
	this.NeedUpdateByUser = null;

	this.temporaryPassword = null;
}
CDocProtect.prototype.Get_Id = function () {
	return this.Id;
};
CDocProtect.prototype.SetNeedUpdate = function(userId)
{
	this.NeedUpdateByUser = userId;
};
CDocProtect.prototype.GetNeedUpdate = function()
{
	return this.NeedUpdateByUser;
};
CDocProtect.prototype.isOnlyView = function () {
	return this.edit === Asc.c_oAscEDocProtect.ReadOnly;
};
CDocProtect.prototype.getEnforcement = function () {
	return this.enforcement;
};
CDocProtect.prototype.getRestrictionType = function () {
	var res = null;
	switch (this.edit) {
		case Asc.c_oAscEDocProtect.Comments:
			res = Asc.c_oAscRestrictionType.OnlyComments;
			break;
		case Asc.c_oAscEDocProtect.Forms:
			res = Asc.c_oAscRestrictionType.OnlyForms;
			break;
		case Asc.c_oAscEDocProtect.ReadOnly:
			res = Asc.c_oAscRestrictionType.View;
			break;
		case Asc.c_oAscEDocProtect.TrackedChanges:
			break;
	}
	return res;
};
CDocProtect.prototype.generateHashParams = function () {
	var params = AscCommon.generateHashParams();

	this.saltValue = params.saltValue;
	this.spinCount = params.spinCount;
	//this.algorithmName = params.algorithmName;
};
CDocProtect.prototype.generateHashParams = function () {
	var params = AscCommon.generateHashParams();

	this.saltValue = params.saltValue;
	this.spinCount = params.spinCount;
	//this.algorithmName = params.algorithmName;
};
CDocProtect.prototype.getAlgorithmNameForCheck = function () {
	if (this.algorithmName) {
		return AscCommon.fromModelAlgorithmName(this.algorithmName);
	} else if (this.cryptAlgorithmSid) {
		return AscCommon.fromModelCryptAlgorithmSid(this.cryptAlgorithmSid);
	}
	return null;
};
CDocProtect.prototype.isPassword = function () {
	return this.algorithmName != null || this.cryptAlgorithmSid != null;
};
CDocProtect.prototype.setProps = function (oProps) {
	let doc = editor && editor.private_GetLogicDocument && editor.private_GetLogicDocument();
	let userId = doc && doc.GetUserId && doc.GetUserId();
	History.Add(new CChangesDocumentProtection(this, this, oProps, userId));
	this.setFromInterface(oProps);
};
CDocProtect.prototype.setFromInterface = function (oProps) {
	this.edit = oProps.edit;
	this.saltValue = oProps.saltValue;
	this.spinCount = oProps.spinCount;
	this.cryptAlgorithmSid = oProps.cryptAlgorithmSid;
	this.hashValue = oProps.hashValue;
	this.cryptProviderType = oProps.cryptProviderType;

	this.enforcement = oProps.enforcement;
};
CDocProtect.prototype.Refresh_RecalcData = function () {
};
CDocProtect.prototype.Copy = function () {
	return AscFormat.ExecuteNoHistory(function () {
		var oDocProtect = new CDocProtect();
		oDocProtect.algorithmName = this.algorithmName;
		oDocProtect.edit = this.edit;
		oDocProtect.enforcement = this.enforcement;
		oDocProtect.formatting = this.formatting;
		oDocProtect.hashValue = this.hashValue;
		oDocProtect.saltValue = this.saltValue;
		oDocProtect.spinCount = this.spinCount;
		oDocProtect.algIdExt = this.algIdExt;
		oDocProtect.algIdExtSource = this.algIdExtSource;
		oDocProtect.cryptAlgorithmClass = this.cryptAlgorithmClass;
		oDocProtect.cryptAlgorithmSid = this.cryptAlgorithmSid;
		oDocProtect.cryptAlgorithmType = this.cryptAlgorithmType;
		oDocProtect.cryptProvider = this.cryptProvider;
		oDocProtect.cryptProviderType = this.cryptProviderType;
		oDocProtect.cryptProviderTypeExt = this.cryptProviderTypeExt;
		oDocProtect.cryptProviderTypeExtSource = this.cryptProviderTypeExtSource;
		return oDocProtect;
	}, this, []);
};
/*CDocProtect.prototype.Write_ToBinary2 = function(Writer)
{
	Writer.WriteLong(AscDFH.historyitem_type_DocumentProtection);
	Writer.WriteString2("" + this.Id);
};
CDocProtect.prototype.Read_FromBinary2 = function(Reader)
{
	this.Id = Reader.GetString2();
};*/
CDocProtect.prototype.asc_getIsPassword = function()
{
	return this.enforcement ? this.isPassword() : null
};
CDocProtect.prototype.asc_getEditType = function()
{
	return this.enforcement !== false ? this.edit : Asc.c_oAscEDocProtect.None;
};
CDocProtect.prototype.asc_setPassword = function(val)
{
	this.temporaryPassword = val;
};
CDocProtect.prototype.asc_setEditType = function(val)
{
	this.edit = val;
	if (this.edit != null) {
		this.enforcement = true;
	}
};


function CWriteProtection() {
	this.algorithmName = null;
	this.recommended = null;
	this.hashValue = null;
	this.saltValue = null;
	this.spinCount = null;

	this.algIdExt = null;
	this.algIdExtSource = null;
	this.cryptAlgorithmClass = null;
	this.cryptAlgorithmSid = null;
	this.cryptAlgorithmType = null;
	this.cryptProvider = null;
	this.cryptProviderType = null;
	this.cryptProviderTypeExt = null;
	this.cryptProviderTypeExtSource = null;
}

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'].CDocProtect = CDocProtect;
prot = CDocProtect.prototype;
prot["asc_getIsPassword"] = prot.asc_getIsPassword;
prot["asc_getEditType"] = prot.asc_getEditType;
prot["asc_setEditType"] = prot.asc_setEditType;
prot["asc_setPassword"] = prot.asc_setPassword;

window['AscCommonWord'].ECryptAlgType = ECryptAlgType;
