/*
 * (c) Copyright Ascensio System SIA 2010-2020
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
/**
 * User: Ilja.Kirillov
 * Date: 25.03.2020
 * Time: 16:06
 */

var c_oAscDefaultPlaceholderName = {
	Text     : "DefaultPlaceholder_TEXT",
	List     : "DefaultPlaceholder_LIST",
	DateTime : "DefaultPlaceholder_DATE",
	Equation : "DefaultPlaceholder_EQUATION",
	TextForm : "DefaultPlaceholder_TEXTFORM"
};

/**
 * Класс для хранения и работы с дополнительными DocContents
 * @param {CDocument} oLogicDocument
 * @constructor
 */
function CGlossaryDocument(oLogicDocument)
{
	this.Id = oLogicDocument.GetIdCounter().Get_NewId();

	this.Lock = new AscCommon.CLock();

	this.LogicDocument = oLogicDocument;

	this.DocParts = {};

	// Инициализировать нужно сразу, чтобы не было проблем с совместным редактированием
	this.DefaultPlaceholder = {
		Text     : this.private_CreateDefaultPlaceholder(c_oAscDefaultPlaceholderName.Text, AscCommon.translateManager.getValue("Your text here")),
		List     : this.private_CreateDefaultPlaceholder(c_oAscDefaultPlaceholderName.List, AscCommon.translateManager.getValue("Choose an item")),
		DateTime : this.private_CreateDefaultPlaceholder(c_oAscDefaultPlaceholderName.DateTime, AscCommon.translateManager.getValue("Enter a date")),
		Equation : this.private_CreateDefaultPlaceholder(c_oAscDefaultPlaceholderName.Equation, AscCommon.translateManager.getValue("Type equation here")),
		TextForm : this.private_CreateDefaultTextFormPlaceholder()
	};

	// TODO: Реализовать работу нумерации, стилей, сносок, заданных в контентах по-нормальному
	this.Numbering = new CNumbering();
	this.Styles    = new CStyles();
	this.Footnotes = new CFootnotesController(oLogicDocument);
	this.Endnotes  = new CEndnotesController(oLogicDocument);

	oLogicDocument.GetTableId().Add(this, this.Id);
}
/**
 * @returns {CDocument}
 */
CGlossaryDocument.prototype.GetLogicDocument = function()
{
	return this.LogicDocument;
};
/**
 * Получаем идентификатор данного класса
 * @returns {string}
 */
CGlossaryDocument.prototype.GetId = function()
{
	return this.Id;
};
CGlossaryDocument.prototype.Get_Id = function()
{
	return this.Id;
};
/**
 * @return {CNumbering}
 */
CGlossaryDocument.prototype.GetNumbering = function()
{
	return this.Numbering;
};
/**
 * @return {CStyles}
 */
CGlossaryDocument.prototype.GetStyles = function()
{
	return this.Styles;
};
/**
 * @return {CFootnotesController}
 */
CGlossaryDocument.prototype.GetFootnotes = function()
{
	return this.Footnotes;
};
/**
 * @return {CEndnotesController}
 */
CGlossaryDocument.prototype.GetEndnotes = function()
{
	return this.Endnotes;
};
/**
 * Создаем новый контент
 * @param {string} sName
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.CreateDocPart = function(sName)
{
	var oDocPart = new CDocPart(this, sName);

	this.DocParts[oDocPart.GetId()] = oDocPart;
	this.LogicDocument.GetHistory().Add(new CChangesGlossaryAddDocPart(this, oDocPart.GetId()));

	return oDocPart;
};
/**
 * Добавляем новый контент
 * @param {CDocPart} oDocPart
 */
CGlossaryDocument.prototype.AddDocPart = function(oDocPart)
{
	this.DocParts[oDocPart.GetId()] = oDocPart;
	this.LogicDocument.GetHistory().Add(new CChangesGlossaryAddDocPart(this, oDocPart.GetId()));
};
/**
 * Ищем контент по имени
 * @param {string} sName
 * @returns {?CDocPart}
 */
CGlossaryDocument.prototype.GetDocPartByName = function(sName)
{
	for (var sId in this.DocParts)
	{
		var oDocPart = this.DocParts[sId];
		if (sName === oDocPart.GetDocPartName())
			return oDocPart;
	}

	return null;
};
/**
 * Получаем дефолтовый контент для плейсхолдера для обычного текста
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.GetDefaultPlaceholderText = function()
{
	return this.DefaultPlaceholder.Text;
};
/**
 * Получаем дефолтовый контент для плейсхолдера для списка
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.GetDefaultPlaceholderList = function()
{
	return this.DefaultPlaceholder.List;
};
/**
 * Получаем дефолтовый контент для плейсхолдера для поля даты-время
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.GetDefaultPlaceholderDateTime = function()
{
	return this.DefaultPlaceholder.DateTime;
};
/**
 * Получаем дефолтовый контент для плейсхолдера для формулы
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.GetDefaultPlaceholderEquation = function()
{
	return this.DefaultPlaceholder.Equation;
};
/**
 * Получаем дефолтовый контент для плейсхолдера для текстовых форм
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.GetDefaultPlaceholderTextForm = function()
{
	return this.DefaultPlaceholder.TextForm;
};
/**
 * @param sName
 * @param sText
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.private_CreateDefaultPlaceholder = function(sName, sText)
{
	var oDocPart = this.CreateDocPart(sName);

	var oParagraph = oDocPart.GetFirstParagraph();
	var oRun       = new ParaRun();
	oParagraph.AddToContent(0, oRun);
	oRun.AddText(sText);

	oDocPart.SetDocPartBehavior(c_oAscDocPartBehavior.Content);
	oDocPart.SetDocPartCategory("Common", c_oAscDocPartGallery.Placeholder);
	oDocPart.AddDocPartType(c_oAscDocPartType.BBPlcHolder);

	return oDocPart;
};
/**
 * @returns {CDocPart}
 */
CGlossaryDocument.prototype.private_CreateDefaultTextFormPlaceholder = function()
{
	var oDocPart = this.CreateDocPart(c_oAscDefaultPlaceholderName.TextForm);

	var oParagraph = oDocPart.GetFirstParagraph();
	var oRun       = new ParaRun();
	oParagraph.AddToContent(0, oRun);
	oRun.AddToContent(0, new ParaText(0x0020));

	oDocPart.SetDocPartBehavior(c_oAscDocPartBehavior.Content);
	oDocPart.SetDocPartCategory("Common", c_oAscDocPartGallery.Placeholder);
	oDocPart.AddDocPartType(c_oAscDocPartType.BBPlcHolder);

	return oDocPart;
};
/**
 * Проверяем залоченность данного класса для совсместного редактирования
 * @param {number} nCheckType
 */
CGlossaryDocument.prototype.Document_Is_SelectionLocked = function(nCheckType)
{
	// GlossaryDocument пока даем редактировать только одному пользователю за раз
	return this.Lock.Check(this.GetId());
};
/**
 * Получем новое уникальное имя
 * @returns {string}
 */
CGlossaryDocument.prototype.GetNewName = function()
{
	return AscCommon.CreateUUID(true);
};
/**
 * Проверяем, является ли заданный контент контентом по умолчанию
 * @param {CDocPart} oDocPart
 * @returns {boolean}
 */
CGlossaryDocument.prototype.IsDefaultDocPart = function(oDocPart)
{
	return (oDocPart === this.DefaultPlaceholder.Text
		|| oDocPart === this.DefaultPlaceholder.DateTime
		|| oDocPart === this.DefaultPlaceholder.List
		|| oDocPart === this.DefaultPlaceholder.Equation
		|| oDocPart === this.DefaultPlaceholder.List);
};
CGlossaryDocument.prototype.Refresh_RecalcData = function(Data)
{

};


/**
 * Класс, представляющий дополнительное содержимое документа (например, для плейсхолдеров документа)
 * @param {CGlossaryDocument} oGlossary
 * @param {string} sName
 * @constructor
 * @extends {CDocumentContent}
 */
function CDocPart(oGlossary, sName)
{
	var oLogicDocument = oGlossary ? oGlossary.GetLogicDocument() : null;

	this.Glossary = oGlossary;
	this.Pr       = new CDocPartPr(sName);

	// Конструктор базового класса должен быть в конце, т.к. там идет добавление класса по Id
	CDocumentContent.call(this, oLogicDocument, oLogicDocument ? oLogicDocument.GetDrawingDocument() : undefined, 0, 0, 0, 0, true, false, false);
}

CDocPart.prototype = Object.create(CDocumentContent.prototype);
CDocPart.prototype.constructor = CDocPart;

/**
 * Делаем копию данного контента
 * @param {string} [sNewName=""] Опционально можно задать новое имя
 * @returns {CDocPart}
 */
CDocPart.prototype.Copy = function(sNewName)
{
	var oDocPart = new CDocPart(this.Glossary);
	oDocPart.Copy2(this);

	if (sNewName)
		oDocPart.SetDocPartName(sNewName);
	else
		oDocPart.SetDocPartName(this.GetDocPartName());

	if (this.Pr.Category)
		oDocPart.SetDocPartCategory(this.Pr.Category.Name, this.Pr.Category.Gallery);

	if (this.Pr.Behavior)
		oDocPart.SetDocPartBehavior(this.Pr.Behavior);

	// TODO: GUID наверное надо новый генерить
	// if (this.Pr.GUID)
	// 	oDocPart.SetDocPartGUID(this.Pr.GUID);

	if (this.Pr.Description)
		oDocPart.SetDocPartDescription(this.Pr.Description);

	if (this.Pr.Types)
		oDocPart.SetDocPartTypes(this.Pr.Types);

	if (this.Pr.Style)
		oDocPart.SetDocPartStyle(this.Pr.Style);

	this.Glossary.AddDocPart(oDocPart);

	return oDocPart;
};
CDocPart.prototype.Refresh_RecalcData2 = function(nIndex, nCurPage)
{
};
CDocPart.prototype.Write_ToBinary2 = function(oWriter)
{
	oWriter.WriteLong(AscDFH.historyitem_type_DocPart);
	oWriter.WriteString2(this.Glossary ? this.Glossary.GetId() : "");
	this.Pr.WriteToBinary(oWriter);
	CDocumentContent.prototype.Write_ToBinary2.call(this, oWriter);
};
CDocPart.prototype.Read_FromBinary2 = function(oReader)
{
	// historyitem_type_DocPart
	// String : Glossary.Id
	// CDocPartPr
	this.Glossary =  AscCommon.g_oTableId.Get_ById(oReader.GetString2());
	this.Pr.ReadFromBinary(oReader);
	oReader.GetLong(); // Должен вернуть historyitem_type_DocumentContent
	CDocumentContent.prototype.Read_FromBinary2.call(this, oReader);
};
CDocPart.prototype.SetDocPartName = function(sName)
{
	if (this.Pr.Name !== sName)
	{
		History.Add(new CChangesDocPartName(this, this.Pr.Name, sName));
		this.Pr.Name = sName;
	}
};
CDocPart.prototype.GetDocPartName = function()
{
	return this.Pr.Name;
};
CDocPart.prototype.SetDocPartStyle = function(sStyle)
{
	if (this.Pr.Style !== sStyle)
	{
		History.Add(new CChangesDocPartStyle(this, this.Pr.Style, sStyle));
		this.Pr.Style = sStyle;
	}
};
CDocPart.prototype.GetDocPartStyle = function()
{
	return this.Pr.Style;
};
CDocPart.prototype.SetDocPartTypes = function(nTypes)
{
	if (this.Pr.Types !== nTypes)
	{
		History.Add(new CChangesDocPartTypes(this, this.Pr.Types, nTypes));
		this.Pr.Types = nTypes;
	}
};
/**
 * @param {c_oAscDocPartType} nType
 */
CDocPart.prototype.AddDocPartType = function(nType)
{
	this.SetDocPartTypes(this.Pr.Types | nType);
};
/**
 * @param {c_oAscDocPartType} nType
 * @returns {boolean}
 */
CDocPart.prototype.CheckDocPartType = function(nType)
{
	if (this.Pr.Types & c_oAscDocPartType.All)
		return true;

	return !!(this.Pr.Types & nType);
};
CDocPart.prototype.SetDocPartDescription = function(sDescription)
{
	if (this.Pr.Description !== sDescription)
	{
		History.Add(new CChangesDocPartDescription(this, this.Pr.Description, sDescription));
		this.Pr.Description = sDescription;
	}
};
CDocPart.prototype.GetDocPartDescription = function()
{
	return this.Pr.Description;
};
CDocPart.prototype.SetDocPartGUID = function(sGUID)
{
	if (this.Pr.GUID !== sGUID)
	{
		History.Add(new CChangesDocPartGUID(this, this.Pr.GUID, sGUID));
		this.Pr.GUID = sGUID;
	}
};
CDocPart.prototype.GetDocPartGUID = function()
{
	return this.Pr.GUID;
};
CDocPart.prototype.SetDocPartCategory = function(sName, nGallery)
{
	var oNewCategory = undefined;
	if (undefined !== sName)
		oNewCategory = new CDocPartCategory(sName, nGallery);

	if ((!this.Pr.Category && oNewCategory)
		|| (this.Pr.Category && !this.Pr.Category.IsEqual(oNewCategory)))
	{
		History.Add(new CChangesDocPartCategory(this, this.Pr.Category, oNewCategory));
		this.Pr.Category = oNewCategory;
	}
};
CDocPart.prototype.GetDocPartCategory = function()
{
	return this.Pr.Category;
};
CDocPart.prototype.SetDocPartBehavior = function(nBehavior)
{
	if (this.Pr.Behavior !== nBehavior)
	{
		History.Add(new CChangesDocPartBehavior(this, this.Pr.Behavior, nBehavior));
		this.Pr.Behavior = nBehavior;
	}
};
/**
 * @param {c_oAscDocPartBehavior} nType
 */
CDocPart.prototype.AddDocPartBehavior = function(nType)
{
	this.SetDocPartBehavior(this.Pr.Behavior | nType);
};
/**
 * @param {c_oAscDocPartBehavior} nType
 * @returns {boolean}
 */
CDocPart.prototype.CheckDocPartBehavior = function(nType)
{
	return !!(this.Pr.Behavior & nType);
};

/** @enum {number} */
var c_oAscDocPartType = {
	Undefined   : 0x0000,
	All         : 0x0001,
	AutoExp     : 0x0002,
	BBPlcHolder : 0x0004,
	FormFld     : 0x0008,
	None        : 0x0010,
	Normal      : 0x0020,
	Speller     : 0x0040,
	Toolbar     : 0x0080
};

/** @enum {number} */
var c_oAscDocPartBehavior = {
	Undefined : 0x00,
	Content   : 0x01,
	P         : 0x02,
	Pg        : 0x04
};

/**
 * Настройки для дополнительного содержимого документа
 * @param {string} sName
 * @constructor
 */
function CDocPartPr(sName)
{
	this.Name        = sName ? sName : undefined;
	this.Style       = undefined;
	this.Types       = c_oAscDocPartType.Undefined;
	this.Description = undefined;
	this.GUID        = undefined;
	this.Category    = undefined;
	this.Behavior    = c_oAscDocPartBehavior.Undefined;
}

CDocPartPr.prototype.WriteToBinary = function(oWriter)
{
	// String           : Name
	// Long             : Types
	// Long             : Behavior
	// Flags            : Флаг, означающий заданы ли следующие поля
	// String           : Style
	// String           : Description
	// String           : GUID
	// CDocPartCategory : Category

	oWriter.WriteString2(this.Name);
	oWriter.WriteLong(this.Types);
	oWriter.WriteLong(this.Behavior);

	var nStartPos = oWriter.GetCurPosition();
	oWriter.Skip(4);
	var nFlags = 0;

	if (undefined !== this.Style)
	{
		nFlags |= 1;
		oWriter.WriteString2(this.Style);
	}

	if (undefined !== this.Description)
	{
		nFlags |= 2;
		oWriter.WriteString2(this.Description);
	}

	if (undefined !== this.GUID)
	{
		nFlags |= 4;
		oWriter.WriteString2(this.GUID);
	}

	if (undefined !== this.Category)
	{
		nFlags |= 8;
		this.Category.WriteToBinary(oWriter);
	}

	var nEndPos = oWriter.GetCurPosition();
	oWriter.Seek(nStartPos);
	oWriter.WriteLong(nFlags);
	oWriter.Seek(nEndPos);
};
CDocPartPr.prototype.ReadFromBinary = function(oReader)
{
	// String           : Name
	// Long             : Types
	// Long             : Behavior
	// Flags            : Флаг, означающий заданы ли следующие поля
	// String           : Style
	// String           : Description
	// String           : GUID
	// CDocPartCategory : Category

	this.Name     = oReader.GetString2();
	this.Types    = oReader.GetLong();
	this.Behavior = oReader.GetLong();

	var nFlags = oReader.GetLong();
	if (nFlags & 1)
		this.Style = oReader.GetString2();

	if (nFlags & 2)
		this.Description = oReader.GetString2();

	if (nFlags & 4)
		this.GUID = oReader.GetString2();

	if (nFlags & 8)
	{
		this.Category = new CDocPartCategory();
		this.Category.ReadFromBinary(oReader);
	}
};

/** @enum {number} */
var c_oAscDocPartGallery = {
	Any               : 0,
	AutoTxt           : 1,
	Bib               : 2,
	CoverPg           : 3,
	CustAutoTxt       : 4,
	CustBib           : 5,
	CustCoverPg       : 6,
	CustEq            : 7,
	CustFtrs          : 8,
	CustHdrs          : 9,
	Custom1           : 10,
	Custom2           : 11,
	Custom3           : 12,
	Custom4           : 13,
	Custom5           : 14,
	CustPgNum         : 15,
	CustPgNumB        : 16,
	CustPgNumMargins  : 17,
	CustPgNumT        : 18,
	CustQuickParts    : 19,
	CustTblOfContents : 20,
	CustTbls          : 21,
	CustTxtBox        : 22,
	CustWatermarks    : 23,
	Default           : 24,
	DocParts          : 25,
	Eq                : 26,
	Ftrs              : 27,
	Hdrs              : 28,
	PgNum             : 29,
	PgNumB            : 30,
	PgNumMargins      : 31,
	PgNumT            : 32,
	Placeholder       : 33,
	TblOfContents     : 34,
	Tbls              : 35,
	TxtBox            : 36,
	Watermarks        : 37
};

/**
 * Класс для определения категории заданного специального содержимого
 * @param {string} sName
 * @param {number} nGallery
 * @constructor
 */
function CDocPartCategory(sName, nGallery)
{
	this.Name    = undefined !== sName ? sName : "";
	this.Gallery = undefined !== nGallery ? nGallery : c_oAscDocPartGallery.Default;
}
CDocPartCategory.prototype.WriteToBinary = function(oWriter)
{
	// String : Name
	// Long   : Gallery
	oWriter.WriteString2(this.Name);
	oWriter.WriteLong(this.Gallery);
};
CDocPartCategory.prototype.ReadFromBinary = function(oReader)
{
	// String : Name
	// Long   : Gallery
	this.Name    = oReader.GetString2();
	this.Gallery = oReader.GetLong();
};
CDocPartCategory.prototype.Write_ToBinary = function(oWriter)
{
	return this.WriteToBinary(oWriter);
};
CDocPartCategory.prototype.Read_FromBinary = function(oReader)
{
	return this.ReadFromBinary(oReader);
};
/**
 * Проверяем на совпадение
 * @param {CDocPartCategory} oCategory
 * @returns {boolean}
 */
CDocPartCategory.prototype.IsEqual = function(oCategory)
{
	if (!oCategory)
		return false;

	return (this.Name === oCategory.Name && this.Gallery === oCategory.Gallery);
};

//------------------------------------------------------------export---------------------------------------------------
window['AscCommon'] = window['AscCommon'] || {};
window['AscCommon'].CDocPart = CDocPart;

window["Asc"] = window["Asc"] || {};

var prot;
prot = window["Asc"]["c_oAscDocPartType"] = c_oAscDocPartType;

prot["Undefined"]   = c_oAscDocPartType.Undefined;
prot["All"]         = c_oAscDocPartType.All;
prot["AutoExp"]     = c_oAscDocPartType.AutoExp;
prot["BBPlcHolder"] = c_oAscDocPartType.BBPlcHolder;
prot["FormFld"]     = c_oAscDocPartType.FormFld;
prot["None"]        = c_oAscDocPartType.None;
prot["Normal"]      = c_oAscDocPartType.Normal;
prot["Speller"]     = c_oAscDocPartType.Speller;
prot["Toolbar"]     = c_oAscDocPartType.Toolbar;


prot = window["Asc"]["c_oAscDocPartGallery"] = c_oAscDocPartGallery;

prot["Any"]               = c_oAscDocPartGallery.Any;
prot["AutoTxt"]           = c_oAscDocPartGallery.AutoTxt;
prot["Bib"]               = c_oAscDocPartGallery.Bib;
prot["CoverPg"]           = c_oAscDocPartGallery.CoverPg;
prot["CustAutoTxt"]       = c_oAscDocPartGallery.CustAutoTxt;
prot["CustBib"]           = c_oAscDocPartGallery.CustBib;
prot["CustCoverPg"]       = c_oAscDocPartGallery.CustCoverPg;
prot["CustEq"]            = c_oAscDocPartGallery.CustEq;
prot["CustFtrs"]          = c_oAscDocPartGallery.CustFtrs;
prot["CustHdrs"]          = c_oAscDocPartGallery.CustHdrs;
prot["Custom1"]           = c_oAscDocPartGallery.Custom1;
prot["Custom2"]           = c_oAscDocPartGallery.Custom2;
prot["Custom3"]           = c_oAscDocPartGallery.Custom3;
prot["Custom4"]           = c_oAscDocPartGallery.Custom4;
prot["Custom5"]           = c_oAscDocPartGallery.Custom5;
prot["CustPgNum"]         = c_oAscDocPartGallery.CustPgNum;
prot["CustPgNumB"]        = c_oAscDocPartGallery.CustPgNumB;
prot["CustPgNumMargins"]  = c_oAscDocPartGallery.CustPgNumMargins;
prot["CustPgNumT"]        = c_oAscDocPartGallery.CustPgNumT;
prot["CustQuickParts"]    = c_oAscDocPartGallery.CustQuickParts;
prot["CustTblOfContents"] = c_oAscDocPartGallery.CustTblOfContents;
prot["CustTbls"]          = c_oAscDocPartGallery.CustTbls;
prot["CustTxtBox"]        = c_oAscDocPartGallery.CustTxtBox;
prot["CustWatermarks"]    = c_oAscDocPartGallery.CustWatermarks;
prot["Default"]           = c_oAscDocPartGallery.Default;
prot["DocParts"]          = c_oAscDocPartGallery.DocParts;
prot["Eq"]                = c_oAscDocPartGallery.Eq;
prot["Ftrs"]              = c_oAscDocPartGallery.Ftrs;
prot["Hdrs"]              = c_oAscDocPartGallery.Hdrs;
prot["PgNum"]             = c_oAscDocPartGallery.PgNum;
prot["PgNumB"]            = c_oAscDocPartGallery.PgNumB;
prot["PgNumMargins"]      = c_oAscDocPartGallery.PgNumMargins;
prot["PgNumT"]            = c_oAscDocPartGallery.PgNumT;
prot["Placeholder"]       = c_oAscDocPartGallery.Placeholder;
prot["TblOfContents"]     = c_oAscDocPartGallery.TblOfContents;
prot["Tbls"]              = c_oAscDocPartGallery.Tbls;
prot["TxtBox"]            = c_oAscDocPartGallery.TxtBox;
prot["Watermarks"]        = c_oAscDocPartGallery.Watermarks;

prot = window["Asc"]["c_oAscDocPartBehavior"] = c_oAscDocPartBehavior;

prot["Undefined"] = c_oAscDocPartBehavior.Undefined;
prot["Content"]   = c_oAscDocPartBehavior.Content;
prot["P"]         = c_oAscDocPartBehavior.P;
prot["Pg"]        = c_oAscDocPartBehavior.Pg;
