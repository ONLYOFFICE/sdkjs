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

var para_Unknown                   = -1; //
var para_RunBase                   = 0x0000; // Базовый элемент, он не должен использоваться как самостоятельный объект
var para_Text                      = 0x0001; // Текст
var para_Space                     = 0x0002; // Пробелы
var para_TextPr                    = 0x0003; // Свойства текста
var para_End                       = 0x0004; // Конец параграфа
var para_NewLine                   = 0x0010; // Новая строка
var para_NewLineRendered           = 0x0011; // Рассчитанный перенос строки
var para_InlineBreak               = 0x0012; // Перенос внутри строки (для обтекания)
var para_PageBreakRendered         = 0x0013; // Рассчитанный перенос страницы
var para_Numbering                 = 0x0014; // Элемент, обозначающий нумерацию для списков
var para_Tab                       = 0x0015; // Табуляция
var para_Drawing                   = 0x0016; // Графика (картинки, автофигуры, диаграммы, графики)
var para_PageNum                   = 0x0017; // Нумерация страницы
var para_FlowObjectAnchor          = 0x0018; // Привязка для "плавающих" объектов
var para_HyperlinkStart            = 0x0019; // Начало гиперссылки
var para_HyperlinkEnd              = 0x0020; // Конец гиперссылки
var para_CollaborativeChangesStart = 0x0021; // Начало изменений другого редактора
var para_CollaborativeChangesEnd   = 0x0022; // Конец изменений другого редактора
var para_CommentStart              = 0x0023; // Начало комментария
var para_CommentEnd                = 0x0024; // Начало комментария
var para_PresentationNumbering     = 0x0025; // Элемент, обозначающий нумерацию для списков в презентациях
var para_Math                      = 0x0026; // Формула
var para_Run                       = 0x0027; // Текстовый элемент
var para_Sym                       = 0x0028; // Символ
var para_Comment                   = 0x0029; // Метка начала или конца комментария
var para_Hyperlink                 = 0x0030; // Гиперссылка
var para_Math_Run                  = 0x0031; // Run в формуле
var para_Math_Placeholder          = 0x0032; // Плейсхолдер
var para_Math_Composition          = 0x0033; // Математический объект (дробь, степень и т.п.)
var para_Math_Text                 = 0x0034; // Текст в формуле
var para_Math_Ampersand            = 0x0035; // &
var para_Field                     = 0x0036; // Поле
var para_Math_BreakOperator        = 0x0037; // break operator в формуле
var para_Math_Content              = 0x0038; // math content
var para_FootnoteReference         = 0x0039; // Ссылка на сноску
var para_FootnoteRef               = 0x0040; // Номер сноски (должен быть только внутри сноски)
var para_Separator                 = 0x0041; // Разделить, который используется для сносок
var para_ContinuationSeparator     = 0x0042; // Большой разделитель, который используется для сносок
var para_PageCount                 = 0x0043; // Количество страниц
var para_InlineLevelSdt            = 0x0044; // Внутристроковый контейнер
var para_FieldChar                 = 0x0045;
var para_InstrText                 = 0x0046;
var para_Bookmark                  = 0x0047;
var para_RevisionMove              = 0x0048;
var para_EndnoteReference          = 0x0049; // Ссылка на сноску
var para_EndnoteRef                = 0x004a; // Номер сноски (должен быть только внутри сноски)
var para_PermStart                 = 0x004b;
var para_PermEnd                   = 0x004c;

(function(window)
{
	function ReadRunElementFromBinary(oReader)
	{
		let oElement = null;
		switch (oReader.GetLong())
		{
			case para_TextPr:
			case para_Drawing:
			case para_HyperlinkStart:
			case para_InlineLevelSdt:
			case para_Bookmark:
			{
				var ElementId = oReader.GetString2();
				oElement       = g_oTableId.Get_ById(ElementId);
				return oElement;
			}
			case para_RunBase               : oElement = new AscWord.CRunElementBase(); break;
			case para_Text                  : oElement = new AscWord.CRunText(); break;
			case para_Space                 : oElement = new AscWord.CRunSpace(); break;
			case para_End                   : oElement = new AscWord.CRunParagraphMark(); break;
			case para_NewLine               : oElement = new AscWord.CRunBreak(); break;
			case para_Numbering             : oElement = new ParaNumbering(); break;
			case para_Tab                   : oElement = new AscWord.CRunTab(); break;
			case para_PageNum               : oElement = new AscWord.CRunPageNum(); break;
			case para_Math_Placeholder      : oElement = new CMathText(); break;
			case para_Math_Text             : oElement = new CMathText(); break;
			case para_Math_BreakOperator    : oElement = new CMathText(); break;
			case para_Math_Ampersand        : oElement = new CMathAmp(); break;
			case para_PresentationNumbering : oElement = new ParaPresentationNumbering(); break;
			case para_FootnoteReference     : oElement = new AscWord.CRunFootnoteReference(); break;
			case para_FootnoteRef           : oElement = new AscWord.CRunFootnoteRef(); break;
			case para_Separator             : oElement = new AscWord.CRunSeparator(); break;
			case para_ContinuationSeparator : oElement = new AscWord.CRunContinuationSeparator(); break;
			case para_PageCount             : oElement = new AscWord.CRunPagesCount(); break;
			case para_FieldChar             : oElement = new ParaFieldChar(); break;
			case para_InstrText             : oElement = new ParaInstrText(); break;
			case para_RevisionMove          : oElement = new AscWord.CRunRevisionMove(); break;
			case para_EndnoteReference      : oElement = new AscWord.CRunEndnoteReference(); break;
			case para_EndnoteRef            : oElement = new AscWord.CRunEndnoteRef(); break;
		}

		if (oElement)
			oElement.Read_FromBinary(oReader);

		return oElement;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].ReadRunElementFromBinary = ReadRunElementFromBinary;

})(window);
