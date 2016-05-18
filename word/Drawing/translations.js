/*
 *
 * (c) Copyright Ascensio System Limited 2010-2016
 *
 * This program is freeware. You can redistribute it and/or modify it under the terms of the GNU 
 * General Public License (GPL) version 3 as published by the Free Software Foundation (https://www.gnu.org/copyleft/gpl.html). 
 * In accordance with Section 7(a) of the GNU GPL its Section 15 shall be amended to the effect that 
 * Ascensio System SIA expressly excludes the warranty of non-infringement of any third-party rights.
 *
 * THIS PROGRAM IS DISTRIBUTED WITHOUT ANY WARRANTY; WITHOUT EVEN THE IMPLIED WARRANTY OF MERCHANTABILITY OR
 * FITNESS FOR A PARTICULAR PURPOSE. For more details, see GNU GPL at https://www.gnu.org/copyleft/gpl.html
 *
 * You can contact Ascensio System SIA by email at sales@onlyoffice.com
 *
 * The interactive user interfaces in modified source and object code versions of ONLYOFFICE must display 
 * Appropriate Legal Notices, as required under Section 5 of the GNU GPL version 3.
 *
 * Pursuant to Section 7  3(b) of the GNU GPL you must retain the original ONLYOFFICE logo which contains 
 * relevant author attributions when distributing the software. If the display of the logo in its graphic 
 * form is not reasonably feasible for technical reasons, you must include the words "Powered by ONLYOFFICE" 
 * in every copy of the program you distribute. 
 * Pursuant to Section 7  3(e) we decline to grant you any rights under trademark law for use of our trademarks.
 *
*/
"use strict";

(function(window, undefined){

var translations_map = {};
var styles_index_map = {};
styles_index_map["Normal"]          = 0;
styles_index_map["No list"]         = 1;
styles_index_map["Heading 1"]       = 2;
styles_index_map["Heading 2"]       = 3;
styles_index_map["Heading 3"]       = 4;
styles_index_map["Heading 4"]       = 5;
styles_index_map["Heading 5"]       = 6;
styles_index_map["Heading 6"]       = 7;
styles_index_map["Heading 7"]       = 8;
styles_index_map["Heading 8"]       = 9;
styles_index_map["Heading 9"]       = 10;
styles_index_map["Paragraph List"]  = 11;
styles_index_map["Normal Table"]    = 12;
styles_index_map["No Spacing"]      = 13;


// en
translations_map["en"] = {};
translations_map["en"].DefaultStyles =
[
    "Normal",
    "No List",
    "Heading 1",
    "Heading 2",
    "Heading 3",
    "Heading 4",
    "Heading 5",
    "Heading 6",
    "Heading 7",
    "Heading 8",
    "Heading 9",
    "Paragraph List",
    "Normal Table",
    "No Spacing"
];
translations_map["en"].StylesText = "AaBbCcDdEeFf";

// ru
translations_map["ru"] = {};
translations_map["ru"].DefaultStyles =
[
    "Обычный",
    "Нет списка",
    "Заголовок 1",
    "Заголовок 2",
    "Заголовок 3",
    "Заголовок 4",
    "Заголовок 5",
    "Заголовок 6",
    "Заголовок 7",
    "Заголовок 8",
    "Заголовок 9",
    "Абзац списка",
    "Обычная таблица",
    "Без интервала"
];
translations_map["ru"].StylesText = "АаБбВвГгДдЕе";

// de
translations_map["de"] = {};
translations_map["de"].DefaultStyles =
[
    "Standard",
    "Keine Liste",
    "Überschrift 1",
    "Überschrift 2",
    "Überschrift 3",
    "Überschrift 4",
    "Überschrift 5",
    "Überschrift 6",
    "Überschrift 7",
    "Überschrift 8",
    "Überschrift 9",
    "Listenabsatz",
    "Normale Tabelle",
    "Kein Leerraum"
];
translations_map["de"].StylesText = "AaBbCcDdEeFf";

// fr
translations_map["fr"] = {};
translations_map["fr"].DefaultStyles =
[
    "Normal",
    "Sans liste",
    "Titre 1",
    "Titre 2",
    "Titre 3",
    "Titre 4",
    "Titre 5",
    "Titre 6",
    "Titre 7",
    "Titre 8",
    "Titre 9",
    "Paragraphe de liste",
    "Tableau normal",
    "Sans interligne"
];
translations_map["fr"].StylesText = "AaBbCcDdEeFf";

// it
translations_map["it"] = {};
translations_map["it"].DefaultStyles =
[
    "Normale",
    "Nessun elenco",
    "Titolo 1",
    "Titolo 2",
    "Titolo 3",
    "Titolo 4",
    "Titolo 5",
    "Titolo 6",
    "Titolo 7",
    "Titolo 8",
    "Titolo 9",
    "Paragrafo elenco",
    "Tabella normale",
    "Nessuna spaziatura"
];
translations_map["it"].StylesText = "AaBbCcDdEeFf";

// es
translations_map["es"] = {};
translations_map["es"].DefaultStyles =
[
    "Normal",
    "No lista",
    "Título 1",
    "Título 2",
    "Título 3",
    "Título 4",
    "Título 5",
    "Título 6",
    "Título 7",
    "Título 8",
    "Título 9",
    "Lista de párrafos",
    "Tabla",
    "Sin espaciado"
];
translations_map["es"].StylesText = "AaBbCcDdEeFf";

// lv
translations_map["lv"] = {};
translations_map["lv"].DefaultStyles =
[
    "Normāls",
    "Nav saraksts",
    "Virsraksts 1",
    "Virsraksts 2",
    "Virsraksts 3",
    "Virsraksts 4",
    "Virsraksts 5",
    "Virsraksts 6",
    "Virsraksts 7",
    "Virsraksts 8",
    "Virsraksts 9",
    "Panta saraksts",
    "Normāla tabula",
    "Bez atstarpes"
];
translations_map["lv"].StylesText = "AaBbCcDdEeFf";

    //--------------------------------------------------------export----------------------------------------------------
    window['AscCommonWord'] = window['AscCommonWord'] || {};
    window['AscCommonWord'].translations_map = translations_map;
})(window);
