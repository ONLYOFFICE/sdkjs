/*
 * (c) Copyright Ascensio System SIA 2010-2019
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

//var window = {};
(function (undefined) {

    var sIdentifier = '(\u0009\|\u0000A\|\u0000D\|[\u0020-\uD7FF]\|[\uE000-\uFFFD]\|[\u10000-\u10FFFF])+';
    var sComparison = '<=\|<>\|>=\|=\|<\|>';
    var sOperator =  "<>\|<=\|>=\|>\|<\|-\|\\^\|\\*\|/\|\\\%|\\+\|=";
    var sColumnName = '([A-Z]){1,2}';
    var sDecimalDigit = '[0-9]';
    var sRowName = sDecimalDigit + '+';
    var sColon = ':';
    var sComma = ',';
    var sFullStop = '\\.';
    var sWholeNumberPart = '([0-9]+)((,[0-9]{3})+[0-9]+)*';
    var sFractionalPart = sDecimalDigit + '+';

    var sConstant =  sWholeNumberPart + sFullStop + sFractionalPart + '\|' + '(' + sWholeNumberPart + '(' + sFullStop +')*)' + '\|(' + sFullStop + sFractionalPart + ')';
    var sCellName = sColumnName + sRowName;
    var sCellCellRange = sCellName + sColon + sCellName;
    var sRowRange = sRowName + sColon + sRowName;
    var sColumnRange = sColumnName + sColon + sColumnName;
    var sCellRange = '(' + sCellCellRange + ')\|(' + sRowRange + ')\|(' + sColumnRange + ')';
    var sCellReference = '(' + sCellRange + ')\|(' + sCellName + ')';
    var sBookmark = "(?:[A-Z\xAA\xBA\xC0-\xD6\xD8-\xDE\u0100\u0102\u0104\u0106\u0108\u010A\u010C\u010E\u0110\u0112\u0114\u0116\u0118\u011A\u011C\u011E\u0120\u0122\u0124\u0126\u0128\u012A\u012C\u012E\u0130\u0132\u0134\u0136\u0139\u013B\u013D\u013F\u0141\u0143\u0145\u0147\u014A\u014C\u014E\u0150\u0152\u0154\u0156\u0158\u015A\u015C\u015E\u0160\u0162\u0164\u0166\u0168\u016A\u016C\u016E\u0170\u0172\u0174\u0176\u0178\u0179\u017B\u017D\u0181\u0182\u0184\u0186\u0187\u0189-\u018B\u018E-\u0191\u0193\u0194\u0196-\u0198\u019C\u019D\u019F\u01A0\u01A2\u01A4\u01A6\u01A7\u01A9\u01AC\u01AE\u01AF\u01B1-\u01B3\u01B5\u01B7\u01B8\u01BC\u01C4\u01C5\u01C7\u01C8\u01CA\u01CB\u01CD\u01CF\u01D1\u01D3\u01D5\u01D7\u01D9\u01DB\u01DE\u01E0\u01E2\u01E4\u01E6\u01E8\u01EA\u01EC\u01EE\u01F1\u01F2\u01F4\u01F6-\u01F8\u01FA\u01FC\u01FE\u0200\u0202\u0204\u0206\u0208\u020A\u020C\u020E\u0210\u0212\u0214\u0216\u0218\u021A\u021C\u021E\u0220\u0222\u0224\u0226\u0228\u022A\u022C\u022E\u0230\u0232\u023A\u023B\u023D\u023E\u0241\u0243-\u0246\u0248\u024A\u024C\u024E\u02B0-\u02C1\u02C6-\u02D1\u02E0-\u02E4\u02EC\u02EE\u0370\u0372\u0376\u0386\u0388-\u038A\u038C\u038E\u038F\u0391-\u03A1\u03A3-\u03AB\u03CF\u03D2-\u03D4\u03D8\u03DA\u03DC\u03DE\u03E0\u03E2\u03E4\u03E6\u03E8\u03EA\u03EC\u03EE\u03F4\u03F7\u03F9\u03FA\u03FD-\u042F\u0460\u0462\u0464\u0466\u0468\u046A\u046C\u046E\u0470\u0472\u0474\u0476\u0478\u047A\u047C\u047E\u0480\u048A\u048C\u048E\u0490\u0492\u0494\u0496\u0498\u049A\u049C\u049E\u04A0\u04A2\u04A4\u04A6\u04A8\u04AA\u04AC\u04AE\u04B0\u04B2\u04B4\u04B6\u04B8\u04BA\u04BC\u04BE\u04C0\u04C1\u04C3\u04C5\u04C7\u04C9\u04CB\u04CD\u04D0\u04D2\u04D4\u04D6\u04D8\u04DA\u04DC\u04DE\u04E0\u04E2\u04E4\u04E6\u04E8\u04EA\u04EC\u04EE\u04F0\u04F2\u04F4\u04F6\u04F8\u04FA\u04FC\u04FE\u0500\u0502\u0504\u0506\u0508\u050A\u050C\u050E\u0510\u0512\u0514\u0516\u0518\u051A\u051C\u051E\u0520\u0522\u0531-\u0556\u05D0-\u05EA\u05F0-\u05F2\u0621-\u064A\u066E\u066F\u0671-\u06D3\u06D5\u06E5\u06E6\u06EE\u06EF\u06FA-\u06FC\u06FF\u0710\u0712-\u072F\u074D-\u07A5\u07B1\u07CA-\u07EA\u07F4\u07F5\u07FA\u0904-\u0939\u093D\u0950\u0958-\u0961\u0971\u0972\u097B-\u097F\u0985-\u098C\u098F\u0990\u0993-\u09A8\u09AA-\u09B0\u09B2\u09B6-\u09B9\u09BD\u09CE\u09DC\u09DD\u09DF-\u09E1\u09F0\u09F1\u0A05-\u0A0A\u0A0F\u0A10\u0A13-\u0A28\u0A2A-\u0A30\u0A32\u0A33\u0A35\u0A36\u0A38\u0A39\u0A59-\u0A5C\u0A5E\u0A72-\u0A74\u0A85-\u0A8D\u0A8F-\u0A91\u0A93-\u0AA8\u0AAA-\u0AB0\u0AB2\u0AB3\u0AB5-\u0AB9\u0ABD\u0AD0\u0AE0\u0AE1\u0B05-\u0B0C\u0B0F\u0B10\u0B13-\u0B28\u0B2A-\u0B30\u0B32\u0B33\u0B35-\u0B39\u0B3D\u0B5C\u0B5D\u0B5F-\u0B61\u0B71\u0B83\u0B85-\u0B8A\u0B8E-\u0B90\u0B92-\u0B95\u0B99\u0B9A\u0B9C\u0B9E\u0B9F\u0BA3\u0BA4\u0BA8-\u0BAA\u0BAE-\u0BB9\u0BD0\u0C05-\u0C0C\u0C0E-\u0C10\u0C12-\u0C28\u0C2A-\u0C33\u0C35-\u0C39\u0C3D\u0C58\u0C59\u0C60\u0C61\u0C85-\u0C8C\u0C8E-\u0C90\u0C92-\u0CA8\u0CAA-\u0CB3\u0CB5-\u0CB9\u0CBD\u0CDE\u0CE0\u0CE1\u0CF1\u0CF2\u0D05-\u0D0C\u0D0E-\u0D10\u0D12-\u0D28\u0D2A-\u0D39\u0D3D\u0D60\u0D61\u0D7A-\u0D7F\u0D85-\u0D96\u0D9A-\u0DB1\u0DB3-\u0DBB\u0DBD\u0DC0-\u0DC6\u0E01-\u0E30\u0E32\u0E33\u0E40-\u0E46\u0E81\u0E82\u0E84\u0E86-\u0E8A\u0E8C-\u0EA3\u0EA5\u0EA7-\u0EB0\u0EB2\u0EB3\u0EBD\u0EC0-\u0EC4\u0EC6\u0EDC\u0EDD\u0F00\u0F40-\u0F47\u0F49-\u0F6C\u0F88-\u0F8C\u1000-\u102A\u103F\u1050-\u1055\u105A-\u105D\u1061\u1065\u1066\u106E-\u1070\u1075-\u1081\u108E\u10A0-\u10C5\u1100-\u11FF\u1780-\u17B3\u17D7\u17DC\u1820-\u1877\u1880-\u1884\u1887-\u18A8\u18AA\u1900-\u191C\u1950-\u196D\u1970-\u1974\u1980-\u19A9\u19B0-\u19C9\u1D62-\u1D6A\u1E00\u1E02\u1E04\u1E06\u1E08\u1E0A\u1E0C\u1E0E\u1E10\u1E12\u1E14\u1E16\u1E18\u1E1A\u1E1C\u1E1E\u1E20\u1E22\u1E24\u1E26\u1E28\u1E2A\u1E2C\u1E2E\u1E30\u1E32\u1E34\u1E36\u1E38\u1E3A\u1E3C\u1E3E\u1E40\u1E42\u1E44\u1E46\u1E48\u1E4A\u1E4C\u1E4E\u1E50\u1E52\u1E54\u1E56\u1E58\u1E5A\u1E5C\u1E5E\u1E60\u1E62\u1E64\u1E66\u1E68\u1E6A\u1E6C\u1E6E\u1E70\u1E72\u1E74\u1E76\u1E78\u1E7A\u1E7C\u1E7E\u1E80\u1E82\u1E84\u1E86\u1E88\u1E8A\u1E8C\u1E8E\u1E90\u1E92\u1E94\u1E9E\u1EA0\u1EA2\u1EA4\u1EA6\u1EA8\u1EAA\u1EAC\u1EAE\u1EB0\u1EB2\u1EB4\u1EB6\u1EB8\u1EBA\u1EBC\u1EBE\u1EC0\u1EC2\u1EC4\u1EC6\u1EC8\u1ECA\u1ECC\u1ECE\u1ED0\u1ED2\u1ED4\u1ED6\u1ED8\u1EDA\u1EDC\u1EDE\u1EE0\u1EE2\u1EE4\u1EE6\u1EE8\u1EEA\u1EEC\u1EEE\u1EF0\u1EF2\u1EF4\u1EF6\u1EF8\u1EFA\u1EFC\u1EFE\u1F08-\u1F0F\u1F18-\u1F1D\u1F28-\u1F2F\u1F38-\u1F3F\u1F48-\u1F4D\u1F59\u1F5B\u1F5D\u1F5F\u1F68-\u1F6F\u1F88-\u1F8F\u1F98-\u1F9F\u1FA8-\u1FAF\u1FB8-\u1FBC\u1FC8-\u1FCC\u1FD8-\u1FDB\u1FE8-\u1FEC\u1FF8-\u1FFC\u2071\u207F\u2102\u2107\u210B-\u210D\u2110-\u2112\u2115\u2119-\u211D\u2124\u2126\u2128\u212A-\u212D\u2130-\u2133\u213E\u213F\u2145\u2183\u2C00-\u2C2E\u2C60\u2C62-\u2C64\u2C67\u2C69\u2C6B\u2C6D-\u2C6F\u2C72\u2C75\u2C7C\u2C80\u2C82\u2C84\u2C86\u2C88\u2C8A\u2C8C\u2C8E\u2C90\u2C92\u2C94\u2C96\u2C98\u2C9A\u2C9C\u2C9E\u2CA0\u2CA2\u2CA4\u2CA6\u2CA8\u2CAA\u2CAC\u2CAE\u2CB0\u2CB2\u2CB4\u2CB6\u2CB8\u2CBA\u2CBC\u2CBE\u2CC0\u2CC2\u2CC4\u2CC6\u2CC8\u2CCA\u2CCC\u2CCE\u2CD0\u2CD2\u2CD4\u2CD6\u2CD8\u2CDA\u2CDC\u2CDE\u2CE0\u2CE2\u3005\u3006\u3031-\u3035\u303B\u303C\u3041-\u3096\u309D-\u309F\u30A1-\u30FA\u30FC-\u30FF\u3105-\u312F\u3131-\u318E\u31A0-\u31BA\u31F0-\u31FF\u3400-\u4DB5\u4E00-\u9FEF\uA000-\uA48C\uA640\uA642\uA644\uA646\uA648\uA64A\uA64C\uA64E\uA650\uA652\uA654\uA656\uA658\uA65A\uA65C\uA65E\uA662\uA664\uA666\uA668\uA66A\uA66C\uA680\uA682\uA684\uA686\uA688\uA68A\uA68C\uA68E\uA690\uA692\uA694\uA696\uA722\uA724\uA726\uA728\uA72A\uA72C\uA72E\uA732\uA734\uA736\uA738\uA73A\uA73C\uA73E\uA740\uA742\uA744\uA746\uA748\uA74A\uA74C\uA74E\uA750\uA752\uA754\uA756\uA758\uA75A\uA75C\uA75E\uA760\uA762\uA764\uA766\uA768\uA76A\uA76C\uA76E\uA779\uA77B\uA77D\uA77E\uA780\uA782\uA784\uA786\uA78B\uA960-\uA97C\uAC00-\uD7A3\uD7B0-\uD7C6\uD7CB-\uD7FB\uF900-\uFA6D\uFA70-\uFAD9\uFB1D\uFB1F-\uFB28\uFB2A-\uFB36\uFB38-\uFB3C\uFB3E\uFB40\uFB41\uFB43\uFB44\uFB46-\uFBB1\uFBD3-\uFD3D\uFD50-\uFD8F\uFD92-\uFDC7\uFDF0-\uFDFB\uFE70-\uFE74\uFE76-\uFEFC\uFF21-\uFF3A\uFF66-\uFF9F]|\uD800[\uDC00-\uDC0B\uDC0D-\uDC26\uDC28-\uDC3A\uDC3C\uDC3D\uDC3F-\uDC4D\uDC50-\uDC5D\uDC80-\uDCFA\uDE80-\uDE9C\uDEA0-\uDED0\uDF00-\uDF1F\uDF2D-\uDF40\uDF42-\uDF49\uDF50-\uDF75\uDF80-\uDF9D\uDFA0-\uDFC3\uDFC8-\uDFCF]|\uD801[\uDC00-\uDC27\uDC50-\uDC9D\uDCB0-\uDCD3\uDD00-\uDD27\uDD30-\uDD63\uDE00-\uDF36\uDF40-\uDF55\uDF60-\uDF67]|\uD802[\uDC00-\uDC05\uDC08\uDC0A-\uDC35\uDC37\uDC38\uDC3C\uDC3F-\uDC55\uDC60-\uDC76\uDC80-\uDC9E\uDCE0-\uDCF2\uDCF4\uDCF5\uDD00-\uDD15\uDD20-\uDD39\uDD80-\uDDB7\uDDBE\uDDBF\uDE00\uDE10-\uDE13\uDE15-\uDE17\uDE19-\uDE35\uDE60-\uDE7C\uDE80-\uDE9C\uDEC0-\uDEC7\uDEC9-\uDEE4\uDF00-\uDF35\uDF40-\uDF55\uDF60-\uDF72\uDF80-\uDF91]|\uD803[\uDC00-\uDC48\uDC80-\uDCB2\uDD00-\uDD23\uDF00-\uDF1C\uDF27\uDF30-\uDF45\uDFE0-\uDFF6]|\uD804[\uDC03-\uDC37\uDC83-\uDCAF\uDCD0-\uDCE8\uDD03-\uDD26\uDD44\uDD50-\uDD72\uDD76\uDD83-\uDDB2\uDDC1-\uDDC4\uDDDA\uDDDC\uDE00-\uDE11\uDE13-\uDE2B\uDE80-\uDE86\uDE88\uDE8A-\uDE8D\uDE8F-\uDE9D\uDE9F-\uDEA8\uDEB0-\uDEDE\uDF05-\uDF0C\uDF0F\uDF10\uDF13-\uDF28\uDF2A-\uDF30\uDF32\uDF33\uDF35-\uDF39\uDF3D\uDF50\uDF5D-\uDF61]|\uD805[\uDC00-\uDC34\uDC47-\uDC4A\uDC5F\uDC80-\uDCAF\uDCC4\uDCC5\uDCC7\uDD80-\uDDAE\uDDD8-\uDDDB\uDE00-\uDE2F\uDE44\uDE80-\uDEAA\uDEB8\uDF00-\uDF1A]|\uD806[\uDC00-\uDC2B\uDCA0-\uDCBF\uDCFF\uDDA0-\uDDA7\uDDAA-\uDDD0\uDDE1\uDDE3\uDE00\uDE0B-\uDE32\uDE3A\uDE50\uDE5C-\uDE89\uDE9D\uDEC0-\uDEF8]|\uD807[\uDC00-\uDC08\uDC0A-\uDC2E\uDC40\uDC72-\uDC8F\uDD00-\uDD06\uDD08\uDD09\uDD0B-\uDD30\uDD46\uDD60-\uDD65\uDD67\uDD68\uDD6A-\uDD89\uDD98\uDEE0-\uDEF2]|\uD808[\uDC00-\uDF99]|\uD809[\uDC80-\uDD43]|[\uD80C\uD81C-\uD820\uD840-\uD868\uD86A-\uD86C\uD86F-\uD872\uD874-\uD879][\uDC00-\uDFFF]|\uD80D[\uDC00-\uDC2E]|\uD811[\uDC00-\uDE46]|\uD81A[\uDC00-\uDE38\uDE40-\uDE5E\uDED0-\uDEED\uDF00-\uDF2F\uDF40-\uDF43\uDF63-\uDF77\uDF7D-\uDF8F]|\uD81B[\uDE40-\uDE5F\uDF00-\uDF4A\uDF50\uDF93-\uDF9F\uDFE0\uDFE1\uDFE3]|\uD821[\uDC00-\uDFF7]|\uD822[\uDC00-\uDEF2]|\uD82C[\uDC00-\uDD1E\uDD50-\uDD52\uDD64-\uDD67\uDD70-\uDEFB]|\uD82F[\uDC00-\uDC6A\uDC70-\uDC7C\uDC80-\uDC88\uDC90-\uDC99]|\uD835[\uDC00-\uDC19\uDC34-\uDC4D\uDC68-\uDC81\uDC9C\uDC9E\uDC9F\uDCA2\uDCA5\uDCA6\uDCA9-\uDCAC\uDCAE-\uDCB5\uDCD0-\uDCE9\uDD04\uDD05\uDD07-\uDD0A\uDD0D-\uDD14\uDD16-\uDD1C\uDD38\uDD39\uDD3B-\uDD3E\uDD40-\uDD44\uDD46\uDD4A-\uDD50\uDD6C-\uDD85\uDDA0-\uDDB9\uDDD4-\uDDED\uDE08-\uDE21\uDE3C-\uDE55\uDE70-\uDE89\uDEA8-\uDEC0\uDEE2-\uDEFA\uDF1C-\uDF34\uDF56-\uDF6E\uDF90-\uDFA8\uDFCA]|\uD838[\uDD00-\uDD2C\uDD37-\uDD3D\uDD4E\uDEC0-\uDEEB]|\uD83A[\uDC00-\uDCC4\uDD00-\uDD21\uDD4B]|\uD83B[\uDE00-\uDE03\uDE05-\uDE1F\uDE21\uDE22\uDE24\uDE27\uDE29-\uDE32\uDE34-\uDE37\uDE39\uDE3B\uDE42\uDE47\uDE49\uDE4B\uDE4D-\uDE4F\uDE51\uDE52\uDE54\uDE57\uDE59\uDE5B\uDE5D\uDE5F\uDE61\uDE62\uDE64\uDE67-\uDE6A\uDE6C-\uDE72\uDE74-\uDE77\uDE79-\uDE7C\uDE7E\uDE80-\uDE89\uDE8B-\uDE9B\uDEA1-\uDEA3\uDEA5-\uDEA9\uDEAB-\uDEBB]|\uD869[\uDC00-\uDED6\uDF00-\uDFFF]|\uD86D[\uDC00-\uDF34\uDF40-\uDFFF]|\uD86E[\uDC00-\uDC1D\uDC20-\uDFFF]|\uD873[\uDC00-\uDEA1\uDEB0-\uDFFF]|\uD87A[\uDC00-\uDFE0]|\uD87E[\uDC00-\uDE1D])((?:[0-9A-Z_\xAA\xBA\xC0-\xD6\xD8-\xDE\u0100\u0102\u0104\u0106\u0108\u010A\u010C\u010E\u0110\u0112\u0114\u0116\u0118\u011A\u011C\u011E\u0120\u0122\u0124\u0126\u0128\u012A\u012C\u012E\u0130\u0132\u0134\u0136\u0139\u013B\u013D\u013F\u0141\u0143\u0145\u0147\u014A\u014C\u014E\u0150\u0152\u0154\u0156\u0158\u015A\u015C\u015E\u0160\u0162\u0164\u0166\u0168\u016A\u016C\u016E\u0170\u0172\u0174\u0176\u0178\u0179\u017B\u017D\u0181\u0182\u0184\u0186\u0187\u0189-\u018B\u018E-\u0191\u0193\u0194\u0196-\u0198\u019C\u019D\u019F\u01A0\u01A2\u01A4\u01A6\u01A7\u01A9\u01AC\u01AE\u01AF\u01B1-\u01B3\u01B5\u01B7\u01B8\u01BC\u01C4\u01C5\u01C7\u01C8\u01CA\u01CB\u01CD\u01CF\u01D1\u01D3\u01D5\u01D7\u01D9\u01DB\u01DE\u01E0\u01E2\u01E4\u01E6\u01E8\u01EA\u01EC\u01EE\u01F1\u01F2\u01F4\u01F6-\u01F8\u01FA\u01FC\u01FE\u0200\u0202\u0204\u0206\u0208\u020A\u020C\u020E\u0210\u0212\u0214\u0216\u0218\u021A\u021C\u021E\u0220\u0222\u0224\u0226\u0228\u022A\u022C\u022E\u0230\u0232\u023A\u023B\u023D\u023E\u0241\u0243-\u0246\u0248\u024A\u024C\u024E\u02B0-\u02C1\u02C6-\u02D1\u02E0-\u02E4\u02EC\u02EE\u0370\u0372\u0376\u0386\u0388-\u038A\u038C\u038E\u038F\u0391-\u03A1\u03A3-\u03AB\u03CF\u03D2-\u03D4\u03D8\u03DA\u03DC\u03DE\u03E0\u03E2\u03E4\u03E6\u03E8\u03EA\u03EC\u03EE\u03F4\u03F7\u03F9\u03FA\u03FD-\u042F\u0460\u0462\u0464\u0466\u0468\u046A\u046C\u046E\u0470\u0472\u0474\u0476\u0478\u047A\u047C\u047E\u0480\u048A\u048C\u048E\u0490\u0492\u0494\u0496\u0498\u049A\u049C\u049E\u04A0\u04A2\u04A4\u04A6\u04A8\u04AA\u04AC\u04AE\u04B0\u04B2\u04B4\u04B6\u04B8\u04BA\u04BC\u04BE\u04C0\u04C1\u04C3\u04C5\u04C7\u04C9\u04CB\u04CD\u04D0\u04D2\u04D4\u04D6\u04D8\u04DA\u04DC\u04DE\u04E0\u04E2\u04E4\u04E6\u04E8\u04EA\u04EC\u04EE\u04F0\u04F2\u04F4\u04F6\u04F8\u04FA\u04FC\u04FE\u0500\u0502\u0504\u0506\u0508\u050A\u050C\u050E\u0510\u0512\u0514\u0516\u0518\u051A\u051C\u051E\u0520\u0522\u0531-\u0556\u05D0-\u05EA\u05F0-\u05F2\u0621-\u064A\u066E\u066F\u0671-\u06D3\u06D5\u06E5\u06E6\u06EE\u06EF\u06FA-\u06FC\u06FF\u0710\u0712-\u072F\u074D-\u07A5\u07B1\u07CA-\u07EA\u07F4\u07F5\u07FA\u0904-\u0939\u093D\u0950\u0958-\u0961\u0971\u0972\u097B-\u097F\u0985-\u098C\u098F\u0990\u0993-\u09A8\u09AA-\u09B0\u09B2\u09B6-\u09B9\u09BD\u09CE\u09DC\u09DD\u09DF-\u09E1\u09F0\u09F1\u0A05-\u0A0A\u0A0F\u0A10\u0A13-\u0A28\u0A2A-\u0A30\u0A32\u0A33\u0A35\u0A36\u0A38\u0A39\u0A59-\u0A5C\u0A5E\u0A72-\u0A74\u0A85-\u0A8D\u0A8F-\u0A91\u0A93-\u0AA8\u0AAA-\u0AB0\u0AB2\u0AB3\u0AB5-\u0AB9\u0ABD\u0AD0\u0AE0\u0AE1\u0B05-\u0B0C\u0B0F\u0B10\u0B13-\u0B28\u0B2A-\u0B30\u0B32\u0B33\u0B35-\u0B39\u0B3D\u0B5C\u0B5D\u0B5F-\u0B61\u0B71\u0B83\u0B85-\u0B8A\u0B8E-\u0B90\u0B92-\u0B95\u0B99\u0B9A\u0B9C\u0B9E\u0B9F\u0BA3\u0BA4\u0BA8-\u0BAA\u0BAE-\u0BB9\u0BD0\u0C05-\u0C0C\u0C0E-\u0C10\u0C12-\u0C28\u0C2A-\u0C33\u0C35-\u0C39\u0C3D\u0C58\u0C59\u0C60\u0C61\u0C85-\u0C8C\u0C8E-\u0C90\u0C92-\u0CA8\u0CAA-\u0CB3\u0CB5-\u0CB9\u0CBD\u0CDE\u0CE0\u0CE1\u0CF1\u0CF2\u0D05-\u0D0C\u0D0E-\u0D10\u0D12-\u0D28\u0D2A-\u0D39\u0D3D\u0D60\u0D61\u0D7A-\u0D7F\u0D85-\u0D96\u0D9A-\u0DB1\u0DB3-\u0DBB\u0DBD\u0DC0-\u0DC6\u0E01-\u0E30\u0E32\u0E33\u0E40-\u0E46\u0E81\u0E82\u0E84\u0E86-\u0E8A\u0E8C-\u0EA3\u0EA5\u0EA7-\u0EB0\u0EB2\u0EB3\u0EBD\u0EC0-\u0EC4\u0EC6\u0EDC\u0EDD\u0F00\u0F40-\u0F47\u0F49-\u0F6C\u0F88-\u0F8C\u1000-\u102A\u103F\u1050-\u1055\u105A-\u105D\u1061\u1065\u1066\u106E-\u1070\u1075-\u1081\u108E\u10A0-\u10C5\u1100-\u11FF\u1780-\u17B3\u17D7\u17DC\u1820-\u1877\u1880-\u1884\u1887-\u18A8\u18AA\u1900-\u191C\u1950-\u196D\u1970-\u1974\u1980-\u19A9\u19B0-\u19C9\u1D62-\u1D6A\u1E00\u1E02\u1E04\u1E06\u1E08\u1E0A\u1E0C\u1E0E\u1E10\u1E12\u1E14\u1E16\u1E18\u1E1A\u1E1C\u1E1E\u1E20\u1E22\u1E24\u1E26\u1E28\u1E2A\u1E2C\u1E2E\u1E30\u1E32\u1E34\u1E36\u1E38\u1E3A\u1E3C\u1E3E\u1E40\u1E42\u1E44\u1E46\u1E48\u1E4A\u1E4C\u1E4E\u1E50\u1E52\u1E54\u1E56\u1E58\u1E5A\u1E5C\u1E5E\u1E60\u1E62\u1E64\u1E66\u1E68\u1E6A\u1E6C\u1E6E\u1E70\u1E72\u1E74\u1E76\u1E78\u1E7A\u1E7C\u1E7E\u1E80\u1E82\u1E84\u1E86\u1E88\u1E8A\u1E8C\u1E8E\u1E90\u1E92\u1E94\u1E9E\u1EA0\u1EA2\u1EA4\u1EA6\u1EA8\u1EAA\u1EAC\u1EAE\u1EB0\u1EB2\u1EB4\u1EB6\u1EB8\u1EBA\u1EBC\u1EBE\u1EC0\u1EC2\u1EC4\u1EC6\u1EC8\u1ECA\u1ECC\u1ECE\u1ED0\u1ED2\u1ED4\u1ED6\u1ED8\u1EDA\u1EDC\u1EDE\u1EE0\u1EE2\u1EE4\u1EE6\u1EE8\u1EEA\u1EEC\u1EEE\u1EF0\u1EF2\u1EF4\u1EF6\u1EF8\u1EFA\u1EFC\u1EFE\u1F08-\u1F0F\u1F18-\u1F1D\u1F28-\u1F2F\u1F38-\u1F3F\u1F48-\u1F4D\u1F59\u1F5B\u1F5D\u1F5F\u1F68-\u1F6F\u1F88-\u1F8F\u1F98-\u1F9F\u1FA8-\u1FAF\u1FB8-\u1FBC\u1FC8-\u1FCC\u1FD8-\u1FDB\u1FE8-\u1FEC\u1FF8-\u1FFC\u2071\u207F\u2102\u2107\u210B-\u210D\u2110-\u2112\u2115\u2119-\u211D\u2124\u2126\u2128\u212A-\u212D\u2130-\u2133\u213E\u213F\u2145\u2183\u2C00-\u2C2E\u2C60\u2C62-\u2C64\u2C67\u2C69\u2C6B\u2C6D-\u2C6F\u2C72\u2C75\u2C7C\u2C80\u2C82\u2C84\u2C86\u2C88\u2C8A\u2C8C\u2C8E\u2C90\u2C92\u2C94\u2C96\u2C98\u2C9A\u2C9C\u2C9E\u2CA0\u2CA2\u2CA4\u2CA6\u2CA8\u2CAA\u2CAC\u2CAE\u2CB0\u2CB2\u2CB4\u2CB6\u2CB8\u2CBA\u2CBC\u2CBE\u2CC0\u2CC2\u2CC4\u2CC6\u2CC8\u2CCA\u2CCC\u2CCE\u2CD0\u2CD2\u2CD4\u2CD6\u2CD8\u2CDA\u2CDC\u2CDE\u2CE0\u2CE2\u3005\u3006\u3031-\u3035\u303B\u303C\u3041-\u3096\u309D-\u309F\u30A1-\u30FA\u30FC-\u30FF\u3105-\u312F\u3131-\u318E\u31A0-\u31BA\u31F0-\u31FF\u3400-\u4DB5\u4E00-\u9FEF\uA000-\uA48C\uA640\uA642\uA644\uA646\uA648\uA64A\uA64C\uA64E\uA650\uA652\uA654\uA656\uA658\uA65A\uA65C\uA65E\uA662\uA664\uA666\uA668\uA66A\uA66C\uA680\uA682\uA684\uA686\uA688\uA68A\uA68C\uA68E\uA690\uA692\uA694\uA696\uA722\uA724\uA726\uA728\uA72A\uA72C\uA72E\uA732\uA734\uA736\uA738\uA73A\uA73C\uA73E\uA740\uA742\uA744\uA746\uA748\uA74A\uA74C\uA74E\uA750\uA752\uA754\uA756\uA758\uA75A\uA75C\uA75E\uA760\uA762\uA764\uA766\uA768\uA76A\uA76C\uA76E\uA779\uA77B\uA77D\uA77E\uA780\uA782\uA784\uA786\uA78B\uA960-\uA97C\uAC00-\uD7A3\uD7B0-\uD7C6\uD7CB-\uD7FB\uF900-\uFA6D\uFA70-\uFAD9\uFB1D\uFB1F-\uFB28\uFB2A-\uFB36\uFB38-\uFB3C\uFB3E\uFB40\uFB41\uFB43\uFB44\uFB46-\uFBB1\uFBD3-\uFD3D\uFD50-\uFD8F\uFD92-\uFDC7\uFDF0-\uFDFB\uFE70-\uFE74\uFE76-\uFEFC\uFF21-\uFF3A\uFF66-\uFF9F]|\uD800[\uDC00-\uDC0B\uDC0D-\uDC26\uDC28-\uDC3A\uDC3C\uDC3D\uDC3F-\uDC4D\uDC50-\uDC5D\uDC80-\uDCFA\uDE80-\uDE9C\uDEA0-\uDED0\uDF00-\uDF1F\uDF2D-\uDF40\uDF42-\uDF49\uDF50-\uDF75\uDF80-\uDF9D\uDFA0-\uDFC3\uDFC8-\uDFCF]|\uD801[\uDC00-\uDC27\uDC50-\uDC9D\uDCB0-\uDCD3\uDD00-\uDD27\uDD30-\uDD63\uDE00-\uDF36\uDF40-\uDF55\uDF60-\uDF67]|\uD802[\uDC00-\uDC05\uDC08\uDC0A-\uDC35\uDC37\uDC38\uDC3C\uDC3F-\uDC55\uDC60-\uDC76\uDC80-\uDC9E\uDCE0-\uDCF2\uDCF4\uDCF5\uDD00-\uDD15\uDD20-\uDD39\uDD80-\uDDB7\uDDBE\uDDBF\uDE00\uDE10-\uDE13\uDE15-\uDE17\uDE19-\uDE35\uDE60-\uDE7C\uDE80-\uDE9C\uDEC0-\uDEC7\uDEC9-\uDEE4\uDF00-\uDF35\uDF40-\uDF55\uDF60-\uDF72\uDF80-\uDF91]|\uD803[\uDC00-\uDC48\uDC80-\uDCB2\uDD00-\uDD23\uDF00-\uDF1C\uDF27\uDF30-\uDF45\uDFE0-\uDFF6]|\uD804[\uDC03-\uDC37\uDC83-\uDCAF\uDCD0-\uDCE8\uDD03-\uDD26\uDD44\uDD50-\uDD72\uDD76\uDD83-\uDDB2\uDDC1-\uDDC4\uDDDA\uDDDC\uDE00-\uDE11\uDE13-\uDE2B\uDE80-\uDE86\uDE88\uDE8A-\uDE8D\uDE8F-\uDE9D\uDE9F-\uDEA8\uDEB0-\uDEDE\uDF05-\uDF0C\uDF0F\uDF10\uDF13-\uDF28\uDF2A-\uDF30\uDF32\uDF33\uDF35-\uDF39\uDF3D\uDF50\uDF5D-\uDF61]|\uD805[\uDC00-\uDC34\uDC47-\uDC4A\uDC5F\uDC80-\uDCAF\uDCC4\uDCC5\uDCC7\uDD80-\uDDAE\uDDD8-\uDDDB\uDE00-\uDE2F\uDE44\uDE80-\uDEAA\uDEB8\uDF00-\uDF1A]|\uD806[\uDC00-\uDC2B\uDCA0-\uDCBF\uDCFF\uDDA0-\uDDA7\uDDAA-\uDDD0\uDDE1\uDDE3\uDE00\uDE0B-\uDE32\uDE3A\uDE50\uDE5C-\uDE89\uDE9D\uDEC0-\uDEF8]|\uD807[\uDC00-\uDC08\uDC0A-\uDC2E\uDC40\uDC72-\uDC8F\uDD00-\uDD06\uDD08\uDD09\uDD0B-\uDD30\uDD46\uDD60-\uDD65\uDD67\uDD68\uDD6A-\uDD89\uDD98\uDEE0-\uDEF2]|\uD808[\uDC00-\uDF99]|\uD809[\uDC80-\uDD43]|[\uD80C\uD81C-\uD820\uD840-\uD868\uD86A-\uD86C\uD86F-\uD872\uD874-\uD879][\uDC00-\uDFFF]|\uD80D[\uDC00-\uDC2E]|\uD811[\uDC00-\uDE46]|\uD81A[\uDC00-\uDE38\uDE40-\uDE5E\uDED0-\uDEED\uDF00-\uDF2F\uDF40-\uDF43\uDF63-\uDF77\uDF7D-\uDF8F]|\uD81B[\uDE40-\uDE5F\uDF00-\uDF4A\uDF50\uDF93-\uDF9F\uDFE0\uDFE1\uDFE3]|\uD821[\uDC00-\uDFF7]|\uD822[\uDC00-\uDEF2]|\uD82C[\uDC00-\uDD1E\uDD50-\uDD52\uDD64-\uDD67\uDD70-\uDEFB]|\uD82F[\uDC00-\uDC6A\uDC70-\uDC7C\uDC80-\uDC88\uDC90-\uDC99]|\uD835[\uDC00-\uDC19\uDC34-\uDC4D\uDC68-\uDC81\uDC9C\uDC9E\uDC9F\uDCA2\uDCA5\uDCA6\uDCA9-\uDCAC\uDCAE-\uDCB5\uDCD0-\uDCE9\uDD04\uDD05\uDD07-\uDD0A\uDD0D-\uDD14\uDD16-\uDD1C\uDD38\uDD39\uDD3B-\uDD3E\uDD40-\uDD44\uDD46\uDD4A-\uDD50\uDD6C-\uDD85\uDDA0-\uDDB9\uDDD4-\uDDED\uDE08-\uDE21\uDE3C-\uDE55\uDE70-\uDE89\uDEA8-\uDEC0\uDEE2-\uDEFA\uDF1C-\uDF34\uDF56-\uDF6E\uDF90-\uDFA8\uDFCA]|\uD838[\uDD00-\uDD2C\uDD37-\uDD3D\uDD4E\uDEC0-\uDEEB]|\uD83A[\uDC00-\uDCC4\uDD00-\uDD21\uDD4B]|\uD83B[\uDE00-\uDE03\uDE05-\uDE1F\uDE21\uDE22\uDE24\uDE27\uDE29-\uDE32\uDE34-\uDE37\uDE39\uDE3B\uDE42\uDE47\uDE49\uDE4B\uDE4D-\uDE4F\uDE51\uDE52\uDE54\uDE57\uDE59\uDE5B\uDE5D\uDE5F\uDE61\uDE62\uDE64\uDE67-\uDE6A\uDE6C-\uDE72\uDE74-\uDE77\uDE79-\uDE7C\uDE7E\uDE80-\uDE89\uDE8B-\uDE9B\uDEA1-\uDEA3\uDEA5-\uDEA9\uDEAB-\uDEBB]|\uD869[\uDC00-\uDED6\uDF00-\uDFFF]|\uD86D[\uDC00-\uDF34\uDF40-\uDFFF]|\uD86E[\uDC00-\uDC1D\uDC20-\uDFFF]|\uD873[\uDC00-\uDEA1\uDEB0-\uDFFF]|\uD87A[\uDC00-\uDFE0]|\uD87E[\uDC00-\uDE1D]){0,39})";
    var sBookmarkCellRef = sBookmark + '( +(' + sCellReference + '))*';
    var sFunctions = "(ABS\|AND\|AVERAGE\|COUNT\|DEFINED\|FALSE\|IF\|INT\|MAX\|MIN\|MOD\|NOT\|OR\|PRODUCT\|ROUND\|SIGN\|SUM\|TRUE)" ;

    var oComparisonOpRegExp = new RegExp(sComparison, 'g');
    var oColumnNameRegExp = new RegExp(sColumnName, 'g');
    var oCellNameRegExp = new RegExp(sCellName, 'g');
    var oRowNameRegExp = new RegExp(sRowName, 'g');
    var oCellRangeRegExp = new RegExp(sCellRange, 'g');
    var oCellCellRangeRegExp = new RegExp(sCellCellRange, 'g');
    var oRowRangeRegExp = new RegExp(sRowRange, 'g');
    var oColRangeRegExp = new RegExp(sColumnRange, 'g');
    var oCellReferenceRegExp = new RegExp(sCellReference, 'g');
    var oIdentifierRegExp = new RegExp(sIdentifier, 'g');
    var oBookmarkNameRegExp = new RegExp(sBookmark, 'g');
    var oBookmarkCellRefRegExp = new RegExp(sBookmarkCellRef, 'g');
    var oConstantRegExp = new RegExp(sConstant, 'g');
    var oOperatorRegExp = new RegExp(sOperator, 'g');
    var oFunctionsRegExp = new RegExp(sFunctions, 'g');


    var oLettersMap = {};
    oLettersMap['A'] = 1;
    oLettersMap['B'] = 2;
    oLettersMap['C'] = 3;
    oLettersMap['D'] = 4;
    oLettersMap['E'] = 5;
    oLettersMap['F'] = 6;
    oLettersMap['G'] = 7;
    oLettersMap['H'] = 8;
    oLettersMap['I'] = 9;
    oLettersMap['J'] = 10;
    oLettersMap['K'] = 11;
    oLettersMap['L'] = 12;
    oLettersMap['M'] = 13;
    oLettersMap['N'] = 14;
    oLettersMap['O'] = 15;
    oLettersMap['P'] = 16;
    oLettersMap['Q'] = 17;
    oLettersMap['R'] = 18;
    oLettersMap['S'] = 19;
    oLettersMap['T'] = 20;
    oLettersMap['U'] = 21;
    oLettersMap['V'] = 22;
    oLettersMap['W'] = 23;
    oLettersMap['X'] = 24;
    oLettersMap['Y'] = 25;
    oLettersMap['Z'] = 26;


  var oDigitMap = {};
    oDigitMap['0'] = 0;
    oDigitMap['1'] = 1;
    oDigitMap['2'] = 2;
    oDigitMap['3'] = 3;
    oDigitMap['4'] = 4;
    oDigitMap['5'] = 5;
    oDigitMap['6'] = 6;
    oDigitMap['7'] = 7;
    oDigitMap['8'] = 8;
    oDigitMap['9'] = 9;


    function CFormulaNode(parseQueue) {
        this.document = null;
        this.result = null;
        this.error = null;
        this.parseQueue = parseQueue;
        this.parent = null;
    }
    CFormulaNode.prototype.argumentsCount = 0;
    CFormulaNode.prototype.supportErrorArgs = function(){
        return false;
    };
    CFormulaNode.prototype.inFunction = function(){
        if(this.isFunction()){
            return this;
        }
        if(!this.parent){
            return this.parent;
        }
        return this.parent.inFunction();
    };

    CFormulaNode.prototype.isCell = function(){
        return false;
    };


    CFormulaNode.prototype.checkSizeFormated = function(_result){
        if(_result.length > 63){
            this.setError((ERROR_TYPE_LARGE_NUMBER), null);
        }
    };
    CFormulaNode.prototype.checkRoundNumber = function(number){
        return fRoundNumber(number, 2);
    };

    CFormulaNode.prototype.checkBraces = function(_result){
        return _result;
    };

    CFormulaNode.prototype.formatResult = function(){
        var sResult = null;
        if(AscFormat.isRealNumber(this.result)){
            var _result = this.result;
            if(_result === Infinity){
                _result = 1.0;
            }
            if(_result === -Infinity){
                _result = -1.0;
            }
            if(this.parseQueue.format){
                return this.parseQueue.format.formatToWord(_result, 14);
            }
            sResult = "";
            if(_result < 0){
                _result = -_result;
            }
            _result = this.checkRoundNumber(_result);
            var i;
            var sRes = _result.toExponential(13);
            var aDigits = sRes.split('e');
            var nPow = parseInt(aDigits[1]);
            var sNum = aDigits[0];
            var t = sNum.split('.');
            if(nPow < 0){
                for(i = t[1].length - 1; i > -1; --i){
                    if(t[1][i] !== '0'){
                        break;
                    }
                }
                if(i > -1){
                    sResult += t[1].slice(0, i + 1);
                }
                sResult = t[0] + sResult;
                nPow = -nPow - 1;
                for(i = 0; i < nPow; ++i){
                    sResult = "0" + sResult;
                }
                sResult = ("0" + this.digitSeparator + sResult);
            }
            else{
                sResult += t[0];

                for(i = 0; i < nPow; ++i){
                    if(t[1] && i < t[1].length){
                        sResult += t[1][i];
                    }
                    else{
                        sResult += '0';
                    }
                }
                if(t[1] && nPow < t[1].length){
                    for(i = t[1].length - 1; i >= nPow; --i){
                        if(t[1][i] !== '0'){
                            break;
                        }
                    }
                    var sStr = "";
                    for(; i >= nPow; --i){
                        sStr = t[1][i] + sStr;
                    }
                    if(sStr !== ""){
                        sResult += (this.digitSeparator + sStr);
                    }
                }
            }
            this.checkSizeFormated(sResult);
            if(this.result < 0){
                sResult = '-' + sResult;
            }
            sResult = this.checkBraces(sResult);
        }
        return sResult;
    };

    CFormulaNode.prototype.calculate = function () {
        this.error = null;
        this.result = null;
        if(!this.parseQueue){
            this.setError(ERROR_TYPE_ERROR, null);
            return;
        }
        var aArgs = [];
        for(var i = 0; i < this.argumentsCount; ++i){
            var oArg = this.parseQueue.getNext();
            if(!oArg){
                this.setError(ERROR_TYPE_MISSING_ARGUMENT, null);
                return;
            }
            oArg.parent = this;
            oArg.calculate();

            aArgs.push(oArg);
        }
        if(!this.supportErrorArgs()){
            for(i = aArgs.length - 1; i > -1; --i){
                oArg = aArgs[i];
                if(oArg.error){
                    this.error = oArg.error;
                    return;
                }
            }
        }
        if(this.isOperator()){

        }
        this._calculate(aArgs);
    };

    CFormulaNode.prototype._calculate = function(aArgs){
        this.setError(ERROR_TYPE_ERROR, null);//not implemented
    };

    CFormulaNode.prototype.isFunction = function () {
        return false;
    };
    CFormulaNode.prototype.isOperator = function () {
        return false;
    };
    CFormulaNode.prototype.setError = function (Type, Data) {
        this.error = new CError(Type, Data);
    };
    CFormulaNode.prototype.setParseQueue = function(oQueue){
        this.parseQueue = oQueue;
    };

    CFormulaNode.prototype.argumentsCount = 0;

    function CNumberNode(parseQueue) {
        CFormulaNode.call(this, parseQueue);
        this.value = null;
    }
    CNumberNode.prototype = Object.create(CFormulaNode.prototype);
    CNumberNode.prototype.precedence = 15;
    CNumberNode.prototype.checkBraces = function(_result){
        var sFormula = this.parseQueue.formula;
        if(sFormula[0] === '(' && sFormula[sFormula.length - 1] === ')'){
            return '(' + _result + ')';
        }
        return _result;
    };
    CNumberNode.prototype.checkSizeFormated = function(_result){
        var sFormula = this.parseQueue.formula;
        if(sFormula[0] === '(' && sFormula[sFormula.length - 1] === ')'){
            CFormulaNode.prototype.checkSizeFormated.call(this, _result);
        }
    };
    CNumberNode.prototype.checkRoundNumber = function(number){
        return number;
    };
    CNumberNode.prototype._calculate = function () {
        if(AscFormat.isRealNumber(this.value)){
            this.result = this.value;
        }
        else{
            this.setError(ERROR_TYPE_ERROR, null);//not a number
        }
        return this.error;
    };

    function CListSeparatorNode(parseQueue) {
        CFormulaNode.call(this, parseQueue);
    }
    CListSeparatorNode.prototype = Object.create(CFormulaNode.prototype);
    CListSeparatorNode.prototype.precedence = 15;

    function COperatorNode(parseQueue){
        CFormulaNode.call(this, parseQueue);
    }

    COperatorNode.prototype = Object.create(CFormulaNode.prototype);
    COperatorNode.prototype.precedence = 0;
    COperatorNode.prototype.argumentsCount = 2;
    COperatorNode.prototype.isOperator = function(){
        return true;
    };
    COperatorNode.prototype.checkCellInFunction = function (aArgs) {
        var i, oArg;
        if(this.parent && this.parent.isFunction() && this.parent.listSupport()){
            for(i = aArgs.length - 1; i > -1; --i){
                oArg = aArgs[i];
                if(oArg.isCell()){
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, oArg.getOwnCellName());
                    return;
                }
            }
        }
    };

    function CUnaryMinusOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CUnaryMinusOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CUnaryMinusOperatorNode.prototype.precedence = 13;
    CUnaryMinusOperatorNode.prototype.argumentsCount = 1;
    CUnaryMinusOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = -aArgs[0].result;
    };

    function CPowersAndRootsOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CPowersAndRootsOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CPowersAndRootsOperatorNode.prototype.precedence = 12;
    CPowersAndRootsOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = Math.pow(aArgs[1].result, aArgs[0].result);
    };
    function CMultiplicationOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CMultiplicationOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CMultiplicationOperatorNode.prototype.precedence = 11;
    CMultiplicationOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = aArgs[0].result * aArgs[1].result;
    };

    function CDivisionOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CDivisionOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CDivisionOperatorNode.prototype.precedence = 11;
    CDivisionOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        if(AscFormat.fApproxEqual(0.0, aArgs[0].result)){
            this.setError(ERROR_TYPE_ZERO_DIVIDE, null);
            return;
        }
        this.result = aArgs[1].result/aArgs[0].result;
    };

    function CPercentageOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CPercentageOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CPercentageOperatorNode.prototype.precedence = 8;
    CPercentageOperatorNode.prototype.argumentsCount = 1;
    CPercentageOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        if(aArgs[0].error){
            this.setError(ERROR_TYPE_SYNTAX_ERROR, "%");
            return;
        }
        this.result = aArgs[0].result / 100.0;
    };
    CPercentageOperatorNode.prototype.supportErrorArgs = function () {
        return true;
    };

    function CAdditionOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CAdditionOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CAdditionOperatorNode.prototype.precedence = 7;
    CAdditionOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = aArgs[1].result + aArgs[0].result;
    };

    function CSubtractionOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CSubtractionOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CSubtractionOperatorNode.prototype.precedence = 7;
    CSubtractionOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = aArgs[1].result - aArgs[0].result;
    };

    function CEqualToOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CEqualToOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CEqualToOperatorNode.prototype.precedence = 6;
    CEqualToOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = AscFormat.fApproxEqual(aArgs[1].result, aArgs[0].result) ? 1.0 : 0.0;
    };

    function CNotEqualToOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CNotEqualToOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CNotEqualToOperatorNode.prototype.precedence = 5;
    CNotEqualToOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = AscFormat.fApproxEqual(aArgs[1].result, aArgs[0].result) ? 0.0 : 1.0;
    };
    function CLessThanOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CLessThanOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CLessThanOperatorNode.prototype.precedence = 4;
    CLessThanOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = (aArgs[1].result < aArgs[0].result) ? 1.0 : 0.0;
    };
    function CLessThanOrEqualToOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CLessThanOrEqualToOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CLessThanOrEqualToOperatorNode.prototype.precedence = 3;
    CLessThanOrEqualToOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = (aArgs[1].result <= aArgs[0].result) ? 1.0 : 0.0;
    };
    function CGreaterThanOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CGreaterThanOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CGreaterThanOperatorNode.prototype.precedence = 2;
    CGreaterThanOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = (aArgs[1].result > aArgs[0].result) ? 1.0 : 0.0;
    };
    function CGreaterThanOrEqualOperatorNode(parseQueue){
        COperatorNode.call(this, parseQueue);
    }

    CGreaterThanOrEqualOperatorNode.prototype = Object.create(COperatorNode.prototype);
    CGreaterThanOrEqualOperatorNode.prototype.precedence = 1;
    CGreaterThanOrEqualOperatorNode.prototype._calculate = function (aArgs) {
        this.checkCellInFunction(aArgs);
        if(this.error){
            return;
        }
        this.result = (aArgs[1].result >= aArgs[0].result) ? 1.0 : 0.0;
    };



    function CLeftParenOperatorNode(parseQueue){
        CFormulaNode.call(this, parseQueue);
    }

    CLeftParenOperatorNode.prototype = Object.create(CFormulaNode.prototype);
    CLeftParenOperatorNode.prototype.precedence = 1;
    // CLeftParenOperatorNode.prototype._calculate = function (aArgs) {
    // };

    function CRightParenOperatorNode(parseQueue){
        CFormulaNode.call(this, parseQueue);
    }

    CRightParenOperatorNode.prototype = Object.create(CFormulaNode.prototype);
    CRightParenOperatorNode.prototype.precedence = 1;
    // CRightParenOperatorNode.prototype._calculate = function (aArgs) {
    // };

    function CLineSeparatorOperatorNode(parseQueue){
        CFormulaNode.call(this, parseQueue);
    }

    CLineSeparatorOperatorNode.prototype = Object.create(CFormulaNode.prototype);
    CLineSeparatorOperatorNode.prototype.precedence = 1;
    // CRightParenOperatorNode.prototype._calculate = function (aArgs) {
    // };

    var oOperatorsMap = {};
    oOperatorsMap["-"] = CUnaryMinusOperatorNode;
    oOperatorsMap["^"] = CPowersAndRootsOperatorNode;
    oOperatorsMap["*"] = CMultiplicationOperatorNode;
    oOperatorsMap["/"] = CDivisionOperatorNode;
    oOperatorsMap["%"] = CPercentageOperatorNode;
    oOperatorsMap["+"] = CAdditionOperatorNode;
    oOperatorsMap["-"] = CSubtractionOperatorNode;
    oOperatorsMap["="] = CEqualToOperatorNode;
    oOperatorsMap["<>"]= CNotEqualToOperatorNode;
    oOperatorsMap["<"] = CLessThanOperatorNode;
    oOperatorsMap["<="]= CLessThanOrEqualToOperatorNode;
    oOperatorsMap[">"] = CGreaterThanOperatorNode;
    oOperatorsMap[">="] = CGreaterThanOrEqualOperatorNode;
    oOperatorsMap["("] = CLeftParenOperatorNode;
    oOperatorsMap[")"] = CRightParenOperatorNode;


    var LEFT = 0;
    var RIGHT = 1;
    var ABOVE = 2;
    var BELOW = 3;

    var sLetters = "ZABCDEFGHIJKLMNOPQRSTUVWXY";
    var sDigits = "0123456789";

    function CCellRangeNode(parseQueue){
        CFormulaNode.call(this, parseQueue);
        this.bookmarkName = null;
        this.c1 = null;
        this.r1 = null;
        this.c2 = null;
        this.r2 = null;
        this.dir = null;
    }
    CCellRangeNode.prototype = Object.create(CFormulaNode.prototype);
    CCellRangeNode.prototype.argumentsCount = 0;

    CCellRangeNode.prototype.getCellName = function(c, r){
        var _c = c + 1 ;
        var _r = r + 1;

        var sColName = sLetters[(_c % 26)];
        _c = ((_c / 26) >> 0);
        while(_c !== 0){
            sColName = sLetters[(_c % 26)] + sColName;
            _c = ((_c / 26) >> 0);
        }
        var sRowName = sDigits[(_r % 10)];
        _r = ((_r / 10) >> 0);
        while(_r!== 0){
            sRowName = sDigits[(_r % 10)] + sRowName;
            _r = ((_r / 10) >> 0);
        }
        return sColName + sRowName;
    };
    CCellRangeNode.prototype.getOwnCellName = function(c, r){
        return this.getCellName(this.c1, this.r1);
    };

    CCellRangeNode.prototype.parseText = function(sText){
        var oParser = new CTextParser(',', this.digitSeparator);
        oParser.setFlag(PARSER_MASK_CLEAN, true);
        oParser.parse(Asc.trim(sText));
        if(oParser.parseQueue){
            oParser.parseQueue.flags = oParser.flags;
            oParser.parseQueue.calculate(false);
            if(!AscFormat.isRealNumber(oParser.parseQueue.result) || oParser.parseQueue.pos > 0){
                var aQueue = oParser.parseQueue.queue;
                var fSumm = 0.0;
                if(aQueue.length > 0){
                    for(var i = 0; i < aQueue.length; ++i){
                        if(aQueue[i] instanceof CNumberNode){
                            fSumm += aQueue[i].value;
                        }
                        else if(aQueue[i] instanceof CLineSeparatorOperatorNode){
                            continue;
                        }
                        else{
                            break;
                        }
                    }
                    if(aQueue.length === i){
                        oParser.parseQueue.result = fSumm;
                    }
                }
            }
        }
        return oParser.parseQueue;
    };

    CCellRangeNode.prototype.getContentValue = function(oContent){
        var sString;
        oContent.SetApplyToAll(true);
        sString = oContent.GetSelectedText(false, {NewLineParagraph : true, NewLine : true});
        oContent.SetApplyToAll(false);
        return this.parseText(sString);
    };

    CCellRangeNode.prototype.getCell = function(oTable, nRow, nCol){
        var Row = oTable.Get_Row(nRow);
        if (!Row)
            return null;

        return Row.Get_Cell(nCol);
    };

    CCellRangeNode.prototype.calculateInRow = function(oTable, nRowIndex, nStart, nEnd){
        var oRow = oTable.GetRow(nRowIndex);
        if(!oRow){
            return;
        }
        var nCellsCount = oRow.Get_CellsCount();
        for(var i = nStart; i <= nEnd && i < nCellsCount; ++i){
            var oCurCell = oRow.Get_Cell(i);
            if(oCurCell){
                var res = this.getContentValue(oCurCell.GetContent());
                if(res && AscFormat.isRealNumber(res.result)){
                    this.result.push(res.result);
                }
                // else{
                //     this.result.push(0);
                // }
            }
        }
    };

    CCellRangeNode.prototype.calculateInCol = function(oTable, nColIndex, nStart, nEnd){
        var nRowsCount = oTable.Get_RowsCount();
        for(var i = nStart; i <= nEnd && i < nRowsCount; ++i){
            var oRow = oTable.Get_Row(i);
            if(oRow){
                var oCurCell = oRow.Get_Cell(nColIndex);
                if(oCurCell){
                    var res = this.getContentValue(oCurCell.GetContent());
                    if(res && AscFormat.isRealNumber(res.result)){
                        this.result.push(res.result);
                    }
                    // else{
                    //     this.result.push(0);
                    // }
                }
            }
        }
    };

    CCellRangeNode.prototype.calculateCellRange = function(oTable){
        this.result = [];

        var nStartCol, nEndCol, nStartRow, nEndRow;
        if(this.c1 !== null && this.c2 !== null){
            nStartCol = this.c1;
            nEndCol = this.c2;
        }
        else{
            nStartCol = 0;
            nEndCol = +Infinity;
        }
        if(this.r1 !== null && this.r2 !== null){
            nStartRow = this.r1;
            nEndRow = this.r2;
        }
        else{
            nStartRow = 0;
            nEndRow = oTable.Get_RowsCount() - 1;
        }
        for(var i = nStartRow; i <= nEndRow; ++i){
            this.calculateInRow(oTable, i, nStartCol, nEndCol);
        }
    };

    CCellRangeNode.prototype.calculateCell = function(oTable){
        var oCell, oContent;
        oCell = this.getCell(oTable, this.r1, this.c1);
        if(!oCell){
            var oFunction = this.inFunction();
            if(!oFunction){
                this.setError(this.getCellName(this.c1, this.r1) + " " + AscCommon.translateManager.getValue(ERROR_TYPE_IS_NOT_IN_TABLE), null);
            }
            else{
                if(this.c1 > 63 && oFunction.listSupport()){
                    this.setError(ERROR_TYPE_INDEX_TOO_LARGE, null);
                }
                else{
                    this.setError(this.getCellName(this.c1, this.r1) + " " + AscCommon.translateManager.getValue(ERROR_TYPE_IS_NOT_IN_TABLE), null);
                }
            }
            return;
        }
        oContent = oCell.GetContent();
        if(!oContent){
            this.result = 0.0;
            return;
        }
        var oRes = this.getContentValue(oContent);
        if(oRes && !AscFormat.isRealNumber(oRes.result)){
            this.result = 0.0;
        }
        else{
            this.result = oRes.result;
        }
    };


    CCellRangeNode.prototype._parseCellVal = function (oCurCell, oRes) {
        if(oCurCell){
            var res = this.getContentValue(oCurCell.GetContent());
            if(!res || !(res.flags & PARSER_MASK_CLEAN)){
                if(oRes.bClean === true && this.result.length > 0){
                    oRes.bBreak = true;
                }
                else{
                    oRes.bClean = false;
                    if(res && res.result !== null){
                        this.result.push(res.result);
                    }
                }
            }
            else{
                if(res.result !== null){
                    this.result.push(res.result);
                }
            }
        }
    };

    CCellRangeNode.prototype._calculate = function(){
        var oTable, oCell, oContent, oRow, i, oCurCell, oCurRow, nCellCount;
        if(this.isCell()){
            oTable = this.parseQueue.getParentTable();
            if(!oTable){
                this.result = 0.0;
                return
            }
            this.calculateCell(oTable);
        }
        else if(this.isDir()){
            oTable = this.parseQueue.getParentTable();
            if(!oTable){
                this.setError(ERROR_TYPE_FORMULA_NOT_IN_TABLE, null);
                return
            }
            oCell = this.parseQueue.getParentCell();
            if(!oCell){
                this.setError(ERROR_TYPE_FORMULA_NOT_IN_TABLE, null);
                return
            }
            oRow = oCell.GetRow();
            if(!oRow){
                this.setError(ERROR_TYPE_FORMULA_NOT_IN_TABLE, null);
                return
            }
            var oResult = {bClean: true, bBreak: false};
            if(this.dir === LEFT){
                if(oCell.Index === 0){
                    this.setError(ERROR_TYPE_INDEX_ZERO, null);
                    return;
                }
                this.result = [];
                for(i = oCell.Index - 1; i > -1; --i){
                    oCurCell = oRow.Get_Cell(i);
                    this._parseCellVal(oCurCell, oResult);
                    if(oResult.bBreak){
                        break;
                    }
                }
            }
            else if(this.dir === ABOVE){
                if(oRow.Index === 0){
                    this.setError(ERROR_TYPE_INDEX_ZERO, null);
                    return;
                }
                this.result = [];
                for(i = oRow.Index - 1; i > -1; --i){
                    oCurRow = oTable.Get_Row(i);
                    oCurCell = oCurRow.Get_Cell(oCell.Index);
                    this._parseCellVal(oCurCell, oResult);
                    if(oResult.bBreak){
                        break;
                    }
                }
            }
            else if(this.dir === RIGHT){
                this.result = [];
                nCellCount = oRow.Get_CellsCount();
                for(i = oCell.Index + 1; i < nCellCount; ++i){
                    oCurCell = oRow.Get_Cell(i);
                    this._parseCellVal(oCurCell, oResult);
                    if(oResult.bBreak){
                        break;
                    }
                }
            }
            else if(this.dir === BELOW){
                this.result = [];
                nCellCount = oTable.Get_RowsCount();
                for(i = oRow.Index + 1; i < nCellCount; ++i){
                    oCurRow = oTable.Get_Row(i);
                    oCurCell = oCurRow.Get_Cell(oCell.Index);
                    this._parseCellVal(oCurCell, oResult);
                    if(oResult.bBreak){
                        break;
                    }
                }
            }
            return;
        }
        else if(this.isCellRange()){
            this.calculateCellRange(this.parseQueue.getParentTable());
        }
        else if(this.isBookmark() || this.isBookmarkCell() || this.isBookmarkCellRange()){
            var oDocument = this.parseQueue.document;
            if(!oDocument || !oDocument.BookmarksManager){
                this.setError(ERROR_TYPE_ERROR, null);
                return;
            }
            oDocument.TurnOff_InterfaceEvents();
            var oSelectionState = oDocument.GetSelectionState();
            if(oDocument.BookmarksManager.SelectBookmark(this.bookmarkName)){
                var oCurrentParagraph = oDocument.GetCurrentParagraph();
                if(oCurrentParagraph.Parent){
                    oCell = oCurrentParagraph.Parent.IsTableCellContent(true);
                    if(oCell){
                        oRow = oCell.GetRow();
                        if(oRow){
                            oTable = oRow.GetTable();
                        }
                    }
                    if(oTable && !oTable.IsCellSelection()){
                        oTable = null;
                    }
                    if(this.isBookmark()){
                        if(!oTable){
                            var sString = oDocument.GetSelectedText(false, {NewLineParagraph : true, NewLine : true});
                            var oRes = this.parseText(sString);
                            if(oRes && !AscFormat.isRealNumber(oRes.result)){
                                this.result = 0.0;
                            }
                            else{
                                this.result = oRes.result;
                            }
                        }
                        else {
                            this.r1 = 0;
                            this.r2 = oTable.Get_RowsCount() - 1;
                            this.calculateCellRange(oTable);
                            this.r1 = null;
                            this.r2 = null;
                            var dResult = 0;
                            for(i = 0; i < this.result.length; ++i){
                                dResult += this.result[i];
                            }
                            this.result = dResult;
                        }
                    }
                    else{
                        if(oTable){
                            if(this.isBookmarkCell()){
                                this.calculateCell(oTable);
                            }
                            else {
                                this.calculateCellRange(oTable);
                            }
                        }
                        else{
                            var sData = "";
                            if(this.isBookmarkCell()){
                                sData = this.getCellName(this.c1, this.r1);
                            }
                            else {
                                if(this.c1 !== null && this.r1 !== null){
                                    sData = this.getCellName(this.c1, this.r1);
                                }
                                else{
                                    sData = ":"
                                }
                            }
                            this.setError(ERROR_TYPE_SYNTAX_ERROR, sData);
                        }
                    }
                }
                else {
                    this.setError(ERROR_TYPE_ERROR, null);
                }
            }
            else{
                this.setError(ERROR_TYPE_UNDEFINED_BOOKMARK, this.bookmarkName);
            }
            oDocument.SetSelectionState(oSelectionState);
            oDocument.TurnOn_InterfaceEvents(false);
        }
        else{
            this.setError(ERROR_TYPE_ERROR, null);
        }
    };

    CCellRangeNode.prototype.isCell = function(){
        return this.bookmarkName === null && this.c1 !== null && this.r1 !== null && this.c2 === null && this.r2 === null;
    };
    CCellRangeNode.prototype.isDir = function(){
        return this.dir !== null;
    };
    CCellRangeNode.prototype.isBookmarkCellRange = function(){
        return this.bookmarkName !== null && (this.c1 !== null || this.r1 !== null) && (this.c2 !== null || this.r2 !== null);
    };
    CCellRangeNode.prototype.isCellRange = function(){
        return this.bookmarkName === null && (this.c1 !== null || this.r1 !== null) && (this.c2 !== null || this.r2 !== null);
    };

    CCellRangeNode.prototype.isBookmarkCell = function(){
        return this.bookmarkName !== null && this.c1 !== null && this.r1 !== null && this.c2 === null && this.r2 === null;
    };
    CCellRangeNode.prototype.isBookmarkCellRange = function(){
        return this.bookmarkName !== null && (this.c1 !== null || this.r1 !== null) && (this.c2 !== null || this.r2 !== null);
    };
    CCellRangeNode.prototype.isBookmark = function(){
        return this.bookmarkName !== null && this.c1 === null && this.r1 === null && this.c2 === null && this.r2 === null;
    };

    function CFunctionNode(parseQueue){
        CFormulaNode.call(this, parseQueue);
        this.operands = [];
    }

    CFunctionNode.prototype = Object.create(CFormulaNode.prototype);
    CFunctionNode.prototype.precedence = 14;
    CFunctionNode.prototype.minArgumentsCount = 0;
    CFunctionNode.prototype.maxArgumentsCount = 0;
    CFunctionNode.prototype.isFunction = function () {
        return true;
    };
    CFunctionNode.prototype.listSupport = function () {
        return false;
    };

    function CABSFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CABSFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CABSFunctionNode.prototype.minArgumentsCount = 1;
    CABSFunctionNode.prototype.maxArgumentsCount = 1;
    CABSFunctionNode.prototype._calculate = function (aArgs) {
        this.result = Math.abs(aArgs[0].result);
    };

    function CANDFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CANDFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CANDFunctionNode.prototype.minArgumentsCount = 2;
    CANDFunctionNode.prototype.maxArgumentsCount = 2;
    CANDFunctionNode.prototype._calculate = function (aArgs) {
        this.result = (AscFormat.fApproxEqual(aArgs[1].result, 0.0) || AscFormat.fApproxEqual(aArgs[0].result, 0.0)) ? 0.0 : 1.0;
    };

    function CAVERAGEFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CAVERAGEFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CAVERAGEFunctionNode.prototype.minArgumentsCount = 2;
    CAVERAGEFunctionNode.prototype.maxArgumentsCount = +Infinity;
    CAVERAGEFunctionNode.prototype.listSupport = function () {
        return true;
    };
    CAVERAGEFunctionNode.prototype._calculate = function (aArgs) {
        var summ = 0.0;
        var count = 0;
        var result;
        for(var i = 0; i < aArgs.length; ++i){
            result = aArgs[i].result;
            if(Array.isArray(result)){
                for(var j = 0; j < result.length; ++j){
                    summ += result[j];
                }
                count += result.length;
            }
            else {
                summ += result;
                ++count;
            }
        }
        this.result = summ/count;
    };

    function CCOUNTFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CCOUNTFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CCOUNTFunctionNode.prototype.minArgumentsCount = 2;
    CCOUNTFunctionNode.prototype.maxArgumentsCount = +Infinity;
    CCOUNTFunctionNode.prototype.listSupport = function () {
        return true;
    };
    CCOUNTFunctionNode.prototype._calculate = function (aArgs) {
        var count = 0;
        var result;
        for(var i = 0; i < aArgs.length; ++i){
            result = aArgs[i].result;
            if(Array.isArray(result)){
                count += result.length;
            }
            else {
                ++count;
            }
        }
        this.result = count;
    };


    function CDEFINEDFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CDEFINEDFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CDEFINEDFunctionNode.prototype.minArgumentsCount = 1;
    CDEFINEDFunctionNode.prototype.maxArgumentsCount = 1;
    CDEFINEDFunctionNode.prototype.supportErrorArgs = function () {
        return true;
    };
    CDEFINEDFunctionNode.prototype._calculate = function (aArgs) {
        if(aArgs[0].error){
            this.result = 0.0;
        }
        else{
            this.result = 1.0;
        }
    };

    function CFALSEFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CFALSEFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CFALSEFunctionNode.prototype.minArgumentsCount = 0;
    CFALSEFunctionNode.prototype.maxArgumentsCount = 0;
    CFALSEFunctionNode.prototype._calculate = function (aArgs) {
        this.result = 0.0;
    };
    function CTRUEFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CTRUEFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CTRUEFunctionNode.prototype.minArgumentsCount = 0;
    CTRUEFunctionNode.prototype.maxArgumentsCount = 0;
    CTRUEFunctionNode.prototype._calculate = function (aArgs) {
        this.result = 1.0;
    };

    function CINTFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CINTFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CINTFunctionNode.prototype.minArgumentsCount = 1;
    CINTFunctionNode.prototype.maxArgumentsCount = 1;
    CINTFunctionNode.prototype._calculate = function (aArgs) {
        this.result = (aArgs[0].result >> 0);
    };
    function CIFFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CIFFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CIFFunctionNode.prototype.minArgumentsCount = 3;
    CIFFunctionNode.prototype.maxArgumentsCount = 3;
    CIFFunctionNode.prototype._calculate = function (aArgs) {
        this.result = ((aArgs[2].result !== 0.0) ? aArgs[1].result : aArgs[0].result);
    };
    function CMAXFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CMAXFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CMAXFunctionNode.prototype.minArgumentsCount = 2;
    CMAXFunctionNode.prototype.maxArgumentsCount = +Infinity;
    CMAXFunctionNode.prototype.listSupport = function () {
        return true;
    };
    CMAXFunctionNode.prototype._calculate = function (aArgs) {
        var ret = null;
        for(var i = 0; i < aArgs.length; ++i){
            if(!aArgs[i].error){
                if(!Array.isArray(aArgs[i].result)){
                    if(ret === null){
                        ret = aArgs[i].result;
                    }
                    else{
                        ret = Math.max(ret, aArgs[i].result);
                    }
                }
                else{
                    for(var j = 0; j < aArgs[i].result.length; ++j){
                        if(ret === null){
                            ret = aArgs[i].result[j];
                        }
                        else{
                            ret = Math.max(ret, aArgs[i].result[j]);
                        }
                    }
                }
            }
        }
        if(ret !== null){
            this.result = ret;
        }
        else{
            this.result = 0.0;
        }
    };

    function CMINFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CMINFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CMINFunctionNode.prototype.minArgumentsCount = 2;
    CMINFunctionNode.prototype.maxArgumentsCount = +Infinity;
    CMINFunctionNode.prototype.listSupport = function () {
        return true;
    };
    CMINFunctionNode.prototype._calculate = function (aArgs) {
        var ret = null;
        for(var i = 0; i < aArgs.length; ++i){
            if(!aArgs[i].error){
                if(!Array.isArray(aArgs[i].result)){
                    if(ret === null){
                        ret = aArgs[i].result;
                    }
                    else{
                        ret = Math.min(ret, aArgs[i].result);
                    }
                }
                else{
                    for(var j = 0; j < aArgs[i].result.length; ++j){
                        if(ret === null){
                            ret = aArgs[i].result[j];
                        }
                        else{
                            ret = Math.min(ret, aArgs[i].result[j]);
                        }
                    }
                }
            }
        }
        if(ret !== null){
            this.result = ret;
        }
        else{
            this.result = 0.0;
        }
    };

    function CMODFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CMODFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CMODFunctionNode.prototype.minArgumentsCount = 2;
    CMODFunctionNode.prototype.maxArgumentsCount = 2;
    CMODFunctionNode.prototype._calculate = function (aArgs) {
        this.result = (aArgs[1].result % aArgs[0].result);
    };

    function CNOTFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CNOTFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CNOTFunctionNode.prototype.minArgumentsCount = 1;
    CNOTFunctionNode.prototype.maxArgumentsCount = 1;
    CNOTFunctionNode.prototype._calculate = function (aArgs) {
        this.result = AscFormat.fApproxEqual(aArgs[0].result, 0.0) ? 1 : 0;
    };

    function CORFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CORFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CORFunctionNode.prototype.minArgumentsCount = 2;
    CORFunctionNode.prototype.maxArgumentsCount = 2;
    CORFunctionNode.prototype._calculate = function (aArgs) {
        this.result = (!AscFormat.fApproxEqual(aArgs[1].result, 0.0) || !AscFormat.fApproxEqual(aArgs[0].result, 0.0)) ? 1.0 : 0.0;
    };

    function CPRODUCTFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CPRODUCTFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CPRODUCTFunctionNode.prototype.minArgumentsCount = 2;
    CPRODUCTFunctionNode.prototype.maxArgumentsCount = +Infinity;
    CPRODUCTFunctionNode.prototype.listSupport = function () {
        return true;
    };
    CPRODUCTFunctionNode.prototype._calculate = function (aArgs) {
        if(aArgs.length === 0){
            return 0.0;
        }
        var ret = null;
        for(var i = 0; i < aArgs.length; ++i){
            if(!aArgs[i].error){
                if(!Array.isArray(aArgs[i].result)){
                    if(ret === null){
                        ret = aArgs[i].result;
                    }
                    else{
                        ret *= aArgs[i].result;
                    }
                }
                else{
                    for(var j = 0; j < aArgs[i].result.length; ++j){
                        if(ret === null){
                            ret = aArgs[i].result[j];
                        }
                        else{
                            ret *= aArgs[i].result[j];
                        }
                    }
                }
            }
        }
        if(ret !== null){
            this.result = ret;
        }
        else{
            this.result = 0.0;
        }
    };

    function fRoundNumber(number, precision){
        if (precision == 0) {
            return Math.round(number);
        }
        var sign = 1;
        if (number < 0) {
            sign = -1;
            number = Math.abs(number);
        }
        number = number.toString().split('e');
        number = Math.round(+(number[0] + 'e' + (number[1] ? (+number[1] + precision) : precision)));
        number = number.toString().split('e');
        return (+(number[0] + 'e' + (number[1] ? (+number[1] - precision) : -precision)) * sign);
    }

    function CROUNDFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CROUNDFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CROUNDFunctionNode.prototype.minArgumentsCount = 2;
    CROUNDFunctionNode.prototype.maxArgumentsCount = 2;
    CROUNDFunctionNode.prototype._calculate = function (aArgs) {
        this.result = fRoundNumber(aArgs[1].result, (aArgs[0].result >> 0));
    };

    function CSIGNFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CSIGNFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CSIGNFunctionNode.prototype.minArgumentsCount = 1;
    CSIGNFunctionNode.prototype.maxArgumentsCount = 1;
    CSIGNFunctionNode.prototype._calculate = function (aArgs) {
        if(aArgs[0].result < 0.0){
            this.result = -1.0;
        }
        else if(aArgs[0].result > 0.0){
            this.result = 1.0;
        }
        else{
            this.result = 0.0;
        }
    };

    function CSUMFunctionNode(parseQueue){
        CFunctionNode.call(this, parseQueue);
    }
    CSUMFunctionNode.prototype = Object.create(CFunctionNode.prototype);
    CSUMFunctionNode.prototype.minArgumentsCount = 2;
    CSUMFunctionNode.prototype.maxArgumentsCount = +Infinity;
    CSUMFunctionNode.prototype.listSupport = function () {
        return true;
    };

    CSUMFunctionNode.prototype._calculate = function (aArgs) {
        if(aArgs.length === 0){
            return 0.0;
        }
        var ret = null;
        for(var i = 0; i < aArgs.length; ++i){
            if(!aArgs[i].error){
                if(!Array.isArray(aArgs[i].result)){
                    if(ret === null){
                        ret = aArgs[i].result;
                    }
                    else{
                        ret += aArgs[i].result;
                    }
                }
                else{
                    for(var j = 0; j < aArgs[i].result.length; ++j){
                        if(ret === null){
                            ret = aArgs[i].result[j];
                        }
                        else{
                            ret += aArgs[i].result[j];
                        }
                    }
                }
            }
        }
        if(ret !== null){
            this.result = ret;
        }
        else{
            this.result = 0.0;
        }
    };

    function CParseQueue() {
        this.queue = [];
        this.result = null;
        this.resultS = null;
        this.format = null;
        this.digitSeparator = null;
        this.pos = -1;
        this.formula = null;
        this.ParentContent = null;
        this.Document = null;
    }
    CParseQueue.prototype.add = function(oToken){
        oToken.setParseQueue(this);
        oToken.digitSeparator = this.digitSeparator;
        this.queue.push(oToken);
        this.pos = this.queue.length - 1;
    };
    CParseQueue.prototype.last = function(){
        return this.queue[this.queue.length - 1];
    };
    CParseQueue.prototype.getNext = function(){
        if(this.pos > -1){
            return this.queue[--this.pos];
        }
        return null;
    };

    CParseQueue.prototype.getParentTable = function(){
        var oCell = this.getParentCell();
        if(oCell){
            var oRow = oCell.Row;
            if(oRow){
                return oRow.Table;
            }
        }
        return null;
    };

    CParseQueue.prototype.getParentCell = function(){
        var oCell = this.ParentContent && this.ParentContent.IsTableCellContent(true);
        if(oCell){
            return oCell;
        }
        return null;
    };
    CParseQueue.prototype.setError = function(Type, Data){
        this.error = new CError(Type, Data);
    };
    CParseQueue.prototype.calculate = function(bFormat){
        this.pos = this.queue.length - 1;

        this.error = null;
        this.result = null;
        if(this.pos < 0){
            this.setError(ERROR_TYPE_ERROR, null);
            return this.error;
        }
        var oLastToken = this.queue[this.pos];
        oLastToken.calculate();
        if(bFormat !== false){
            this.resultS = oLastToken.formatResult();
        }
        this.error = oLastToken.error;
        this.result = oLastToken.result;
        return this.error;
    };
    function CError(Type, Data){
        this.Type = AscCommon.translateManager.getValue(Type);
        this.Data = Data;
    }


    var oFuncMap = {};
    oFuncMap['ABS'] = CABSFunctionNode;
    oFuncMap['AND'] = CANDFunctionNode;
    oFuncMap['AVERAGE'] = CAVERAGEFunctionNode;
    oFuncMap['COUNT'] = CCOUNTFunctionNode;
    oFuncMap['DEFINED'] = CDEFINEDFunctionNode;
    oFuncMap['FALSE'] = CFALSEFunctionNode;
    oFuncMap['IF'] = CIFFunctionNode;
    oFuncMap['INT'] = CINTFunctionNode;
    oFuncMap['MAX'] = CMAXFunctionNode;
    oFuncMap['MIN'] = CMINFunctionNode;
    oFuncMap['MOD'] = CMODFunctionNode;
    oFuncMap['NOT'] = CNOTFunctionNode;
    oFuncMap['OR'] = CORFunctionNode;
    oFuncMap['PRODUCT'] = CPRODUCTFunctionNode;
    oFuncMap['ROUND'] = CROUNDFunctionNode;
    oFuncMap['SIGN'] = CSIGNFunctionNode;
    oFuncMap['SUM'] = CSUMFunctionNode;
    oFuncMap['TRUE'] = CTRUEFunctionNode;

    var PARSER_MASK_LEFT_PAREN      = 1;
    var PARSER_MASK_RIGHT_PAREN     = 2;
    var PARSER_MASK_LIST_SEPARATOR  = 4;
    var PARSER_MASK_BINARY_OPERATOR = 8;
    var PARSER_MASK_UNARY_OPERATOR  = 16;
    var PARSER_MASK_NUMBER          = 32;
    var PARSER_MASK_FUNCTION        = 64;
    var PARSER_MASK_CELL_NAME       = 128;
    var PARSER_MASK_CELL_RANGE      = 256;
    var PARSER_MASK_BOOKMARK        = 512;
    var PARSER_MASK_BOOKMARK_CELL_REF = 1024;
    var PARSER_MASK_CLEAN = 2048;

    var ERROR_TYPE_SYNTAX_ERROR = "Syntax Error";
    var ERROR_TYPE_MISSING_OPERATOR = "Missing Operator";
    var ERROR_TYPE_MISSING_ARGUMENT = "Missing Argument";
    var ERROR_TYPE_LARGE_NUMBER = "Number Too Large To Format";
    var ERROR_TYPE_ZERO_DIVIDE = "Zero Divide";
    var ERROR_TYPE_IS_NOT_IN_TABLE = "Is Not In Table";
    var ERROR_TYPE_INDEX_TOO_LARGE = "Index Too Large";
    var ERROR_TYPE_FORMULA_NOT_IN_TABLE = "The Formula Not In Table";
    var ERROR_TYPE_INDEX_ZERO = "Table Index Cannot be Zero";
    var ERROR_TYPE_UNDEFINED_BOOKMARK = "Undefined Bookmark";
    var ERROR_TYPE_UNEXPECTED_END = "Unexpected End of Formula";
    var ERROR_TYPE_ERROR = "ERROR";

    function CFormulaParser(sListSeparator, sDigitSeparator){
        this.listSeparator = sListSeparator;
        this.digitSeparator = sDigitSeparator;

        this.formula = null;
        this.pos = 0;
        this.parseQueue = null;
        this.error = null;
        this.flags = 0;//[unary opearor, binary operator,]
    }

    CFormulaParser.prototype.setFlag = function(nMask, bVal){
        if(bVal){
            this.flags |= nMask;
        }
        else{
            this.flags &= (~nMask);
        }
    };

    CFormulaParser.prototype.checkExpression = function(oRegExp, fCallback){
        oRegExp.lastIndex = this.pos;
        var oRes = oRegExp.exec(this.formula);
        if(oRes && oRes.index === this.pos){
            var ret = fCallback.call(this, this.pos, oRegExp.lastIndex);
            this.pos = oRegExp.lastIndex;
            return ret;
        }
        return null;
    };


    CFormulaParser.prototype.parseNumber = function(nStartPos, nEndPos){
        var sNum = this.formula.slice(nStartPos, nEndPos);
        sNum = sNum.replace(sComma, '');
        var number = parseFloat(sNum);
        if(AscFormat.isRealNumber(number)){
            var ret = new CNumberNode(this.parseQueue);
            ret.value = number;
            return ret;
        }
        return null;
    };


    CFormulaParser.prototype.parseCoord = function(nStartPos, nEndPos, oMap, nBase){
        var nRet = 0;
        var nMultiplicator = 1;
        for(var i = nEndPos - 1; i >= nStartPos; --i){
            nRet += oMap[this.formula[i]]*nMultiplicator;
            nMultiplicator *= nBase;
        }
        return nRet;
    };

    CFormulaParser.prototype.parseCol = function(nStartPos, nEndPos){
        return this.parseCoord(nStartPos, nEndPos, oLettersMap, 26) - 1;
    };

    CFormulaParser.prototype.parseRow = function(nStartPos, nEndPos){
        return this.parseCoord(nStartPos, nEndPos, oDigitMap, 10) - 1;
    };

    CFormulaParser.prototype.parseCellName = function(){
        var c, r;
        c = this.checkExpression(oColumnNameRegExp, this.parseCol);
        if(c === null){
            return null;
        }
        r = this.checkExpression(oRowNameRegExp, this.parseRow);
        if(r === null){
            return null;
        }
        var oRet = new CCellRangeNode(this.parseQueue);
        oRet.c1 = c;
        oRet.r1 = r;
        return oRet;
    };

    CFormulaParser.prototype.parseCellCellRange = function (nStart, nEndPos) {

        var oFirstCell, oSecondCell;
        oFirstCell = this.checkExpression(oCellNameRegExp, this.parseCellName);
        if(oFirstCell === null){
            return null;
        }
        while (this.formula[this.pos] === ' '){
            ++this.pos;
        }
        ++this.pos;
        while (this.formula[this.pos] === ' '){
            ++this.pos;
        }
        oSecondCell = this.checkExpression(oCellNameRegExp, this.parseCellName);
        if(oSecondCell === null){
            return null;
        }
        var r1, r2, c1, c2;
        r1 = Math.min(oFirstCell.r1, oSecondCell.r1);
        r2 = Math.max(oFirstCell.r1, oSecondCell.r1);
        c1 = Math.min(oFirstCell.c1, oSecondCell.c1);
        c2 = Math.max(oFirstCell.c1, oSecondCell.c1);
        oFirstCell.r1 = r1;
        oFirstCell.r2 = r2;
        oFirstCell.c1 = c1;
        oFirstCell.c2 = c2;
        return oFirstCell;
    };

    CFormulaParser.prototype.parseRowRange = function(nStartPos, nEndPos){
        var r1, r2;
        r1 = this.checkExpression(oRowNameRegExp, this.parseRow);
        if(r1 === null){
            return null;
        }
        while (this.formula[this.pos] === ' '){
            ++this.pos;
        }
        ++this.pos;
        while (this.formula[this.pos] === ' '){
            ++this.pos;
        }
        r2 = this.checkExpression(oRowNameRegExp, this.parseRow);
        if(r2 === null){
            return null;
        }
        var oRet = new CCellRangeNode(this.parseQueue);
        oRet.r1 = Math.min(r1, r2);
        oRet.r2 = Math.max(r1, r2);
        return oRet;
    };

    CFormulaParser.prototype.parseColRange = function(nStartPos, nEndPos){
        var c1, c2;
        c1 = this.checkExpression(oColumnNameRegExp, this.parseCol);
        if(c1 === null){
            return null;
        }
        while (this.formula[this.pos] === ' '){
            ++this.pos;
        }
        ++this.pos;
        while (this.formula[this.pos] === ' '){
            ++this.pos;
        }
        c2 = this.checkExpression(oColumnNameRegExp, this.parseCol);
        if(c2 === null){
            return null;
        }
        var oRet = new CCellRangeNode(this.parseQueue);
        oRet.c1 = Math.min(c1, c2);
        oRet.c2 = Math.max(c1, c2);
        return oRet;
    };

    CFormulaParser.prototype.parseCellRange = function (nStartPos, nEndPos) {
        var oRet;
        oRet = this.checkExpression(oCellCellRangeRegExp, this.parseCellCellRange);
        if(oRet){
            return oRet;
        }
        oRet = this.checkExpression(oRowRangeRegExp, this.parseRowRange);
        if(oRet){
            return oRet;
        }
        oRet = this.checkExpression(oColRangeRegExp, this.parseColRange);
        if(oRet){
            return oRet;
        }
        return null;
    };

    CFormulaParser.prototype.parseCellRef = function(nStartPos, nEndPos){
        var oRet;
        oRet = this.checkExpression(oCellRangeRegExp, this.parseCellRange);
        if(oRet){
            return oRet;
        }
        oRet = this.checkExpression(oCellNameRegExp, this.parseCellName);
        if(oRet){
            return oRet;
        }
        return null;
    };

    CFormulaParser.prototype.parseBookmark = function (nStartPos, nEndPos) {
        var oRet = new CCellRangeNode(this.parseQueue);
        oRet.bookmarkName = this.formula.slice(nStartPos, nEndPos);
        return oRet;
    };

    CFormulaParser.prototype.parseBookmarkCellRef = function(nStartPos, nEndPos){

        var oResult = this.checkExpression(oBookmarkNameRegExp, this.parseBookmark);
        if(oResult === null){
            return null;
        }
        if(this.pos < nEndPos){
            while(this.formula[this.pos] === ' '){
                ++this.pos;
            }
            var oRes = this.checkExpression(oCellReferenceRegExp, this.parseCellRef);
            if(oRes){
                oRes.bookmarkName = oResult.bookmarkName;
                return oRes;
            }
        }
        return oResult;
    };


    CFormulaParser.prototype.parseOperator = function(nStartPos, nEndPos){
        var sOperator = this.formula.slice(nStartPos, nEndPos);
        if(sOperator === '-'){
            if(this.flags & PARSER_MASK_UNARY_OPERATOR){
                return new CUnaryMinusOperatorNode(this.parseQueue);
            }
            return new CSubtractionOperatorNode(this.parseQueue);
        }
        if(oOperatorsMap[sOperator]){
            return new oOperatorsMap[sOperator]();
        }
        return null;
    };

    CFormulaParser.prototype.parseFunction = function(nStartPos, nEndPos){
        var sFunction = this.formula.slice(nStartPos, nEndPos).toUpperCase();
        if(oFuncMap[sFunction]){
            return new oFuncMap[sFunction]();
        }
        return null;
    };

    CFormulaParser.prototype.parseCurrent = function () {
        //TODO: R1C1
        while(this.formula[this.pos] === ' '){
            ++this.pos;
        }
        if(this.pos >= this.formula.length){
            return null;
        }
        //check parentheses
        if(this.formula[this.pos] === '('){
            ++this.pos;
            return new CLeftParenOperatorNode(this.parseQueue);
        }
        if(this.formula[this.pos] === ')'){
            ++this.pos;
            return new CRightParenOperatorNode(this.parseQueue);
        }
        if(this.formula[this.pos] === this.listSeparator){
            ++this.pos;
            return new CListSeparatorNode(this.parseQueue);
        }
        var oRet;
        //check operators
        oRet = this.checkExpression(oOperatorRegExp, this.parseOperator);
        if(oRet){
            return oRet;
        }
        //check function
        oRet = this.checkExpression(oFunctionsRegExp, this.parseFunction);
        if(oRet){
            return oRet;
        }
        //directions
        if(this.formula.indexOf('LEFT', this.pos) === this.pos){
            this.pos += 'LEFT'.length;
            oRet = new CCellRangeNode(this.parseQueue);
            oRet.dir = LEFT;
            return oRet;
        }
        if(this.formula.indexOf('ABOVE', this.pos) === this.pos){
            this.pos += 'ABOVE'.length;
            oRet = new CCellRangeNode(this.parseQueue);
            oRet.dir = ABOVE;
            return oRet;
        }
        if(this.formula.indexOf('BELOW', this.pos) === this.pos){
            this.pos += 'BELOW'.length;
            oRet = new CCellRangeNode(this.parseQueue);
            oRet.dir = BELOW;
            return oRet;
        }
        if(this.formula.indexOf('RIGHT', this.pos) === this.pos){
            this.pos += 'RIGHT'.length;
            oRet = new CCellRangeNode(this.parseQueue);
            oRet.dir = RIGHT;
            return oRet;
        }
        //check cell reference
        var oRes = this.checkExpression(oCellReferenceRegExp, this.parseCellRef);
        if(oRes){
            return oRes;
        }
        //check number
        oRet = this.checkExpression(oConstantRegExp, this.parseNumber);
        if(oRet){
            return oRet;
        }
        //check bookmark
        oRet = this.checkExpression(oBookmarkCellRefRegExp, this.parseBookmarkCellRef);
        if(oRet){
            return oRet;
        }
        return null;
    };

    CFormulaParser.prototype.setError = function(Type, Data){
        this.parseQueue = null;
        this.error = new CError(Type, Data);
    };

    CFormulaParser.prototype.getErrorString = function(startPos, endPos){
        var nStartPos = startPos;
        while (nStartPos < this.formula.length){
            if(this.formula[nStartPos] === ' '){
                nStartPos++;
            }
            else{
                break;
            }
        }
        if(nStartPos < endPos){
            return this.formula.slice(nStartPos, endPos);
        }
        return "";
    };

    CFormulaParser.prototype.checkSingularToken = function(oToken){
        if(oToken instanceof CNumberNode || oToken instanceof CCellRangeNode
            || oToken instanceof CRightParenOperatorNode || oToken instanceof CFALSEFunctionNode
            || oToken instanceof CTRUEFunctionNode || oToken instanceof CPercentageOperatorNode){
            return true;
        }
        return false;
    };

    CFormulaParser.prototype.parse = function(sFormula, oParentContent){
        if(typeof sFormula !== "string"){
            this.setError(ERROR_TYPE_ERROR, null);
            return;
        }
        this.pos = 0;
        this.formula = sFormula.toUpperCase();
        this.parseQueue = null;
        this.error = null;


        this.parseQueue = new CParseQueue();
        this.parseQueue.formula = this.formula;
        this.parseQueue.digitSeparator = this.digitSeparator;
        this.parseQueue.ParentContent = oParentContent;
        this.parseQueue.document = editor.WordControl.m_oLogicDocument;
        var oCurToken;
        var aStack = [];
        var aFunctionsStack = [];
        var oLastToken = null;
        var oLastFunction = null;
        var oToken;
        this.setFlag(PARSER_MASK_LEFT_PAREN, true);
        this.setFlag(PARSER_MASK_RIGHT_PAREN, false);
        this.setFlag(PARSER_MASK_LIST_SEPARATOR, false);
        this.setFlag(PARSER_MASK_BINARY_OPERATOR, false);
        this.setFlag(PARSER_MASK_UNARY_OPERATOR, true);
        this.setFlag(PARSER_MASK_NUMBER, true);
        this.setFlag(PARSER_MASK_FUNCTION, true);
        this.setFlag(PARSER_MASK_CELL_NAME, true);
        this.setFlag(PARSER_MASK_CELL_RANGE, false);
        this.setFlag(PARSER_MASK_BOOKMARK, true);
        this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, false);
        var nStartPos = this.pos;
        if(sFormula === ''){
            this.setFlag(PARSER_MASK_CLEAN, false);
        }
        while (oCurToken = this.parseCurrent()){
            if(oCurToken instanceof CNumberNode || oCurToken instanceof CFALSEFunctionNode || oCurToken instanceof CTRUEFunctionNode){
                if(this.checkSingularToken(oLastToken)){
                    if(oCurToken instanceof CNumberNode && oLastToken instanceof CNumberNode){
                        this.setError(ERROR_TYPE_MISSING_OPERATOR, null);
                    }
                    else{
                        this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                    }
                    return;
                }
                if(this.flags & PARSER_MASK_NUMBER){
                    this.parseQueue.add(oCurToken);
                    oCurToken.calculate();
                    if(oCurToken.error){
                        this.error = oCurToken.error;
                        return;
                    }
                    this.setFlag(PARSER_MASK_NUMBER, true);
                    this.setFlag(PARSER_MASK_UNARY_OPERATOR, false);
                    this.setFlag(PARSER_MASK_LEFT_PAREN, false);
                    this.setFlag(PARSER_MASK_RIGHT_PAREN, true);
                    this.setFlag(PARSER_MASK_BINARY_OPERATOR, true);
                    this.setFlag(PARSER_MASK_FUNCTION, false);
                    this.setFlag(PARSER_MASK_LIST_SEPARATOR, aFunctionsStack.length > 0);
                    this.setFlag(PARSER_MASK_CELL_NAME, false);
                    this.setFlag(PARSER_MASK_CELL_RANGE, false);
                    this.setFlag(PARSER_MASK_BOOKMARK, true);
                    this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, false);
                }
                else{
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                    return;
                }
            }
            else if(oCurToken instanceof CCellRangeNode){
                if((oCurToken.isDir() || oCurToken.isCellRange()) && !(this.flags & PARSER_MASK_CELL_RANGE)){
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, ':');
                    return;
                }
                if(oCurToken.isBookmarkCellRange() && !(this.flags & PARSER_MASK_BOOKMARK_CELL_REF)){
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, ':');
                    return;
                }

                if((oCurToken.isCell() || oCurToken.isBookmark()) && this.checkSingularToken(oLastToken)){
                    if(oLastToken instanceof CPercentageOperatorNode){
                        this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                    }
                    else{
                        this.setError(ERROR_TYPE_MISSING_OPERATOR, null);
                    }
                    return;
                }
                this.parseQueue.add(oCurToken);

                if(aFunctionsStack.length === 0){
                    oCurToken.calculate();

                    if(oCurToken.error){
                        this.error = oCurToken.error;
                        return;
                    }
                }
                else{
                    // var oFunction = aFunctionsStack[aFunctionsStack.length - 1];
                    // if(oCurToken.isCell() && !oFunction.listSupport()){
                    //     oCurToken.calculate();
                    //     if(oCurToken.error){
                    //         this.error = oCurToken.error;
                    //         return;
                    //     }
                    // }
                }
                this.setFlag(PARSER_MASK_NUMBER, true);
                this.setFlag(PARSER_MASK_UNARY_OPERATOR, false);
                this.setFlag(PARSER_MASK_LEFT_PAREN, false);
                this.setFlag(PARSER_MASK_RIGHT_PAREN, true);
                this.setFlag(PARSER_MASK_BINARY_OPERATOR, oCurToken.isCell() || oCurToken.isBookmark());
                this.setFlag(PARSER_MASK_FUNCTION, false);
                this.setFlag(PARSER_MASK_LIST_SEPARATOR, aFunctionsStack.length > 0);
                this.setFlag(PARSER_MASK_CELL_NAME, false);
                this.setFlag(PARSER_MASK_CELL_RANGE, false);
                this.setFlag(PARSER_MASK_BOOKMARK, false);
                this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, false);
            }
            else if(oCurToken.isFunction()){
                if(this.flags & PARSER_MASK_FUNCTION){
                    aStack.push(oCurToken);
                    this.setFlag(PARSER_MASK_RIGHT_PAREN, false);
                    if(oCurToken.maxArgumentsCount < 1){
                        this.setFlag(PARSER_MASK_LEFT_PAREN, false);
                    }
                    else{
                        this.setFlag(PARSER_MASK_LEFT_PAREN, true);
                        this.setFlag(PARSER_MASK_RIGHT_PAREN, false);
                        this.setFlag(PARSER_MASK_LIST_SEPARATOR, false);
                        this.setFlag(PARSER_MASK_BINARY_OPERATOR, false);
                        this.setFlag(PARSER_MASK_UNARY_OPERATOR, false);
                        this.setFlag(PARSER_MASK_NUMBER, false);
                        this.setFlag(PARSER_MASK_FUNCTION, false);
                        this.setFlag(PARSER_MASK_CELL_NAME, false);
                        this.setFlag(PARSER_MASK_CELL_RANGE, false);
                        this.setFlag(PARSER_MASK_BOOKMARK, false);
                        this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, false);
                    }
                }
                else{
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                    return;
                }
            }
            else if(oCurToken instanceof CListSeparatorNode){
                if(!(this.flags & PARSER_MASK_LIST_SEPARATOR)){
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                    return;
                }
                else{
                    if(aFunctionsStack.length > 0){

                        while(aStack.length > 0 && !(aStack[aStack.length-1] instanceof CLeftParenOperatorNode)){
                            oToken = aStack.pop();

                            this.parseQueue.add(oToken);
                            oToken.calculate();
                            if(oToken.error){
                                this.error = oToken.error;
                                return;
                            }
                        }
                        if(aStack.length === 0){
                            this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                            return;
                        }
                        aFunctionsStack[aFunctionsStack.length-1].operands.push(this.parseQueue.last());
                        if(aFunctionsStack[aFunctionsStack.length-1].operands.length >= aFunctionsStack[aFunctionsStack.length-1].maxArgumentsCount){
                            this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                            return;
                        }
                    }
                    else{
                        this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                        return;
                    }

                    this.setFlag(PARSER_MASK_LEFT_PAREN, true);
                    this.setFlag(PARSER_MASK_RIGHT_PAREN, false);
                    this.setFlag(PARSER_MASK_LIST_SEPARATOR, false);
                    this.setFlag(PARSER_MASK_BINARY_OPERATOR, false);
                    this.setFlag(PARSER_MASK_UNARY_OPERATOR, true);
                    this.setFlag(PARSER_MASK_NUMBER, true);
                    this.setFlag(PARSER_MASK_FUNCTION, true);
                    this.setFlag(PARSER_MASK_CELL_NAME, true);
                    this.setFlag(PARSER_MASK_CELL_RANGE, aFunctionsStack.length > 0 && aFunctionsStack[aFunctionsStack.length-1].listSupport());
                    this.setFlag(PARSER_MASK_BOOKMARK, true);
                    this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, aFunctionsStack.length > 0 && aFunctionsStack[aFunctionsStack.length-1].listSupport());
                }
            }
            else if(oCurToken instanceof CLeftParenOperatorNode){
                if(this.flags && PARSER_MASK_LEFT_PAREN){
                    if(oLastToken && oLastToken.isFunction(oLastToken)){
                        aFunctionsStack.push(oLastToken);
                    }
                    this.setFlag(PARSER_MASK_LEFT_PAREN, true);
                    this.setFlag(PARSER_MASK_RIGHT_PAREN, true);
                    this.setFlag(PARSER_MASK_LIST_SEPARATOR, false);
                    this.setFlag(PARSER_MASK_BINARY_OPERATOR, false);
                    this.setFlag(PARSER_MASK_UNARY_OPERATOR, true);
                    this.setFlag(PARSER_MASK_NUMBER, true);
                    this.setFlag(PARSER_MASK_FUNCTION, true);
                    this.setFlag(PARSER_MASK_CELL_NAME, true);
                    this.setFlag(PARSER_MASK_CELL_RANGE, aFunctionsStack.length > 0 && aFunctionsStack[aFunctionsStack.length-1].listSupport());
                    this.setFlag(PARSER_MASK_BOOKMARK, false);
                    this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, aFunctionsStack.length > 0 && aFunctionsStack[aFunctionsStack.length-1].listSupport());
                }
                else{
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                    return;
                }
                aStack.push(oCurToken);
            }
            else if(oCurToken instanceof CRightParenOperatorNode){
                while(aStack.length > 0 && !(aStack[aStack.length-1] instanceof CLeftParenOperatorNode)){
                    oToken = aStack.pop();
                    this.parseQueue.add(oToken);
                    oToken.calculate();
                    if(oToken.error){
                        this.error = oToken.error;
                        return;
                    }
                }

                if(aStack.length === 0){
                    this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                    return;
                }
                aStack.pop();//remove left paren
                if(aStack[aStack.length-1] && aStack[aStack.length-1].isFunction()){
                    aFunctionsStack.pop();
                    aStack[aStack.length-1].operands.push(this.parseQueue.last());
                    oLastFunction = aStack[aStack.length-1];
                    if(oLastFunction.operands.length > oLastFunction.maxArgumentsCount){
                        this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                        return;
                    }
                    if(oLastFunction.operands.length < oLastFunction.minArgumentsCount){
                        if(!oLastFunction.listSupport()){
                            this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                            return;
                        }
                        else{
                            for(var i = 0; i < oLastFunction.operands.length; ++i){
                                var oOperand = oLastFunction.operands[i];
                                if(oOperand instanceof CCellRangeNode){
                                    if(oOperand.isCell() || oOperand.isCellRange() || oOperand.isBookmarkCellRange() || oOperand.isDir()){
                                        break;
                                    }
                                }
                            }
                            if(i === oLastFunction.operands.length){
                                this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                                return;
                            }
                        }
                    } else if (oLastFunction instanceof CSUMFunctionNode && oLastFunction.operands.length === 1 && oLastFunction.operands[0] instanceof CNumberNode) {
                      this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                      return;
                    }
                    oLastFunction.argumentsCount = oLastFunction.operands.length;
                    oToken = aStack.pop();
                    this.parseQueue.add(oToken);
                    oToken.calculate();
                    if(oToken.error){
                        this.error = oToken.error;
                        return;
                    }
                }
                this.setFlag(PARSER_MASK_NUMBER, true);
                this.setFlag(PARSER_MASK_UNARY_OPERATOR, false);
                this.setFlag(PARSER_MASK_LEFT_PAREN, false);
                this.setFlag(PARSER_MASK_RIGHT_PAREN, true);
                this.setFlag(PARSER_MASK_BINARY_OPERATOR, true);
                this.setFlag(PARSER_MASK_FUNCTION, false);
                this.setFlag(PARSER_MASK_LIST_SEPARATOR, aFunctionsStack.length > 0);
                this.setFlag(PARSER_MASK_CELL_NAME, false);
                this.setFlag(PARSER_MASK_CELL_RANGE, false);
                this.setFlag(PARSER_MASK_BOOKMARK, false);
                this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, false);
            }
            else if(oCurToken.isOperator()){
                if(oCurToken instanceof CUnaryMinusOperatorNode){
                    if(!(this.flags & PARSER_MASK_UNARY_OPERATOR)){
                        this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                        return;
                    }
                    this.setFlag(PARSER_MASK_UNARY_OPERATOR, false);
                }
                else{
                    if(!(this.flags & PARSER_MASK_BINARY_OPERATOR)){
                        this.setError(ERROR_TYPE_SYNTAX_ERROR, this.getErrorString(nStartPos, this.pos));
                        return;
                    }

                    this.setFlag(PARSER_MASK_UNARY_OPERATOR, true);
                }
                while(aStack.length > 0 && (!(aStack[aStack.length-1] instanceof CLeftParenOperatorNode) && aStack[aStack.length-1].precedence >= oCurToken.precedence)){
                    oToken = aStack.pop();
                    this.parseQueue.add(oToken);
                    oToken.calculate();
                    if(oToken.error){
                        this.error = oToken.error;
                        return;
                    }
                }
                if(oCurToken instanceof CPercentageOperatorNode){
                    this.parseQueue.add(oCurToken);
                    oCurToken.calculate();
                    if(oCurToken.error){
                        this.error = oCurToken.error;
                        return;
                    }
                    this.setFlag(PARSER_MASK_NUMBER, true);
                    this.setFlag(PARSER_MASK_UNARY_OPERATOR, false);
                    this.setFlag(PARSER_MASK_LEFT_PAREN, false);
                    this.setFlag(PARSER_MASK_RIGHT_PAREN, true);
                    this.setFlag(PARSER_MASK_BINARY_OPERATOR, true);
                    this.setFlag(PARSER_MASK_FUNCTION, false);
                    this.setFlag(PARSER_MASK_LIST_SEPARATOR, aFunctionsStack.length > 0);
                    this.setFlag(PARSER_MASK_CELL_NAME, false);
                    this.setFlag(PARSER_MASK_CELL_RANGE, false);
                    this.setFlag(PARSER_MASK_BOOKMARK, true);
                    this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, false);
                }
                else{
                    this.setFlag(PARSER_MASK_NUMBER, true);
                    this.setFlag(PARSER_MASK_LEFT_PAREN, true);
                    this.setFlag(PARSER_MASK_RIGHT_PAREN, false);
                    this.setFlag(PARSER_MASK_BINARY_OPERATOR, false);
                    this.setFlag(PARSER_MASK_FUNCTION, true);
                    this.setFlag(PARSER_MASK_LIST_SEPARATOR, false);
                    this.setFlag(PARSER_MASK_CELL_NAME, true);
                    this.setFlag(PARSER_MASK_CELL_RANGE, false);
                    this.setFlag(PARSER_MASK_BOOKMARK, true);
                    this.setFlag(PARSER_MASK_BOOKMARK_CELL_REF, false);
                    aStack.push(oCurToken);
                }
            }
            if(!(oCurToken instanceof CLineSeparatorOperatorNode)){
                oLastToken = oCurToken;
            }
            nStartPos = this.pos;
        }
        if(this.pos < this.formula.length){
            this.setFlag(PARSER_MASK_CLEAN, false);
            this.error = new CError(ERROR_TYPE_SYNTAX_ERROR, this.formula[this.pos]);
            return;
        }
        while (aStack.length > 0){
            oCurToken = aStack.pop();
            if(oCurToken instanceof CLeftParenOperatorNode){
                this.setError(ERROR_TYPE_UNEXPECTED_END, null);
                return;
            } else
            if(oCurToken instanceof CRightParenOperatorNode){
                this.setError(ERROR_TYPE_SYNTAX_ERROR, '');
                return;
            }
            this.parseQueue.add(oCurToken);
            oCurToken.calculate();
            if(oCurToken.error){
                this.error = oCurToken.error;
                return;
            }
        }
    };
    window['AscCommonWord'] = window['AscCommonWord'] || {};
    window['AscCommonWord'].CFormulaParser = CFormulaParser;


    function CTextParser(sListSeparator, sDigitSeparator){
        CFormulaParser.call(this, sListSeparator, sDigitSeparator);//TODO: take list separator and digits separator from settings
        this.clean =  true;
    }
    CTextParser.prototype = Object.create(CFormulaParser.prototype);
    CTextParser.prototype.checkSingularToken = function(oToken){
        return false;
    };

    CTextParser.prototype.parseCurrent = function () {
        while(this.formula[this.pos] === ' '){
            ++this.pos;
            this.setFlag(PARSER_MASK_CLEAN, false);
        }
        if(this.pos >= this.formula.length){
            return null;
        }
        //check parentheses
        if(this.formula[this.pos] === '('){
            ++this.pos;
            this.setFlag(PARSER_MASK_CLEAN, false);
            return new CLeftParenOperatorNode(this.parseQueue);
        }
        if(this.formula[this.pos] === ')'){
            ++this.pos;
            this.setFlag(PARSER_MASK_CLEAN, false);
            return new CRightParenOperatorNode(this.parseQueue);
        }
        if(this.formula[this.pos] === '\n' || this.formula[this.pos] === '\t' || this.formula[this.pos] === '\r'){
            ++this.pos;
            this.setFlag(PARSER_MASK_CLEAN, false);
            return new CLineSeparatorOperatorNode(this.parseQueue);
        }
        var oRet;
        //check operators
        oRet = this.checkExpression(oOperatorRegExp, this.parseOperator);
        if(oRet){
            if(!(oRet instanceof CUnaryMinusOperatorNode)){
                this.setFlag(PARSER_MASK_CLEAN, false);
            }
            return oRet;
        }

        //check bookmark
        oRet = this.checkExpression(oBookmarkCellRefRegExp, this.parseBookmarkCellRef);
        if(oRet){
            this.setFlag(PARSER_MASK_CLEAN, false);
            return new CLineSeparatorOperatorNode(this.parseQueue);
        }

        //check number
        oRet = this.checkExpression(oConstantRegExp, this.parseNumber);
        var oRes;
        if(oRet){
            return oRet;
        }
        return null;
    };
})();
//window['AscCommonWord'].createExpression();
