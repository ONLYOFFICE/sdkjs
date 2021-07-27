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
    var sBookmark = "[\\u{41}-\\u{5a}\\u{aa}\\u{ba}\\u{c0}-\\u{d6}\\u{d8}-\\u{de}\\u{100}\\u{102}\\u{104}\\u{106}\\u{108}\\u{10a}\\u{10c}\\u{10e}\\u{110}\\u{112}\\u{114}\\u{116}\\u{118}\\u{11a}\\u{11c}\\u{11e}\\u{120}\\u{122}\\u{124}\\u{126}\\u{128}\\u{12a}\\u{12c}\\u{12e}\\u{130}\\u{132}\\u{134}\\u{136}\\u{139}\\u{13b}\\u{13d}\\u{13f}\\u{141}\\u{143}\\u{145}\\u{147}\\u{14a}\\u{14c}\\u{14e}\\u{150}\\u{152}\\u{154}\\u{156}\\u{158}\\u{15a}\\u{15c}\\u{15e}\\u{160}\\u{162}\\u{164}\\u{166}\\u{168}\\u{16a}\\u{16c}\\u{16e}\\u{170}\\u{172}\\u{174}\\u{176}\\u{178}-\\u{179}\\u{17b}\\u{17d}\\u{181}-\\u{182}\\u{184}\\u{186}-\\u{187}\\u{189}-\\u{18b}\\u{18e}-\\u{191}\\u{193}-\\u{194}\\u{196}-\\u{198}\\u{19c}-\\u{19d}\\u{19f}-\\u{1a0}\\u{1a2}\\u{1a4}\\u{1a6}-\\u{1a7}\\u{1a9}\\u{1ac}\\u{1ae}-\\u{1af}\\u{1b1}-\\u{1b3}\\u{1b5}\\u{1b7}-\\u{1b8}\\u{1bc}\\u{1c4}-\\u{1c5}\\u{1c7}-\\u{1c8}\\u{1ca}-\\u{1cb}\\u{1cd}\\u{1cf}\\u{1d1}\\u{1d3}\\u{1d5}\\u{1d7}\\u{1d9}\\u{1db}\\u{1de}\\u{1e0}\\u{1e2}\\u{1e4}\\u{1e6}\\u{1e8}\\u{1ea}\\u{1ec}\\u{1ee}\\u{1f1}-\\u{1f2}\\u{1f4}\\u{1f6}-\\u{1f8}\\u{1fa}\\u{1fc}\\u{1fe}\\u{200}\\u{202}\\u{204}\\u{206}\\u{208}\\u{20a}\\u{20c}\\u{20e}\\u{210}\\u{212}\\u{214}\\u{216}\\u{218}\\u{21a}\\u{21c}\\u{21e}\\u{220}\\u{222}\\u{224}\\u{226}\\u{228}\\u{22a}\\u{22c}\\u{22e}\\u{230}\\u{232}\\u{23a}-\\u{23b}\\u{23d}-\\u{23e}\\u{241}\\u{243}-\\u{246}\\u{248}\\u{24a}\\u{24c}\\u{24e}\\u{2b0}-\\u{2c1}\\u{2c6}-\\u{2d1}\\u{2e0}-\\u{2e4}\\u{2ec}\\u{2ee}\\u{370}\\u{372}\\u{376}\\u{386}\\u{388}-\\u{38a}\\u{38c}\\u{38e}-\\u{38f}\\u{391}-\\u{3a1}\\u{3a3}-\\u{3ab}\\u{3cf}\\u{3d2}-\\u{3d4}\\u{3d8}\\u{3da}\\u{3dc}\\u{3de}\\u{3e0}\\u{3e2}\\u{3e4}\\u{3e6}\\u{3e8}\\u{3ea}\\u{3ec}\\u{3ee}\\u{3f4}\\u{3f7}\\u{3f9}-\\u{3fa}\\u{3fd}-\\u{42f}\\u{460}\\u{462}\\u{464}\\u{466}\\u{468}\\u{46a}\\u{46c}\\u{46e}\\u{470}\\u{472}\\u{474}\\u{476}\\u{478}\\u{47a}\\u{47c}\\u{47e}\\u{480}\\u{48a}\\u{48c}\\u{48e}\\u{490}\\u{492}\\u{494}\\u{496}\\u{498}\\u{49a}\\u{49c}\\u{49e}\\u{4a0}\\u{4a2}\\u{4a4}\\u{4a6}\\u{4a8}\\u{4aa}\\u{4ac}\\u{4ae}\\u{4b0}\\u{4b2}\\u{4b4}\\u{4b6}\\u{4b8}\\u{4ba}\\u{4bc}\\u{4be}\\u{4c0}-\\u{4c1}\\u{4c3}\\u{4c5}\\u{4c7}\\u{4c9}\\u{4cb}\\u{4cd}\\u{4d0}\\u{4d2}\\u{4d4}\\u{4d6}\\u{4d8}\\u{4da}\\u{4dc}\\u{4de}\\u{4e0}\\u{4e2}\\u{4e4}\\u{4e6}\\u{4e8}\\u{4ea}\\u{4ec}\\u{4ee}\\u{4f0}\\u{4f2}\\u{4f4}\\u{4f6}\\u{4f8}\\u{4fa}\\u{4fc}\\u{4fe}\\u{500}\\u{502}\\u{504}\\u{506}\\u{508}\\u{50a}\\u{50c}\\u{50e}\\u{510}\\u{512}\\u{514}\\u{516}\\u{518}\\u{51a}\\u{51c}\\u{51e}\\u{520}\\u{522}\\u{531}-\\u{556}\\u{5d0}-\\u{5ea}\\u{5f0}-\\u{5f2}\\u{621}-\\u{64a}\\u{66e}-\\u{66f}\\u{671}-\\u{6d3}\\u{6d5}\\u{6e5}-\\u{6e6}\\u{6ee}-\\u{6ef}\\u{6fa}-\\u{6fc}\\u{6ff}\\u{710}\\u{712}-\\u{72f}\\u{74d}-\\u{7a5}\\u{7b1}\\u{7ca}-\\u{7ea}\\u{7f4}-\\u{7f5}\\u{7fa}\\u{904}-\\u{939}\\u{93d}\\u{950}\\u{958}-\\u{961}\\u{971}-\\u{972}\\u{97b}-\\u{97f}\\u{985}-\\u{98c}\\u{98f}-\\u{990}\\u{993}-\\u{9a8}\\u{9aa}-\\u{9b0}\\u{9b2}\\u{9b6}-\\u{9b9}\\u{9bd}\\u{9ce}\\u{9dc}-\\u{9dd}\\u{9df}-\\u{9e1}\\u{9f0}-\\u{9f1}\\u{a05}-\\u{a0a}\\u{a0f}-\\u{a10}\\u{a13}-\\u{a28}\\u{a2a}-\\u{a30}\\u{a32}-\\u{a33}\\u{a35}-\\u{a36}\\u{a38}-\\u{a39}\\u{a59}-\\u{a5c}\\u{a5e}\\u{a72}-\\u{a74}\\u{a85}-\\u{a8d}\\u{a8f}-\\u{a91}\\u{a93}-\\u{aa8}\\u{aaa}-\\u{ab0}\\u{ab2}-\\u{ab3}\\u{ab5}-\\u{ab9}\\u{abd}\\u{ad0}\\u{ae0}-\\u{ae1}\\u{b05}-\\u{b0c}\\u{b0f}-\\u{b10}\\u{b13}-\\u{b28}\\u{b2a}-\\u{b30}\\u{b32}-\\u{b33}\\u{b35}-\\u{b39}\\u{b3d}\\u{b5c}-\\u{b5d}\\u{b5f}-\\u{b61}\\u{b71}\\u{b83}\\u{b85}-\\u{b8a}\\u{b8e}-\\u{b90}\\u{b92}-\\u{b95}\\u{b99}-\\u{b9a}\\u{b9c}\\u{b9e}-\\u{b9f}\\u{ba3}-\\u{ba4}\\u{ba8}-\\u{baa}\\u{bae}-\\u{bb9}\\u{bd0}\\u{c05}-\\u{c0c}\\u{c0e}-\\u{c10}\\u{c12}-\\u{c28}\\u{c2a}-\\u{c33}\\u{c35}-\\u{c39}\\u{c3d}\\u{c58}-\\u{c59}\\u{c60}-\\u{c61}\\u{c85}-\\u{c8c}\\u{c8e}-\\u{c90}\\u{c92}-\\u{ca8}\\u{caa}-\\u{cb3}\\u{cb5}-\\u{cb9}\\u{cbd}\\u{cde}\\u{ce0}-\\u{ce1}\\u{cf1}-\\u{cf2}\\u{d05}-\\u{d0c}\\u{d0e}-\\u{d10}\\u{d12}-\\u{d28}\\u{d2a}-\\u{d39}\\u{d3d}\\u{d60}-\\u{d61}\\u{d7a}-\\u{d7f}\\u{d85}-\\u{d96}\\u{d9a}-\\u{db1}\\u{db3}-\\u{dbb}\\u{dbd}\\u{dc0}-\\u{dc6}\\u{e01}-\\u{e30}\\u{e32}-\\u{e33}\\u{e40}-\\u{e46}\\u{e81}-\\u{e82}\\u{e84}\\u{e86}-\\u{e8a}\\u{e8c}-\\u{ea3}\\u{ea5}\\u{ea7}-\\u{eb0}\\u{eb2}-\\u{eb3}\\u{ebd}\\u{ec0}-\\u{ec4}\\u{ec6}\\u{edc}-\\u{edd}\\u{f00}\\u{f40}-\\u{f47}\\u{f49}-\\u{f6c}\\u{f88}-\\u{f8c}\\u{1000}-\\u{102a}\\u{103f}\\u{1050}-\\u{1055}\\u{105a}-\\u{105d}\\u{1061}\\u{1065}-\\u{1066}\\u{106e}-\\u{1070}\\u{1075}-\\u{1081}\\u{108e}\\u{10a0}-\\u{10c5}\\u{1100}-\\u{11ff}\\u{1780}-\\u{17b3}\\u{17d7}\\u{17dc}\\u{1820}-\\u{1877}\\u{1880}-\\u{1884}\\u{1887}-\\u{18a8}\\u{18aa}\\u{1900}-\\u{191c}\\u{1950}-\\u{196d}\\u{1970}-\\u{1974}\\u{1980}-\\u{19a9}\\u{19b0}-\\u{19c9}\\u{1d62}-\\u{1d6a}\\u{1e00}\\u{1e02}\\u{1e04}\\u{1e06}\\u{1e08}\\u{1e0a}\\u{1e0c}\\u{1e0e}\\u{1e10}\\u{1e12}\\u{1e14}\\u{1e16}\\u{1e18}\\u{1e1a}\\u{1e1c}\\u{1e1e}\\u{1e20}\\u{1e22}\\u{1e24}\\u{1e26}\\u{1e28}\\u{1e2a}\\u{1e2c}\\u{1e2e}\\u{1e30}\\u{1e32}\\u{1e34}\\u{1e36}\\u{1e38}\\u{1e3a}\\u{1e3c}\\u{1e3e}\\u{1e40}\\u{1e42}\\u{1e44}\\u{1e46}\\u{1e48}\\u{1e4a}\\u{1e4c}\\u{1e4e}\\u{1e50}\\u{1e52}\\u{1e54}\\u{1e56}\\u{1e58}\\u{1e5a}\\u{1e5c}\\u{1e5e}\\u{1e60}\\u{1e62}\\u{1e64}\\u{1e66}\\u{1e68}\\u{1e6a}\\u{1e6c}\\u{1e6e}\\u{1e70}\\u{1e72}\\u{1e74}\\u{1e76}\\u{1e78}\\u{1e7a}\\u{1e7c}\\u{1e7e}\\u{1e80}\\u{1e82}\\u{1e84}\\u{1e86}\\u{1e88}\\u{1e8a}\\u{1e8c}\\u{1e8e}\\u{1e90}\\u{1e92}\\u{1e94}\\u{1e9e}\\u{1ea0}\\u{1ea2}\\u{1ea4}\\u{1ea6}\\u{1ea8}\\u{1eaa}\\u{1eac}\\u{1eae}\\u{1eb0}\\u{1eb2}\\u{1eb4}\\u{1eb6}\\u{1eb8}\\u{1eba}\\u{1ebc}\\u{1ebe}\\u{1ec0}\\u{1ec2}\\u{1ec4}\\u{1ec6}\\u{1ec8}\\u{1eca}\\u{1ecc}\\u{1ece}\\u{1ed0}\\u{1ed2}\\u{1ed4}\\u{1ed6}\\u{1ed8}\\u{1eda}\\u{1edc}\\u{1ede}\\u{1ee0}\\u{1ee2}\\u{1ee4}\\u{1ee6}\\u{1ee8}\\u{1eea}\\u{1eec}\\u{1eee}\\u{1ef0}\\u{1ef2}\\u{1ef4}\\u{1ef6}\\u{1ef8}\\u{1efa}\\u{1efc}\\u{1efe}\\u{1f08}-\\u{1f0f}\\u{1f18}-\\u{1f1d}\\u{1f28}-\\u{1f2f}\\u{1f38}-\\u{1f3f}\\u{1f48}-\\u{1f4d}\\u{1f59}\\u{1f5b}\\u{1f5d}\\u{1f5f}\\u{1f68}-\\u{1f6f}\\u{1f88}-\\u{1f8f}\\u{1f98}-\\u{1f9f}\\u{1fa8}-\\u{1faf}\\u{1fb8}-\\u{1fbc}\\u{1fc8}-\\u{1fcc}\\u{1fd8}-\\u{1fdb}\\u{1fe8}-\\u{1fec}\\u{1ff8}-\\u{1ffc}\\u{2071}\\u{207f}\\u{2102}\\u{2107}\\u{210b}-\\u{210d}\\u{2110}-\\u{2112}\\u{2115}\\u{2119}-\\u{211d}\\u{2124}\\u{2126}\\u{2128}\\u{212a}-\\u{212d}\\u{2130}-\\u{2133}\\u{213e}-\\u{213f}\\u{2145}\\u{2183}\\u{2c00}-\\u{2c2e}\\u{2c60}\\u{2c62}-\\u{2c64}\\u{2c67}\\u{2c69}\\u{2c6b}\\u{2c6d}-\\u{2c6f}\\u{2c72}\\u{2c75}\\u{2c7c}\\u{2c80}\\u{2c82}\\u{2c84}\\u{2c86}\\u{2c88}\\u{2c8a}\\u{2c8c}\\u{2c8e}\\u{2c90}\\u{2c92}\\u{2c94}\\u{2c96}\\u{2c98}\\u{2c9a}\\u{2c9c}\\u{2c9e}\\u{2ca0}\\u{2ca2}\\u{2ca4}\\u{2ca6}\\u{2ca8}\\u{2caa}\\u{2cac}\\u{2cae}\\u{2cb0}\\u{2cb2}\\u{2cb4}\\u{2cb6}\\u{2cb8}\\u{2cba}\\u{2cbc}\\u{2cbe}\\u{2cc0}\\u{2cc2}\\u{2cc4}\\u{2cc6}\\u{2cc8}\\u{2cca}\\u{2ccc}\\u{2cce}\\u{2cd0}\\u{2cd2}\\u{2cd4}\\u{2cd6}\\u{2cd8}\\u{2cda}\\u{2cdc}\\u{2cde}\\u{2ce0}\\u{2ce2}\\u{3005}-\\u{3006}\\u{3031}-\\u{3035}\\u{303b}-\\u{303c}\\u{3041}-\\u{3096}\\u{309d}-\\u{309f}\\u{30a1}-\\u{30fa}\\u{30fc}-\\u{30ff}\\u{3105}-\\u{312f}\\u{3131}-\\u{318e}\\u{31a0}-\\u{31ba}\\u{31f0}-\\u{31ff}\\u{3400}-\\u{4db5}\\u{4e00}-\\u{9fef}\\u{a000}-\\u{a48c}\\u{a640}\\u{a642}\\u{a644}\\u{a646}\\u{a648}\\u{a64a}\\u{a64c}\\u{a64e}\\u{a650}\\u{a652}\\u{a654}\\u{a656}\\u{a658}\\u{a65a}\\u{a65c}\\u{a65e}\\u{a662}\\u{a664}\\u{a666}\\u{a668}\\u{a66a}\\u{a66c}\\u{a680}\\u{a682}\\u{a684}\\u{a686}\\u{a688}\\u{a68a}\\u{a68c}\\u{a68e}\\u{a690}\\u{a692}\\u{a694}\\u{a696}\\u{a722}\\u{a724}\\u{a726}\\u{a728}\\u{a72a}\\u{a72c}\\u{a72e}\\u{a732}\\u{a734}\\u{a736}\\u{a738}\\u{a73a}\\u{a73c}\\u{a73e}\\u{a740}\\u{a742}\\u{a744}\\u{a746}\\u{a748}\\u{a74a}\\u{a74c}\\u{a74e}\\u{a750}\\u{a752}\\u{a754}\\u{a756}\\u{a758}\\u{a75a}\\u{a75c}\\u{a75e}\\u{a760}\\u{a762}\\u{a764}\\u{a766}\\u{a768}\\u{a76a}\\u{a76c}\\u{a76e}\\u{a779}\\u{a77b}\\u{a77d}-\\u{a77e}\\u{a780}\\u{a782}\\u{a784}\\u{a786}\\u{a78b}\\u{a960}-\\u{a97c}\\u{ac00}-\\u{d7a3}\\u{d7b0}-\\u{d7c6}\\u{d7cb}-\\u{d7fb}\\u{f900}-\\u{fa6d}\\u{fa70}-\\u{fad9}\\u{fb1d}\\u{fb1f}-\\u{fb28}\\u{fb2a}-\\u{fb36}\\u{fb38}-\\u{fb3c}\\u{fb3e}\\u{fb40}-\\u{fb41}\\u{fb43}-\\u{fb44}\\u{fb46}-\\u{fbb1}\\u{fbd3}-\\u{fd3d}\\u{fd50}-\\u{fd8f}\\u{fd92}-\\u{fdc7}\\u{fdf0}-\\u{fdfb}\\u{fe70}-\\u{fe74}\\u{fe76}-\\u{fefc}\\u{ff21}-\\u{ff3a}\\u{ff66}-\\u{ff9f}\\u{10000}-\\u{1000b}\\u{1000d}-\\u{10026}\\u{10028}-\\u{1003a}\\u{1003c}-\\u{1003d}\\u{1003f}-\\u{1004d}\\u{10050}-\\u{1005d}\\u{10080}-\\u{100fa}\\u{10280}-\\u{1029c}\\u{102a0}-\\u{102d0}\\u{10300}-\\u{1031f}\\u{1032d}-\\u{10340}\\u{10342}-\\u{10349}\\u{10350}-\\u{10375}\\u{10380}-\\u{1039d}\\u{103a0}-\\u{103c3}\\u{103c8}-\\u{103cf}\\u{10400}-\\u{10427}\\u{10450}-\\u{1049d}\\u{104b0}-\\u{104d3}\\u{10500}-\\u{10527}\\u{10530}-\\u{10563}\\u{10600}-\\u{10736}\\u{10740}-\\u{10755}\\u{10760}-\\u{10767}\\u{10800}-\\u{10805}\\u{10808}\\u{1080a}-\\u{10835}\\u{10837}-\\u{10838}\\u{1083c}\\u{1083f}-\\u{10855}\\u{10860}-\\u{10876}\\u{10880}-\\u{1089e}\\u{108e0}-\\u{108f2}\\u{108f4}-\\u{108f5}\\u{10900}-\\u{10915}\\u{10920}-\\u{10939}\\u{10980}-\\u{109b7}\\u{109be}-\\u{109bf}\\u{10a00}\\u{10a10}-\\u{10a13}\\u{10a15}-\\u{10a17}\\u{10a19}-\\u{10a35}\\u{10a60}-\\u{10a7c}\\u{10a80}-\\u{10a9c}\\u{10ac0}-\\u{10ac7}\\u{10ac9}-\\u{10ae4}\\u{10b00}-\\u{10b35}\\u{10b40}-\\u{10b55}\\u{10b60}-\\u{10b72}\\u{10b80}-\\u{10b91}\\u{10c00}-\\u{10c48}\\u{10c80}-\\u{10cb2}\\u{10d00}-\\u{10d23}\\u{10f00}-\\u{10f1c}\\u{10f27}\\u{10f30}-\\u{10f45}\\u{10fe0}-\\u{10ff6}\\u{11003}-\\u{11037}\\u{11083}-\\u{110af}\\u{110d0}-\\u{110e8}\\u{11103}-\\u{11126}\\u{11144}\\u{11150}-\\u{11172}\\u{11176}\\u{11183}-\\u{111b2}\\u{111c1}-\\u{111c4}\\u{111da}\\u{111dc}\\u{11200}-\\u{11211}\\u{11213}-\\u{1122b}\\u{11280}-\\u{11286}\\u{11288}\\u{1128a}-\\u{1128d}\\u{1128f}-\\u{1129d}\\u{1129f}-\\u{112a8}\\u{112b0}-\\u{112de}\\u{11305}-\\u{1130c}\\u{1130f}-\\u{11310}\\u{11313}-\\u{11328}\\u{1132a}-\\u{11330}\\u{11332}-\\u{11333}\\u{11335}-\\u{11339}\\u{1133d}\\u{11350}\\u{1135d}-\\u{11361}\\u{11400}-\\u{11434}\\u{11447}-\\u{1144a}\\u{1145f}\\u{11480}-\\u{114af}\\u{114c4}-\\u{114c5}\\u{114c7}\\u{11580}-\\u{115ae}\\u{115d8}-\\u{115db}\\u{11600}-\\u{1162f}\\u{11644}\\u{11680}-\\u{116aa}\\u{116b8}\\u{11700}-\\u{1171a}\\u{11800}-\\u{1182b}\\u{118a0}-\\u{118bf}\\u{118ff}\\u{119a0}-\\u{119a7}\\u{119aa}-\\u{119d0}\\u{119e1}\\u{119e3}\\u{11a00}\\u{11a0b}-\\u{11a32}\\u{11a3a}\\u{11a50}\\u{11a5c}-\\u{11a89}\\u{11a9d}\\u{11ac0}-\\u{11af8}\\u{11c00}-\\u{11c08}\\u{11c0a}-\\u{11c2e}\\u{11c40}\\u{11c72}-\\u{11c8f}\\u{11d00}-\\u{11d06}\\u{11d08}-\\u{11d09}\\u{11d0b}-\\u{11d30}\\u{11d46}\\u{11d60}-\\u{11d65}\\u{11d67}-\\u{11d68}\\u{11d6a}-\\u{11d89}\\u{11d98}\\u{11ee0}-\\u{11ef2}\\u{12000}-\\u{12399}\\u{12480}-\\u{12543}\\u{13000}-\\u{1342e}\\u{14400}-\\u{14646}\\u{16800}-\\u{16a38}\\u{16a40}-\\u{16a5e}\\u{16ad0}-\\u{16aed}\\u{16b00}-\\u{16b2f}\\u{16b40}-\\u{16b43}\\u{16b63}-\\u{16b77}\\u{16b7d}-\\u{16b8f}\\u{16e40}-\\u{16e5f}\\u{16f00}-\\u{16f4a}\\u{16f50}\\u{16f93}-\\u{16f9f}\\u{16fe0}-\\u{16fe1}\\u{16fe3}\\u{17000}-\\u{187f7}\\u{18800}-\\u{18af2}\\u{1b000}-\\u{1b11e}\\u{1b150}-\\u{1b152}\\u{1b164}-\\u{1b167}\\u{1b170}-\\u{1b2fb}\\u{1bc00}-\\u{1bc6a}\\u{1bc70}-\\u{1bc7c}\\u{1bc80}-\\u{1bc88}\\u{1bc90}-\\u{1bc99}\\u{1d400}-\\u{1d419}\\u{1d434}-\\u{1d44d}\\u{1d468}-\\u{1d481}\\u{1d49c}\\u{1d49e}-\\u{1d49f}\\u{1d4a2}\\u{1d4a5}-\\u{1d4a6}\\u{1d4a9}-\\u{1d4ac}\\u{1d4ae}-\\u{1d4b5}\\u{1d4d0}-\\u{1d4e9}\\u{1d504}-\\u{1d505}\\u{1d507}-\\u{1d50a}\\u{1d50d}-\\u{1d514}\\u{1d516}-\\u{1d51c}\\u{1d538}-\\u{1d539}\\u{1d53b}-\\u{1d53e}\\u{1d540}-\\u{1d544}\\u{1d546}\\u{1d54a}-\\u{1d550}\\u{1d56c}-\\u{1d585}\\u{1d5a0}-\\u{1d5b9}\\u{1d5d4}-\\u{1d5ed}\\u{1d608}-\\u{1d621}\\u{1d63c}-\\u{1d655}\\u{1d670}-\\u{1d689}\\u{1d6a8}-\\u{1d6c0}\\u{1d6e2}-\\u{1d6fa}\\u{1d71c}-\\u{1d734}\\u{1d756}-\\u{1d76e}\\u{1d790}-\\u{1d7a8}\\u{1d7ca}\\u{1e100}-\\u{1e12c}\\u{1e137}-\\u{1e13d}\\u{1e14e}\\u{1e2c0}-\\u{1e2eb}\\u{1e800}-\\u{1e8c4}\\u{1e900}-\\u{1e921}\\u{1e94b}\\u{1ee00}-\\u{1ee03}\\u{1ee05}-\\u{1ee1f}\\u{1ee21}-\\u{1ee22}\\u{1ee24}\\u{1ee27}\\u{1ee29}-\\u{1ee32}\\u{1ee34}-\\u{1ee37}\\u{1ee39}\\u{1ee3b}\\u{1ee42}\\u{1ee47}\\u{1ee49}\\u{1ee4b}\\u{1ee4d}-\\u{1ee4f}\\u{1ee51}-\\u{1ee52}\\u{1ee54}\\u{1ee57}\\u{1ee59}\\u{1ee5b}\\u{1ee5d}\\u{1ee5f}\\u{1ee61}-\\u{1ee62}\\u{1ee64}\\u{1ee67}-\\u{1ee6a}\\u{1ee6c}-\\u{1ee72}\\u{1ee74}-\\u{1ee77}\\u{1ee79}-\\u{1ee7c}\\u{1ee7e}\\u{1ee80}-\\u{1ee89}\\u{1ee8b}-\\u{1ee9b}\\u{1eea1}-\\u{1eea3}\\u{1eea5}-\\u{1eea9}\\u{1eeab}-\\u{1eebb}\\u{20000}-\\u{2a6d6}\\u{2a700}-\\u{2b734}\\u{2b740}-\\u{2b81d}\\u{2b820}-\\u{2cea1}\\u{2ceb0}-\\u{2ebe0}\\u{2f800}-\\u{2fa1d}]([\\u{41}-\\u{5a}\\u{aa}\\u{ba}\\u{c0}-\\u{d6}\\u{d8}-\\u{de}\\u{100}\\u{102}\\u{104}\\u{106}\\u{108}\\u{10a}\\u{10c}\\u{10e}\\u{110}\\u{112}\\u{114}\\u{116}\\u{118}\\u{11a}\\u{11c}\\u{11e}\\u{120}\\u{122}\\u{124}\\u{126}\\u{128}\\u{12a}\\u{12c}\\u{12e}\\u{130}\\u{132}\\u{134}\\u{136}\\u{139}\\u{13b}\\u{13d}\\u{13f}\\u{141}\\u{143}\\u{145}\\u{147}\\u{14a}\\u{14c}\\u{14e}\\u{150}\\u{152}\\u{154}\\u{156}\\u{158}\\u{15a}\\u{15c}\\u{15e}\\u{160}\\u{162}\\u{164}\\u{166}\\u{168}\\u{16a}\\u{16c}\\u{16e}\\u{170}\\u{172}\\u{174}\\u{176}\\u{178}-\\u{179}\\u{17b}\\u{17d}\\u{181}-\\u{182}\\u{184}\\u{186}-\\u{187}\\u{189}-\\u{18b}\\u{18e}-\\u{191}\\u{193}-\\u{194}\\u{196}-\\u{198}\\u{19c}-\\u{19d}\\u{19f}-\\u{1a0}\\u{1a2}\\u{1a4}\\u{1a6}-\\u{1a7}\\u{1a9}\\u{1ac}\\u{1ae}-\\u{1af}\\u{1b1}-\\u{1b3}\\u{1b5}\\u{1b7}-\\u{1b8}\\u{1bc}\\u{1c4}-\\u{1c5}\\u{1c7}-\\u{1c8}\\u{1ca}-\\u{1cb}\\u{1cd}\\u{1cf}\\u{1d1}\\u{1d3}\\u{1d5}\\u{1d7}\\u{1d9}\\u{1db}\\u{1de}\\u{1e0}\\u{1e2}\\u{1e4}\\u{1e6}\\u{1e8}\\u{1ea}\\u{1ec}\\u{1ee}\\u{1f1}-\\u{1f2}\\u{1f4}\\u{1f6}-\\u{1f8}\\u{1fa}\\u{1fc}\\u{1fe}\\u{200}\\u{202}\\u{204}\\u{206}\\u{208}\\u{20a}\\u{20c}\\u{20e}\\u{210}\\u{212}\\u{214}\\u{216}\\u{218}\\u{21a}\\u{21c}\\u{21e}\\u{220}\\u{222}\\u{224}\\u{226}\\u{228}\\u{22a}\\u{22c}\\u{22e}\\u{230}\\u{232}\\u{23a}-\\u{23b}\\u{23d}-\\u{23e}\\u{241}\\u{243}-\\u{246}\\u{248}\\u{24a}\\u{24c}\\u{24e}\\u{2b0}-\\u{2c1}\\u{2c6}-\\u{2d1}\\u{2e0}-\\u{2e4}\\u{2ec}\\u{2ee}\\u{370}\\u{372}\\u{376}\\u{386}\\u{388}-\\u{38a}\\u{38c}\\u{38e}-\\u{38f}\\u{391}-\\u{3a1}\\u{3a3}-\\u{3ab}\\u{3cf}\\u{3d2}-\\u{3d4}\\u{3d8}\\u{3da}\\u{3dc}\\u{3de}\\u{3e0}\\u{3e2}\\u{3e4}\\u{3e6}\\u{3e8}\\u{3ea}\\u{3ec}\\u{3ee}\\u{3f4}\\u{3f7}\\u{3f9}-\\u{3fa}\\u{3fd}-\\u{42f}\\u{460}\\u{462}\\u{464}\\u{466}\\u{468}\\u{46a}\\u{46c}\\u{46e}\\u{470}\\u{472}\\u{474}\\u{476}\\u{478}\\u{47a}\\u{47c}\\u{47e}\\u{480}\\u{48a}\\u{48c}\\u{48e}\\u{490}\\u{492}\\u{494}\\u{496}\\u{498}\\u{49a}\\u{49c}\\u{49e}\\u{4a0}\\u{4a2}\\u{4a4}\\u{4a6}\\u{4a8}\\u{4aa}\\u{4ac}\\u{4ae}\\u{4b0}\\u{4b2}\\u{4b4}\\u{4b6}\\u{4b8}\\u{4ba}\\u{4bc}\\u{4be}\\u{4c0}-\\u{4c1}\\u{4c3}\\u{4c5}\\u{4c7}\\u{4c9}\\u{4cb}\\u{4cd}\\u{4d0}\\u{4d2}\\u{4d4}\\u{4d6}\\u{4d8}\\u{4da}\\u{4dc}\\u{4de}\\u{4e0}\\u{4e2}\\u{4e4}\\u{4e6}\\u{4e8}\\u{4ea}\\u{4ec}\\u{4ee}\\u{4f0}\\u{4f2}\\u{4f4}\\u{4f6}\\u{4f8}\\u{4fa}\\u{4fc}\\u{4fe}\\u{500}\\u{502}\\u{504}\\u{506}\\u{508}\\u{50a}\\u{50c}\\u{50e}\\u{510}\\u{512}\\u{514}\\u{516}\\u{518}\\u{51a}\\u{51c}\\u{51e}\\u{520}\\u{522}\\u{531}-\\u{556}\\u{5d0}-\\u{5ea}\\u{5f0}-\\u{5f2}\\u{621}-\\u{64a}\\u{66e}-\\u{66f}\\u{671}-\\u{6d3}\\u{6d5}\\u{6e5}-\\u{6e6}\\u{6ee}-\\u{6ef}\\u{6fa}-\\u{6fc}\\u{6ff}\\u{710}\\u{712}-\\u{72f}\\u{74d}-\\u{7a5}\\u{7b1}\\u{7ca}-\\u{7ea}\\u{7f4}-\\u{7f5}\\u{7fa}\\u{904}-\\u{939}\\u{93d}\\u{950}\\u{958}-\\u{961}\\u{971}-\\u{972}\\u{97b}-\\u{97f}\\u{985}-\\u{98c}\\u{98f}-\\u{990}\\u{993}-\\u{9a8}\\u{9aa}-\\u{9b0}\\u{9b2}\\u{9b6}-\\u{9b9}\\u{9bd}\\u{9ce}\\u{9dc}-\\u{9dd}\\u{9df}-\\u{9e1}\\u{9f0}-\\u{9f1}\\u{a05}-\\u{a0a}\\u{a0f}-\\u{a10}\\u{a13}-\\u{a28}\\u{a2a}-\\u{a30}\\u{a32}-\\u{a33}\\u{a35}-\\u{a36}\\u{a38}-\\u{a39}\\u{a59}-\\u{a5c}\\u{a5e}\\u{a72}-\\u{a74}\\u{a85}-\\u{a8d}\\u{a8f}-\\u{a91}\\u{a93}-\\u{aa8}\\u{aaa}-\\u{ab0}\\u{ab2}-\\u{ab3}\\u{ab5}-\\u{ab9}\\u{abd}\\u{ad0}\\u{ae0}-\\u{ae1}\\u{b05}-\\u{b0c}\\u{b0f}-\\u{b10}\\u{b13}-\\u{b28}\\u{b2a}-\\u{b30}\\u{b32}-\\u{b33}\\u{b35}-\\u{b39}\\u{b3d}\\u{b5c}-\\u{b5d}\\u{b5f}-\\u{b61}\\u{b71}\\u{b83}\\u{b85}-\\u{b8a}\\u{b8e}-\\u{b90}\\u{b92}-\\u{b95}\\u{b99}-\\u{b9a}\\u{b9c}\\u{b9e}-\\u{b9f}\\u{ba3}-\\u{ba4}\\u{ba8}-\\u{baa}\\u{bae}-\\u{bb9}\\u{bd0}\\u{c05}-\\u{c0c}\\u{c0e}-\\u{c10}\\u{c12}-\\u{c28}\\u{c2a}-\\u{c33}\\u{c35}-\\u{c39}\\u{c3d}\\u{c58}-\\u{c59}\\u{c60}-\\u{c61}\\u{c85}-\\u{c8c}\\u{c8e}-\\u{c90}\\u{c92}-\\u{ca8}\\u{caa}-\\u{cb3}\\u{cb5}-\\u{cb9}\\u{cbd}\\u{cde}\\u{ce0}-\\u{ce1}\\u{cf1}-\\u{cf2}\\u{d05}-\\u{d0c}\\u{d0e}-\\u{d10}\\u{d12}-\\u{d28}\\u{d2a}-\\u{d39}\\u{d3d}\\u{d60}-\\u{d61}\\u{d7a}-\\u{d7f}\\u{d85}-\\u{d96}\\u{d9a}-\\u{db1}\\u{db3}-\\u{dbb}\\u{dbd}\\u{dc0}-\\u{dc6}\\u{e01}-\\u{e30}\\u{e32}-\\u{e33}\\u{e40}-\\u{e46}\\u{e81}-\\u{e82}\\u{e84}\\u{e86}-\\u{e8a}\\u{e8c}-\\u{ea3}\\u{ea5}\\u{ea7}-\\u{eb0}\\u{eb2}-\\u{eb3}\\u{ebd}\\u{ec0}-\\u{ec4}\\u{ec6}\\u{edc}-\\u{edd}\\u{f00}\\u{f40}-\\u{f47}\\u{f49}-\\u{f6c}\\u{f88}-\\u{f8c}\\u{1000}-\\u{102a}\\u{103f}\\u{1050}-\\u{1055}\\u{105a}-\\u{105d}\\u{1061}\\u{1065}-\\u{1066}\\u{106e}-\\u{1070}\\u{1075}-\\u{1081}\\u{108e}\\u{10a0}-\\u{10c5}\\u{1100}-\\u{11ff}\\u{1780}-\\u{17b3}\\u{17d7}\\u{17dc}\\u{1820}-\\u{1877}\\u{1880}-\\u{1884}\\u{1887}-\\u{18a8}\\u{18aa}\\u{1900}-\\u{191c}\\u{1950}-\\u{196d}\\u{1970}-\\u{1974}\\u{1980}-\\u{19a9}\\u{19b0}-\\u{19c9}\\u{1d62}-\\u{1d6a}\\u{1e00}\\u{1e02}\\u{1e04}\\u{1e06}\\u{1e08}\\u{1e0a}\\u{1e0c}\\u{1e0e}\\u{1e10}\\u{1e12}\\u{1e14}\\u{1e16}\\u{1e18}\\u{1e1a}\\u{1e1c}\\u{1e1e}\\u{1e20}\\u{1e22}\\u{1e24}\\u{1e26}\\u{1e28}\\u{1e2a}\\u{1e2c}\\u{1e2e}\\u{1e30}\\u{1e32}\\u{1e34}\\u{1e36}\\u{1e38}\\u{1e3a}\\u{1e3c}\\u{1e3e}\\u{1e40}\\u{1e42}\\u{1e44}\\u{1e46}\\u{1e48}\\u{1e4a}\\u{1e4c}\\u{1e4e}\\u{1e50}\\u{1e52}\\u{1e54}\\u{1e56}\\u{1e58}\\u{1e5a}\\u{1e5c}\\u{1e5e}\\u{1e60}\\u{1e62}\\u{1e64}\\u{1e66}\\u{1e68}\\u{1e6a}\\u{1e6c}\\u{1e6e}\\u{1e70}\\u{1e72}\\u{1e74}\\u{1e76}\\u{1e78}\\u{1e7a}\\u{1e7c}\\u{1e7e}\\u{1e80}\\u{1e82}\\u{1e84}\\u{1e86}\\u{1e88}\\u{1e8a}\\u{1e8c}\\u{1e8e}\\u{1e90}\\u{1e92}\\u{1e94}\\u{1e9e}\\u{1ea0}\\u{1ea2}\\u{1ea4}\\u{1ea6}\\u{1ea8}\\u{1eaa}\\u{1eac}\\u{1eae}\\u{1eb0}\\u{1eb2}\\u{1eb4}\\u{1eb6}\\u{1eb8}\\u{1eba}\\u{1ebc}\\u{1ebe}\\u{1ec0}\\u{1ec2}\\u{1ec4}\\u{1ec6}\\u{1ec8}\\u{1eca}\\u{1ecc}\\u{1ece}\\u{1ed0}\\u{1ed2}\\u{1ed4}\\u{1ed6}\\u{1ed8}\\u{1eda}\\u{1edc}\\u{1ede}\\u{1ee0}\\u{1ee2}\\u{1ee4}\\u{1ee6}\\u{1ee8}\\u{1eea}\\u{1eec}\\u{1eee}\\u{1ef0}\\u{1ef2}\\u{1ef4}\\u{1ef6}\\u{1ef8}\\u{1efa}\\u{1efc}\\u{1efe}\\u{1f08}-\\u{1f0f}\\u{1f18}-\\u{1f1d}\\u{1f28}-\\u{1f2f}\\u{1f38}-\\u{1f3f}\\u{1f48}-\\u{1f4d}\\u{1f59}\\u{1f5b}\\u{1f5d}\\u{1f5f}\\u{1f68}-\\u{1f6f}\\u{1f88}-\\u{1f8f}\\u{1f98}-\\u{1f9f}\\u{1fa8}-\\u{1faf}\\u{1fb8}-\\u{1fbc}\\u{1fc8}-\\u{1fcc}\\u{1fd8}-\\u{1fdb}\\u{1fe8}-\\u{1fec}\\u{1ff8}-\\u{1ffc}\\u{2071}\\u{207f}\\u{2102}\\u{2107}\\u{210b}-\\u{210d}\\u{2110}-\\u{2112}\\u{2115}\\u{2119}-\\u{211d}\\u{2124}\\u{2126}\\u{2128}\\u{212a}-\\u{212d}\\u{2130}-\\u{2133}\\u{213e}-\\u{213f}\\u{2145}\\u{2183}\\u{2c00}-\\u{2c2e}\\u{2c60}\\u{2c62}-\\u{2c64}\\u{2c67}\\u{2c69}\\u{2c6b}\\u{2c6d}-\\u{2c6f}\\u{2c72}\\u{2c75}\\u{2c7c}\\u{2c80}\\u{2c82}\\u{2c84}\\u{2c86}\\u{2c88}\\u{2c8a}\\u{2c8c}\\u{2c8e}\\u{2c90}\\u{2c92}\\u{2c94}\\u{2c96}\\u{2c98}\\u{2c9a}\\u{2c9c}\\u{2c9e}\\u{2ca0}\\u{2ca2}\\u{2ca4}\\u{2ca6}\\u{2ca8}\\u{2caa}\\u{2cac}\\u{2cae}\\u{2cb0}\\u{2cb2}\\u{2cb4}\\u{2cb6}\\u{2cb8}\\u{2cba}\\u{2cbc}\\u{2cbe}\\u{2cc0}\\u{2cc2}\\u{2cc4}\\u{2cc6}\\u{2cc8}\\u{2cca}\\u{2ccc}\\u{2cce}\\u{2cd0}\\u{2cd2}\\u{2cd4}\\u{2cd6}\\u{2cd8}\\u{2cda}\\u{2cdc}\\u{2cde}\\u{2ce0}\\u{2ce2}\\u{3005}-\\u{3006}\\u{3031}-\\u{3035}\\u{303b}-\\u{303c}\\u{3041}-\\u{3096}\\u{309d}-\\u{309f}\\u{30a1}-\\u{30fa}\\u{30fc}-\\u{30ff}\\u{3105}-\\u{312f}\\u{3131}-\\u{318e}\\u{31a0}-\\u{31ba}\\u{31f0}-\\u{31ff}\\u{3400}-\\u{4db5}\\u{4e00}-\\u{9fef}\\u{a000}-\\u{a48c}\\u{a640}\\u{a642}\\u{a644}\\u{a646}\\u{a648}\\u{a64a}\\u{a64c}\\u{a64e}\\u{a650}\\u{a652}\\u{a654}\\u{a656}\\u{a658}\\u{a65a}\\u{a65c}\\u{a65e}\\u{a662}\\u{a664}\\u{a666}\\u{a668}\\u{a66a}\\u{a66c}\\u{a680}\\u{a682}\\u{a684}\\u{a686}\\u{a688}\\u{a68a}\\u{a68c}\\u{a68e}\\u{a690}\\u{a692}\\u{a694}\\u{a696}\\u{a722}\\u{a724}\\u{a726}\\u{a728}\\u{a72a}\\u{a72c}\\u{a72e}\\u{a732}\\u{a734}\\u{a736}\\u{a738}\\u{a73a}\\u{a73c}\\u{a73e}\\u{a740}\\u{a742}\\u{a744}\\u{a746}\\u{a748}\\u{a74a}\\u{a74c}\\u{a74e}\\u{a750}\\u{a752}\\u{a754}\\u{a756}\\u{a758}\\u{a75a}\\u{a75c}\\u{a75e}\\u{a760}\\u{a762}\\u{a764}\\u{a766}\\u{a768}\\u{a76a}\\u{a76c}\\u{a76e}\\u{a779}\\u{a77b}\\u{a77d}-\\u{a77e}\\u{a780}\\u{a782}\\u{a784}\\u{a786}\\u{a78b}\\u{a960}-\\u{a97c}\\u{ac00}-\\u{d7a3}\\u{d7b0}-\\u{d7c6}\\u{d7cb}-\\u{d7fb}\\u{f900}-\\u{fa6d}\\u{fa70}-\\u{fad9}\\u{fb1d}\\u{fb1f}-\\u{fb28}\\u{fb2a}-\\u{fb36}\\u{fb38}-\\u{fb3c}\\u{fb3e}\\u{fb40}-\\u{fb41}\\u{fb43}-\\u{fb44}\\u{fb46}-\\u{fbb1}\\u{fbd3}-\\u{fd3d}\\u{fd50}-\\u{fd8f}\\u{fd92}-\\u{fdc7}\\u{fdf0}-\\u{fdfb}\\u{fe70}-\\u{fe74}\\u{fe76}-\\u{fefc}\\u{ff21}-\\u{ff3a}\\u{ff66}-\\u{ff9f}\\u{10000}-\\u{1000b}\\u{1000d}-\\u{10026}\\u{10028}-\\u{1003a}\\u{1003c}-\\u{1003d}\\u{1003f}-\\u{1004d}\\u{10050}-\\u{1005d}\\u{10080}-\\u{100fa}\\u{10280}-\\u{1029c}\\u{102a0}-\\u{102d0}\\u{10300}-\\u{1031f}\\u{1032d}-\\u{10340}\\u{10342}-\\u{10349}\\u{10350}-\\u{10375}\\u{10380}-\\u{1039d}\\u{103a0}-\\u{103c3}\\u{103c8}-\\u{103cf}\\u{10400}-\\u{10427}\\u{10450}-\\u{1049d}\\u{104b0}-\\u{104d3}\\u{10500}-\\u{10527}\\u{10530}-\\u{10563}\\u{10600}-\\u{10736}\\u{10740}-\\u{10755}\\u{10760}-\\u{10767}\\u{10800}-\\u{10805}\\u{10808}\\u{1080a}-\\u{10835}\\u{10837}-\\u{10838}\\u{1083c}\\u{1083f}-\\u{10855}\\u{10860}-\\u{10876}\\u{10880}-\\u{1089e}\\u{108e0}-\\u{108f2}\\u{108f4}-\\u{108f5}\\u{10900}-\\u{10915}\\u{10920}-\\u{10939}\\u{10980}-\\u{109b7}\\u{109be}-\\u{109bf}\\u{10a00}\\u{10a10}-\\u{10a13}\\u{10a15}-\\u{10a17}\\u{10a19}-\\u{10a35}\\u{10a60}-\\u{10a7c}\\u{10a80}-\\u{10a9c}\\u{10ac0}-\\u{10ac7}\\u{10ac9}-\\u{10ae4}\\u{10b00}-\\u{10b35}\\u{10b40}-\\u{10b55}\\u{10b60}-\\u{10b72}\\u{10b80}-\\u{10b91}\\u{10c00}-\\u{10c48}\\u{10c80}-\\u{10cb2}\\u{10d00}-\\u{10d23}\\u{10f00}-\\u{10f1c}\\u{10f27}\\u{10f30}-\\u{10f45}\\u{10fe0}-\\u{10ff6}\\u{11003}-\\u{11037}\\u{11083}-\\u{110af}\\u{110d0}-\\u{110e8}\\u{11103}-\\u{11126}\\u{11144}\\u{11150}-\\u{11172}\\u{11176}\\u{11183}-\\u{111b2}\\u{111c1}-\\u{111c4}\\u{111da}\\u{111dc}\\u{11200}-\\u{11211}\\u{11213}-\\u{1122b}\\u{11280}-\\u{11286}\\u{11288}\\u{1128a}-\\u{1128d}\\u{1128f}-\\u{1129d}\\u{1129f}-\\u{112a8}\\u{112b0}-\\u{112de}\\u{11305}-\\u{1130c}\\u{1130f}-\\u{11310}\\u{11313}-\\u{11328}\\u{1132a}-\\u{11330}\\u{11332}-\\u{11333}\\u{11335}-\\u{11339}\\u{1133d}\\u{11350}\\u{1135d}-\\u{11361}\\u{11400}-\\u{11434}\\u{11447}-\\u{1144a}\\u{1145f}\\u{11480}-\\u{114af}\\u{114c4}-\\u{114c5}\\u{114c7}\\u{11580}-\\u{115ae}\\u{115d8}-\\u{115db}\\u{11600}-\\u{1162f}\\u{11644}\\u{11680}-\\u{116aa}\\u{116b8}\\u{11700}-\\u{1171a}\\u{11800}-\\u{1182b}\\u{118a0}-\\u{118bf}\\u{118ff}\\u{119a0}-\\u{119a7}\\u{119aa}-\\u{119d0}\\u{119e1}\\u{119e3}\\u{11a00}\\u{11a0b}-\\u{11a32}\\u{11a3a}\\u{11a50}\\u{11a5c}-\\u{11a89}\\u{11a9d}\\u{11ac0}-\\u{11af8}\\u{11c00}-\\u{11c08}\\u{11c0a}-\\u{11c2e}\\u{11c40}\\u{11c72}-\\u{11c8f}\\u{11d00}-\\u{11d06}\\u{11d08}-\\u{11d09}\\u{11d0b}-\\u{11d30}\\u{11d46}\\u{11d60}-\\u{11d65}\\u{11d67}-\\u{11d68}\\u{11d6a}-\\u{11d89}\\u{11d98}\\u{11ee0}-\\u{11ef2}\\u{12000}-\\u{12399}\\u{12480}-\\u{12543}\\u{13000}-\\u{1342e}\\u{14400}-\\u{14646}\\u{16800}-\\u{16a38}\\u{16a40}-\\u{16a5e}\\u{16ad0}-\\u{16aed}\\u{16b00}-\\u{16b2f}\\u{16b40}-\\u{16b43}\\u{16b63}-\\u{16b77}\\u{16b7d}-\\u{16b8f}\\u{16e40}-\\u{16e5f}\\u{16f00}-\\u{16f4a}\\u{16f50}\\u{16f93}-\\u{16f9f}\\u{16fe0}-\\u{16fe1}\\u{16fe3}\\u{17000}-\\u{187f7}\\u{18800}-\\u{18af2}\\u{1b000}-\\u{1b11e}\\u{1b150}-\\u{1b152}\\u{1b164}-\\u{1b167}\\u{1b170}-\\u{1b2fb}\\u{1bc00}-\\u{1bc6a}\\u{1bc70}-\\u{1bc7c}\\u{1bc80}-\\u{1bc88}\\u{1bc90}-\\u{1bc99}\\u{1d400}-\\u{1d419}\\u{1d434}-\\u{1d44d}\\u{1d468}-\\u{1d481}\\u{1d49c}\\u{1d49e}-\\u{1d49f}\\u{1d4a2}\\u{1d4a5}-\\u{1d4a6}\\u{1d4a9}-\\u{1d4ac}\\u{1d4ae}-\\u{1d4b5}\\u{1d4d0}-\\u{1d4e9}\\u{1d504}-\\u{1d505}\\u{1d507}-\\u{1d50a}\\u{1d50d}-\\u{1d514}\\u{1d516}-\\u{1d51c}\\u{1d538}-\\u{1d539}\\u{1d53b}-\\u{1d53e}\\u{1d540}-\\u{1d544}\\u{1d546}\\u{1d54a}-\\u{1d550}\\u{1d56c}-\\u{1d585}\\u{1d5a0}-\\u{1d5b9}\\u{1d5d4}-\\u{1d5ed}\\u{1d608}-\\u{1d621}\\u{1d63c}-\\u{1d655}\\u{1d670}-\\u{1d689}\\u{1d6a8}-\\u{1d6c0}\\u{1d6e2}-\\u{1d6fa}\\u{1d71c}-\\u{1d734}\\u{1d756}-\\u{1d76e}\\u{1d790}-\\u{1d7a8}\\u{1d7ca}\\u{1e100}-\\u{1e12c}\\u{1e137}-\\u{1e13d}\\u{1e14e}\\u{1e2c0}-\\u{1e2eb}\\u{1e800}-\\u{1e8c4}\\u{1e900}-\\u{1e921}\\u{1e94b}\\u{1ee00}-\\u{1ee03}\\u{1ee05}-\\u{1ee1f}\\u{1ee21}-\\u{1ee22}\\u{1ee24}\\u{1ee27}\\u{1ee29}-\\u{1ee32}\\u{1ee34}-\\u{1ee37}\\u{1ee39}\\u{1ee3b}\\u{1ee42}\\u{1ee47}\\u{1ee49}\\u{1ee4b}\\u{1ee4d}-\\u{1ee4f}\\u{1ee51}-\\u{1ee52}\\u{1ee54}\\u{1ee57}\\u{1ee59}\\u{1ee5b}\\u{1ee5d}\\u{1ee5f}\\u{1ee61}-\\u{1ee62}\\u{1ee64}\\u{1ee67}-\\u{1ee6a}\\u{1ee6c}-\\u{1ee72}\\u{1ee74}-\\u{1ee77}\\u{1ee79}-\\u{1ee7c}\\u{1ee7e}\\u{1ee80}-\\u{1ee89}\\u{1ee8b}-\\u{1ee9b}\\u{1eea1}-\\u{1eea3}\\u{1eea5}-\\u{1eea9}\\u{1eeab}-\\u{1eebb}\\u{20000}-\\u{2a6d6}\\u{2a700}-\\u{2b734}\\u{2b740}-\\u{2b81d}\\u{2b820}-\\u{2cea1}\\u{2ceb0}-\\u{2ebe0}\\u{2f800}-\\u{2fa1d}0-9_]{0,39})";
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
    var oBookmarkNameRegExp = new RegExp(sBookmark, 'gu');
    var oBookmarkCellRefRegExp = new RegExp(sBookmarkCellRef, 'gu');
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
    CSUMFunctionNode.prototype.minArgumentsCount = 1;
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
                                if(oLastFunction.operands[i] instanceof CCellRangeNode){
                                    if(oLastFunction.operands[i].isCellRange() || oLastFunction.operands[i].isBookmarkCellRange() || oLastFunction.operands[i].isDir()){
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
