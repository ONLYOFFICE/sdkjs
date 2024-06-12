import re
import os

#https://www.unicode.org/Public/UNIDATA/UnicodeData.txt


data = open("UnicodeData.txt", "rt")
out = open("unicode-data.js", "wt")

header = "/*\n\
 * (c) Copyright Ascensio System SIA 2010-2024\n\
 *\n\
 * This program is a free software product. You can redistribute it and/or\n\
 * modify it under the terms of the GNU Affero General Public License (AGPL)\n\
 * version 3 as published by the Free Software Foundation. In accordance with\n\
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect\n\
 * that Ascensio System SIA expressly excludes the warranty of non-infringement\n\
 * of any third-party rights.\n\
 *\n\
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied\n\
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For\n\
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html\n\
 *\n\
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish\n\
 * street, Riga, Latvia, EU, LV-1050.\n\
 *\n\
 * The  interactive user interfaces in modified source and object code versions\n\
 * of the Program must display Appropriate Legal Notices, as required under\n\
 * Section 5 of the GNU AGPL version 3.\n\
 *\n\
 * Pursuant to Section 7(b) of the License you must retain the original Product\n\
 * logo when distributing the program. Pursuant to Section 7(e) we decline to\n\
 * grant you any rights under trademark law for use of our trademarks.\n\
 *\n\
 * All the Product's GUI elements, including illustrations and icon sets, as\n\
 * well as technical writing content are licensed under the terms of the\n\
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License\n\
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode\n\
 *\n\
 */\n\
\n\
\"use strict\";\n\
\n\
(function()\n\
{\n\
	const L   = AscBidi.TYPE.L;\n\
	const R   = AscBidi.TYPE.R;\n\
	const AL  = AscBidi.TYPE.AL;\n\
	const EN  = AscBidi.TYPE.EN;\n\
	const ES  = AscBidi.TYPE.ES;\n\
	const ET  = AscBidi.TYPE.ET;\n\
	const AN  = AscBidi.TYPE.AN;\n\
	const CS  = AscBidi.TYPE.CS;\n\
	const NSM = AscBidi.TYPE.NSM;\n\
	const BN  = AscBidi.TYPE.BN;\n\
	const B   = AscBidi.TYPE.B;\n\
	const S   = AscBidi.TYPE.S;\n\
	const WS  = AscBidi.TYPE.WS;\n\
	const ON  = AscBidi.TYPE.ON;\n\
	const LRE = AscBidi.TYPE.LRE;\n\
	const LRO = AscBidi.TYPE.LRO;\n\
	const RLE = AscBidi.TYPE.RLE;\n\
	const RLO = AscBidi.TYPE.RLO;\n\
	const PDF = AscBidi.TYPE.PDF;\n\
	const LRI = AscBidi.TYPE.LRI;\n\
	const RLI = AscBidi.TYPE.RLI;\n\
	const FSI = AscBidi.TYPE.FSI;\n\
	const PDI = AscBidi.TYPE.PDI;\n\
	const unicodeTable = new Uint32Array([\n		"

mid = "	]);\n\
	const unicodeTable2 = {\n		"

footer = "\n	};\n\
	AscBidi.getType = function(codePoint)\n\
	{\n\
		if (codePoint < 0xFFFE)\n\
			return unicodeTable[codePoint];\n\
		if (undefined !== unicodeTable2[codePoint])\n\
			return unicodeTable2[codePoint];\n\
		if (0x10800 <= codePoint && codePoint < 0x11000)\n\
			return R;\n\
		if (0xE0000 <= codePoint && codePoint < 0xE1000)\n\
			return BN;\n\
		return L;\n\
	}\n\
})(window);\n"

allTypes = {
    "L" : 0,
    "R" : 0,
    "AL" : 0,
    "EN" : 0,
    "ES" : 0,
    "ET" : 0,
    "AN" : 0,
    "CS" : 0,
    "NSM" : 0,
    "BN" : 0,
    "B" : 0,
    "S" : 0,
    "WS" : 0,
    "ON" : 0,
    "LRE" : 0,
    "LRO" : 0,
    "RLE" : 0,
    "RLO" : 0,
    "PDF" : 0,
    "LRI" : 0,
    "RLI" : 0,
    "FSI" : 0,
    "PDI" : 0
}


h = header.split('\n')

out.writelines(header)

table = ["L"] * 0xFFFF

# base case
for c in range(0x0590, 0x0600):
    table[c] = "R"
for c in range(0x07C0, 0x0900):
    table[c] = "R"
for c in range(0xFB1D, 0xFB50):
    table[c] = "R"

for c in range(0x0600, 0x07C0):
    table[c] = "AL"
for c in range(0xFB50, 0xFDD0):
    table[c] = "AL"
for c in range(0xFDF0, 0xFE00):
    table[c] = "AL"
for c in range(0xFE70, 0xFF00):
    table[c] = "AL"

for c in range(0x2060, 0x2070):
    table[c] = "BN"
for c in range(0xFDD0, 0xFDF0):
    table[c] = "BN"
for c in range(0xFFF0, 0xFFF9):
    table[c] = "BN"

table2 = {}

for c in range(0xFFFF, 0x110000, 0x10000):
    table2[c - 1] = "BN"
    table2[c]     = "BN"

for line in data:
    cpData = line.split(';')
    unicode = int(cpData[0], 16)
    bidiType = cpData[4]

    if bidiType in allTypes:
        allTypes[bidiType] += 1
    else:
        print("Another type: " + bidiType)
        allTypes[bidiType] = 1

    if unicode >= 0xFFFE:
        table2[unicode] = bidiType
    else:
        table[unicode] = bidiType

for i in range(0xFFFE):
    out.write(table[i])
    if i != len(table) - 1:
        out.write(",")

out.writelines(mid)

addComma = False
for i in table2:
    if addComma:
        out.write(",")
    out.write(str(i))
    out.write(":")
    out.write(table2[i])
    addComma = True

out.writelines(footer)

out.close()
data.close()

for t in sorted(allTypes):
    print(t + " : " + str(allTypes[t]))







