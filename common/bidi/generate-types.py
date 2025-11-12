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
	const m = [\n\
		AscBidi.TYPE.L,\n\
		AscBidi.TYPE.R,\n\
		AscBidi.TYPE.AL,\n\
		AscBidi.TYPE.EN,\n\
		AscBidi.TYPE.ES,\n\
		AscBidi.TYPE.ET,\n\
		AscBidi.TYPE.AN,\n\
		AscBidi.TYPE.CS,\n\
		AscBidi.TYPE.NSM,\n\
		AscBidi.TYPE.BN,\n\
		AscBidi.TYPE.B,\n\
		AscBidi.TYPE.S,\n\
		AscBidi.TYPE.WS,\n\
		AscBidi.TYPE.ON,\n\
		AscBidi.TYPE.LRE,\n\
		AscBidi.TYPE.LRO,\n\
		AscBidi.TYPE.RLE,\n\
		AscBidi.TYPE.RLO,\n\
		AscBidi.TYPE.PDF,\n\
		AscBidi.TYPE.LRI,\n\
		AscBidi.TYPE.RLI,\n\
		AscBidi.TYPE.FSI,\n\
		AscBidi.TYPE.PDI\n\
	];\n\
	const t = new Uint8Array(0x110000).fill(0);\n"

footer = "	\n	AscBidi.getType = function(codePoint)\n\
	{\n\
		if (codePoint >= 0x110000 || codePoint < 0)\n\
			return AscBidi.TYPE.L;\n\
		return m[t[codePoint]];\n\
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

mVal = {
    "L" : 0,
    "R" : 1,
    "AL" : 2,
    "EN" : 3,
    "ES" : 4,
    "ET" : 5,
    "AN" : 6,
    "CS" : 7,
    "NSM" : 8,
    "BN" : 9,
    "B" : 10,
    "S" : 11,
    "WS" : 12,
    "ON" : 13,
    "LRE" : 14,
    "LRO" : 15,
    "RLE" : 16,
    "RLO" : 17,
    "PDF" : 18,
    "LRI" : 19,
    "RLI" : 20,
    "FSI" : 21,
    "PDI" : 22
}


h = header.split('\n')

out.writelines(header)

tableLen = 0x110000
table = ["L"] * tableLen

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

for c in range(0xFFFF, 0x110000, 0x10000):
    table[c - 1] = "BN"
    table[c]     = "BN"

for line in data:
    cpData = line.split(';')
    unicode = int(cpData[0], 16)
    bidiType = cpData[4]

    if bidiType in allTypes:
        allTypes[bidiType] += 1
    else:
        print("Another type: " + bidiType)
        allTypes[bidiType] = 1

    table[unicode] = bidiType

val      = mVal["L"]
count    = tableLen
startPos = 0

def writeLine():
    if val == mVal["L"] or 0 == count:
        return

    out.write("	t.fill(")
    out.write(str(val))
    out.write(",")
    out.write(str(startPos))
    out.write(",")
    out.write(str(startPos + count))
    out.write(");\n")

writeLine()

val      = mVal[table[0]]
count    = 1
startPos = 0

for i in range(1, tableLen):
    _val = mVal[table[i]]
    if val != _val:
        writeLine()
        val = _val
        count = 1
        startPos = i
    else:
        count += 1

writeLine()

out.writelines(footer)

out.close()
data.close()

for t in sorted(allTypes):
    print(t + " : " + str(allTypes[t]))







