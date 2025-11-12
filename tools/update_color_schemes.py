#!/usr/bin/env python3

import os
import zipfile
from xml.dom.minidom import parse

def main():
    print('Generate color schemes')

    license = '''/*
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
 */'''

    script_dir = os.path.dirname(os.path.realpath(__file__))
    schemes_script = script_dir + "/../common/Shapes/ColorSchemes.js"
    schemes_file = open(schemes_script, "w")
    schemes_file.write(license)
    schemes_file.write("\n")
    schemes_file.write("\"use strict\";")
    schemes_file.write("\n")
    schemes_file.write("(function(window, undefined){\n")
    schemes_file.write("\n")
    schemes_file.write("\tlet CColor = AscCommon.CColor;\n")
    schemes_file.write("\tlet CAscColorScheme = AscCommon.CAscColorScheme;\n")
    schemes_file.write("\tlet g_oUserColorScheme = [];\n")
    schemes_file.write("\n")
    schemes_file.write("\tlet scheme;\n")
    directory =  os.path.join(script_dir, 'colors')
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        if os.path.isfile(f):
            with zipfile.ZipFile(f,"r") as zip_ref:
                for name in zip_ref.namelist():
                    if name.endswith('theme1.xml'):
                        theme = parse(zip_ref.open(name))
                        scheme = theme.getElementsByTagName("a:clrScheme")[0]
                        name = scheme.getAttribute("name")
                        
                        schemes_file.write("\tscheme = new CAscColorScheme();\n")
                        schemes_file.write("\tscheme.name = \"" + name + "\";\n")
                        colors = scheme.childNodes
                        color_val = ""
                        for colorEl in colors:
                            color = colorEl.firstChild
                            if color.localName == "sysClr":
                                color_val = "0x" + color.getAttribute("lastClr")
                            if color.localName == "srgbClr":
                                color_val = "0x" + color.getAttribute("val")
                            RGBint = int(color_val, 16)
                            B =  RGBint & 255
                            G = (RGBint >> 8) & 255
                            R =   (RGBint >> 16) & 255
                            schemes_file.write("\tscheme.putColor(new CColor("+ hex(R) +", "+hex(G)+", "+hex(B) +"));\n")
                        schemes_file.write("\tg_oUserColorScheme.push(scheme);\n") 
                        schemes_file.write("\n")


                #zip_ref.extractall("targetdir")
                
    schemes_file.write("\twindow['AscCommon'].g_oUserColorScheme = g_oUserColorScheme;\n")
    schemes_file.write("})(window);\n")
    schemes_file.close();

if __name__ == "__main__":
    main()