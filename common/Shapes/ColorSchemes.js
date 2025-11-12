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
(function(window, undefined){

	let CColor = AscCommon.CColor;
	let CAscColorScheme = AscCommon.CAscColorScheme;
	let g_oUserColorScheme = [];

	let scheme;
	scheme = new CAscColorScheme();
	scheme.name = "Aspect";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x32, 0x32, 0x32));
	scheme.putColor(new CColor(0xe3, 0xde, 0xd1));
	scheme.putColor(new CColor(0xf0, 0x7f, 0x9));
	scheme.putColor(new CColor(0x9f, 0x29, 0x36));
	scheme.putColor(new CColor(0x1b, 0x58, 0x7c));
	scheme.putColor(new CColor(0x4e, 0x85, 0x42));
	scheme.putColor(new CColor(0x60, 0x48, 0x78));
	scheme.putColor(new CColor(0xc1, 0x98, 0x59));
	scheme.putColor(new CColor(0x6b, 0x9f, 0x25));
	scheme.putColor(new CColor(0xb2, 0x6b, 0x2));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Blue Green";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x37, 0x35, 0x45));
	scheme.putColor(new CColor(0xce, 0xdb, 0xe6));
	scheme.putColor(new CColor(0x34, 0x94, 0xba));
	scheme.putColor(new CColor(0x58, 0xb6, 0xc0));
	scheme.putColor(new CColor(0x75, 0xbd, 0xa7));
	scheme.putColor(new CColor(0x7a, 0x8c, 0x8e));
	scheme.putColor(new CColor(0x84, 0xac, 0xb6));
	scheme.putColor(new CColor(0x26, 0x83, 0xc6));
	scheme.putColor(new CColor(0x6b, 0x9f, 0x25));
	scheme.putColor(new CColor(0x9f, 0x67, 0x15));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Blue II";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x33, 0x5b, 0x74));
	scheme.putColor(new CColor(0xdf, 0xe3, 0xe5));
	scheme.putColor(new CColor(0x1c, 0xad, 0xe4));
	scheme.putColor(new CColor(0x26, 0x83, 0xc6));
	scheme.putColor(new CColor(0x27, 0xce, 0xd7));
	scheme.putColor(new CColor(0x42, 0xba, 0x97));
	scheme.putColor(new CColor(0x3e, 0x88, 0x53));
	scheme.putColor(new CColor(0x62, 0xa3, 0x9f));
	scheme.putColor(new CColor(0x6e, 0xac, 0x1c));
	scheme.putColor(new CColor(0xb2, 0x6b, 0x2));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Blue Warm";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x24, 0x28, 0x52));
	scheme.putColor(new CColor(0xac, 0xcb, 0xf9));
	scheme.putColor(new CColor(0x4a, 0x66, 0xac));
	scheme.putColor(new CColor(0x62, 0x9d, 0xd1));
	scheme.putColor(new CColor(0x29, 0x7f, 0xd5));
	scheme.putColor(new CColor(0x7f, 0x8f, 0xa9));
	scheme.putColor(new CColor(0x5a, 0xa2, 0xae));
	scheme.putColor(new CColor(0x9d, 0x90, 0xa0));
	scheme.putColor(new CColor(0x94, 0x54, 0xc3));
	scheme.putColor(new CColor(0x3e, 0xbb, 0xf0));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Blue";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x17, 0x40, 0x6d));
	scheme.putColor(new CColor(0xdb, 0xef, 0xf9));
	scheme.putColor(new CColor(0xf, 0x6f, 0xc6));
	scheme.putColor(new CColor(0x0, 0x9d, 0xd9));
	scheme.putColor(new CColor(0xb, 0xd0, 0xd9));
	scheme.putColor(new CColor(0x10, 0xcf, 0x9b));
	scheme.putColor(new CColor(0x7c, 0xca, 0x62));
	scheme.putColor(new CColor(0xa5, 0xc2, 0x49));
	scheme.putColor(new CColor(0xf4, 0x91, 0x0));
	scheme.putColor(new CColor(0x85, 0xdf, 0xd0));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Grayscale";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xf8, 0xf8, 0xf8));
	scheme.putColor(new CColor(0xdd, 0xdd, 0xdd));
	scheme.putColor(new CColor(0xb2, 0xb2, 0xb2));
	scheme.putColor(new CColor(0x96, 0x96, 0x96));
	scheme.putColor(new CColor(0x80, 0x80, 0x80));
	scheme.putColor(new CColor(0x5f, 0x5f, 0x5f));
	scheme.putColor(new CColor(0x4d, 0x4d, 0x4d));
	scheme.putColor(new CColor(0x5f, 0x5f, 0x5f));
	scheme.putColor(new CColor(0x91, 0x91, 0x91));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Green Yellow";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x45, 0x5f, 0x51));
	scheme.putColor(new CColor(0xe2, 0xdf, 0xcc));
	scheme.putColor(new CColor(0x99, 0xcb, 0x38));
	scheme.putColor(new CColor(0x63, 0xa5, 0x37));
	scheme.putColor(new CColor(0x37, 0xa7, 0x6f));
	scheme.putColor(new CColor(0x44, 0xc1, 0xa3));
	scheme.putColor(new CColor(0x4e, 0xb3, 0xcf));
	scheme.putColor(new CColor(0x51, 0xc3, 0xf9));
	scheme.putColor(new CColor(0xee, 0x7b, 0x8));
	scheme.putColor(new CColor(0x97, 0x7b, 0x2d));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Green";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x45, 0x5f, 0x51));
	scheme.putColor(new CColor(0xe3, 0xde, 0xd1));
	scheme.putColor(new CColor(0x54, 0x9e, 0x39));
	scheme.putColor(new CColor(0x8a, 0xb8, 0x33));
	scheme.putColor(new CColor(0xc0, 0xcf, 0x3a));
	scheme.putColor(new CColor(0x2, 0x96, 0x76));
	scheme.putColor(new CColor(0x4a, 0xb5, 0xc4));
	scheme.putColor(new CColor(0x9, 0x89, 0xb1));
	scheme.putColor(new CColor(0x6b, 0x9f, 0x25));
	scheme.putColor(new CColor(0xba, 0x69, 0x6));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Marquee";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x5e, 0x5e, 0x5e));
	scheme.putColor(new CColor(0xdd, 0xdd, 0xdd));
	scheme.putColor(new CColor(0x41, 0x8a, 0xb3));
	scheme.putColor(new CColor(0xa6, 0xb7, 0x27));
	scheme.putColor(new CColor(0xf6, 0x92, 0x0));
	scheme.putColor(new CColor(0x83, 0x83, 0x83));
	scheme.putColor(new CColor(0xfe, 0xc3, 0x6));
	scheme.putColor(new CColor(0xdf, 0x53, 0x27));
	scheme.putColor(new CColor(0xf5, 0x9e, 0x0));
	scheme.putColor(new CColor(0xb2, 0xb2, 0xb2));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Median";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x77, 0x5f, 0x55));
	scheme.putColor(new CColor(0xeb, 0xdd, 0xc3));
	scheme.putColor(new CColor(0x94, 0xb6, 0xd2));
	scheme.putColor(new CColor(0xdd, 0x80, 0x47));
	scheme.putColor(new CColor(0xa5, 0xab, 0x81));
	scheme.putColor(new CColor(0xd8, 0xb2, 0x5c));
	scheme.putColor(new CColor(0x7b, 0xa7, 0x9d));
	scheme.putColor(new CColor(0x96, 0x8c, 0x8c));
	scheme.putColor(new CColor(0xf7, 0xb6, 0x15));
	scheme.putColor(new CColor(0x70, 0x44, 0x4));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Office 2007 - 2010";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x1f, 0x49, 0x7d));
	scheme.putColor(new CColor(0xee, 0xec, 0xe1));
	scheme.putColor(new CColor(0x4f, 0x81, 0xbd));
	scheme.putColor(new CColor(0xc0, 0x50, 0x4d));
	scheme.putColor(new CColor(0x9b, 0xbb, 0x59));
	scheme.putColor(new CColor(0x80, 0x64, 0xa2));
	scheme.putColor(new CColor(0x4b, 0xac, 0xc6));
	scheme.putColor(new CColor(0xf7, 0x96, 0x46));
	scheme.putColor(new CColor(0x0, 0x0, 0xff));
	scheme.putColor(new CColor(0x80, 0x0, 0x80));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Office 2013 - 2022";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x44, 0x54, 0x6a));
	scheme.putColor(new CColor(0xe7, 0xe6, 0xe6));
	scheme.putColor(new CColor(0x44, 0x72, 0xc4));
	scheme.putColor(new CColor(0xed, 0x7d, 0x31));
	scheme.putColor(new CColor(0xa5, 0xa5, 0xa5));
	scheme.putColor(new CColor(0xff, 0xc0, 0x0));
	scheme.putColor(new CColor(0x5b, 0x9b, 0xd5));
	scheme.putColor(new CColor(0x70, 0xad, 0x47));
	scheme.putColor(new CColor(0x5, 0x63, 0xc1));
	scheme.putColor(new CColor(0x95, 0x4f, 0x72));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Office";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0xe, 0x28, 0x41));
	scheme.putColor(new CColor(0xe8, 0xe8, 0xe8));
	scheme.putColor(new CColor(0x15, 0x60, 0x82));
	scheme.putColor(new CColor(0xe9, 0x71, 0x32));
	scheme.putColor(new CColor(0x19, 0x6b, 0x24));
	scheme.putColor(new CColor(0xf, 0x9e, 0xd5));
	scheme.putColor(new CColor(0xa0, 0x2b, 0x93));
	scheme.putColor(new CColor(0x4e, 0xa7, 0x2e));
	scheme.putColor(new CColor(0x46, 0x78, 0x86));
	scheme.putColor(new CColor(0x96, 0x60, 0x7d));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Orange Red";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x69, 0x64, 0x64));
	scheme.putColor(new CColor(0xe9, 0xe5, 0xdc));
	scheme.putColor(new CColor(0xd3, 0x48, 0x17));
	scheme.putColor(new CColor(0x9b, 0x2d, 0x1f));
	scheme.putColor(new CColor(0xa2, 0x8e, 0x6a));
	scheme.putColor(new CColor(0x95, 0x62, 0x51));
	scheme.putColor(new CColor(0x91, 0x84, 0x85));
	scheme.putColor(new CColor(0x85, 0x5d, 0x5d));
	scheme.putColor(new CColor(0xcc, 0x99, 0x0));
	scheme.putColor(new CColor(0x96, 0xa9, 0xa9));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Orange";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x63, 0x70, 0x52));
	scheme.putColor(new CColor(0xcc, 0xdd, 0xea));
	scheme.putColor(new CColor(0xe4, 0x83, 0x12));
	scheme.putColor(new CColor(0xbd, 0x58, 0x2c));
	scheme.putColor(new CColor(0x86, 0x56, 0x40));
	scheme.putColor(new CColor(0x9b, 0x83, 0x57));
	scheme.putColor(new CColor(0xc2, 0xbc, 0x80));
	scheme.putColor(new CColor(0x94, 0xa0, 0x88));
	scheme.putColor(new CColor(0x29, 0x98, 0xe3));
	scheme.putColor(new CColor(0x8c, 0x8c, 0x8c));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Paper";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x44, 0x4d, 0x26));
	scheme.putColor(new CColor(0xfe, 0xfa, 0xc9));
	scheme.putColor(new CColor(0xa5, 0xb5, 0x92));
	scheme.putColor(new CColor(0xf3, 0xa4, 0x47));
	scheme.putColor(new CColor(0xe7, 0xbc, 0x29));
	scheme.putColor(new CColor(0xd0, 0x92, 0xa7));
	scheme.putColor(new CColor(0x9c, 0x85, 0xc0));
	scheme.putColor(new CColor(0x80, 0x9e, 0xc2));
	scheme.putColor(new CColor(0x8e, 0x58, 0xb6));
	scheme.putColor(new CColor(0x7f, 0x6f, 0x6f));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Red Orange";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x50, 0x50, 0x46));
	scheme.putColor(new CColor(0xee, 0xec, 0xe1));
	scheme.putColor(new CColor(0xe8, 0x4c, 0x22));
	scheme.putColor(new CColor(0xff, 0xbd, 0x47));
	scheme.putColor(new CColor(0xb6, 0x49, 0x26));
	scheme.putColor(new CColor(0xff, 0x84, 0x27));
	scheme.putColor(new CColor(0xcc, 0x99, 0x0));
	scheme.putColor(new CColor(0xb2, 0x26, 0x0));
	scheme.putColor(new CColor(0xcc, 0x99, 0x0));
	scheme.putColor(new CColor(0x66, 0x66, 0x99));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Red Violet";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x45, 0x45, 0x51));
	scheme.putColor(new CColor(0xd8, 0xd9, 0xdc));
	scheme.putColor(new CColor(0xe3, 0x2d, 0x91));
	scheme.putColor(new CColor(0xc8, 0x30, 0xcc));
	scheme.putColor(new CColor(0x4e, 0xa6, 0xdc));
	scheme.putColor(new CColor(0x47, 0x75, 0xe7));
	scheme.putColor(new CColor(0x89, 0x71, 0xe1));
	scheme.putColor(new CColor(0xd5, 0x47, 0x73));
	scheme.putColor(new CColor(0x6b, 0x9f, 0x25));
	scheme.putColor(new CColor(0x8c, 0x8c, 0x8c));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Red";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x32, 0x32, 0x32));
	scheme.putColor(new CColor(0xe5, 0xc2, 0x43));
	scheme.putColor(new CColor(0xa5, 0x30, 0xf));
	scheme.putColor(new CColor(0xd5, 0x58, 0x16));
	scheme.putColor(new CColor(0xe1, 0x98, 0x25));
	scheme.putColor(new CColor(0xb1, 0x9c, 0x7d));
	scheme.putColor(new CColor(0x7f, 0x5f, 0x52));
	scheme.putColor(new CColor(0xb2, 0x7d, 0x49));
	scheme.putColor(new CColor(0x6b, 0x9f, 0x25));
	scheme.putColor(new CColor(0xb2, 0x6b, 0x2));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Slipstream";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x21, 0x27, 0x45));
	scheme.putColor(new CColor(0xb4, 0xdc, 0xfa));
	scheme.putColor(new CColor(0x4e, 0x67, 0xc8));
	scheme.putColor(new CColor(0x5e, 0xcc, 0xf3));
	scheme.putColor(new CColor(0xa7, 0xea, 0x52));
	scheme.putColor(new CColor(0x5d, 0xce, 0xaf));
	scheme.putColor(new CColor(0xff, 0x80, 0x21));
	scheme.putColor(new CColor(0xf1, 0x41, 0x24));
	scheme.putColor(new CColor(0x56, 0xc7, 0xaa));
	scheme.putColor(new CColor(0x59, 0xa8, 0xd1));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Violet II";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x63, 0x2e, 0x62));
	scheme.putColor(new CColor(0xea, 0xe5, 0xeb));
	scheme.putColor(new CColor(0x92, 0x27, 0x8f));
	scheme.putColor(new CColor(0x9b, 0x57, 0xd3));
	scheme.putColor(new CColor(0x75, 0x5d, 0xd9));
	scheme.putColor(new CColor(0x66, 0x5e, 0xb8));
	scheme.putColor(new CColor(0x45, 0xa5, 0xed));
	scheme.putColor(new CColor(0x59, 0x82, 0xdb));
	scheme.putColor(new CColor(0x0, 0x66, 0xff));
	scheme.putColor(new CColor(0x66, 0x66, 0x99));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Violet";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x37, 0x35, 0x45));
	scheme.putColor(new CColor(0xdc, 0xd8, 0xdc));
	scheme.putColor(new CColor(0xad, 0x84, 0xc6));
	scheme.putColor(new CColor(0x87, 0x84, 0xc7));
	scheme.putColor(new CColor(0x5d, 0x73, 0x9a));
	scheme.putColor(new CColor(0x69, 0x97, 0xaf));
	scheme.putColor(new CColor(0x84, 0xac, 0xb6));
	scheme.putColor(new CColor(0x6f, 0x81, 0x83));
	scheme.putColor(new CColor(0x69, 0xa0, 0x20));
	scheme.putColor(new CColor(0x8c, 0x8c, 0x8c));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Yellow Orange";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x4e, 0x3b, 0x30));
	scheme.putColor(new CColor(0xfb, 0xee, 0xc9));
	scheme.putColor(new CColor(0xf0, 0xa2, 0x2e));
	scheme.putColor(new CColor(0xa5, 0x64, 0x4e));
	scheme.putColor(new CColor(0xb5, 0x8b, 0x80));
	scheme.putColor(new CColor(0xc3, 0x98, 0x6d));
	scheme.putColor(new CColor(0xa1, 0x95, 0x74));
	scheme.putColor(new CColor(0xc1, 0x75, 0x29));
	scheme.putColor(new CColor(0xad, 0x1f, 0x1f));
	scheme.putColor(new CColor(0xff, 0xc4, 0x2f));
	g_oUserColorScheme.push(scheme);

	scheme = new CAscColorScheme();
	scheme.name = "Yellow";
	scheme.putColor(new CColor(0x0, 0x0, 0x0));
	scheme.putColor(new CColor(0xff, 0xff, 0xff));
	scheme.putColor(new CColor(0x39, 0x30, 0x2a));
	scheme.putColor(new CColor(0xe5, 0xde, 0xdb));
	scheme.putColor(new CColor(0xff, 0xca, 0x8));
	scheme.putColor(new CColor(0xf8, 0x93, 0x1d));
	scheme.putColor(new CColor(0xce, 0x8d, 0x3e));
	scheme.putColor(new CColor(0xec, 0x70, 0x16));
	scheme.putColor(new CColor(0xe6, 0x48, 0x23));
	scheme.putColor(new CColor(0x9c, 0x6a, 0x6a));
	scheme.putColor(new CColor(0x29, 0x98, 0xe3));
	scheme.putColor(new CColor(0x7f, 0x72, 0x3d));
	g_oUserColorScheme.push(scheme);

	window['AscCommon'].g_oUserColorScheme = g_oUserColorScheme;
})(window);
