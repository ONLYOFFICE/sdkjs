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

(function()
{
	const m = [
		AscBidi.TYPE.L,
		AscBidi.TYPE.R,
		AscBidi.TYPE.AL,
		AscBidi.TYPE.EN,
		AscBidi.TYPE.ES,
		AscBidi.TYPE.ET,
		AscBidi.TYPE.AN,
		AscBidi.TYPE.CS,
		AscBidi.TYPE.NSM,
		AscBidi.TYPE.BN,
		AscBidi.TYPE.B,
		AscBidi.TYPE.S,
		AscBidi.TYPE.WS,
		AscBidi.TYPE.ON,
		AscBidi.TYPE.LRE,
		AscBidi.TYPE.LRO,
		AscBidi.TYPE.RLE,
		AscBidi.TYPE.RLO,
		AscBidi.TYPE.PDF,
		AscBidi.TYPE.LRI,
		AscBidi.TYPE.RLI,
		AscBidi.TYPE.FSI,
		AscBidi.TYPE.PDI
	];
	const t = new Uint8Array(0x110000).fill(0);
	t.fill(9,0,9);
	t.fill(11,9,10);
	t.fill(10,10,11);
	t.fill(11,11,12);
	t.fill(12,12,13);
	t.fill(10,13,14);
	t.fill(9,14,28);
	t.fill(10,28,31);
	t.fill(11,31,32);
	t.fill(12,32,33);
	t.fill(13,33,35);
	t.fill(5,35,38);
	t.fill(13,38,43);
	t.fill(4,43,44);
	t.fill(7,44,45);
	t.fill(4,45,46);
	t.fill(7,46,48);
	t.fill(3,48,58);
	t.fill(7,58,59);
	t.fill(13,59,65);
	t.fill(13,91,97);
	t.fill(13,123,127);
	t.fill(9,127,133);
	t.fill(10,133,134);
	t.fill(9,134,160);
	t.fill(7,160,161);
	t.fill(13,161,162);
	t.fill(5,162,166);
	t.fill(13,166,170);
	t.fill(13,171,173);
	t.fill(9,173,174);
	t.fill(13,174,176);
	t.fill(5,176,178);
	t.fill(3,178,180);
	t.fill(13,180,181);
	t.fill(13,182,185);
	t.fill(3,185,186);
	t.fill(13,187,192);
	t.fill(13,215,216);
	t.fill(13,247,248);
	t.fill(13,697,699);
	t.fill(13,706,720);
	t.fill(13,722,736);
	t.fill(13,741,750);
	t.fill(13,751,768);
	t.fill(8,768,880);
	t.fill(13,884,886);
	t.fill(13,894,895);
	t.fill(13,900,902);
	t.fill(13,903,904);
	t.fill(13,1014,1015);
	t.fill(8,1155,1162);
	t.fill(13,1418,1419);
	t.fill(13,1421,1423);
	t.fill(5,1423,1424);
	t.fill(1,1424,1425);
	t.fill(8,1425,1470);
	t.fill(1,1470,1471);
	t.fill(8,1471,1472);
	t.fill(1,1472,1473);
	t.fill(8,1473,1475);
	t.fill(1,1475,1476);
	t.fill(8,1476,1478);
	t.fill(1,1478,1479);
	t.fill(8,1479,1480);
	t.fill(1,1480,1536);
	t.fill(6,1536,1542);
	t.fill(13,1542,1544);
	t.fill(2,1544,1545);
	t.fill(5,1545,1547);
	t.fill(2,1547,1548);
	t.fill(7,1548,1549);
	t.fill(2,1549,1550);
	t.fill(13,1550,1552);
	t.fill(8,1552,1563);
	t.fill(2,1563,1611);
	t.fill(8,1611,1632);
	t.fill(6,1632,1642);
	t.fill(5,1642,1643);
	t.fill(6,1643,1645);
	t.fill(2,1645,1648);
	t.fill(8,1648,1649);
	t.fill(2,1649,1750);
	t.fill(8,1750,1757);
	t.fill(6,1757,1758);
	t.fill(13,1758,1759);
	t.fill(8,1759,1765);
	t.fill(2,1765,1767);
	t.fill(8,1767,1769);
	t.fill(13,1769,1770);
	t.fill(8,1770,1774);
	t.fill(2,1774,1776);
	t.fill(3,1776,1786);
	t.fill(2,1786,1809);
	t.fill(8,1809,1810);
	t.fill(2,1810,1840);
	t.fill(8,1840,1867);
	t.fill(2,1867,1958);
	t.fill(8,1958,1969);
	t.fill(2,1969,1984);
	t.fill(1,1984,2027);
	t.fill(8,2027,2036);
	t.fill(1,2036,2038);
	t.fill(13,2038,2042);
	t.fill(1,2042,2045);
	t.fill(8,2045,2046);
	t.fill(1,2046,2070);
	t.fill(8,2070,2074);
	t.fill(1,2074,2075);
	t.fill(8,2075,2084);
	t.fill(1,2084,2085);
	t.fill(8,2085,2088);
	t.fill(1,2088,2089);
	t.fill(8,2089,2094);
	t.fill(1,2094,2137);
	t.fill(8,2137,2140);
	t.fill(1,2140,2144);
	t.fill(2,2144,2155);
	t.fill(1,2155,2160);
	t.fill(2,2160,2191);
	t.fill(1,2191,2192);
	t.fill(6,2192,2194);
	t.fill(1,2194,2200);
	t.fill(8,2200,2208);
	t.fill(2,2208,2250);
	t.fill(8,2250,2274);
	t.fill(6,2274,2275);
	t.fill(8,2275,2307);
	t.fill(8,2362,2363);
	t.fill(8,2364,2365);
	t.fill(8,2369,2377);
	t.fill(8,2381,2382);
	t.fill(8,2385,2392);
	t.fill(8,2402,2404);
	t.fill(8,2433,2434);
	t.fill(8,2492,2493);
	t.fill(8,2497,2501);
	t.fill(8,2509,2510);
	t.fill(8,2530,2532);
	t.fill(5,2546,2548);
	t.fill(5,2555,2556);
	t.fill(8,2558,2559);
	t.fill(8,2561,2563);
	t.fill(8,2620,2621);
	t.fill(8,2625,2627);
	t.fill(8,2631,2633);
	t.fill(8,2635,2638);
	t.fill(8,2641,2642);
	t.fill(8,2672,2674);
	t.fill(8,2677,2678);
	t.fill(8,2689,2691);
	t.fill(8,2748,2749);
	t.fill(8,2753,2758);
	t.fill(8,2759,2761);
	t.fill(8,2765,2766);
	t.fill(8,2786,2788);
	t.fill(5,2801,2802);
	t.fill(8,2810,2816);
	t.fill(8,2817,2818);
	t.fill(8,2876,2877);
	t.fill(8,2879,2880);
	t.fill(8,2881,2885);
	t.fill(8,2893,2894);
	t.fill(8,2901,2903);
	t.fill(8,2914,2916);
	t.fill(8,2946,2947);
	t.fill(8,3008,3009);
	t.fill(8,3021,3022);
	t.fill(13,3059,3065);
	t.fill(5,3065,3066);
	t.fill(13,3066,3067);
	t.fill(8,3072,3073);
	t.fill(8,3076,3077);
	t.fill(8,3132,3133);
	t.fill(8,3134,3137);
	t.fill(8,3142,3145);
	t.fill(8,3146,3150);
	t.fill(8,3157,3159);
	t.fill(8,3170,3172);
	t.fill(13,3192,3199);
	t.fill(8,3201,3202);
	t.fill(8,3260,3261);
	t.fill(8,3276,3278);
	t.fill(8,3298,3300);
	t.fill(8,3328,3330);
	t.fill(8,3387,3389);
	t.fill(8,3393,3397);
	t.fill(8,3405,3406);
	t.fill(8,3426,3428);
	t.fill(8,3457,3458);
	t.fill(8,3530,3531);
	t.fill(8,3538,3541);
	t.fill(8,3542,3543);
	t.fill(8,3633,3634);
	t.fill(8,3636,3643);
	t.fill(5,3647,3648);
	t.fill(8,3655,3663);
	t.fill(8,3761,3762);
	t.fill(8,3764,3773);
	t.fill(8,3784,3791);
	t.fill(8,3864,3866);
	t.fill(8,3893,3894);
	t.fill(8,3895,3896);
	t.fill(8,3897,3898);
	t.fill(13,3898,3902);
	t.fill(8,3953,3967);
	t.fill(8,3968,3973);
	t.fill(8,3974,3976);
	t.fill(8,3981,3992);
	t.fill(8,3993,4029);
	t.fill(8,4038,4039);
	t.fill(8,4141,4145);
	t.fill(8,4146,4152);
	t.fill(8,4153,4155);
	t.fill(8,4157,4159);
	t.fill(8,4184,4186);
	t.fill(8,4190,4193);
	t.fill(8,4209,4213);
	t.fill(8,4226,4227);
	t.fill(8,4229,4231);
	t.fill(8,4237,4238);
	t.fill(8,4253,4254);
	t.fill(8,4957,4960);
	t.fill(13,5008,5018);
	t.fill(13,5120,5121);
	t.fill(12,5760,5761);
	t.fill(13,5787,5789);
	t.fill(8,5906,5909);
	t.fill(8,5938,5940);
	t.fill(8,5970,5972);
	t.fill(8,6002,6004);
	t.fill(8,6068,6070);
	t.fill(8,6071,6078);
	t.fill(8,6086,6087);
	t.fill(8,6089,6100);
	t.fill(5,6107,6108);
	t.fill(8,6109,6110);
	t.fill(13,6128,6138);
	t.fill(13,6144,6155);
	t.fill(8,6155,6158);
	t.fill(9,6158,6159);
	t.fill(8,6159,6160);
	t.fill(8,6277,6279);
	t.fill(8,6313,6314);
	t.fill(8,6432,6435);
	t.fill(8,6439,6441);
	t.fill(8,6450,6451);
	t.fill(8,6457,6460);
	t.fill(13,6464,6465);
	t.fill(13,6468,6470);
	t.fill(13,6622,6656);
	t.fill(8,6679,6681);
	t.fill(8,6683,6684);
	t.fill(8,6742,6743);
	t.fill(8,6744,6751);
	t.fill(8,6752,6753);
	t.fill(8,6754,6755);
	t.fill(8,6757,6765);
	t.fill(8,6771,6781);
	t.fill(8,6783,6784);
	t.fill(8,6832,6863);
	t.fill(8,6912,6916);
	t.fill(8,6964,6965);
	t.fill(8,6966,6971);
	t.fill(8,6972,6973);
	t.fill(8,6978,6979);
	t.fill(8,7019,7028);
	t.fill(8,7040,7042);
	t.fill(8,7074,7078);
	t.fill(8,7080,7082);
	t.fill(8,7083,7086);
	t.fill(8,7142,7143);
	t.fill(8,7144,7146);
	t.fill(8,7149,7150);
	t.fill(8,7151,7154);
	t.fill(8,7212,7220);
	t.fill(8,7222,7224);
	t.fill(8,7376,7379);
	t.fill(8,7380,7393);
	t.fill(8,7394,7401);
	t.fill(8,7405,7406);
	t.fill(8,7412,7413);
	t.fill(8,7416,7418);
	t.fill(8,7616,7680);
	t.fill(13,8125,8126);
	t.fill(13,8127,8130);
	t.fill(13,8141,8144);
	t.fill(13,8157,8160);
	t.fill(13,8173,8176);
	t.fill(13,8189,8191);
	t.fill(12,8192,8203);
	t.fill(9,8203,8206);
	t.fill(1,8207,8208);
	t.fill(13,8208,8232);
	t.fill(12,8232,8233);
	t.fill(10,8233,8234);
	t.fill(14,8234,8235);
	t.fill(16,8235,8236);
	t.fill(18,8236,8237);
	t.fill(15,8237,8238);
	t.fill(17,8238,8239);
	t.fill(7,8239,8240);
	t.fill(5,8240,8245);
	t.fill(13,8245,8260);
	t.fill(7,8260,8261);
	t.fill(13,8261,8287);
	t.fill(12,8287,8288);
	t.fill(9,8288,8294);
	t.fill(19,8294,8295);
	t.fill(20,8295,8296);
	t.fill(21,8296,8297);
	t.fill(22,8297,8298);
	t.fill(9,8298,8304);
	t.fill(3,8304,8305);
	t.fill(3,8308,8314);
	t.fill(4,8314,8316);
	t.fill(13,8316,8319);
	t.fill(3,8320,8330);
	t.fill(4,8330,8332);
	t.fill(13,8332,8335);
	t.fill(5,8352,8385);
	t.fill(8,8400,8433);
	t.fill(13,8448,8450);
	t.fill(13,8451,8455);
	t.fill(13,8456,8458);
	t.fill(13,8468,8469);
	t.fill(13,8470,8473);
	t.fill(13,8478,8484);
	t.fill(13,8485,8486);
	t.fill(13,8487,8488);
	t.fill(13,8489,8490);
	t.fill(5,8494,8495);
	t.fill(13,8506,8508);
	t.fill(13,8512,8517);
	t.fill(13,8522,8526);
	t.fill(13,8528,8544);
	t.fill(13,8585,8588);
	t.fill(13,8592,8722);
	t.fill(4,8722,8723);
	t.fill(5,8723,8724);
	t.fill(13,8724,9014);
	t.fill(13,9083,9109);
	t.fill(13,9110,9255);
	t.fill(13,9280,9291);
	t.fill(13,9312,9352);
	t.fill(3,9352,9372);
	t.fill(13,9450,9900);
	t.fill(13,9901,10240);
	t.fill(13,10496,11124);
	t.fill(13,11126,11158);
	t.fill(13,11159,11264);
	t.fill(13,11493,11499);
	t.fill(8,11503,11506);
	t.fill(13,11513,11520);
	t.fill(8,11647,11648);
	t.fill(8,11744,11776);
	t.fill(13,11776,11870);
	t.fill(13,11904,11930);
	t.fill(13,11931,12020);
	t.fill(13,12032,12246);
	t.fill(13,12272,12288);
	t.fill(12,12288,12289);
	t.fill(13,12289,12293);
	t.fill(13,12296,12321);
	t.fill(8,12330,12334);
	t.fill(13,12336,12337);
	t.fill(13,12342,12344);
	t.fill(13,12349,12352);
	t.fill(8,12441,12443);
	t.fill(13,12443,12445);
	t.fill(13,12448,12449);
	t.fill(13,12539,12540);
	t.fill(13,12736,12772);
	t.fill(13,12783,12784);
	t.fill(13,12829,12831);
	t.fill(13,12880,12896);
	t.fill(13,12924,12927);
	t.fill(13,12977,12992);
	t.fill(13,13004,13008);
	t.fill(13,13175,13179);
	t.fill(13,13278,13280);
	t.fill(13,13311,13312);
	t.fill(13,19904,19968);
	t.fill(13,42128,42183);
	t.fill(13,42509,42512);
	t.fill(8,42607,42611);
	t.fill(13,42611,42612);
	t.fill(8,42612,42622);
	t.fill(13,42622,42624);
	t.fill(8,42654,42656);
	t.fill(8,42736,42738);
	t.fill(13,42752,42786);
	t.fill(13,42888,42889);
	t.fill(8,43010,43011);
	t.fill(8,43014,43015);
	t.fill(8,43019,43020);
	t.fill(8,43045,43047);
	t.fill(13,43048,43052);
	t.fill(8,43052,43053);
	t.fill(5,43064,43066);
	t.fill(13,43124,43128);
	t.fill(8,43204,43206);
	t.fill(8,43232,43250);
	t.fill(8,43263,43264);
	t.fill(8,43302,43310);
	t.fill(8,43335,43346);
	t.fill(8,43392,43395);
	t.fill(8,43443,43444);
	t.fill(8,43446,43450);
	t.fill(8,43452,43454);
	t.fill(8,43493,43494);
	t.fill(8,43561,43567);
	t.fill(8,43569,43571);
	t.fill(8,43573,43575);
	t.fill(8,43587,43588);
	t.fill(8,43596,43597);
	t.fill(8,43644,43645);
	t.fill(8,43696,43697);
	t.fill(8,43698,43701);
	t.fill(8,43703,43705);
	t.fill(8,43710,43712);
	t.fill(8,43713,43714);
	t.fill(8,43756,43758);
	t.fill(8,43766,43767);
	t.fill(13,43882,43884);
	t.fill(8,44005,44006);
	t.fill(8,44008,44009);
	t.fill(8,44013,44014);
	t.fill(1,64285,64286);
	t.fill(8,64286,64287);
	t.fill(1,64287,64297);
	t.fill(4,64297,64298);
	t.fill(1,64298,64336);
	t.fill(2,64336,64830);
	t.fill(13,64830,64848);
	t.fill(2,64848,64975);
	t.fill(13,64975,64976);
	t.fill(9,64976,65008);
	t.fill(2,65008,65021);
	t.fill(13,65021,65024);
	t.fill(8,65024,65040);
	t.fill(13,65040,65050);
	t.fill(8,65056,65072);
	t.fill(13,65072,65104);
	t.fill(7,65104,65105);
	t.fill(13,65105,65106);
	t.fill(7,65106,65107);
	t.fill(13,65108,65109);
	t.fill(7,65109,65110);
	t.fill(13,65110,65119);
	t.fill(5,65119,65120);
	t.fill(13,65120,65122);
	t.fill(4,65122,65124);
	t.fill(13,65124,65127);
	t.fill(13,65128,65129);
	t.fill(5,65129,65131);
	t.fill(13,65131,65132);
	t.fill(2,65136,65279);
	t.fill(9,65279,65280);
	t.fill(13,65281,65283);
	t.fill(5,65283,65286);
	t.fill(13,65286,65291);
	t.fill(4,65291,65292);
	t.fill(7,65292,65293);
	t.fill(4,65293,65294);
	t.fill(7,65294,65296);
	t.fill(3,65296,65306);
	t.fill(7,65306,65307);
	t.fill(13,65307,65313);
	t.fill(13,65339,65345);
	t.fill(13,65371,65382);
	t.fill(5,65504,65506);
	t.fill(13,65506,65509);
	t.fill(5,65509,65511);
	t.fill(13,65512,65519);
	t.fill(9,65520,65529);
	t.fill(13,65529,65534);
	t.fill(9,65534,65536);
	t.fill(13,65793,65794);
	t.fill(13,65856,65933);
	t.fill(13,65936,65949);
	t.fill(13,65952,65953);
	t.fill(8,66045,66046);
	t.fill(8,66272,66273);
	t.fill(3,66273,66300);
	t.fill(8,66422,66427);
	t.fill(1,67584,67590);
	t.fill(1,67592,67593);
	t.fill(1,67594,67638);
	t.fill(1,67639,67641);
	t.fill(1,67644,67645);
	t.fill(1,67647,67670);
	t.fill(1,67671,67743);
	t.fill(1,67751,67760);
	t.fill(1,67808,67827);
	t.fill(1,67828,67830);
	t.fill(1,67835,67868);
	t.fill(13,67871,67872);
	t.fill(1,67872,67898);
	t.fill(1,67903,67904);
	t.fill(1,67968,68024);
	t.fill(1,68028,68048);
	t.fill(1,68050,68097);
	t.fill(8,68097,68100);
	t.fill(8,68101,68103);
	t.fill(8,68108,68112);
	t.fill(1,68112,68116);
	t.fill(1,68117,68120);
	t.fill(1,68121,68150);
	t.fill(8,68152,68155);
	t.fill(8,68159,68160);
	t.fill(1,68160,68169);
	t.fill(1,68176,68185);
	t.fill(1,68192,68256);
	t.fill(1,68288,68325);
	t.fill(8,68325,68327);
	t.fill(1,68331,68343);
	t.fill(1,68352,68406);
	t.fill(13,68409,68416);
	t.fill(1,68416,68438);
	t.fill(1,68440,68467);
	t.fill(1,68472,68498);
	t.fill(1,68505,68509);
	t.fill(1,68521,68528);
	t.fill(1,68608,68681);
	t.fill(1,68736,68787);
	t.fill(1,68800,68851);
	t.fill(1,68858,68864);
	t.fill(2,68864,68900);
	t.fill(8,68900,68904);
	t.fill(6,68912,68922);
	t.fill(6,69216,69247);
	t.fill(1,69248,69290);
	t.fill(8,69291,69293);
	t.fill(1,69293,69294);
	t.fill(1,69296,69298);
	t.fill(8,69373,69376);
	t.fill(1,69376,69416);
	t.fill(2,69424,69446);
	t.fill(8,69446,69457);
	t.fill(2,69457,69466);
	t.fill(1,69488,69506);
	t.fill(8,69506,69510);
	t.fill(1,69510,69514);
	t.fill(1,69552,69580);
	t.fill(1,69600,69623);
	t.fill(8,69633,69634);
	t.fill(8,69688,69703);
	t.fill(13,69714,69734);
	t.fill(8,69744,69745);
	t.fill(8,69747,69749);
	t.fill(8,69759,69762);
	t.fill(8,69811,69815);
	t.fill(8,69817,69819);
	t.fill(8,69826,69827);
	t.fill(8,69888,69891);
	t.fill(8,69927,69932);
	t.fill(8,69933,69941);
	t.fill(8,70003,70004);
	t.fill(8,70016,70018);
	t.fill(8,70070,70079);
	t.fill(8,70089,70093);
	t.fill(8,70095,70096);
	t.fill(8,70191,70194);
	t.fill(8,70196,70197);
	t.fill(8,70198,70200);
	t.fill(8,70206,70207);
	t.fill(8,70209,70210);
	t.fill(8,70367,70368);
	t.fill(8,70371,70379);
	t.fill(8,70400,70402);
	t.fill(8,70459,70461);
	t.fill(8,70464,70465);
	t.fill(8,70502,70509);
	t.fill(8,70512,70517);
	t.fill(8,70712,70720);
	t.fill(8,70722,70725);
	t.fill(8,70726,70727);
	t.fill(8,70750,70751);
	t.fill(8,70835,70841);
	t.fill(8,70842,70843);
	t.fill(8,70847,70849);
	t.fill(8,70850,70852);
	t.fill(8,71090,71094);
	t.fill(8,71100,71102);
	t.fill(8,71103,71105);
	t.fill(8,71132,71134);
	t.fill(8,71219,71227);
	t.fill(8,71229,71230);
	t.fill(8,71231,71233);
	t.fill(13,71264,71277);
	t.fill(8,71339,71340);
	t.fill(8,71341,71342);
	t.fill(8,71344,71350);
	t.fill(8,71351,71352);
	t.fill(8,71453,71456);
	t.fill(8,71458,71462);
	t.fill(8,71463,71468);
	t.fill(8,71727,71736);
	t.fill(8,71737,71739);
	t.fill(8,71995,71997);
	t.fill(8,71998,71999);
	t.fill(8,72003,72004);
	t.fill(8,72148,72152);
	t.fill(8,72154,72156);
	t.fill(8,72160,72161);
	t.fill(8,72193,72199);
	t.fill(8,72201,72203);
	t.fill(8,72243,72249);
	t.fill(8,72251,72255);
	t.fill(8,72263,72264);
	t.fill(8,72273,72279);
	t.fill(8,72281,72284);
	t.fill(8,72330,72343);
	t.fill(8,72344,72346);
	t.fill(8,72752,72759);
	t.fill(8,72760,72766);
	t.fill(8,72850,72872);
	t.fill(8,72874,72881);
	t.fill(8,72882,72884);
	t.fill(8,72885,72887);
	t.fill(8,73009,73015);
	t.fill(8,73018,73019);
	t.fill(8,73020,73022);
	t.fill(8,73023,73030);
	t.fill(8,73031,73032);
	t.fill(8,73104,73106);
	t.fill(8,73109,73110);
	t.fill(8,73111,73112);
	t.fill(8,73459,73461);
	t.fill(8,73472,73474);
	t.fill(8,73526,73531);
	t.fill(8,73536,73537);
	t.fill(8,73538,73539);
	t.fill(13,73685,73693);
	t.fill(5,73693,73697);
	t.fill(13,73697,73714);
	t.fill(8,78912,78913);
	t.fill(8,78919,78934);
	t.fill(8,92912,92917);
	t.fill(8,92976,92983);
	t.fill(8,94031,94032);
	t.fill(8,94095,94099);
	t.fill(13,94178,94179);
	t.fill(8,94180,94181);
	t.fill(8,113821,113823);
	t.fill(9,113824,113828);
	t.fill(8,118528,118574);
	t.fill(8,118576,118599);
	t.fill(8,119143,119146);
	t.fill(9,119155,119163);
	t.fill(8,119163,119171);
	t.fill(8,119173,119180);
	t.fill(8,119210,119214);
	t.fill(13,119273,119275);
	t.fill(13,119296,119362);
	t.fill(8,119362,119365);
	t.fill(13,119365,119366);
	t.fill(13,119552,119639);
	t.fill(13,120539,120540);
	t.fill(13,120597,120598);
	t.fill(13,120655,120656);
	t.fill(13,120713,120714);
	t.fill(13,120771,120772);
	t.fill(3,120782,120832);
	t.fill(8,121344,121399);
	t.fill(8,121403,121453);
	t.fill(8,121461,121462);
	t.fill(8,121476,121477);
	t.fill(8,121499,121504);
	t.fill(8,121505,121520);
	t.fill(8,122880,122887);
	t.fill(8,122888,122905);
	t.fill(8,122907,122914);
	t.fill(8,122915,122917);
	t.fill(8,122918,122923);
	t.fill(8,123023,123024);
	t.fill(8,123184,123191);
	t.fill(8,123566,123567);
	t.fill(8,123628,123632);
	t.fill(5,123647,123648);
	t.fill(8,124140,124144);
	t.fill(1,124928,125125);
	t.fill(1,125127,125136);
	t.fill(8,125136,125143);
	t.fill(1,125184,125252);
	t.fill(8,125252,125259);
	t.fill(1,125259,125260);
	t.fill(1,125264,125274);
	t.fill(1,125278,125280);
	t.fill(2,126065,126133);
	t.fill(2,126209,126270);
	t.fill(2,126464,126468);
	t.fill(2,126469,126496);
	t.fill(2,126497,126499);
	t.fill(2,126500,126501);
	t.fill(2,126503,126504);
	t.fill(2,126505,126515);
	t.fill(2,126516,126520);
	t.fill(2,126521,126522);
	t.fill(2,126523,126524);
	t.fill(2,126530,126531);
	t.fill(2,126535,126536);
	t.fill(2,126537,126538);
	t.fill(2,126539,126540);
	t.fill(2,126541,126544);
	t.fill(2,126545,126547);
	t.fill(2,126548,126549);
	t.fill(2,126551,126552);
	t.fill(2,126553,126554);
	t.fill(2,126555,126556);
	t.fill(2,126557,126558);
	t.fill(2,126559,126560);
	t.fill(2,126561,126563);
	t.fill(2,126564,126565);
	t.fill(2,126567,126571);
	t.fill(2,126572,126579);
	t.fill(2,126580,126584);
	t.fill(2,126585,126589);
	t.fill(2,126590,126591);
	t.fill(2,126592,126602);
	t.fill(2,126603,126620);
	t.fill(2,126625,126628);
	t.fill(2,126629,126634);
	t.fill(2,126635,126652);
	t.fill(13,126704,126706);
	t.fill(13,126976,127020);
	t.fill(13,127024,127124);
	t.fill(13,127136,127151);
	t.fill(13,127153,127168);
	t.fill(13,127169,127184);
	t.fill(13,127185,127222);
	t.fill(3,127232,127243);
	t.fill(13,127243,127248);
	t.fill(13,127279,127280);
	t.fill(13,127338,127344);
	t.fill(13,127405,127406);
	t.fill(13,127584,127590);
	t.fill(13,127744,128728);
	t.fill(13,128732,128749);
	t.fill(13,128752,128765);
	t.fill(13,128768,128887);
	t.fill(13,128891,128986);
	t.fill(13,128992,129004);
	t.fill(13,129008,129009);
	t.fill(13,129024,129036);
	t.fill(13,129040,129096);
	t.fill(13,129104,129114);
	t.fill(13,129120,129160);
	t.fill(13,129168,129198);
	t.fill(13,129200,129202);
	t.fill(13,129280,129620);
	t.fill(13,129632,129646);
	t.fill(13,129648,129661);
	t.fill(13,129664,129673);
	t.fill(13,129680,129726);
	t.fill(13,129727,129734);
	t.fill(13,129742,129756);
	t.fill(13,129760,129769);
	t.fill(13,129776,129785);
	t.fill(13,129792,129939);
	t.fill(13,129940,129995);
	t.fill(3,130032,130042);
	t.fill(9,131070,131072);
	t.fill(9,196606,196608);
	t.fill(9,262142,262144);
	t.fill(9,327678,327680);
	t.fill(9,393214,393216);
	t.fill(9,458750,458752);
	t.fill(9,524286,524288);
	t.fill(9,589822,589824);
	t.fill(9,655358,655360);
	t.fill(9,720894,720896);
	t.fill(9,786430,786432);
	t.fill(9,851966,851968);
	t.fill(9,917502,917504);
	t.fill(9,917505,917506);
	t.fill(9,917536,917632);
	t.fill(8,917760,918000);
	t.fill(9,983038,983040);
	t.fill(9,1048574,1048576);
	t.fill(9,1114110,1114112);
	
	AscBidi.getType = function(codePoint)
	{
		if (codePoint >= 0x110000 || codePoint < 0)
			return AscBidi.TYPE.L;
		return m[t[codePoint]];
	}
})(window);
