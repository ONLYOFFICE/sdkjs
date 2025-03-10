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

var lcid_unknown = 0x0000; // Unknown
var lcid_ar = 0x0001; // Arabic
var lcid_bg = 0x0002; // Bulgarian
var lcid_ca = 0x0003; // Catalan
var lcid_zhHans = 0x0004; // Chinese, Han (Simplified variant)
var lcid_cs = 0x0005; // Czech
var lcid_da = 0x0006; // Danish
var lcid_de = 0x0007; // German
var lcid_el = 0x0008; // Modern Greek (1453-)
var lcid_en = 0x0009; // English
var lcid_es = 0x000a; // Spanish
var lcid_fi = 0x000b; // Finnish
var lcid_fr = 0x000c; // French
var lcid_he = 0x000d; // Hebrew
var lcid_hu = 0x000e; // Hungarian
var lcid_is = 0x000f; // Icelandic
var lcid_it = 0x0010; // Italian
var lcid_ja = 0x0011; // Japanese
var lcid_ko = 0x0012; // Korean
var lcid_nl = 0x0013; // Dutch
var lcid_no = 0x0014; // Norwegian
var lcid_pl = 0x0015; // Polish
var lcid_pt = 0x0016; // Portuguese
var lcid_rm = 0x0017; // Romansh
var lcid_ro = 0x0018; // Romanian
var lcid_ru = 0x0019; // Russian
var lcid_hr = 0x001a; // Croatian
var lcid_sk = 0x001b; // Slovak
var lcid_sq = 0x001c; // Albanian
var lcid_sv = 0x001d; // Swedish
var lcid_th = 0x001e; // Thai
var lcid_tr = 0x001f; // Turkish
var lcid_ur = 0x0020; // Urdu
var lcid_id = 0x0021; // Indonesian
var lcid_uk = 0x0022; // Ukrainian
var lcid_be = 0x0023; // Belarusian
var lcid_sl = 0x0024; // Slovenian
var lcid_et = 0x0025; // Estonian
var lcid_lv = 0x0026; // Latvian
var lcid_lt = 0x0027; // Lithuanian
var lcid_tg = 0x0028; // Tajik
var lcid_fa = 0x0029; // Persian
var lcid_vi = 0x002a; // Vietnamese
var lcid_hy = 0x002b; // Armenian
var lcid_az = 0x002c; // Azerbaijani
var lcid_eu = 0x002d; // Basque
var lcid_hsb = 0x002e; // Upper Sorbian
var lcid_mk = 0x002f; // Macedonian
var lcid_tn = 0x0032; // Tswana
var lcid_xh = 0x0034; // Xhosa
var lcid_zu = 0x0035; // Zulu
var lcid_af = 0x0036; // Afrikaans
var lcid_ka = 0x0037; // Georgian
var lcid_fo = 0x0038; // Faroese
var lcid_hi = 0x0039; // Hindi
var lcid_mt = 0x003a; // Maltese
var lcid_se = 0x003b; // Northern Sami
var lcid_ga = 0x003c; // Irish
var lcid_ms = 0x003e; // Malay (macrolanguage)
var lcid_kk = 0x003f; // Kazakh
var lcid_ky = 0x0040; // Kirghiz
var lcid_sw = 0x0041; // Swahili (macrolanguage)
var lcid_tk = 0x0042; // Turkmen
var lcid_uz = 0x0043; // Uzbek
var lcid_tt = 0x0044; // Tatar
var lcid_bn = 0x0045; // Bengali
var lcid_pa = 0x0046; // Panjabi
var lcid_gu = 0x0047; // Gujarati
var lcid_or = 0x0048; // Oriya
var lcid_ta = 0x0049; // Tamil
var lcid_te = 0x004a; // Telugu
var lcid_kn = 0x004b; // Kannada
var lcid_ml = 0x004c; // Malayalam
var lcid_as = 0x004d; // Assamese
var lcid_mr = 0x004e; // Marathi
var lcid_sa = 0x004f; // Sanskrit
var lcid_mn = 0x0050; // Mongolian
var lcid_bo = 0x0051; // Tibetan
var lcid_cy = 0x0052; // Welsh
var lcid_km = 0x0053; // Central Khmer
var lcid_lo = 0x0054; // Lao
var lcid_gl = 0x0056; // Galician
var lcid_kok = 0x0057; // Konkani (macrolanguage)
var lcid_syr = 0x005a; // Syriac
var lcid_si = 0x005b; // Sinhala
var lcid_iu = 0x005d; // Inuktitut
var lcid_am = 0x005e; // Amharic
var lcid_tzm = 0x005f; // Central Atlas Tamazight
var lcid_ne = 0x0061; // Nepali
var lcid_fy = 0x0062; // Western Frisian
var lcid_ps = 0x0063; // Pushto
var lcid_fil = 0x0064; // Filipino
var lcid_dv = 0x0065; // Dhivehi
var lcid_ha = 0x0068; // Hausa
var lcid_yo = 0x006a; // Yoruba
var lcid_quz = 0x006b; // Cusco Quechua
var lcid_nso = 0x006c; // Pedi
var lcid_ba = 0x006d; // Bashkir
var lcid_lb = 0x006e; // Luxembourgish
var lcid_kl = 0x006f; // Kalaallisut
var lcid_ig = 0x0070; // Igbo
var lcid_ii = 0x0078; // Sichuan Yi
var lcid_arn = 0x007a; // Mapudungun
var lcid_moh = 0x007c; // Mohawk
var lcid_br = 0x007e; // Breton
var lcid_ug = 0x0080; // Uighur
var lcid_mi = 0x0081; // Maori
var lcid_oc = 0x0082; // Occitan (post 1500)
var lcid_co = 0x0083; // Corsican
var lcid_gsw = 0x0084; // Swiss German
var lcid_sah = 0x0085; // Yakut
var lcid_qut = 0x0086; //
var lcid_rw = 0x0087; // Kinyarwanda
var lcid_wo = 0x0088; // Wolof
var lcid_prs = 0x008c; // Dari
var lcid_gd = 0x0091; // Scottish Gaelic
var lcid_arSA = 0x0401; // Arabic, Saudi Arabia
var lcid_bgBG = 0x0402; // Bulgarian, Bulgaria
var lcid_caES = 0x0403; // Catalan, Spain
var lcid_zhTW = 0x0404; // Chinese, Taiwan, Province of China
var lcid_csCZ = 0x0405; // Czech, Czech Republic
var lcid_daDK = 0x0406; // Danish, Denmark
var lcid_deDE = 0x0407; // German, Germany
var lcid_elGR = 0x0408; // Modern Greek (1453-), Greece
var lcid_enUS = 0x0409; // English, United States
var lcid_esES_tradnl = 0x040a; // Spanish
var lcid_fiFI = 0x040b; // Finnish, Finland
var lcid_frFR = 0x040c; // French, France
var lcid_heIL = 0x040d; // Hebrew, Israel
var lcid_huHU = 0x040e; // Hungarian, Hungary
var lcid_isIS = 0x040f; // Icelandic, Iceland
var lcid_itIT = 0x0410; // Italian, Italy
var lcid_jaJP = 0x0411; // Japanese, Japan
var lcid_koKR = 0x0412; // Korean, Republic of Korea
var lcid_nlNL = 0x0413; // Dutch, Netherlands
var lcid_nbNO = 0x0414; // Norwegian Bokmal, Norway
var lcid_plPL = 0x0415; // Polish, Poland
var lcid_ptBR = 0x0416; // Portuguese, Brazil
var lcid_rmCH = 0x0417; // Romansh, Switzerland
var lcid_roRO = 0x0418; // Romanian, Romania
var lcid_ruRU = 0x0419; // Russian, Russian Federation
var lcid_hrHR = 0x041a; // Croatian, Croatia
var lcid_skSK = 0x041b; // Slovak, Slovakia
var lcid_sqAL = 0x041c; // Albanian, Albania
var lcid_svSE = 0x041d; // Swedish, Sweden
var lcid_thTH = 0x041e; // Thai, Thailand
var lcid_trTR = 0x041f; // Turkish, Turkey
var lcid_urPK = 0x0420; // Urdu, Pakistan
var lcid_idID = 0x0421; // Indonesian, Indonesia
var lcid_ukUA = 0x0422; // Ukrainian, Ukraine
var lcid_beBY = 0x0423; // Belarusian, Belarus
var lcid_slSI = 0x0424; // Slovenian, Slovenia
var lcid_etEE = 0x0425; // Estonian, Estonia
var lcid_lvLV = 0x0426; // Latvian, Latvia
var lcid_ltLT = 0x0427; // Lithuanian, Lithuania
var lcid_tgCyrlTJ = 0x0428; // Tajik, Cyrillic, Tajikistan
var lcid_faIR = 0x0429; // Persian, Islamic Republic of Iran
var lcid_viVN = 0x042a; // Vietnamese, Viet Nam
var lcid_hyAM = 0x042b; // Armenian, Armenia
var lcid_azLatnAZ = 0x042c; // Azerbaijani, Latin, Azerbaijan
var lcid_euES = 0x042d; // Basque, Spain
var lcid_wenDE = 0x042e; // Sorbian languages, Germany
var lcid_mkMK = 0x042f; // Macedonian, The Former Yugoslav Republic of Macedonia
var lcid_stZA = 0x0430; // Southern Sotho, South Africa
var lcid_tsZA = 0x0431; // Tsonga, South Africa
var lcid_tnZA = 0x0432; // Tswana, South Africa
var lcid_venZA = 0x0433; // South Africa
var lcid_xhZA = 0x0434; // Xhosa, South Africa
var lcid_zuZA = 0x0435; // Zulu, South Africa
var lcid_afZA = 0x0436; // Afrikaans, South Africa
var lcid_kaGE = 0x0437; // Georgian, Georgia
var lcid_foFO = 0x0438; // Faroese, Faroe Islands
var lcid_hiIN = 0x0439; // Hindi, India
var lcid_mtMT = 0x043a; // Maltese, Malta
var lcid_seNO = 0x043b; // Northern Sami, Norway
var lcid_msMY = 0x043e; // Malay (macrolanguage), Malaysia
var lcid_kkKZ = 0x043f; // Kazakh, Kazakhstan
var lcid_kyKG = 0x0440; // Kirghiz, Kyrgyzstan
var lcid_swKE = 0x0441; // Swahili (macrolanguage), Kenya
var lcid_tkTM = 0x0442; // Turkmen, Turkmenistan
var lcid_uzLatnUZ = 0x0443; // Uzbek, Latin, Uzbekistan
var lcid_ttRU = 0x0444; // Tatar, Russian Federation
var lcid_bnIN = 0x0445; // Bengali, India
var lcid_paIN = 0x0446; // Panjabi, India
var lcid_guIN = 0x0447; // Gujarati, India
var lcid_orIN = 0x0448; // Oriya, India
var lcid_taIN = 0x0449; // Tamil, India
var lcid_teIN = 0x044a; // Telugu, India
var lcid_knIN = 0x044b; // Kannada, India
var lcid_mlIN = 0x044c; // Malayalam, India
var lcid_asIN = 0x044d; // Assamese, India
var lcid_mrIN = 0x044e; // Marathi, India
var lcid_saIN = 0x044f; // Sanskrit, India
var lcid_mnMN = 0x0450; // Mongolian, Mongolia
var lcid_boCN = 0x0451; // Tibetan, China
var lcid_cyGB = 0x0452; // Welsh, United Kingdom
var lcid_kmKH = 0x0453; // Central Khmer, Cambodia
var lcid_loLA = 0x0454; // Lao, Lao People's Democratic Republic
var lcid_myMM = 0x0455; // Burmese, Myanmar
var lcid_glES = 0x0456; // Galician, Spain
var lcid_kokIN = 0x0457; // Konkani (macrolanguage), India
var lcid_mni = 0x0458; // Manipuri
var lcid_sdIN = 0x0459; // Sindhi, India
var lcid_syrSY = 0x045a; // Syriac, Syrian Arab Republic
var lcid_siLK = 0x045b; // Sinhala, Sri Lanka
var lcid_chrUS = 0x045c; // Cherokee, United States
var lcid_iuCansCA = 0x045d; // Inuktitut, Unified Canadian Aboriginal Syllabics, Canada
var lcid_amET = 0x045e; // Amharic, Ethiopia
var lcid_tmz = 0x045f; // Tamanaku
var lcid_neNP = 0x0461; // Nepali, Nepal
var lcid_fyNL = 0x0462; // Western Frisian, Netherlands
var lcid_psAF = 0x0463; // Pushto, Afghanistan
var lcid_filPH = 0x0464; // Filipino, Philippines
var lcid_dvMV = 0x0465; // Dhivehi, Maldives
var lcid_binNG = 0x0466; // Bini, Nigeria
var lcid_fuvNG = 0x0467; // Nigerian Fulfulde, Nigeria
var lcid_haLatnNG = 0x0468; // Hausa, Latin, Nigeria
var lcid_ibbNG = 0x0469; // Ibibio, Nigeria
var lcid_yoNG = 0x046a; // Yoruba, Nigeria
var lcid_quzBO = 0x046b; // Cusco Quechua, Bolivia
var lcid_nsoZA = 0x046c; // Pedi, South Africa
var lcid_baRU = 0x046d; // Bashkir, Russian Federation
var lcid_lbLU = 0x046e; // Luxembourgish, Luxembourg
var lcid_klGL = 0x046f; // Kalaallisut, Greenland
var lcid_igNG = 0x0470; // Igbo, Nigeria
var lcid_krNG = 0x0471; // Kanuri, Nigeria
var lcid_gazET = 0x0472; // West Central Oromo, Ethiopia
var lcid_tiER = 0x0473; // Tigrinya, Eritrea
var lcid_gnPY = 0x0474; // Guarani, Paraguay
var lcid_hawUS = 0x0475; // Hawaiian, United States
var lcid_soSO = 0x0477; // Somali, Somalia
var lcid_iiCN = 0x0478; // Sichuan Yi, China
var lcid_papAN = 0x0479; // Papiamento, Netherlands Antilles
var lcid_arnCL = 0x047a; // Mapudungun, Chile
var lcid_mohCA = 0x047c; // Mohawk, Canada
var lcid_brFR = 0x047e; // Breton, France
var lcid_ugCN = 0x0480; // Uighur, China
var lcid_miNZ = 0x0481; // Maori, New Zealand
var lcid_ocFR = 0x0482; // Occitan (post 1500), France
var lcid_coFR = 0x0483; // Corsican, France
var lcid_gswFR = 0x0484; // Swiss German, France
var lcid_sahRU = 0x0485; // Yakut, Russian Federation
var lcid_qutGT = 0x0486; // Guatemala
var lcid_rwRW = 0x0487; // Kinyarwanda, Rwanda
var lcid_woSN = 0x0488; // Wolof, Senegal
var lcid_prsAF = 0x048c; // Dari, Afghanistan
var lcid_pltMG = 0x048d; // Plateau Malagasy, Madagascar
var lcid_gdGB = 0x0491; // Scottish Gaelic, United Kingdom
var lcid_arIQ = 0x0801; // Arabic, Iraq
var lcid_zhCN = 0x0804; // Chinese, China
var lcid_deCH = 0x0807; // German, Switzerland
var lcid_enGB = 0x0809; // English, United Kingdom
var lcid_esMX = 0x080a; // Spanish, Mexico
var lcid_frBE = 0x080c; // French, Belgium
var lcid_itCH = 0x0810; // Italian, Switzerland
var lcid_nlBE = 0x0813; // Dutch, Belgium
var lcid_nnNO = 0x0814; // Norwegian Nynorsk, Norway
var lcid_ptPT = 0x0816; // Portuguese, Portugal
var lcid_roMO = 0x0818; // Romanian, Macao
var lcid_ruMO = 0x0819; // Russian, Macao
var lcid_srLatnCS = 0x081a; // Serbian, Latin, Serbia and Montenegro
var lcid_svFI = 0x081d; // Swedish, Finland
var lcid_urIN = 0x0820; // Urdu, India
var lcid_azCyrlAZ = 0x082c; // Azerbaijani, Cyrillic, Azerbaijan
var lcid_dsbDE = 0x082e; // Lower Sorbian, Germany
var lcid_seSE = 0x083b; // Northern Sami, Sweden
var lcid_gaIE = 0x083c; // Irish, Ireland
var lcid_msBN = 0x083e; // Malay (macrolanguage), Brunei Darussalam
var lcid_uzCyrlUZ = 0x0843; // Uzbek, Cyrillic, Uzbekistan
var lcid_bnBD = 0x0845; // Bengali, Bangladesh
var lcid_paPK = 0x0846; // Panjabi, Pakistan
var lcid_mnMongCN = 0x0850; // Mongolian, Mongolian, China
var lcid_boBT = 0x0851; // Tibetan, Bhutan
var lcid_sdPK = 0x0859; // Sindhi, Pakistan
var lcid_iuLatnCA = 0x085d; // Inuktitut, Latin, Canada
var lcid_tzmLatnDZ = 0x085f; // Central Atlas Tamazight, Latin, Algeria
var lcid_neIN = 0x0861; // Nepali, India
var lcid_quzEC = 0x086b; // Cusco Quechua, Ecuador
var lcid_tiET = 0x0873; // Tigrinya, Ethiopia
var lcid_arEG = 0x0c01; // Arabic, Egypt
var lcid_zhHK = 0x0c04; // Chinese, Hong Kong
var lcid_deAT = 0x0c07; // German, Austria
var lcid_enAU = 0x0c09; // English, Australia
var lcid_esES = 0x0c0a; // Spanish, Spain
var lcid_frCA = 0x0c0c; // French, Canada
var lcid_srCyrlCS = 0x0c1a; // Serbian, Cyrillic, Serbia and Montenegro
var lcid_seFI = 0x0c3b; // Northern Sami, Finland
var lcid_tmzMA = 0x0c5f; // Tamanaku, Morocco
var lcid_quzPE = 0x0c6b; // Cusco Quechua, Peru
var lcid_arLY = 0x1001; // Arabic, Libyan Arab Jamahiriya
var lcid_zhSG = 0x1004; // Chinese, Singapore
var lcid_deLU = 0x1007; // German, Luxembourg
var lcid_enCA = 0x1009; // English, Canada
var lcid_esGT = 0x100a; // Spanish, Guatemala
var lcid_frCH = 0x100c; // French, Switzerland
var lcid_hrBA = 0x101a; // Croatian, Bosnia and Herzegovina
var lcid_smjNO = 0x103b; // Lule Sami, Norway
var lcid_arDZ = 0x1401; // Arabic, Algeria
var lcid_zhMO = 0x1404; // Chinese, Macao
var lcid_deLI = 0x1407; // German, Liechtenstein
var lcid_enNZ = 0x1409; // English, New Zealand
var lcid_esCR = 0x140a; // Spanish, Costa Rica
var lcid_frLU = 0x140c; // French, Luxembourg
var lcid_bsLatnBA = 0x141a; // Bosnian, Latin, Bosnia and Herzegovina
var lcid_smjSE = 0x143b; // Lule Sami, Sweden
var lcid_arMA = 0x1801; // Arabic, Morocco
var lcid_enIE = 0x1809; // English, Ireland
var lcid_esPA = 0x180a; // Spanish, Panama
var lcid_frMC = 0x180c; // French, Monaco
var lcid_srLatnBA = 0x181a; // Serbian, Latin, Bosnia and Herzegovina
var lcid_smaNO = 0x183b; // Southern Sami, Norway
var lcid_arTN = 0x1c01; // Arabic, Tunisia
var lcid_enZA = 0x1c09; // English, South Africa
var lcid_esDO = 0x1c0a; // Spanish, Dominican Republic
var lcid_frWest = 0x1c0c; // French
var lcid_srCyrlBA = 0x1c1a; // Serbian, Cyrillic, Bosnia and Herzegovina
var lcid_smaSE = 0x1c3b; // Southern Sami, Sweden
var lcid_arOM = 0x2001; // Arabic, Oman
var lcid_enJM = 0x2009; // English, Jamaica
var lcid_esVE = 0x200a; // Spanish, Venezuela
var lcid_frRE = 0x200c; // French, Reunion
var lcid_bsCyrlBA = 0x201a; // Bosnian, Cyrillic, Bosnia and Herzegovina
var lcid_smsFI = 0x203b; // Skolt Sami, Finland
var lcid_arYE = 0x2401; // Arabic, Yemen
var lcid_enCB = 0x2409; // English
var lcid_esCO = 0x240a; // Spanish, Colombia
var lcid_frCG = 0x240c; // French, Congo
var lcid_srLatnRS = 0x241a; // Serbian, Latin, Serbia
var lcid_smnFI = 0x243b; // Inari Sami, Finland
var lcid_arSY = 0x2801; // Arabic, Syrian Arab Republic
var lcid_enBZ = 0x2809; // English, Belize
var lcid_esPE = 0x280a; // Spanish, Peru
var lcid_frSN = 0x280c; // French, Senegal
var lcid_srCyrlRS = 0x281a; // Serbian, Cyrillic, Serbia
var lcid_arJO = 0x2c01; // Arabic, Jordan
var lcid_enTT = 0x2c09; // English, Trinidad and Tobago
var lcid_esAR = 0x2c0a; // Spanish, Argentina
var lcid_frCM = 0x2c0c; // French, Cameroon
var lcid_srLatnME = 0x2c1a; // Serbian, Latin, Montenegro
var lcid_arLB = 0x3001; // Arabic, Lebanon
var lcid_enZW = 0x3009; // English, Zimbabwe
var lcid_esEC = 0x300a; // Spanish, Ecuador
var lcid_frCI = 0x300c; // French, Cote d'Ivoire
var lcid_srCyrlME = 0x301a; // Serbian, Cyrillic, Montenegro
var lcid_arKW = 0x3401; // Arabic, Kuwait
var lcid_enPH = 0x3409; // English, Philippines
var lcid_esCL = 0x340a; // Spanish, Chile
var lcid_frML = 0x340c; // French, Mali
var lcid_arAE = 0x3801; // Arabic, United Arab Emirates
var lcid_enID = 0x3809; // English, Indonesia
var lcid_esUY = 0x380a; // Spanish, Uruguay
var lcid_frMA = 0x380c; // French, Morocco
var lcid_arBH = 0x3c01; // Arabic, Bahrain
var lcid_enHK = 0x3c09; // English, Hong Kong
var lcid_esPY = 0x3c0a; // Spanish, Paraguay
var lcid_frHT = 0x3c0c; // French, Haiti
var lcid_arQA = 0x4001; // Arabic, Qatar
var lcid_enIN = 0x4009; // English, India
var lcid_esBO = 0x400a; // Spanish, Bolivia
var lcid_enMY = 0x4409; // English, Malaysia
var lcid_esSV = 0x440a; // Spanish, El Salvador
var lcid_enSG = 0x4809; // English, Singapore
var lcid_esHN = 0x480a; // Spanish, Honduras
var lcid_esNI = 0x4c0a; // Spanish, Nicaragua
var lcid_esPR = 0x500a; // Spanish, Puerto Rico
var lcid_esUS = 0x540a; // Spanish, United States
var lcid_bsCyrl = 0x641a; // Bosnian, Cyrillic
var lcid_bsLatn = 0x681a; // Bosnian, Latin
var lcid_srCyrl = 0x6c1a; // Serbian, Cyrillic
var lcid_srLatn = 0x701a; // Serbian, Latin
var lcid_smn = 0x703b; // Inari Sami
var lcid_azCyrl = 0x742c; // Azerbaijani, Cyrillic
var lcid_sms = 0x743b; // Skolt Sami
var lcid_zh = 0x7804; // Chinese
var lcid_nn = 0x7814; // Norwegian Nynorsk
var lcid_bs = 0x781a; // Bosnian
var lcid_azLatn = 0x782c; // Azerbaijani, Latin
var lcid_sma = 0x783b; // Southern Sami
var lcid_uzCyrl = 0x7843; // Uzbek, Cyrillic
var lcid_mnCyrl = 0x7850; // Mongolian, Cyrillic
var lcid_iuCans = 0x785d; // Inuktitut, Unified Canadian Aboriginal Syllabics
var lcid_zhHant = 0x7c04; // Chinese, Han (Traditional variant)
var lcid_nb = 0x7c14; // Norwegian Bokmal
var lcid_sr = 0x7c1a; // Serbian
var lcid_tgCyrl = 0x7c28; // Tajik, Cyrillic
var lcid_dsb = 0x7c2e; // Lower Sorbian
var lcid_smj = 0x7c3b; // Lule Sami
var lcid_uzLatn = 0x7c43; // Uzbek, Latin
var lcid_mnMong = 0x7c50; // Mongolian, Mongolian
var lcid_iuLatn = 0x7c5d; // Inuktitut, Latin
var lcid_tzmLatn = 0x7c5f; // Central Atlas Tamazight, Latin
var lcid_haLatn = 0x7c68; // Hausa, Latin

//Defines are only allowed in the global scope.
//https://github.com/google/closure-compiler/wiki/Annotating-JavaScript-for-the-Closure-Compiler#define-type-description
window['AscCommon'] = window.AscCommon = window['AscCommon'] || {};
/** @define {string} */
window.AscCommon.g_cCompanyName  = "onlyoffice";
/** @define {string} */
window.AscCommon.g_cProductVersion  = "0.0.0";
/** @define {string} */
window.AscCommon.g_cBuildNumber = "0";
/** @define {string} */
window.AscCommon.g_cIsBeta = "false";

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
	function(window, undefined)
{
	window['Asc'] = window['Asc'] || {};
	
	var g_cCharDelimiter      = String.fromCharCode(5);
	var g_cGeneralFormat      = 'General';
	var FONT_THUMBNAIL_HEIGHT = (window["AscDesktopEditor"] && window["AscDesktopEditor"]["GetFontThumbnailHeight"]) ? window["AscDesktopEditor"]["GetFontThumbnailHeight"]() : 28;
	var c_oAscMaxColumnWidth  = 255;
	var c_oAscMaxRowHeight    = 409.5;
	var c_nMaxConversionTime  = 900000;//depends on config
	var c_nCommonRequestTime  = 30000;//30sec
	var c_nMaxDownloadTitleLen= 255;
	var c_nVersionNoBase64 = 10;
	var c_dMaxParaRunContentLength = 256;
	var c_nMaxHyperlinkLength = 2083;
	var c_sNativeViewerFormats = '.pdf.xps.oxps.djvu';
	var c_sShardKeyName = 'shardkey';
	var c_sWopiSrcName = 'WOPISrc';
	var c_sUserSessionIdName = 'usid';

	//files type for Saving & DownloadAs
	var c_oAscFileType = {
		UNKNOWN : 0,
		PDF     : 0x0201,
		PDFA    : 0x0209,
		DJVU    : 0x0203,
		XPS     : 0x0204,

		// Word
		DOCX : 0x0041,
		DOC  : 0x0042,
		ODT  : 0x0043,
		RTF  : 0x0044,
		TXT  : 0x0045,
		HTML : 0x0046,
		MHT  : 0x0047,
		EPUB : 0x0048,
		FB2  : 0x0049,
		MOBI : 0x004a,
		DOCM : 0x004b,
		DOTX : 0x004c,
		DOTM : 0x004d,
		FODT : 0x004e,
		OTT  : 0x004f,
		DOC_FLAT  : 0x0050,
		DOCX_FLAT  : 0x0051,
		HTML_IN_CONTAINER  : 0x0052,
		DOCX_PACKAGE  : 0x0054,
		OFORM  : 0x0055,
		DOCXF  : 0x0056,
		DOCY : 0x1001,
		CANVAS_WORD : 0x2001,
		JSON : 0x0808,	// Для mail-merge

		// Excel
		XLSX : 0x0101,
		XLS  : 0x0102,
		ODS  : 0x0103,
		CSV  : 0x0104,
		XLSM : 0x0105,
		XLTX : 0x0106,
		XLTM : 0x0107,
		XLSB : 0x0108,
		FODS : 0x0109,
		OTS  : 0x010a,
		XLSX_FLAT  : 0x010b,
		XLSX_PACKAGE  : 0x010c,
		XLSY : 0x1002,
		CANVAS_SPREADSHEET : 0x2002,

		// PowerPoint
		PPTX : 0x0081,
		PPT  : 0x0082,
		ODP  : 0x0083,
		PPSX : 0x0084,
		PPTM : 0x0085,
		PPSM : 0x0086,
		POTX : 0x0087,
		POTM : 0x0088,
		FODP : 0x0089,
		OTP  : 0x008a,
		PPTX_PACKAGE  : 0x008b,
		PPTY : 0x1003,
		CANVAS_PRESENTATION : 0x2003,

		//Draw
		VSDX : 0x4001,
		VSSX : 0x4002,
		VSTX : 0x4003,
		VSDM : 0x4004,
		VSSM : 0x4005,
		VSTM : 0x4006,

		//image
		IMG  : 0x0400,
		JPG  : 0x0401,
		TIFF : 0x0402,
		TGA  : 0x0403,
		GIF  : 0x0404,
		PNG  : 0x0405,
		EMF  : 0x0406,
		WMF  : 0x0407,
		BMP  : 0x0408,
		CR2  : 0x0409,
		PCX  : 0x040a,
		RAS  : 0x040b,
		PSD  : 0x040c,
		ICO  : 0x040d
	};

	var c_oAscTextAssociation = {
		BlockChar : 0,
		BlockLine : 1,
		PlainLine : 2,
		PlainParagraph : 3
	};

	var c_oAscAsyncAction = {
		Open               : 0,  // открытие документа
		Save               : 1,  // сохранение
		LoadDocumentFonts  : 2,  // загружаем фонты документа (сразу после открытия)
		LoadDocumentImages : 3,  // загружаем картинки документа (сразу после загрузки шрифтов)
		LoadFont           : 4,  // подгрузка нужного шрифта
		LoadImage          : 5,  // подгрузка картинки
		DownloadAs         : 6,  // cкачать
		Print              : 7,  // конвертация в PDF и сохранение у пользователя
		UploadImage        : 8,  // загрузка картинки

		ApplyChanges : 9,  // применение изменений от другого пользователя.

		SlowOperation     : 11, // медленная операция
		LoadTheme         : 12, // загрузка темы
		MailMergeLoadFile : 13, // загрузка файла для mail merge
		DownloadMerge     : 14, // cкачать файл с mail merge
		SendMailMerge     : 15,  // рассылка mail merge по почте
		ForceSaveButton   : 16,
		ForceSaveTimeout  : 17,
		Waiting	: 18,
		Submit : 19,
		Disconnect :20,
		RefreshFile: 21
	};

	var c_oAscAdvancedOptionsID = {
		CSV : 0,
		TXT : 1,
		DRM : 2
	};

	var c_oAscAdvancedOptionsAction = {
		None : 0,
		Open : 1,
		Save : 2
	};

	var c_oAscRestrictionType = {
		None           : 0x00,
		OnlyForms      : 0x01,
		OnlyComments   : 0x02,
		OnlySignatures : 0x04,
		View           : 0x80 // the difference between this restriction and the simple ViewMode -
		                      // the editor opens as a full editor, but we forbid ANY editing
		                      // in the ViewMode - opens viewer
	};

	var c_oAscLocalRestrictionType = {
		None		: 0x00,
		ReadOnly	: 0x01,
		Locked		: 0x02,
		Nosafe		: 0x04
	};

	// Режимы отрисовки
	var c_oAscFontRenderingModeType = {
		noHinting             : 1,
		hinting               : 2,
		hintingAndSubpixeling : 3
	};

	var c_oAscAsyncActionType = {
		Information      : 0,
		BlockInteraction : 1,
		Empty            : 2
	};

	var DownloadType = {
		None      : '',
		Download  : 'asc_onDownloadUrl',
		Print     : 'asc_onPrintUrl',
		MailMerge : 'asc_onSaveMailMerge'
	};

	var c_oAscFrameDataType = {
		SendImageUrls: 0,
		GetLoadedImages: 1,
		OpenFrame: 2,
		ShowImageDialogInFrame: 3,
		GetUrlsFromImageDialog: 4,
		SkipStartEndAction: 5,
		StartUploadImageAction: 6
	};

	var CellValueType = {
		Number : 0,
		String : 1,
		Bool   : 2,
		Error  : 3
	};

	/**
	 * NumFormat defines
	 * @enum {number}
	 */
	var c_oAscNumFormatType = {
		None       : -1,
		General    : 0,
		Number     : 1,
		Scientific : 2,
		Accounting : 3,
		Currency   : 4,
		Date       : 5,
		LongDate   : 6,
		Time       : 7,
		Percent    : 8,
		Fraction   : 9,
		Text       : 10,
		Custom     : 11
	};

	var c_oAscDrawingLayerType = {
		BringToFront : 0,
		BringForward : 1,
		SendToBack   : 2,
		SendBackward : 3
	};

	var c_oAscCellAnchorType = {
		cellanchorAbsolute : 0,
		cellanchorOneCell  : 1,
		cellanchorTwoCell  : 2
	};

	var c_oAscChartDefines = {
		defaultChartWidth  : 478,
		defaultChartHeight : 286
	};

	var c_oAscStyleImage = {
		Default  : 0,
		Document : 1
	};

	var c_oAscTypeSelectElement = {
		Paragraph         : 0,
		Table             : 1,
		Image             : 2,
		Header            : 3,
		Hyperlink         : 4,
		SpellCheck        : 5,
		Shape             : 6,
		Slide             : 7,
		Chart             : 8,
		Math              : 9,
		MailMerge         : 10,
		ContentControl    : 11,
		Animation         : 12,
		UnProtectedRegion : 13,
		Text              : 14, // viewer
		Annot             : 15,
		PdfPage           : 16
	};

	var c_oAscLineDrawingRule = {
		Left   : 0,
		Center : 1,
		Right  : 2,
		Top    : 0,
		Bottom : 2
	};

	var align_Right   = 0;
	var align_Left    = 1;
	var align_Center  = 2;
	var align_Justify = 3;
	var align_Distributed = 4;
	var align_CenterContinuous = 5;


	var linerule_AtLeast = 0x00;
	var linerule_Auto    = 0x01;
	var linerule_Exact   = 0x02;

	var c_oAscShdClear = 0;
	var c_oAscShdNil   = 1;

	var c_oAscShd = {
		Clear                 : 0,
		Nil                   : 1,
		DiagCross             : 2,
		DiagStripe            : 3,
		HorzCross             : 4,
		HorzStripe            : 5,
		Pct10                 : 6,
		Pct12                 : 7,
		Pct15                 : 8,
		Pct20                 : 9,
		Pct25                 : 10,
		Pct30                 : 11,
		Pct35                 : 12,
		Pct37                 : 13,
		Pct40                 : 14,
		Pct45                 : 15,
		Pct5                  : 16,
		Pct50                 : 17,
		Pct55                 : 18,
		Pct60                 : 19,
		Pct62                 : 20,
		Pct65                 : 21,
		Pct70                 : 22,
		Pct75                 : 23,
		Pct80                 : 24,
		Pct85                 : 25,
		Pct87                 : 26,
		Pct90                 : 27,
		Pct95                 : 28,
		ReverseDiagStripe     : 29,
		Solid                 : 30,
		ThinDiagCross         : 31,
		ThinDiagStripe        : 32,
		ThinHorzCross         : 33,
		ThinHorzStripe        : 34,
		ThinReverseDiagStripe : 35,
		ThinVertStripe        : 36,
		VertStripe            : 37
	};

	var vertalign_Baseline    = 0;
	var vertalign_SubScript   = 1;
	var vertalign_SuperScript = 2;

	var hdrftr_Header         = 0x01;
	var hdrftr_Footer         = 0x02;

	var vaKSize  =  0.65;  // Коэффициент изменения размера текста для верхнего и нижнего индексов
	var vaKSuper =  0.35;  // Позиция верхнего индекса (относительно размера текста)
	var vaKSub   = -0.141; // Позиция нижнего индекса (относительно размера текста)

	var c_oAscDropCap = {
		None   : 0x00,
		Drop   : 0x01,
		Margin : 0x02
	};


	var c_oAscChartTitleShowSettings = {
		none      : 0,
		overlay   : 1,
		noOverlay : 2
	};

	var c_oAscChartHorAxisLabelShowSettings = {
		none      : 0,
		noOverlay : 1
	};

	var c_oAscChartVertAxisLabelShowSettings = {
		none       : 0,
		rotated    : 1,
		vertical   : 2,
		horizontal : 3
	};

	var c_oAscChartLegendShowSettings = {
		none         : 0,
		left         : 1,
		top          : 2,
		right        : 3,
		bottom       : 4,
		leftOverlay  : 5,
		rightOverlay : 6,
		layout       : 7,
		topRight     : 8 // ToDo добавить в меню
	};

	var c_oAscChartDataLabelsPos = {
		none    : 0,
		b       : 1,
		bestFit : 2,
		ctr     : 3,
		inBase  : 4,
		inEnd   : 5,
		l       : 6,
		outEnd  : 7,
		r       : 8,
		t       : 9,
		show    : 10
	};

	var c_oAscGridLinesSettings = {
		none       : 0,
		major      : 1,
		minor      : 2,
		majorMinor : 3
	};

	var c_oAscChartTypeSettings = {
		barNormal              : 0,
		barStacked             : 1,
		barStackedPer          : 2,
		barNormal3d            : 3,
		barStacked3d           : 4,
		barStackedPer3d        : 5,
		barNormal3dPerspective : 6,
		lineNormal             : 7,
		lineStacked            : 8,
		lineStackedPer         : 9,
		lineNormalMarker       : 10,
		lineStackedMarker      : 11,
		lineStackedPerMarker   : 12,
		line3d                 : 13,
		pie                    : 14,
		pie3d                  : 15,
		hBarNormal             : 16,
		hBarStacked            : 17,
		hBarStackedPer         : 18,
		hBarNormal3d           : 19,
		hBarStacked3d          : 20,
		hBarStackedPer3d       : 21,
		areaNormal             : 22,
		areaStacked            : 23,
		areaStackedPer         : 24,
		doughnut               : 25,
		stock                  : 26,
		scatter                : 27,
		scatterLine            : 28,
		scatterLineMarker      : 29,
		scatterMarker          : 30,
		scatterNone            : 31,
		scatterSmooth          : 32,
		scatterSmoothMarker    : 33,
		surfaceNormal          : 34,
		surfaceWireframe       : 35,
		contourNormal          : 36,
		contourWireframe       : 37,
		comboCustom            : 38,
		comboBarLine           : 39,
		comboBarLineSecondary  : 40,
		comboAreaBar           : 41,
		radar                  : 42,
		radarMarker            : 43,
		radarFilled            : 44,
		unknown                : 45
	};

	var c_oAscValAxisRule = {
		auto  : 0,
		fixed : 1
	};

	var c_oAscValAxUnits = {
		none              : 0,
		BILLIONS          : 1,
		HUNDRED_MILLIONS  : 2,
		HUNDREDS          : 3,
		HUNDRED_THOUSANDS : 4,
		MILLIONS          : 5,
		TEN_MILLIONS      : 6,
		TEN_THOUSANDS     : 7,
		TRILLIONS         : 8,
		CUSTOM            : 9,
		THOUSANDS         : 10

	};

	var c_oAscTickMark = {
		TICK_MARK_CROSS : 0,
		TICK_MARK_IN    : 1,
		TICK_MARK_NONE  : 2,
		TICK_MARK_OUT   : 3
	};

	var c_oAscTickLabelsPos = {
		TICK_LABEL_POSITION_HIGH    : 0,
		TICK_LABEL_POSITION_LOW     : 1,
		TICK_LABEL_POSITION_NEXT_TO : 2,
		TICK_LABEL_POSITION_NONE    : 3
	};

	var c_oAscCrossesRule = {
		auto     : 0,
		maxValue : 1,
		value    : 2,
		minValue : 3
	};

	var c_oAscBetweenLabelsRule = {
		auto   : 0,
		manual : 1
	};

	var c_oAscLabelsPosition = {
		byDivisions      : 0,
		betweenDivisions : 1
	};

	var c_oAscAxisType = {
		date : 0,
		cat  : 1,
		val  : 2,
		ser  : 3
	};

	var c_oAscHAnchor = {
		Margin : 0x00,
		Page   : 0x01,
		Text   : 0x02
	};

	var c_oAscXAlign = {
		Center  : 0x00,
		Inside  : 0x01,
		Left    : 0x02,
		Outside : 0x03,
		Right   : 0x04
	};
	var c_oAscYAlign = {
		Bottom  : 0x00,
		Center  : 0x01,
		Inline  : 0x02,
		Inside  : 0x03,
		Outside : 0x04,
		Top     : 0x05
	};

	var c_oAscVAnchor = {
		Margin : 0x00,
		Page   : 0x01,
		Text   : 0x02
	};

	var c_oAscRelativeFromH = {
		Character     : 0x00,
		Column        : 0x01,
		InsideMargin  : 0x02,
		LeftMargin    : 0x03,
		Margin        : 0x04,
		OutsideMargin : 0x05,
		Page          : 0x06,
		RightMargin   : 0x07
	};

	var c_oAscSizeRelFromH = {
		sizerelfromhMargin        : 0,
		sizerelfromhPage          : 1,
		sizerelfromhLeftMargin    : 2,
		sizerelfromhRightMargin   : 3,
		sizerelfromhInsideMargin  : 4,
		sizerelfromhOutsideMargin : 5
	};

	var c_oAscSizeRelFromV = {
		sizerelfromvMargin        : 0,
		sizerelfromvPage          : 1,
		sizerelfromvTopMargin     : 2,
		sizerelfromvBottomMargin  : 3,
		sizerelfromvInsideMargin  : 4,
		sizerelfromvOutsideMargin : 5
	};

	var c_oAscRelativeFromV = {
		BottomMargin  : 0x00,
		InsideMargin  : 0x01,
		Line          : 0x02,
		Margin        : 0x03,
		OutsideMargin : 0x04,
		Page          : 0x05,
		Paragraph     : 0x06,
		TopMargin     : 0x07
	};

	var c_oAscRectAlignType = {
		b:   0,
		bl:  1,
		br:  2,
		ctr: 3,
		l:   4,
		r:   5,
		t:   6,
		tl:  7,
		tr:  8
	};

	var ST_HueDir = {
		Ccw: 0,
		Cw:  1
	};

	var ST_ClrAppMethod = {
		cycle:  0,
		repeat: 1,
		span:   2
	}

	var ST_AnimLvlStr = {
		ctr:  0,
		lvl:  1,
		none: 2
	}

	var ST_AnimOneStr = {
		branch: 0,
		none:   1,
		one:    2
	}

	var ST_Direction = {
		norm: 0,
		rev:  1
	}
	
	var ST_HierBranchStyle = {
		hang: 0,
		init: 1,
		l:    2,
		r:    3,
		std:  4
	}

	var ST_ResizeHandlesStr = {
		exact: 0,
		rel:   1
	}

	var ST_PtType = {
		node:     0,
		asst:     1,
		doc:      2,
		pres:     3,
		parTrans: 4,
		sibTrans: 5
	}

	var ST_ChildOrderType = {
		b: 0,
		t: 1
	}

	var ST_AlgorithmType = {
		composite: 0,
		conn:      1,
		cycle:     2,
		hierChild: 3,
		hierRoot:  4,
		pyra:      5,
		lin:       6,
		sp:        7,
		tx:        8,
		snake:     9
	}

	var ST_ConstraintRelationship = {
		self: 0,
		ch:   1,
		des:  2
	}

	var ST_BoolOperator = {
		none: 0,
		equ:  1,
		gte:  2,
		lte:  3
	}

	var ST_ElementType = {
		all:      0,
		doc:      1,
		node:     2,
		norm:     3,
		nonNorm:  4,
		asst:     5,
		nonAsst:  6,
		parTrans: 7,
		pres:     8,
		sibTrans: 9
	}

	var ST_ConstraintType = {
		alignOff: 1,
		b: 5,
		begMarg: 2,
		begPad: 4,
		bendDist: 3,
		bMarg: 6,
		bOff: 7,
		connDist: 12,
		ctrX: 8,
		ctrXOff: 9,
		ctrY: 10,
		ctrYOff: 11,
		diam: 13,
		endMarg: 14,
		endPad: 15,
		h: 16,
		hArH: 17,
		hOff: 63, // TODO: add to constr type in x2t
		l: 18,
		lMarg: 19,
		lOff: 20,
		none: 0,
		primFontSz: 24,
		pyraAcctRatio: 25,
		r: 21,
		rMarg: 22,
		rOff: 23,
		secFontSz: 26,
		secSibSp: 28,
		sibSp: 27,
		sp: 29,
		stemThick: 30,
		t: 31,
		tMarg: 32,
		tOff: 33,
		userA: 34,
		userB: 35,
		userC: 36,
		userD: 37,
		userE: 38,
		userF: 39,
		userG: 40,
		userH: 41,
		userI: 42,
		userJ: 43,
		userK: 44,
		userL: 45,
		userM: 46,
		userN: 47,
		userO: 48,
		userP: 49,
		userQ: 50,
		userR: 51,
		userS: 52,
		userT: 53,
		userU: 54,
		userV: 55,
		userW: 56,
		userX: 57,
		userY: 58,
		userZ: 59,
		w: 60,
		wArH: 61,
		wOff: 62
	}

	var ST_VariableType = {
		animLvl:       0,
		animOne:       1,
		bulEnabled:    2,
		chMax:         3,
		chPref:        4,
		dir:           5,
		hierBranch:    6,
		none:          7,
		orgChart:      8,
		resizeHandles: 9
	}
	
	var ST_AxisType = {
		ancst: 6,
		ancstOrSelf: 7,
		ch: 2,
		des: 3,
		desOrSelf: 4,
		follow: 10,
		followSib: 8,
		none: 0,
		par: 5,
		preced: 11,
		precedSib: 9,
		root: 12,
		self: 1
	}

	var ST_FunctionType = {
		cnt: 0,
		depth: 6,
		maxDepth: 7,
		pos: 1,
		posEven: 3,
		posOdd: 4,
		revPos: 2,
		var: 5
	}

	var ST_FunctionOperator = {
		equ: 0,
		gt: 2,
		gte: 4,
		lt: 3,
		lte: 5,
		neq: 1
	}

	var ST_LayoutShapeType = {
		conn: 0,
		none: 1,
		accentBorderCallout1: 2,
		accentBorderCallout2: 3,
		accentBorderCallout3: 4,
		accentCallout1: 5,
		accentCallout2: 6,
		accentCallout3: 7,
		actionButtonBackPrevious: 8,
		actionButtonBeginning: 9,
		actionButtonBlank: 10,
		actionButtonDocument: 11,
		actionButtonEnd: 12,
		actionButtonForwardNext: 13,
		actionButtonHelp: 14,
		actionButtonHome: 15,
		actionButtonInformation: 16,
		actionButtonMovie: 17,
		actionButtonReturn: 18,
		actionButtonSound: 19,
		arc: 20,
		bentArrow: 21,
		bentConnector2: 22,
		bentConnector3: 23,
		bentConnector4: 24,
		bentConnector5: 25,
		bentUpArrow: 26,
		bevel: 27,
		blockArc: 28,
		borderCallout1: 29,
		borderCallout2: 30,
		borderCallout3: 31,
		bracePair: 32,
		bracketPair: 33,
		callout1: 34,
		callout2: 35,
		callout3: 36,
		can: 37,
		chartPlus: 38,
		chartStar: 39,
		chartX: 40,
		chevron: 41,
		chord: 42,
		circularArrow: 43,
		cloud: 44,
		cloudCallout: 45,
		corner: 46,
		cornerTabs: 47,
		cube: 48,
		curvedConnector2: 49,
		curvedConnector3: 50,
		curvedConnector4: 51,
		curvedConnector5: 52,
		curvedDownArrow: 53,
		curvedLeftArrow: 54,
		curvedRightArrow: 55,
		curvedUpArrow: 56,
		decagon: 57,
		diagStripe: 58,
		diamond: 59,
		dodecagon: 60,
		donut: 61,
		doubleWave: 62,
		downArrow: 63,
		downArrowCallout: 64,
		ellipse: 65,
		ellipseRibbon: 66,
		ellipseRibbon2: 67,
		flowChartAlternateProcess: 68,
		flowChartCollate: 69,
		flowChartConnector: 70,
		flowChartDecision: 71,
		flowChartDelay: 72,
		flowChartDisplay: 73,
		flowChartDocument: 74,
		flowChartExtract: 75,
		flowChartInputOutput: 76,
		flowChartInternalStorage: 77,
		flowChartMagneticDisk: 78,
		flowChartMagneticDrum: 79,
		flowChartMagneticTape: 80,
		flowChartManualInput: 81,
		flowChartManualOperation: 82,
		flowChartMerge: 83,
		flowChartMultidocument: 84,
		flowChartOfflineStorage: 85,
		flowChartOffpageConnector: 86,
		flowChartOnlineStorage: 87,
		flowChartOr: 88,
		flowChartPredefinedProcess: 89,
		flowChartPreparation: 90,
		flowChartProcess: 91,
		flowChartPunchedCard: 92,
		flowChartPunchedTape: 93,
		flowChartSort: 94,
		flowChartSummingJunction: 95,
		flowChartTerminator: 96,
		foldedCorner: 97,
		frame: 98,
		funnel: 99,
		gear6: 100,
		gear9: 101,
		halfFrame: 102,
		heart: 103,
		heptagon: 104,
		hexagon: 105,
		homePlate: 106,
		horizontalScroll: 107,
		irregularSeal1: 108,
		irregularSeal2: 109,
		leftArrow: 110,
		leftArrowCallout: 111,
		leftBrace: 112,
		leftBracket: 113,
		leftCircularArrow: 114,
		leftRightArrow: 115,
		leftRightArrowCallout: 116,
		leftRightCircularArrow: 117,
		leftRightRibbon: 118,
		leftRightUpArrow: 119,
		leftUpArrow: 120,
		lightningBolt: 121,
		line: 122,
		lineInv: 123,
		mathDivide: 124,
		mathEqual: 125,
		mathMinus: 126,
		mathMultiply: 127,
		mathNotEqual: 128,
		mathPlus: 129,
		moon: 130,
		nonIsoscelesTrapezoid: 131,
		noSmoking: 132,
		notchedRightArrow: 133,
		octagon: 134,
		parallelogram: 135,
		pentagon: 136,
		pie: 137,
		pieWedge: 138,
		plaque: 139,
		plaqueTabs: 140,
		plus: 141,
		quadArrow: 142,
		quadArrowCallout: 143,
		rect: 144,
		ribbon: 145,
		ribbon2: 146,
		rightArrow: 147,
		rightArrowCallout: 148,
		rightBrace: 149,
		rightBracket: 150,
		round1Rect: 151,
		round2DiagRect: 152,
		round2SameRect: 153,
		roundRect: 154,
		rtTriangle: 155,
		smileyFace: 156,
		snip1Rect: 157,
		snip2DiagRect: 158,
		snip2SameRect: 159,
		snipRoundRect: 160,
		squareTabs: 161,
		star10: 162,
		star12: 163,
		star16: 164,
		star24: 165,
		star32: 166,
		star4: 167,
		star5: 168,
		star6: 169,
		star7: 170,
		star8: 171,
		straightConnector1: 172,
		stripedRightArrow: 173,
		sun: 174,
		swooshArrow: 175,
		teardrop: 176,
		trapezoid: 177,
		triangle: 178,
		upArrow: 179,
		upArrowCallout: 180,
		upDownArrow: 181,
		upDownArrowCallout: 182,
		uturnArrow: 183,
		verticalScroll: 184,
		wave: 185,
		wedgeEllipseCallout: 186,
		wedgeRectCallout: 187,
		wedgeRoundRectCallout: 188
	}


	var ST_ParameterId = {
		alignTx: 0,
		ar: 1,
		autoTxRot: 2,
		begPts: 3,
		begSty: 4,
		bendPt: 5,
		bkpt: 6,
		bkPtFixedVal: 7,
		chAlign: 8,
		chDir: 9,
		connRout: 10,
		contDir: 11,
		ctrShpMap: 12,
		dim: 13,
		dstNode: 14,
		endPts: 15,
		endSty: 16,
		fallback: 17,
		flowDir: 18,
		grDir: 19,
		hierAlign: 20,
		horzAlign: 21,
		linDir: 22,
		lnSpAfChP: 23,
		lnSpAfParP: 24,
		lnSpCh: 25,
		lnSpPar: 26,
		nodeHorzAlign: 27,
		nodeVertAlign: 28,
		off: 29,
		parTxLTRAlign: 30,
		parTxRTLAlign: 31,
		pyraAcctBkgdNode: 32,
		pyraAcctPos: 33,
		pyraAcctTxMar: 34,
		pyraAcctTxNode: 35,
		pyraLvlNode: 36,
		rotPath: 37,
		rtShortDist: 38,
		secChAlign: 39,
		secLinDir: 40,
		shpTxLTRAlignCh: 41,
		shpTxRTLAlignCh: 42,
		spanAng: 43,
		srcNode: 44,
		stAng: 45,
		stBulletLvl: 46,
		stElem: 47,
		txAnchorHorz: 48,
		txAnchorHorzCh: 49,
		txAnchorVert: 50,
		txAnchorVertCh: 51,
		txBlDir: 52,
		txDir: 53,
		vertAlign: 54
	}

	var ST_PresetCameraType = {
		isometricBottomDown: 0,
		isometricBottomUp: 1,
		isometricLeftDown: 2,
		isometricLeftUp: 3,
		isometricOffAxis1Left: 4,
		isometricOffAxis1Right: 5,
		isometricOffAxis1Top: 6,
		isometricOffAxis2Left: 7,
		isometricOffAxis2Right: 8,
		isometricOffAxis2Top: 9,
		isometricOffAxis3Bottom: 10,
		isometricOffAxis3Left: 11,
		isometricOffAxis3Right: 12,
		isometricOffAxis4Bottom: 13,
		isometricOffAxis4Left: 14,
		isometricOffAxis4Right: 15,
		isometricRightDown: 16,
		isometricRightUp: 17,
		isometricTopDown: 18,
		isometricTopUp: 19,
		legacyObliqueBottom: 20,
		legacyObliqueBottomLeft: 21,
		legacyObliqueBottomRight: 22,
		legacyObliqueFront: 23,
		legacyObliqueLeft: 24,
		legacyObliqueRight: 25,
		legacyObliqueTop: 26,
		legacyObliqueTopLeft: 27,
		legacyObliqueTopRight: 28,
		legacyPerspectiveBottom: 29,
		legacyPerspectiveBottomLeft: 30,
		legacyPerspectiveBottomRight: 31,
		legacyPerspectiveFront: 32,
		legacyPerspectiveLeft: 33,
		legacyPerspectiveRight: 34,
		legacyPerspectiveTop: 35,
		legacyPerspectiveTopLeft: 36,
		legacyPerspectiveTopRight: 37,
		obliqueBottom: 38,
		obliqueBottomLeft: 39,
		obliqueBottomRight: 40,
		obliqueLeft: 41,
		obliqueRight: 42,
		obliqueTop: 43,
		obliqueTopLeft: 44,
		obliqueTopRight: 45,
		orthographicFront: 46,
		perspectiveAbove: 47,
		perspectiveAboveLeftFacing: 48,
		perspectiveAboveRightFacing: 49,
		perspectiveBelow: 50,
		perspectiveContrastingLeftFacing: 51,
		perspectiveContrastingRightFacing: 52,
		perspectiveFront: 53,
		perspectiveHeroicExtremeLeftFacing: 54,
		perspectiveHeroicExtremeRightFacing: 55,
		perspectiveHeroicLeftFacing: 56,
		perspectiveHeroicRightFacing: 57,
		perspectiveLeft: 58,
		perspectiveRelaxed: 59,
		perspectiveRelaxedModerately: 60,
		perspectiveRight: 61
	}

	var ST_LightRigDirection = {
		b: 0,
		bl: 1,
		br: 2,
		l: 4,
		r: 5,
		t: 6,
		tl: 7,
		tr: 8
	}

	var ST_LightRigType = {
		balanced: 0,
		brightRoom: 1,
		chilly: 2,
		contrasting: 3,
		flat: 4,
		flood: 5,
		freezing: 6,
		glow: 7,
		harsh: 8,
		legacyFlat1: 9,
		legacyFlat2: 10,
		legacyFlat3: 11,
		legacyFlat4: 12,
		legacyHarsh1: 13,
		legacyHarsh2: 14,
		legacyHarsh3: 15,
		legacyHarsh4: 16,
		legacyNormal1: 17,
		legacyNormal2: 18,
		legacyNormal3: 19,
		legacyNormal4: 20,
		morning: 21,
		soft: 22,
		sunrise: 23,
		sunset: 24,
		threePt: 25,
		twoPt: 26
	}

	var ST_BevelPresetType = {
		angle: 0,
		artDeco: 1,
		circle: 2,
		convex: 3,
		coolSlant: 4,
		cross: 5,
		divot: 6,
		hardEdge: 7,
		relaxedInset: 8,
		riblet: 9,
		slope: 10,
		softRound: 11
	}

	var ST_TLAnimateEffectTransition = {
		in:   0,
		out:  1,
		none: 2
	}

	var ST_PresetMaterialType = {
		clear: 0,
		dkEdge: 1,
		flat: 2,
		legacyMatte: 3,
		legacyMetal: 4,
		legacyPlastic: 5,
		legacyWireframe: 6,
		matte: 7,
		metal: 8,
		plastic: 9,
		powder: 10,
		softEdge: 11,
		softmetal: 12,
		translucentPowder: 13,
		warmMatte: 14
	}

	var ST_CxnType = {
		parOf: 0,
		presOf: 1,
		presParOf: 2,
		unknownRelationShip: 3
	}
	
	var c_oAscOleObjectTypes = {
		document:    1,
		spreadsheet: 2,
		formula:     4
	}

	// image wrap style
	var c_oAscWrapStyle = {
		Inline : 0,
		Flow   : 1
	};

	// Толщина бордера
	var c_oAscBorderWidth     = {
		None   : 0,	// 0px
		Thin   : 1,	// 1px
		Medium : 2,	// 2px
		Thick  : 3		// 3px
	};
	/**
	 * Располагаются в порядке значимости для отрисовки
	 * @enum {number}
	 */
	var c_oAscBorderStyles    = {
		None             : 0,
		Double           : 1,
		Hair             : 2,
		DashDotDot       : 3,
		DashDot          : 4,
		Dotted           : 5,
		Dashed           : 6,
		Thin             : 7,
		MediumDashDotDot : 8,
		SlantDashDot     : 9,
		MediumDashDot    : 10,
		MediumDashed     : 11,
		Medium           : 12,
		Thick            : 13
	};
	var c_oAscBorderType      = {
		Hor  : 1,
		Ver  : 2,
		Diag : 3
	};
	// PageOrientation
	var c_oAscPageOrientation = {
		PagePortrait  : 0x00,
		PageLandscape : 0x01
	};
	/**
	 * lock types
	 * @enum {number}
	 */
	var c_oAscLockTypes       = {
		kLockTypeNone   : 1, // никто не залочил данный объект
		kLockTypeMine   : 2, // данный объект залочен текущим пользователем
		kLockTypeOther  : 3, // данный объект залочен другим(не текущим) пользователем
		kLockTypeOther2 : 4, // данный объект залочен другим(не текущим) пользователем (обновления уже пришли)
		kLockTypeOther3 : 5  // данный объект был залочен (обновления пришли) и снова стал залочен
	};

	var c_oAscFormatPainterState = {
		kOff      : 0,
		kOn       : 1,
		kMultiple : 2
	};

	var c_oAscSaveTypes = {
		PartStart   : 0,
		Part        : 1,
		Complete    : 2,
		CompleteAll : 3
	};

	var c_oAscColor = {
		COLOR_TYPE_NONE   : 0,
		COLOR_TYPE_SRGB   : 1,
		COLOR_TYPE_PRST   : 2,
		COLOR_TYPE_SCHEME : 3,
		COLOR_TYPE_SYS    : 4,
		COLOR_TYPE_STYLE  : 5
	};

	var c_oAscFill = {
		FILL_TYPE_NONE   : 0,
		FILL_TYPE_BLIP   : 1,
		FILL_TYPE_NOFILL : 2,
		FILL_TYPE_SOLID  : 3,
		FILL_TYPE_GRAD   : 4,
		FILL_TYPE_PATT   : 5,
		FILL_TYPE_GRP    : 6
	};

	// Chart defines
	var c_oAscChartType    = {
		line     : "Line",
		bar      : "Bar",
		hbar     : "HBar",
		area     : "Area",
		pie      : "Pie",
		scatter  : "Scatter",
		stock    : "Stock",
		doughnut : "Doughnut"
	};
	var c_oAscChartSubType = {
		normal     : "normal",
		stacked    : "stacked",
		stackedPer : "stackedPer"
	};

	var c_oAscFillGradType = {
		GRAD_LINEAR : 1,
		GRAD_PATH   : 2
	};
	var c_oAscFillBlipType = {
		STRETCH : 1,
		TILE    : 2
	};
	var c_oAscStrokeType   = {
		STROKE_NONE  : 0,
		STROKE_COLOR : 1
	};

	var c_oAscVAlign = {
		Bottom : 0, // (Text Anchor Enum ( Bottom ))
		Center : 1, // (Text Anchor Enum ( Center ))
		Dist   : 2, // (Text Anchor Enum ( Distributed ))
		Just   : 3, // (Text Anchor Enum ( Justified ))
		Top    : 4  // Top
	};

	var c_oAscVertDrawingText = {
		normal  : 1,
		vert    : 3,
		vert270 : 4
	};
	var c_oAscLineJoinType    = {
		Round : 1,
		Bevel : 2,
		Miter : 3
	};
	var c_oAscLineCapType     = {
		Flat   : 0,
		Round  : 1,
		Square : 2
	};
	var c_oAscLineBeginType   = {
		None     : 0,
		Arrow    : 1,
		Diamond  : 2,
		Oval     : 3,
		Stealth  : 4,
		Triangle : 5
	};
	var c_oAscLineBeginSize   = {
		small_small : 0,
		small_mid   : 1,
		small_large : 2,
		mid_small   : 3,
		mid_mid     : 4,
		mid_large   : 5,
		large_small : 6,
		large_mid   : 7,
		large_large : 8
	};
	var c_oAscCsvDelimiter    = {
		None      : 0,
		Tab       : 1,
		Semicolon : 2,
		Colon     : 3,
		Comma     : 4,
		Space     : 5
	};
	var c_oAscUrlType         = {
		Invalid : 0,
		Http    : 1,
		Email   : 2,
		Unsafe   : 3
	};

	var c_oAscCellTextDirection = {
		LRTB : 0x00,
		TBRL : 0x01,
		BTLR : 0x02
	};

	var c_oAscDocumentUnits = {
		Millimeter : 0,
		Inch       : 1,
		Point      : 2
	};

	var c_oAscMouseMoveDataTypes = {
		Common       : 0,
		Hyperlink    : 1,
		LockedObject : 2,
		Footnote     : 3,
		Form         : 4,
		Review       : 5,
		Eyedropper   : 6,
		Placeholder  : 7,
		EffectInfo   : 8,
	};

	// selection type
	var c_oAscSelectionType = {
		RangeCells     : 1,
		RangeCol       : 2,
		RangeRow       : 3,
		RangeMax       : 4,
		RangeImage     : 5,
		RangeChart     : 6,
		RangeShape     : 7,
		RangeSlicer    : 8,
		RangeShapeText : 9,
		RangeChartText : 10,
		RangeFrozen    : 11
	};
	var c_oAscInsertOptions = {
		InsertCellsAndShiftRight : 1,
		InsertCellsAndShiftDown  : 2,
		InsertColumns            : 3,
		InsertRows               : 4,
		InsertTableRowAbove      : 5,
		InsertTableRowBelow      : 6,
		InsertTableColLeft       : 7,
		InsertTableColRight      : 8
	};

	var c_oAscDeleteOptions = {
		DeleteCellsAndShiftLeft : 1,
		DeleteCellsAndShiftTop  : 2,
		DeleteColumns           : 3,
		DeleteRows              : 4,
		DeleteTable             : 5
	};

	// Print default options (in mm)
	var c_oAscPrintDefaultSettings = {
		// Размеры страницы при печати
		PageWidth       : 210,
		PageHeight      : 297,
		PageOrientation : c_oAscPageOrientation.PagePortrait,

		// Поля для страницы при печати
		PageLeftField   : 17.8,
		PageRightField  : 17.8,
		PageTopField    : 19.1,
		PageBottomField : 19.1,
		PageHeaderField : 7.62,
		PageFooterField : 7.62,
		MinPageLeftField	: 0.17,
		MinPageRightField	: 0.17,
		MinPageTopField		: 0.17,
		MinPageBottomField	: 0.17,

		PageGridLines : false,
		PageHeadings  : false
	};

	// Тип печати
	var c_oAscPrintType = {
		ActiveSheets: 0,	// Активные листы
		EntireWorkbook: 1,	// Всю книгу
		Selection: 2		// Выделенный фрагмент
	};

	var c_oZoomType = {
		FitToPage  : 1,
		FitToWidth : 2,
		CustomMode : 3
	};

	var c_oNotifyType = {
		Dirty: 0,
		Shift: 1,
		Move: 2,
		Delete: 3,
		RenameTableColumn: 4,
		ChangeDefName: 5,
		ChangeSheet: 6,
		DelColumnTable: 7,
		Prepare: 8,
		ChangeExternalLink: 9
	};

	var c_oNotifyParentType = {
		Change: 0,
		ChangeFormula: 1,
		EndCalculate: 2,
		GetRangeCell: 3,
		IsDefName: 4,
		Shared: 5,
		ProcessNotify: 6
	};

	var c_oDashType = {
		dash          : 0,
		dashDot       : 1,
		dot           : 2,
		lgDash        : 3,
		lgDashDot     : 4,
		lgDashDotDot  : 5,
		solid         : 6,
		sysDash       : 7,
		sysDashDot    : 8,
		sysDashDotDot : 9,
		sysDot        : 10,
		// visio types
		vsdxTransparent		: 11, // 0 visio value
		vsdxSolid			: 12,
		vsdxDash			: 13,
		vsdxDot				: 14,
		vsdxDashDot			: 15,
		vsdxDashDotDot		: 16,
		vsdxDashDashDot		: 17,
		vsdxLongDashShortDash   		: 18,
		vsdxLongDashShortDashShortDash  : 19,
		vsdxHalfDash  			: 20,
		vsdxHalfDot				: 21,
		vsdxHalfDashDot			: 22,
		vsdxHalfDashDotDot		: 23,
		vsdxHalfDashDashDot   	: 24,
		vsdxHalfLongDashShortDash   		 : 25,
		vsdxHalfLongDashShortDashShortDash   : 26,
		vsdxDoubleDash   		: 27,
		vsdxDoubleDot   		: 28,
		vsdxDoubleDashDot   	: 29,
		vsdxDoubleDashDotDot   	: 30,
		vsdxDoubleDashDashDot   : 31,
		vsdxDoubleLongDashShortDash   			: 32,
		vsdxDoubleLongDashShortDashShortDash    : 33,
		vsdxHalfHalfDash   		: 34,

	};


    /** @enum {number} */
    var c_oAscMathInterfaceType = {
        Common        : 0x00,
        Fraction      : 0x01,
        Script        : 0x02,
        Radical       : 0x03,
        LargeOperator : 0x04,
        Delimiter     : 0x05,
        Function      : 0x06,
        Accent        : 0x07,
        BorderBox     : 0x08,
        Bar           : 0x09,
        Box           : 0x0a,
        Limit         : 0x0b,
        GroupChar     : 0x0c,
        Matrix        : 0x0d,
        EqArray       : 0x0e,
        Phantom       : 0x0f
    };


	/** @enum {number} */
	var c_oAscMathInterfaceBarPos = {
		Top    : 0,
		Bottom : 1
	};

	/** @enum {number} */
	var c_oAscMathInterfaceScript = {
		None      : 0x000,  // Удаление скрипта
		Sup       : 0x001,
		Sub       : 0x002,
		SubSup    : 0x003,
		PreSubSup : 0x004
	};

	/** @enum {number} */
	var c_oAscMathInterfaceFraction = {
		Bar    : 0x001,
		Skewed : 0x002,
		Linear : 0x003,
		NoBar  : 0x004
	};

	/** @enum {number} */
	var c_oAscMathInterfaceLimitPos = {
		None   : -1,  // Удаление предела
		Top    : 0,
		Bottom : 1
	};

	/** @enum {number} */
	var c_oAscMathInterfaceMatrixMatrixAlign = {
		Top    : 0,
		Center : 1,
		Bottom : 2
	};

	/** @enum {number} */
	var c_oAscMathInterfaceMatrixColumnAlign = {
		Left   : 0,
		Center : 1,
		Right  : 2
	};

	/** @enum {number} */
	var c_oAscMathInterfaceEqArrayAlign = {
		Top    : 0,
		Center : 1,
		Bottom : 2
	};

	/** @enum {number} */
	var c_oAscMathInterfaceNaryLimitLocation = {
		UndOvr : 0,
		SubSup : 1
	};

	/** @enum {number} */
	var c_oAscMathInterfaceGroupCharPos = {
		None   : -1,  // Удаление GroupChar
		Top    : 0,
		Bottom : 1
	};

	var c_oAscTabType = {
		Bar     : 0x00,
		Center  : 0x01,
		Clear   : 0x02,
		Decimal : 0x03,
		Num     : 0x05,
		Right   : 0x07,
		Left    : 0x08
	};

	var c_oAscTabLeader = {
		Dot        : 0x00,
		Heavy      : 0x01,
		Hyphen     : 0x02,
		MiddleDot  : 0x03,
		None       : 0x04,
		Underscore : 0x05
	};

	var c_oAscEncodings    = [
		[0, 28596, "ISO-8859-6", "Arabic (ISO 8859-6)"],
		[1, 720, "DOS-720", "Arabic (OEM 720)"],
		[2, 1256, "windows-1256", "Arabic (Windows)"],

		[3, 28594, "ISO-8859-4", "Baltic (ISO 8859-4)"],
		[4, 28603, "ISO-8859-13", "Baltic (ISO 8859-13)"],
		[5, 775, "IBM775", "Baltic (OEM 775)"],
		[6, 1257, "windows-1257", "Baltic (Windows)"],

		[7, 28604, "ISO-8859-14", "Celtic (ISO 8859-14)"],

		[8, 28595, "ISO-8859-5", "Cyrillic (ISO 8859-5)"],
		[9, 20866, "KOI8-R", "Cyrillic (KOI8-R)"],
		[10, 21866, "KOI8-U", "Cyrillic (KOI8-U)"],
		[11, 10007, "x-mac-cyrillic", "Cyrillic (Mac)"],
		[12, 855, "IBM855", "Cyrillic (OEM 855)"],
		[13, 866, "cp866", "Cyrillic (OEM 866)"],
		[14, 1251, "windows-1251", "Cyrillic (Windows)"],

		[15, 852, "IBM852", "Central European (OEM 852)"],
		[16, 1250, "windows-1250", "Central European (Windows)"],

		[17, 950, "Big5", "Chinese (Big5 Traditional)"],
		[18, 936, "GB2312", "Central (GB2312 Simplified)"],

		[19, 28592, "ISO-8859-2", "Eastern European (ISO 8859-2)"],

		[20, 28597, "ISO-8859-7", "Greek (ISO 8859-7)"],
		[21, 737, "IBM737", "Greek (OEM 737)"],
		[22, 869, "IBM869", "Greek (OEM 869)"],
		[23, 1253, "windows-1253", "Greek (Windows)"],

		[24, 28598, "ISO-8859-8", "Hebrew (ISO 8859-8)"],
		[25, 862, "DOS-862", "Hebrew (OEM 862)"],
		[26, 1255, "windows-1255", "Hebrew (Windows)"],

		[27, 932, "Shift_JIS", "Japanese (Shift-JIS)"],
		[52, 950, "EUC-JP", "Japanese (EUC-JP)"],

		[28, 949, "KS_C_5601-1987", "Korean (Windows)"],
		[29, 51949, "EUC-KR", "Korean (EUC)"],

		[30, 861, "IBM861", "North European (Icelandic OEM 861)"],
		[31, 865, "IBM865", "North European (Nordic OEM 865)"],

		[32, 874, "windows-874", "Thai (TIS-620)"],

		[33, 28593, "ISO-8859-3", "Turkish (ISO 8859-3)"],
		[34, 28599, "ISO-8859-9", "Turkish (ISO 8859-9)"],
		[35, 857, "IBM857", "Turkish (OEM 857)"],
		[36, 1254, "windows-1254", "Turkish (Windows)"],

		[37, 28591, "ISO-8859-1", "Western European (ISO-8859-1)"],
		[38, 28605, "ISO-8859-15", "Western European (ISO-8859-15)"],
		[39, 850, "IBM850", "Western European (OEM 850)"],
		[40, 858, "IBM858", "Western European (OEM 858)"],
		[41, 860, "IBM860", "Western European (OEM 860 : Portuguese)"],
		[42, 863, "IBM863", "Western European (OEM 863 : French)"],
		[43, 437, "IBM437", "Western European (OEM-US)"],
		[44, 1252, "windows-1252", "Western European (Windows)"],

		[45, 1258, "windows-1258", "Vietnamese (Windows)"],

		[46, 65001, "UTF-8", "Unicode (UTF-8)"],
		[47, 65000, "UTF-7", "Unicode (UTF-7)"],

		[48, 1200, "UTF-16LE", "Unicode (UTF-16)"],
		[49, 1201, "UTF-16BE", "Unicode (UTF-16 Big Endian)"],

		[50, 12000, "UTF-32LE", "Unicode (UTF-32)"],
		[51, 12001, "UTF-32BE", "Unicode (UTF-32 Big Endian)"]
	];
	var c_oAscEncodingsMap = {
		"437"   : 43, "720" : 1, "737" : 21, "775" : 5, "850" : 39, "852" : 15, "855" : 12, "857" : 35, "858" : 40, "860" : 41, "861" : 30, "862" : 25, "863" : 42, "865" : 31, "866" : 13, "869" : 22, "874" : 32, "932" : 27, "936" : 18, "949" : 28, "950" : 17, "1200" : 48, "1201" : 49, "1250" : 16, "1251" : 14, "1252" : 44, "1253" : 23, "1254" : 36, "1255" : 26, "1256" : 2, "1257" : 6, "1258" : 45, "10007" : 11, "12000" : 50, "12001" : 51, "20866" : 9, "21866" : 10, "28591" : 37, "28592" : 19,
		"28593" : 33, "28594" : 3, "28595" : 8, "28596" : 0, "28597" : 20, "28598" : 24, "28599" : 34, "28603" : 4, "28604" : 7, "28605" : 38, "51949" : 29, "65000" : 47, "65001" : 46
	};
	var c_oAscCodePageNone = -1;
	var c_oAscCodePageUtf7 = 47;//65000
	var c_oAscCodePageUtf8 = 46;//65001
	var c_oAscCodePageUtf16 = 48;//1200
	var c_oAscCodePageUtf16BE = 49;//1201
	var c_oAscCodePageUtf32 = 50;//12000
	var c_oAscCodePageUtf32BE = 51;//12001

	// https://support.office.com/en-us/article/Excel-specifications-and-limits-16c69c74-3d6a-4aaf-ba35-e6eb276e8eaa?ui=en-US&rs=en-US&ad=US&fromAR=1
	var c_oAscMaxTooltipLength       = 256;
	var c_oAscMaxCellOrCommentLength = 32767;
	var c_oAscMaxFormulaLength       = 8192;
	var c_oAscMaxHeaderFooterLength  = 255;
	var c_oAscMaxFilterListLength    = 10000;
	var c_oAscMaxFormulaReferenceLength = 2048;
	var c_oAscMaxTableColumnTextLength  = 256;

	var changestype_None                      = 0; // Ничего не происходит с выделенным элементом (проверка идет через дополнительный параметр)
	var changestype_Paragraph_Content         = 1; // Добавление/удаление элементов в параграф
	var changestype_Paragraph_Properties      = 2; // Изменение свойств параграфа
	var changestype_Paragraph_AddText         = 3; // Добавление текста
	var changestype_Paragraph_TextProperties  = 4; // Изменение настроек текста
	var changestype_Document_Content          = 10; // Добавление/удаление элементов в Document или в DocumentContent
	var changestype_Document_Content_Add      = 11; // Добавление элемента в класс Document или в класс DocumentContent
	var changestype_Document_SectPr           = 12; // Изменения свойств данной секции (размер страницы, поля и ориентация)
	var changestype_Document_Styles           = 13; // Изменяем стили документа (добавление/удаление/модифицирование)
	var changestype_Table_Properties          = 20; // Любые изменения в таблице
	var changestype_Table_RemoveCells         = 21; // Удаление ячеек (строк или столбцов)
	var changestype_Image_Properties          = 23; // Изменения настроек картинки
	var changestype_ContentControl_Remove     = 24; // Удаление контейнера целиком
	var changestype_ContentControl_Properties = 25; // Изменение свойств контейнера
	var changestype_ContentControl_Add        = 26; // Добавление контейнера
	var changestype_HdrFtr                    = 30; // Изменения в колонтитуле (любые изменения)
	var changestype_Remove                    = 40; // Удаление, через кнопку backspace (Удаление назад)
	var changestype_Delete                    = 41; // Удаление, через кнопку delete (Удаление вперед)
	var changestype_Drawing_Props             = 51; // Изменение свойств фигуры
	var changestype_ColorScheme               = 60; // Изменение свойств фигуры
	var changestype_Text_Props                = 61; // Изменение свойств фигуры
	var changestype_RemoveSlide               = 62; // Изменение свойств фигуры
	var changestype_PresentationProps         = 63; // Изменение темы, цветовой схемы, размера слайда;
	var changestype_Theme                     = 64; // Изменение темы;
	var changestype_SlideSize                 = 65; // Изменение цветовой схемы;
	var changestype_SlideBg                   = 66; // Изменение цветовой схемы;
	var changestype_SlideTransition           = 67; // Изменение цветовой схемы;
	var changestype_MoveComment               = 68;
	var changestype_AddSp                     = 69;
	var changestype_AddComment                = 70;
	var changestype_Layout                    = 71;
	var changestype_AddShape                  = 72;
	var changestype_AddShapes                 = 73;
	var changestype_PresDefaultLang           = 74;
	var changestype_SlideHide                 = 75;
	var changestype_CorePr                    = 76;
	var changestype_Document_Settings         = 77; // Изменение общих настроек документа Document.Settings
	var changestype_Timing                    = 78;
	var changestype_ViewPr                    = 79;
	var changestype_DocumentProtection        = 80;
	var changestype_CustomPr                  = 81;

	var changestype_2_InlineObjectMove       = 1; // Передвигаем объект в заданную позцию (проверяем место, в которое пытаемся передвинуть)
	var changestype_2_HdrFtr                 = 2; // Изменения с колонтитулом
	var changestype_2_Comment                = 3; // Работает с комментариями
	var changestype_2_Element_and_Type       = 4; // Проверяем возможно ли сделать изменение заданного типа с заданным элементом(а не с текущим)
	var changestype_2_ElementsArray_and_Type = 5; // Аналогично предыдущему, только идет массив элементов
	var changestype_2_AdditionalTypes        = 6; // Дополнительные проверки типа 1
	var changestype_2_Element_and_Type_Array = 7; // Проверяем возможно ли сделать изменения заданного типа с заданными элементами (для каждого элемента свое изменение)

	var contentchanges_Add    = 1;
	var contentchanges_Remove = 2;

	var PUNCTUATION_FLAG_BASE               = 0x0001;
	var PUNCTUATION_FLAG_CANT_BE_AT_BEGIN   = 0x0010;
	var PUNCTUATION_FLAG_CANT_BE_AT_END     = 0x0020;
	var PUNCTUATION_FLAG_EAST_ASIAN         = 0x0100;
	var PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E = 0x0002;
	var PUNCTUATION_FLAG_CANT_BE_AT_END_E   = 0x0004;

	var g_aPunctuation = [];
	g_aPunctuation[0x0021] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // !
	g_aPunctuation[0x0022] = PUNCTUATION_FLAG_BASE;                                     // "
	g_aPunctuation[0x0023] = PUNCTUATION_FLAG_BASE;                                     // #
	g_aPunctuation[0x0024] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // $
	g_aPunctuation[0x0025] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // %
	g_aPunctuation[0x0026] = PUNCTUATION_FLAG_BASE;                                     // &
	g_aPunctuation[0x0027] = PUNCTUATION_FLAG_BASE;                                     // '
	g_aPunctuation[0x0028] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END | PUNCTUATION_FLAG_CANT_BE_AT_END_E;   // (
	g_aPunctuation[0x0029] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // )
	g_aPunctuation[0x002A] = PUNCTUATION_FLAG_BASE;                                     // *
	g_aPunctuation[0x002B] = PUNCTUATION_FLAG_BASE;                                     // +
	g_aPunctuation[0x002C] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ,
	g_aPunctuation[0x002D] = PUNCTUATION_FLAG_BASE;                                     // -
	g_aPunctuation[0x002E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // .
	g_aPunctuation[0x002F] = PUNCTUATION_FLAG_BASE;                                     // /
	g_aPunctuation[0x003A] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // :
	g_aPunctuation[0x003B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ;
	g_aPunctuation[0x003C] = PUNCTUATION_FLAG_BASE;                                     // <
	g_aPunctuation[0x003D] = PUNCTUATION_FLAG_BASE;                                     // =
	g_aPunctuation[0x003E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // >
	g_aPunctuation[0x003F] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ?
	g_aPunctuation[0x0040] = PUNCTUATION_FLAG_BASE;                                     // @
	g_aPunctuation[0x005B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END | PUNCTUATION_FLAG_CANT_BE_AT_END_E;   // [
	g_aPunctuation[0x005C] = PUNCTUATION_FLAG_BASE;                                     // \
	g_aPunctuation[0x005D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ]
	g_aPunctuation[0x005E] = PUNCTUATION_FLAG_BASE;                                     // ^
	g_aPunctuation[0x005F] = PUNCTUATION_FLAG_BASE;                                     // _
	g_aPunctuation[0x0060] = PUNCTUATION_FLAG_BASE;                                     // `
	g_aPunctuation[0x007B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END | PUNCTUATION_FLAG_CANT_BE_AT_END_E;   // {
	g_aPunctuation[0x007C] = PUNCTUATION_FLAG_BASE;                                     // |
	g_aPunctuation[0x007D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // }
	g_aPunctuation[0x007E] = PUNCTUATION_FLAG_BASE;                                     // ~

	g_aPunctuation[0x00A1] = PUNCTUATION_FLAG_BASE;                                     // ¡
	g_aPunctuation[0x00A2] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ¢
	g_aPunctuation[0x00A3] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // £
	g_aPunctuation[0x00A4] = PUNCTUATION_FLAG_BASE;                                     // ¤
	g_aPunctuation[0x00A5] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // ¥
	g_aPunctuation[0x00A6] = PUNCTUATION_FLAG_BASE;                                     // ¦
	g_aPunctuation[0x00A7] = PUNCTUATION_FLAG_BASE;                                     // §
	g_aPunctuation[0x00A8] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ¨
	g_aPunctuation[0x00A9] = PUNCTUATION_FLAG_BASE;                                     // ©
	g_aPunctuation[0x00AA] = PUNCTUATION_FLAG_BASE;                                     // ª
	g_aPunctuation[0x00AB] = PUNCTUATION_FLAG_BASE;                                     // «
	g_aPunctuation[0x00AC] = PUNCTUATION_FLAG_BASE;                                     // ¬
	g_aPunctuation[0x00AD] = PUNCTUATION_FLAG_BASE;                                     // ­
	g_aPunctuation[0x00AE] = PUNCTUATION_FLAG_BASE;                                     // ®
	g_aPunctuation[0x00AF] = PUNCTUATION_FLAG_BASE;                                     // ¯
	g_aPunctuation[0x00B0] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // °
	g_aPunctuation[0x00B1] = PUNCTUATION_FLAG_BASE;                                     // ±
	g_aPunctuation[0x00B4] = PUNCTUATION_FLAG_BASE;                                     // ´
	g_aPunctuation[0x00B6] = PUNCTUATION_FLAG_BASE;                                     // ¶
	g_aPunctuation[0x00B7] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ·
	g_aPunctuation[0x00B8] = PUNCTUATION_FLAG_BASE;                                     // ¸
	g_aPunctuation[0x00BA] = PUNCTUATION_FLAG_BASE;                                     // º
	g_aPunctuation[0x00BB] = PUNCTUATION_FLAG_BASE;                                     // »
	g_aPunctuation[0x00BB] = PUNCTUATION_FLAG_BASE;                                     // »
	g_aPunctuation[0x00BF] = PUNCTUATION_FLAG_BASE;                                     // ¿

	g_aPunctuation[0x2010] = PUNCTUATION_FLAG_BASE;                                     // ‐
	g_aPunctuation[0x2011] = PUNCTUATION_FLAG_BASE;                                     // ‑
	g_aPunctuation[0x2012] = PUNCTUATION_FLAG_BASE;                                     // ‒
	g_aPunctuation[0x2013] = PUNCTUATION_FLAG_BASE;                                     // –
	g_aPunctuation[0x2014] = PUNCTUATION_FLAG_BASE;                                     // —
	g_aPunctuation[0x2015] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ―
	g_aPunctuation[0x2016] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ‖
	g_aPunctuation[0x2017] = PUNCTUATION_FLAG_BASE;                                     // ‗
	g_aPunctuation[0x2018] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // ‘
	g_aPunctuation[0x2019] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ’
	g_aPunctuation[0x201A] = PUNCTUATION_FLAG_BASE;                                     // ‚
	g_aPunctuation[0x201B] = PUNCTUATION_FLAG_BASE;                                     // ‛
	g_aPunctuation[0x201C] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // “
	g_aPunctuation[0x201D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ”
	g_aPunctuation[0x201E] = PUNCTUATION_FLAG_BASE;                                     // „
	g_aPunctuation[0x201F] = PUNCTUATION_FLAG_BASE;                                     // ‟
	g_aPunctuation[0x2020] = PUNCTUATION_FLAG_BASE;                                     // †
	g_aPunctuation[0x2021] = PUNCTUATION_FLAG_BASE;                                     // ‡
	g_aPunctuation[0x2022] = PUNCTUATION_FLAG_BASE;                                     // •
	g_aPunctuation[0x2023] = PUNCTUATION_FLAG_BASE;                                     // ‣
	g_aPunctuation[0x2024] = PUNCTUATION_FLAG_BASE;                                     // ․
	g_aPunctuation[0x2025] = PUNCTUATION_FLAG_BASE;                                     // ‥
	g_aPunctuation[0x2026] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // …
	g_aPunctuation[0x2027] = PUNCTUATION_FLAG_BASE;                                     // ‧
	g_aPunctuation[0x2030] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ‰
	g_aPunctuation[0x2031] = PUNCTUATION_FLAG_BASE;                                     // ‱
	g_aPunctuation[0x2032] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ′
	g_aPunctuation[0x2033] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ″
	g_aPunctuation[0x2034] = PUNCTUATION_FLAG_BASE;                                     // ‴
	g_aPunctuation[0x2035] = PUNCTUATION_FLAG_BASE;                                     // ‵
	g_aPunctuation[0x2036] = PUNCTUATION_FLAG_BASE;                                     // ‶
	g_aPunctuation[0x2037] = PUNCTUATION_FLAG_BASE;                                     // ‷
	g_aPunctuation[0x2038] = PUNCTUATION_FLAG_BASE;                                     // ‸
	g_aPunctuation[0x2039] = PUNCTUATION_FLAG_BASE;                                     // ‹
	g_aPunctuation[0x203A] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ›
	g_aPunctuation[0x203B] = PUNCTUATION_FLAG_BASE;                                     // ※
	g_aPunctuation[0x203C] = PUNCTUATION_FLAG_BASE;                                     // ‼
	g_aPunctuation[0x203D] = PUNCTUATION_FLAG_BASE;                                     // ‽
	g_aPunctuation[0x203E] = PUNCTUATION_FLAG_BASE;                                     // ‾
	g_aPunctuation[0x203F] = PUNCTUATION_FLAG_BASE;                                     // ‿
	g_aPunctuation[0x2040] = PUNCTUATION_FLAG_BASE;                                     // ⁀
	g_aPunctuation[0x2041] = PUNCTUATION_FLAG_BASE;                                     // ⁁
	g_aPunctuation[0x2042] = PUNCTUATION_FLAG_BASE;                                     // ⁂
	g_aPunctuation[0x2043] = PUNCTUATION_FLAG_BASE;                                     // ⁃
	g_aPunctuation[0x2044] = PUNCTUATION_FLAG_BASE;                                     // ⁄
	g_aPunctuation[0x2045] = PUNCTUATION_FLAG_BASE;                                     // ⁅
	g_aPunctuation[0x2046] = PUNCTUATION_FLAG_BASE;                                     // ⁆
	g_aPunctuation[0x2047] = PUNCTUATION_FLAG_BASE;                                     // ⁇
	g_aPunctuation[0x2048] = PUNCTUATION_FLAG_BASE;                                     // ⁈
	g_aPunctuation[0x2049] = PUNCTUATION_FLAG_BASE;                                     // ⁉
	g_aPunctuation[0x204A] = PUNCTUATION_FLAG_BASE;                                     // ⁊
	g_aPunctuation[0x204B] = PUNCTUATION_FLAG_BASE;                                     // ⁋
	g_aPunctuation[0x204C] = PUNCTUATION_FLAG_BASE;                                     // ⁌
	g_aPunctuation[0x204D] = PUNCTUATION_FLAG_BASE;                                     // ⁍
	g_aPunctuation[0x204E] = PUNCTUATION_FLAG_BASE;                                     // ⁎
	g_aPunctuation[0x204F] = PUNCTUATION_FLAG_BASE;                                     // ⁏
	g_aPunctuation[0x2050] = PUNCTUATION_FLAG_BASE;                                     // ⁐
	g_aPunctuation[0x2051] = PUNCTUATION_FLAG_BASE;                                     // ⁑
	g_aPunctuation[0x2052] = PUNCTUATION_FLAG_BASE;                                     // ⁒
	g_aPunctuation[0x2053] = PUNCTUATION_FLAG_BASE;                                     // ⁓
	g_aPunctuation[0x2054] = PUNCTUATION_FLAG_BASE;                                     // ⁔
	g_aPunctuation[0x2055] = PUNCTUATION_FLAG_BASE;                                     // ⁕
	g_aPunctuation[0x2056] = PUNCTUATION_FLAG_BASE;                                     // ⁖
	g_aPunctuation[0x2057] = PUNCTUATION_FLAG_BASE;                                     // ⁗
	g_aPunctuation[0x2058] = PUNCTUATION_FLAG_BASE;                                     // ⁘
	g_aPunctuation[0x2059] = PUNCTUATION_FLAG_BASE;                                     // ⁙
	g_aPunctuation[0x205A] = PUNCTUATION_FLAG_BASE;                                     // ⁚
	g_aPunctuation[0x205B] = PUNCTUATION_FLAG_BASE;                                     // ⁛
	g_aPunctuation[0x205C] = PUNCTUATION_FLAG_BASE;                                     // ⁜
	g_aPunctuation[0x205D] = PUNCTUATION_FLAG_BASE;                                     // ⁝
	g_aPunctuation[0x205E] = PUNCTUATION_FLAG_BASE;                                     // ⁞
	g_aPunctuation[0x2420] = PUNCTUATION_FLAG_BASE;                                     // ␠
	g_aPunctuation[0x2421] = PUNCTUATION_FLAG_BASE;                                     // ␡
	g_aPunctuation[0x2422] = PUNCTUATION_FLAG_BASE;                                     // ␢
	g_aPunctuation[0x2423] = PUNCTUATION_FLAG_BASE;                                     // ␣
	g_aPunctuation[0x2424] = PUNCTUATION_FLAG_BASE;                                     // ␤
	g_aPunctuation[0x2425] = PUNCTUATION_FLAG_BASE;                                     // ␥
	g_aPunctuation[0x2426] = PUNCTUATION_FLAG_BASE;                                     // ␦

	// Не смотря на то что следующий набор символов идет в блоке CJK Symbols and Punctuation
	// Word не считает их как EastAsian script (w:lang->w:eastAsian)

	g_aPunctuation[0x3001] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 、
	g_aPunctuation[0x3002] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 。
	g_aPunctuation[0x3003] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 〃
	g_aPunctuation[0x3004] = PUNCTUATION_FLAG_BASE;                                     // 〄
	g_aPunctuation[0x3005] = PUNCTUATION_FLAG_BASE;                                     // 々
	g_aPunctuation[0x3006] = PUNCTUATION_FLAG_BASE;                                     // 〆
	g_aPunctuation[0x3007] = PUNCTUATION_FLAG_BASE;                                     // 〇
	g_aPunctuation[0x3008] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // 〈
	g_aPunctuation[0x3009] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 〉
	g_aPunctuation[0x300A] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // 《
	g_aPunctuation[0x300B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 》
	g_aPunctuation[0x300C] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // 「
	g_aPunctuation[0x300D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 」
	g_aPunctuation[0x300E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // 『
	g_aPunctuation[0x300F] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 』
	g_aPunctuation[0x3010] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // 【
	g_aPunctuation[0x3011] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 】
	g_aPunctuation[0x3012] = PUNCTUATION_FLAG_BASE;                                     // 〒
	g_aPunctuation[0x3013] = PUNCTUATION_FLAG_BASE;                                     // 〓
	g_aPunctuation[0x3014] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   //〔
	g_aPunctuation[0x3015] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 〕
	g_aPunctuation[0x3016] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   //〖
	g_aPunctuation[0x3017] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 〗
	g_aPunctuation[0x3018] = PUNCTUATION_FLAG_BASE;                                     // 〘
	g_aPunctuation[0x3019] = PUNCTUATION_FLAG_BASE;                                     // 〙
	g_aPunctuation[0x301A] = PUNCTUATION_FLAG_BASE;                                     // 〚
	g_aPunctuation[0x301B] = PUNCTUATION_FLAG_BASE;                                     // 〛
	g_aPunctuation[0x301C] = PUNCTUATION_FLAG_BASE;                                     // 〜
	g_aPunctuation[0x301D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_END;   // 〝
	g_aPunctuation[0x301E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // 〞
	g_aPunctuation[0x301F] = PUNCTUATION_FLAG_BASE;                                     // 〟

	g_aPunctuation[0xFF01] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ！
	g_aPunctuation[0xFF02] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ＂
	g_aPunctuation[0xFF03] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＃
	g_aPunctuation[0xFF04] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_END;   // ＄
	g_aPunctuation[0xFF05] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ％
	g_aPunctuation[0xFF06] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＆
	g_aPunctuation[0xFF07] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ＇
	g_aPunctuation[0xFF08] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_END | PUNCTUATION_FLAG_CANT_BE_AT_END_E;   // （
	g_aPunctuation[0xFF09] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // )
	g_aPunctuation[0xFF0A] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＊
	g_aPunctuation[0xFF0B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＋
	g_aPunctuation[0xFF0C] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ，
	g_aPunctuation[0xFF0D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // －
	g_aPunctuation[0xFF0E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ．
	g_aPunctuation[0xFF0F] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ／
	g_aPunctuation[0xFF1A] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ：
	g_aPunctuation[0xFF1B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ；
	g_aPunctuation[0xFF1C] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＜
	g_aPunctuation[0xFF1D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＝
	g_aPunctuation[0xFF1E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＞
	g_aPunctuation[0xFF1F] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ？
	g_aPunctuation[0xFF20] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＠
	g_aPunctuation[0xFF3B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_END | PUNCTUATION_FLAG_CANT_BE_AT_END_E;   // ［
	g_aPunctuation[0xFF3C] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＼
	g_aPunctuation[0xFF3D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ］
	g_aPunctuation[0xFF3E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＾
	g_aPunctuation[0xFF3F] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ＿
	g_aPunctuation[0xFF40] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ｀
	g_aPunctuation[0xFF5B] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_END | PUNCTUATION_FLAG_CANT_BE_AT_END_E;   // ｛
	g_aPunctuation[0xFF5C] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ｜
	g_aPunctuation[0xFF5D] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E; // ｝
	g_aPunctuation[0xFF5E] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ～
	g_aPunctuation[0xFF5F] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ｟
	g_aPunctuation[0xFF60] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ｠
	g_aPunctuation[0xFF61] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ｡
	g_aPunctuation[0xFF62] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ｢
	g_aPunctuation[0xFF63] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ｣
	g_aPunctuation[0xFF64] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ､
	g_aPunctuation[0xFF65] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ･
	g_aPunctuation[0xFFE0] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_BEGIN; // ￠
	g_aPunctuation[0xFFE1] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_END;   // ￡
	g_aPunctuation[0xFFE2] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￢
	g_aPunctuation[0xFFE3] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￣
	g_aPunctuation[0xFFE4] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￤
	g_aPunctuation[0xFFE5] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN | PUNCTUATION_FLAG_CANT_BE_AT_END;   // ￥
	g_aPunctuation[0xFFE6] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￦
	g_aPunctuation[0xFFE8] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￨
	g_aPunctuation[0xFFE9] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￩
	g_aPunctuation[0xFFEA] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￪
	g_aPunctuation[0xFFEB] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￫
	g_aPunctuation[0xFFEC] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￬
	g_aPunctuation[0xFFED] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￭
	g_aPunctuation[0xFFEE] = PUNCTUATION_FLAG_BASE | PUNCTUATION_FLAG_EAST_ASIAN;                                     // ￮


	var offlineMode = '_offline_';
	var chartMode = '_chart_';
	var dataMode = '_data_';

	var c_oSpecialPasteProps = {
		paste: 0,
		pasteOnlyFormula: 1,
		formulaNumberFormat: 2,
		formulaAllFormatting: 3,
		formulaWithoutBorders: 4,
		formulaColumnWidth: 5,
		mergeConditionalFormating: 6,
		pasteOnlyValues: 7,
		valueNumberFormat: 8,
		valueAllFormating: 9,
		pasteOnlyFormating: 10,
		transpose: 11,
		link: 12,
		picture: 13,
		linkedPicture: 14,

		sourceformatting: 15,
		destinationFormatting: 16,
		
		mergeFormatting: 17,

		uniteList: 18,
		doNotUniteList: 19,

		insertAsNestedTable: 20,
		uniteIntoTable: 21,
		insertAsNewRows: 22,
		keepTextOnly: 23,
		overwriteCells : 24,

		useTextImport: 25,

		comments: 26,
		columnWidth: 27
	};

	var c_oSpecialPasteOperation = {
		none: 0,
		add: 1,
		subtract: 2,
		multiply: 3,
		divide: 4
	};

	var c_oAbstractNumMultiLvlTypes = {
		HybridMultiLevel: 0,
		MultiLevel: 1,
		SingleLevel: 2
	};


	var c_oAscSmartArtTypes = {
		AccentedPicture: 0, // Акцентируемый рисунок
		Balance: 1, // Баланс
		TitledPictureBlocks: 2, // Блоки рисунков с названиями
		PictureAccentBlocks: 3, // Блоки со смещенными рисунками
		BlockCycle: 4, // Блочный цикл
		StackedVenn: 5, // Венна в столбик
		VerticalEquation: 6, // Вертикальное уравнение
		VerticalBlockList: 7, // Вертикальный блочный список
		VerticalBendingProcess: 8, // Вертикальный ломаный процесс
		VerticalBulletList: 9, // Вертикальный маркированный список
		VerticalCurvedList: 10, // Вертикальный нелинейный список
		VerticalProcess: 11, // Вертикальный процесс
		VerticalBoxList: 12, // Вертикальный список
		VerticalPictureList: 13, // Вертикальный список рисунков
		VerticalCircleList: 14, // Вертикальный список с кругами
		VerticalPictureAccentList: 15, // Вертикальный список со смещенными рисунками
		VerticalArrowList: 16, // Вертикальный список со стрелкой
		VerticalChevronList: 17, // Вертикальный уголковый список
		VerticalAccentList: 18, // Вертикальный уголковый список2
		NestedTarget: 19, // Вложенная целевая
		Funnel: 20, // Воронка
		UpwardArrow: 21, // Восходящая стрелка
		IncreasingArrowsProcess: 22, // Восходящая стрелка процесса
		StepUpProcess: 23, // Восходящий процесс
		CircularPictureCallout: 24, // Выноска с круглыми рисунками
		HorizontalHierarchy: 25, // Горизонтальная иерархия
		HorizontalLabeledHierarchy: 26, // Горизонтальная иерархия с подписями
		HorizontalMultiLevelHierarchy: 27, // Горизонтальная многоуровневая иерархия
		HorizontalOrganizationChart: 28, // Горизонтальная организационная диаграмма
		HorizontalBulletList: 29, // Горизонтальный маркированный список
		HorizontalPictureList: 30, // Горизонтальный список рисунков
		ClosedChevronProcess: 31, // Закрытый уголковый процесс
		HierarchyList: 32, // Иерархический список
		Hierarchy: 33, // Иерархия
		CirclePictureHierarchy: 34, // Иерархия с круглыми рисунками
		LabeledHierarchy: 35, // Иерархия с подписями
		InvertedPyramid: 36, // Инвертированная пирамида
		HexagonCluster: 37, // Кластер шестиугольников
		CircleRelationship: 38, // Круг связей
		CircleAccentTimeline: 39, // Круглая временная шкала
		CircularBendingProcess: 40, // Круглый ломаный процесс
		ArrowRibbon: 41, // Лента со стрелками
		LinearVenn: 42, // Линейная Венна
		PictureLineup: 43, // Линия рисунков
		TitlePictureLineup: 44, // Линия рисунков с названиями
		BendingPictureCaptionList: 45, // Ломаный список рисунков с подписями
		BendingPictureAccentList: 46, // Ломаный список со смещенными рисунками
		TitledMatrix: 47, // Матрица с заголовками
		IncreasingCircleProcess: 48, // Нарастающий процесс с кругами
		BendingPictureBlocks: 49, // Нелинейные рисунки с блоками
		BendingPictureCaption: 50, // Нелинейные рисунки с подписями
		BendingPictureSemiTransparentText: 51, // Нелинейные рисунки с полупрозрачным текстом
		NonDirectionalCycle: 52, // Ненаправленный цикл
		ContinuousBlockProcess: 53, // Непрерывный блочный процесс
		ContinuousPictureList: 54, // Непрерывный список с рисунками
		ContinuousCycle: 55, // Непрерывный цикл
		DescendingBlockList: 56, // Нисходящий блочный список
		StepDownProcess: 57, // Нисходящий процесс
		ReverseList: 58, // Обратный список
		OrganizationChart: 59, // Организационная диаграмма
		NameAndTitleOrganizationChart: 60, // Организационная диаграмма с именами и должностями
		AlternatingFlow: 61, // Переменный поток
		PyramidList: 62, // Пирамидальный список
		PlusAndMinus: 63, // Плюс и минус
		RepeatingBendingProcess: 64, // Повторяющийся ломаный процесс
		CaptionedPictures: 65, // Подписанные рисунки
		DetailedProcess: 66, // Подробный процесс
		PictureStrips: 67, // Полосы рисунков
		HalfCircleOrganizationChart: 68, // Полукруглая организационная диаграмма
		PhasedProcess: 69, // Поэтапный процесс
		BasicVenn: 70, // Простая Венна
		BasicTimeline: 71, // Простая временная шкала
		BasicPie: 72, // Простая круговая
		BasicMatrix: 73, // Простая матрица
		BasicPyramid: 74, // Простая пирамида
		BasicRadial: 75, // Простая радиальная
		BasicTarget: 76, // Простая целевая
		BasicBlockList: 77, // Простой блочный список
		BasicBendingProcess: 78, // Простой ломаный процесс
		BasicProcess: 79, // Простой процесс
		BasicChevronProcess: 80, // Простой уголковый процесс
		BasicCycle: 81, // Простой цикл
		OpposingIdeas: 82, // Противоположные идеи
		OpposingArrows: 83, // Противостоящие стрелки
		RandomToResultProcess: 84, // Процесс от случайности к результату
		SubStepProcess: 85, // Процесс с вложенными шагами
		PieProcess: 86, // Процесс с круговой диаграммой
		AccentProcess: 87, // Процесс со смещением
		AscendingPictureAccentProcess: 88, // Процесс со смещенными по возрастанию рисунками
		PictureAccentProcess: 89, // Процесс со смещенными рисунками
		RadialVenn: 90, // Радиальная Венна
		RadialCycle: 91, // Радиальная циклическая
		RadialCluster: 92, // Радиальный кластер
		RadialList: 93, // Радиальный список
		MultiDirectionalCycle: 94, // Разнонаправленный цикл
		DivergingRadial: 95, // Расходящаяся радиальная
		DivergingArrows: 96, // Расходящиеся стрелки
		FramedTextPicture: 97, // Рисунок с текстом в рамке
		GroupedList: 98, // Сгруппированный список
		SegmentedPyramid: 99, // Сегментированная пирамида
		SegmentedProcess: 100, // Сегментированный процесс
		SegmentedCycle: 101, // Сегментированный цикл
		PictureGrid: 102, // Сетка рисунков
		GridMatrix: 103, // Сетчатая матрица
		SpiralPicture: 104, // Спираль рисунков
		StackedList: 105, // Список в столбик
		PictureCaptionList: 106, // Список названий рисунков
		ProcessList: 107, // Список процессов
		BubblePictureList: 108, // Список рисунков с выносками
		SquareAccentList: 109, // Список с квадратиками
		LinedList: 110, // Список с линиями
		PictureAccentList: 111, // Список со смещенными рисунками
		TitledPictureAccentList: 112, // Список со смещенными рисунками и заголовком
		SnapshotPictureList: 113, // Список со снимками
		ContinuousArrowProcess: 114, // Стрелка непрерывного процесса
		CircleArrowProcess: 115, // Стрелка процесса с кругами
		ProcessArrows: 116, // Стрелки процесса
		StaggeredProcess: 117, // Ступенчатый процесс
		ConvergingRadial: 118, // Сходящаяся радиальная
		ConvergingArrows: 119, // Сходящиеся стрелки
		TableHierarchy: 120, // Табличная иерархия
		TableList: 121, // Табличный список
		TextCycle: 122, // Текстовый цикл
		TrapezoidList: 123, // Трапецевидный список
		DescendingProcess: 124, // Убывающий процесс
		ChevronList: 125, // Уголковый список
		Equation: 126, // Уравнение
		CounterbalanceArrows: 127, // Уравновешивающие стрелки
		TargetList: 128, // Целевой список
		CycleMatrix: 129, // Циклическая матрица
		AlternatingPictureBlocks: 130, // Чередующиеся блоки рисунков
		AlternatingPictureCircles: 131, // Чередующиеся круги рисунков
		AlternatingHexagonList: 132, // Чередующиеся шестиугольники
		Gear: 133, // Шестеренки

		// Office.com
		ArchitectureLayout: 134,
		ChevronAccentProcess: 135,
		CircleProcess: 136,
		ConvergingText: 137,
		HexagonRadial: 138,
		InterconnectedBlockProcess: 139,
		InterconnectedRings: 140,
		PictureFrame: 141,
		PictureOrganizationChart: 142,
		RadialPictureList: 143,
		TabList: 144,
		TabbedArc: 145,
		ThemePictureAccent: 146,
		ThemePictureAlternatingAccent: 147,
		ThemePictureGrid: 148,
		VaryingWidthList: 149,
		VerticalBracketList: 150
	};

	var c_oAscSmartArtListTypes = [
		c_oAscSmartArtTypes.BasicBlockList,
		c_oAscSmartArtTypes.AlternatingHexagonList,
		c_oAscSmartArtTypes.PictureCaptionList,
		c_oAscSmartArtTypes.LinedList,
		c_oAscSmartArtTypes.VerticalBulletList,
		c_oAscSmartArtTypes.VerticalBoxList,
		c_oAscSmartArtTypes.VerticalBracketList,
		c_oAscSmartArtTypes.VaryingWidthList,
		c_oAscSmartArtTypes.TabList,
		c_oAscSmartArtTypes.HorizontalBulletList,
		c_oAscSmartArtTypes.SquareAccentList,
		c_oAscSmartArtTypes.PictureAccentList,
		c_oAscSmartArtTypes.BendingPictureAccentList,
		c_oAscSmartArtTypes.StackedList,
		c_oAscSmartArtTypes.IncreasingCircleProcess,
		c_oAscSmartArtTypes.PieProcess,
		c_oAscSmartArtTypes.DetailedProcess,
		c_oAscSmartArtTypes.GroupedList,
		c_oAscSmartArtTypes.HorizontalPictureList,
		c_oAscSmartArtTypes.ContinuousPictureList,
		c_oAscSmartArtTypes.PictureStrips,
		c_oAscSmartArtTypes.VerticalPictureList,
		c_oAscSmartArtTypes.AlternatingPictureBlocks,
		c_oAscSmartArtTypes.VerticalPictureAccentList,
		c_oAscSmartArtTypes.TitledPictureAccentList,
		c_oAscSmartArtTypes.VerticalBlockList,
		c_oAscSmartArtTypes.VerticalChevronList,
		c_oAscSmartArtTypes.VerticalAccentList,
		c_oAscSmartArtTypes.VerticalArrowList,
		c_oAscSmartArtTypes.TrapezoidList,
		c_oAscSmartArtTypes.DescendingBlockList,
		c_oAscSmartArtTypes.TableList,
		c_oAscSmartArtTypes.SegmentedProcess,
		c_oAscSmartArtTypes.VerticalCurvedList,
		c_oAscSmartArtTypes.PyramidList,
		c_oAscSmartArtTypes.TargetList,
		c_oAscSmartArtTypes.HierarchyList,
		c_oAscSmartArtTypes.VerticalCircleList,
		c_oAscSmartArtTypes.TableHierarchy,
		c_oAscSmartArtTypes.ArchitectureLayout
	];

	var c_oAscSmartArtProcessTypes = [
		c_oAscSmartArtTypes.BasicProcess,
		c_oAscSmartArtTypes.StepUpProcess,
		c_oAscSmartArtTypes.StepDownProcess,
		c_oAscSmartArtTypes.AccentProcess,
		c_oAscSmartArtTypes.PictureAccentProcess,
		c_oAscSmartArtTypes.AlternatingFlow,
		c_oAscSmartArtTypes.IncreasingCircleProcess,
		c_oAscSmartArtTypes.PieProcess,
		c_oAscSmartArtTypes.ContinuousBlockProcess,
		c_oAscSmartArtTypes.IncreasingArrowsProcess,
		c_oAscSmartArtTypes.InterconnectedBlockProcess,
		c_oAscSmartArtTypes.ContinuousArrowProcess,
		c_oAscSmartArtTypes.ConvergingText,
		c_oAscSmartArtTypes.ProcessArrows,
		c_oAscSmartArtTypes.CircleAccentTimeline,
		c_oAscSmartArtTypes.BasicTimeline,
		c_oAscSmartArtTypes.CircleProcess,
		c_oAscSmartArtTypes.BasicChevronProcess,
		c_oAscSmartArtTypes.ChevronAccentProcess,
		c_oAscSmartArtTypes.ClosedChevronProcess,
		c_oAscSmartArtTypes.ChevronList,
		c_oAscSmartArtTypes.VerticalChevronList,
		c_oAscSmartArtTypes.SubStepProcess,
		c_oAscSmartArtTypes.PhasedProcess,
		c_oAscSmartArtTypes.RandomToResultProcess,
		c_oAscSmartArtTypes.VerticalProcess,
		c_oAscSmartArtTypes.StaggeredProcess,
		c_oAscSmartArtTypes.ProcessList,
		c_oAscSmartArtTypes.SegmentedProcess,
		c_oAscSmartArtTypes.CircleArrowProcess,
		c_oAscSmartArtTypes.BasicBendingProcess,
		c_oAscSmartArtTypes.RepeatingBendingProcess,
		c_oAscSmartArtTypes.VerticalBendingProcess,
		c_oAscSmartArtTypes.ContinuousPictureList,
		c_oAscSmartArtTypes.DetailedProcess,
		c_oAscSmartArtTypes.VerticalArrowList,
		c_oAscSmartArtTypes.AscendingPictureAccentProcess,
		c_oAscSmartArtTypes.UpwardArrow,
		c_oAscSmartArtTypes.DescendingProcess,
		c_oAscSmartArtTypes.CircularBendingProcess,
		c_oAscSmartArtTypes.Equation,
		c_oAscSmartArtTypes.VerticalEquation,
		c_oAscSmartArtTypes.Funnel,
		c_oAscSmartArtTypes.Gear,
		c_oAscSmartArtTypes.ArrowRibbon,
		c_oAscSmartArtTypes.OpposingArrows,
		c_oAscSmartArtTypes.ConvergingArrows,
		c_oAscSmartArtTypes.DivergingArrows
	];
	var c_oAscSmartArtCycleTypes = [
		c_oAscSmartArtTypes.BasicCycle,
		c_oAscSmartArtTypes.TextCycle,
		c_oAscSmartArtTypes.BlockCycle,
		c_oAscSmartArtTypes.NonDirectionalCycle,
		c_oAscSmartArtTypes.ContinuousCycle,
		c_oAscSmartArtTypes.MultiDirectionalCycle,
		c_oAscSmartArtTypes.SegmentedCycle,
		c_oAscSmartArtTypes.BasicPie,
		c_oAscSmartArtTypes.HexagonRadial,
		c_oAscSmartArtTypes.RadialCycle,
		c_oAscSmartArtTypes.BasicRadial,
		c_oAscSmartArtTypes.DivergingRadial,
		c_oAscSmartArtTypes.RadialVenn,
		c_oAscSmartArtTypes.CycleMatrix,
		c_oAscSmartArtTypes.Gear,
		c_oAscSmartArtTypes.RadialCluster,
		c_oAscSmartArtTypes.CircleArrowProcess

];
	var c_oAscSmartArtHierarchyTypes = [
		c_oAscSmartArtTypes.OrganizationChart,
		c_oAscSmartArtTypes.PictureOrganizationChart,
		c_oAscSmartArtTypes.NameAndTitleOrganizationChart,
		c_oAscSmartArtTypes.HalfCircleOrganizationChart,
		c_oAscSmartArtTypes.CirclePictureHierarchy,
		c_oAscSmartArtTypes.Hierarchy,
		c_oAscSmartArtTypes.LabeledHierarchy,
		c_oAscSmartArtTypes.TableHierarchy,
		c_oAscSmartArtTypes.HorizontalOrganizationChart,
		c_oAscSmartArtTypes.ArchitectureLayout,
		c_oAscSmartArtTypes.HorizontalMultiLevelHierarchy,
		c_oAscSmartArtTypes.HorizontalHierarchy,
		c_oAscSmartArtTypes.HorizontalLabeledHierarchy,
		c_oAscSmartArtTypes.HierarchyList,
		c_oAscSmartArtTypes.LinedList

	];
	var c_oAscSmartArtRelationshipTypes = [
		c_oAscSmartArtTypes.Balance,
		c_oAscSmartArtTypes.CircleRelationship,
		c_oAscSmartArtTypes.Funnel,
		c_oAscSmartArtTypes.Gear,
		c_oAscSmartArtTypes.HexagonCluster,
		c_oAscSmartArtTypes.OpposingIdeas,
		c_oAscSmartArtTypes.PlusAndMinus,
		c_oAscSmartArtTypes.ReverseList,
		c_oAscSmartArtTypes.ArrowRibbon,
		c_oAscSmartArtTypes.CounterbalanceArrows,
		c_oAscSmartArtTypes.ConvergingArrows,
		c_oAscSmartArtTypes.DivergingArrows,
		c_oAscSmartArtTypes.OpposingArrows,
		c_oAscSmartArtTypes.SegmentedPyramid,
		c_oAscSmartArtTypes.TableHierarchy,
		c_oAscSmartArtTypes.ArchitectureLayout,
		c_oAscSmartArtTypes.TargetList,
		c_oAscSmartArtTypes.NestedTarget,
		c_oAscSmartArtTypes.GroupedList,
		c_oAscSmartArtTypes.ContinuousPictureList,
		c_oAscSmartArtTypes.HierarchyList,
		c_oAscSmartArtTypes.PictureAccentList,
		c_oAscSmartArtTypes.Equation,
		c_oAscSmartArtTypes.VerticalEquation,
		c_oAscSmartArtTypes.ConvergingRadial,
		c_oAscSmartArtTypes.RadialCluster,
		c_oAscSmartArtTypes.RadialList,
		c_oAscSmartArtTypes.TabbedArc,
		c_oAscSmartArtTypes.RadialCycle,
		c_oAscSmartArtTypes.BasicRadial,
		c_oAscSmartArtTypes.DivergingRadial,
		c_oAscSmartArtTypes.NonDirectionalCycle,
		c_oAscSmartArtTypes.BasicTarget,
		c_oAscSmartArtTypes.CycleMatrix,
		c_oAscSmartArtTypes.BasicPie,
		c_oAscSmartArtTypes.BasicVenn,
		c_oAscSmartArtTypes.LinearVenn,
		c_oAscSmartArtTypes.StackedVenn,
		c_oAscSmartArtTypes.RadialVenn,
		c_oAscSmartArtTypes.InterconnectedRings
	];
	var c_oAscSmartArtMatrixTypes = [
		c_oAscSmartArtTypes.BasicMatrix,
		c_oAscSmartArtTypes.TitledMatrix,
		c_oAscSmartArtTypes.GridMatrix,
		c_oAscSmartArtTypes.CycleMatrix
	];
	var c_oAscSmartArtPyramidTypes = [
		c_oAscSmartArtTypes.BasicPyramid,
		c_oAscSmartArtTypes.InvertedPyramid,
		c_oAscSmartArtTypes.PyramidList,
		c_oAscSmartArtTypes.SegmentedPyramid
	];
	var c_oAscSmartArtPictureTypes = [
		c_oAscSmartArtTypes.AccentedPicture,
		c_oAscSmartArtTypes.CircularPictureCallout,
		c_oAscSmartArtTypes.PictureCaptionList,
		c_oAscSmartArtTypes.RadialPictureList,
		c_oAscSmartArtTypes.SnapshotPictureList,
		c_oAscSmartArtTypes.SpiralPicture,
		c_oAscSmartArtTypes.CaptionedPictures,
		c_oAscSmartArtTypes.BendingPictureCaption,
		c_oAscSmartArtTypes.PictureFrame,
		c_oAscSmartArtTypes.BendingPictureSemiTransparentText,
		c_oAscSmartArtTypes.BendingPictureBlocks,
		c_oAscSmartArtTypes.BendingPictureCaptionList,
		c_oAscSmartArtTypes.TitledPictureBlocks,
		c_oAscSmartArtTypes.PictureGrid,
		c_oAscSmartArtTypes.PictureAccentBlocks,
		c_oAscSmartArtTypes.PictureStrips,
		c_oAscSmartArtTypes.ThemePictureAccent,
		c_oAscSmartArtTypes.ThemePictureGrid,
		c_oAscSmartArtTypes.ThemePictureAlternatingAccent,
		c_oAscSmartArtTypes.TitledPictureAccentList,
		c_oAscSmartArtTypes.AlternatingPictureBlocks,
		c_oAscSmartArtTypes.AscendingPictureAccentProcess,
		c_oAscSmartArtTypes.AlternatingPictureCircles,
		c_oAscSmartArtTypes.TitlePictureLineup,
		c_oAscSmartArtTypes.PictureLineup,
		c_oAscSmartArtTypes.FramedTextPicture,
		c_oAscSmartArtTypes.HexagonCluster,
		c_oAscSmartArtTypes.BubblePictureList,
		c_oAscSmartArtTypes.CirclePictureHierarchy,
		c_oAscSmartArtTypes.HorizontalPictureList,
		c_oAscSmartArtTypes.ContinuousPictureList,
		c_oAscSmartArtTypes.VerticalPictureList,
		c_oAscSmartArtTypes.VerticalPictureAccentList,
		c_oAscSmartArtTypes.BendingPictureAccentList,
		c_oAscSmartArtTypes.PictureAccentList,
		c_oAscSmartArtTypes.PictureAccentProcess
	];
	var c_oAscSmartArtOfficeComTypes = [
		c_oAscSmartArtTypes.PictureOrganizationChart,
		c_oAscSmartArtTypes.ChevronAccentProcess,
		c_oAscSmartArtTypes.RadialPictureList,
		c_oAscSmartArtTypes.VerticalBracketList,
		c_oAscSmartArtTypes.InterconnectedBlockProcess,
		c_oAscSmartArtTypes.TabbedArc,
		c_oAscSmartArtTypes.ThemePictureAccent,
		c_oAscSmartArtTypes.VaryingWidthList,
		c_oAscSmartArtTypes.ConvergingText,
		c_oAscSmartArtTypes.InterconnectedRings,
		c_oAscSmartArtTypes.ArchitectureLayout,
		c_oAscSmartArtTypes.ThemePictureAlternatingAccent,
		c_oAscSmartArtTypes.ThemePictureGrid,
		c_oAscSmartArtTypes.CircleProcess,
		c_oAscSmartArtTypes.HexagonRadial,
		c_oAscSmartArtTypes.PictureFrame,
		c_oAscSmartArtTypes.TabList
	];

	var c_oAscSmartArtSectionNames = {
		List:         0,
		Process:      1,
		Cycle:        2,
		Hierarchy:    3,
		Relationship: 4,
		Matrix:       5,
		Pyramid:      6,
		Picture:      7,
		OfficeCom:    8
	};

	var c_oAscSmartArtSections = {};
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.List]         = c_oAscSmartArtListTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.Process]      = c_oAscSmartArtProcessTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.Cycle]        = c_oAscSmartArtCycleTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.Hierarchy]    = c_oAscSmartArtHierarchyTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.Relationship] = c_oAscSmartArtRelationshipTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.Matrix]       = c_oAscSmartArtMatrixTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.Pyramid]      = c_oAscSmartArtPyramidTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.Picture]      = c_oAscSmartArtPictureTypes;
	c_oAscSmartArtSections[c_oAscSmartArtSectionNames.OfficeCom]    = c_oAscSmartArtOfficeComTypes;


	var c_oAscJSONNumberingType = {
		Remove : "remove",
		Bullet : "bullet",
		Number : "number",
		Hybrid : "hybrid"
	};

	/** @enum {number} */
	var c_oAscNumberingFormat = {
		Aiueo                        :  0,
		AiueoFullWidth               :  1,
		ArabicAbjad                  :  2,
		ArabicAlpha                  :  3,
		BahtText                     :  4,
		Bullet                       :  5,
		CardinalText                 :  6,
		Chicago                      :  7,
		ChineseCounting              :  8,
		ChineseCountingThousand      :  9,
		ChineseLegalSimplified       : 10,
		Chosung                      : 11,
		Custom                       : 12,
		Decimal                      : 13,
		DecimalEnclosedCircle        : 14,
		DecimalEnclosedCircleChinese : 15,
		DecimalEnclosedFullstop      : 16,
		DecimalEnclosedParen         : 17,
		DecimalFullWidth             : 18,
		DecimalFullWidth2            : 19,
		DecimalHalfWidth             : 20,
		DecimalZero                  : 21,
		DollarText                   : 22,
		Ganada                       : 23,
		Hebrew1                      : 24,
		Hebrew2                      : 25,
		Hex                          : 26,
		HindiConsonants              : 27,
		HindiCounting                : 28,
		HindiNumbers                 : 29,
		HindiVowels                  : 30,
		IdeographDigital             : 31,
		IdeographEnclosedCircle      : 32,
		IdeographLegalTraditional    : 33,
		IdeographTraditional         : 34,
		IdeographZodiac              : 35,
		IdeographZodiacTraditional   : 36,
		Iroha                        : 37,
		IrohaFullWidth               : 38,
		JapaneseCounting             : 39,
		JapaneseDigitalTenThousand   : 40,
		JapaneseLegal                : 41,
		KoreanCounting               : 42,
		KoreanDigital                : 43,
		KoreanDigital2               : 44,
		KoreanLegal                  : 45,
		LowerLetter                  : 46,
		LowerRoman                   : 47,
		None                         : 48,
		NumberInDash                 : 49,
		Ordinal                      : 50,
		OrdinalText                  : 51,
		RussianLower                 : 52,
		RussianUpper                 : 53,
		TaiwaneseCounting            : 54,
		TaiwaneseCountingThousand    : 55,
		TaiwaneseDigital             : 56,
		ThaiCounting                 : 57,
		ThaiLetters                  : 58,
		ThaiNumbers                  : 59,
		UpperLetter                  : 60,
		UpperRoman                   : 61,
		VietnameseCounting           : 62,

		Ea1JpnKor                    : 0x3000,
		CircleNumWdBlack             : 0x3001,
		Ea1JpnChsDb                  : 0x3002,
		Ea1Cht                       : 0x3003,
		CircleNumWdWhitePlain        : 0x3004,

		CustomGreece                 : 0x4000,
		CustomDecimalFourZero        : 0x4001,
		CustomDecimalThreeZero       : 0x4002,
		CustomDecimalTwoZero         : 0x4003,
		CustomUpperTurkish           : 0x4004,
		CustomLowerTurkish           : 0x4005
	};

	var c_oAscCustomNumberingFormatAssociation = {
		'α, β, γ, ...': c_oAscNumberingFormat.CustomGreece,
		'00001, 00002, 00003, ...': c_oAscNumberingFormat.CustomDecimalFourZero,
		'0001, 0002, 0003, ...': c_oAscNumberingFormat.CustomDecimalThreeZero,
		'001, 002, 003, ...': c_oAscNumberingFormat.CustomDecimalTwoZero,
		'A, Ç, Ĝ, ...': c_oAscNumberingFormat.CustomUpperTurkish,
		'a, ç, ĝ, ...': c_oAscNumberingFormat.CustomLowerTurkish
	};

	var c_oAscAllNumberingTypes = [c_oAscNumberingFormat.Aiueo, c_oAscNumberingFormat.AiueoFullWidth, c_oAscNumberingFormat.ArabicAbjad, c_oAscNumberingFormat.ArabicAlpha, c_oAscNumberingFormat.BahtText, c_oAscNumberingFormat.CardinalText, c_oAscNumberingFormat.Chicago, c_oAscNumberingFormat.ChineseCounting, c_oAscNumberingFormat.ChineseCountingThousand, c_oAscNumberingFormat.ChineseLegalSimplified, c_oAscNumberingFormat.Chosung, c_oAscNumberingFormat.Decimal, c_oAscNumberingFormat.DecimalEnclosedCircle, c_oAscNumberingFormat.DecimalEnclosedCircleChinese, c_oAscNumberingFormat.DecimalEnclosedFullstop, c_oAscNumberingFormat.DecimalEnclosedParen, c_oAscNumberingFormat.DecimalFullWidth, c_oAscNumberingFormat.DecimalFullWidth2, c_oAscNumberingFormat.DecimalHalfWidth, c_oAscNumberingFormat.DecimalZero, c_oAscNumberingFormat.Ganada, c_oAscNumberingFormat.Hebrew1, c_oAscNumberingFormat.Hebrew2, c_oAscNumberingFormat.Hex, c_oAscNumberingFormat.HindiConsonants, c_oAscNumberingFormat.HindiCounting, c_oAscNumberingFormat.HindiNumbers, c_oAscNumberingFormat.HindiVowels, c_oAscNumberingFormat.IdeographDigital, c_oAscNumberingFormat.IdeographEnclosedCircle, c_oAscNumberingFormat.IdeographLegalTraditional, c_oAscNumberingFormat.IdeographTraditional, c_oAscNumberingFormat.IdeographZodiac, c_oAscNumberingFormat.IdeographZodiacTraditional, c_oAscNumberingFormat.Iroha, c_oAscNumberingFormat.IrohaFullWidth, c_oAscNumberingFormat.JapaneseCounting, c_oAscNumberingFormat.JapaneseDigitalTenThousand, c_oAscNumberingFormat.JapaneseLegal, c_oAscNumberingFormat.KoreanCounting, c_oAscNumberingFormat.KoreanDigital, c_oAscNumberingFormat.KoreanDigital2, c_oAscNumberingFormat.KoreanLegal, c_oAscNumberingFormat.LowerLetter, c_oAscNumberingFormat.LowerRoman, c_oAscNumberingFormat.NumberInDash, c_oAscNumberingFormat.Ordinal, c_oAscNumberingFormat.OrdinalText, c_oAscNumberingFormat.RussianLower, c_oAscNumberingFormat.RussianUpper, c_oAscNumberingFormat.TaiwaneseCounting, c_oAscNumberingFormat.TaiwaneseCountingThousand, c_oAscNumberingFormat.TaiwaneseDigital, c_oAscNumberingFormat.ThaiCounting, c_oAscNumberingFormat.ThaiLetters, c_oAscNumberingFormat.ThaiNumbers, c_oAscNumberingFormat.UpperLetter, c_oAscNumberingFormat.UpperRoman, c_oAscNumberingFormat.VietnameseCounting, c_oAscNumberingFormat.CustomGreece, c_oAscNumberingFormat.CustomDecimalFourZero, c_oAscNumberingFormat.CustomDecimalThreeZero, c_oAscNumberingFormat.CustomDecimalTwoZero, c_oAscNumberingFormat.CustomUpperTurkish, c_oAscNumberingFormat.CustomLowerTurkish];

	/** enum {number} */
	var c_oAscNumberingSuff = {
		None  : 0x00,
		Space : 0x01,
		Tab   : 0x02
	};

	var c_oAscNumberingLvlTextType = {
		Text : 0x00,
		Num  : 0x01
	};
	
	var c_oAscChapterSep = {
		Colon : 0,
		EmDash : 1,
		EnDash : 2,
		Hyphen : 3,
		Period : 4
	};

	var c_oAscSdtAppearance = {
		Frame  : 1,
		Hidden : 2
	};


	var c_oAscObjectsAlignType = {
		Selected: 0,
		Slide: 1,
		Page: 2,
		Margin: 3
	};

	var c_oAscItemType = {
		Default: 0,
		Avg: 1,
		Count: 2,
		CountA: 3,
		Max: 4,
		Min: 5,
		Product: 6,
		StdDev: 7,
		StdDevP: 8,
		Sum: 9,
		Var: 10,
		VarP: 11,
		Data: 12,
		Grand: 13,
		Blank: 14
	};

	var c_oAscRevisionsMove = {
		NoMove   : 0,
		MoveTo   : 1,
		MoveFrom : 2
	};

	/** @enum {number} */
	var c_oAscRevisionsChangeType = {
		Unknown    : 0x00,
		TextAdd    : 0x01,
		TextRem    : 0x02,
		ParaAdd    : 0x03,
		ParaRem    : 0x04,
		TextPr     : 0x05,
		ParaPr     : 0x06,
		TablePr    : 0x07,
		RowsAdd    : 0x08,
		RowsRem    : 0x09,
		TableRowPr : 0x0A,

		MoveMark       : 0xFE, // специальный внутренний тип, для обозначения меток переноса
		MoveMarkRemove : 0xFF  // внутреннний тип, для удаления отметок переноса внутри параграфов и таблиц
	};


	/** @enum {number} */
	var c_oAscSectionBreakType = {
		NextPage   : 0x00,
		OddPage    : 0x01,
		EvenPage   : 0x02,
		Continuous : 0x03,
		Column     : 0x04
	};


	var c_oAscSdtLockType = {
		ContentLocked    : 0x00,
		SdtContentLocked : 0x01,
		SdtLocked        : 0x02,
		Unlocked         : 0x03
	};


	/**
	 * Типы горизонтального прилегания для автофигур.
	 * @type {{Center: number, Inside: number, Left: number, Outside: number, Right: number}}
	 * @enum {number}
	 */
	var c_oAscAlignH = {
		Center  : 0x00,
		Inside  : 0x01,
		Left    : 0x02,
		Outside : 0x03,
		Right   : 0x04
	};

	/**
	 * Типы вертикального прилегания для автофигур.
	 * @type {{Bottom: number, Center: number, Inside: number, Outside: number, Top: number}}
	 * @enum {number}
	 */
	var c_oAscAlignV = {
		Bottom  : 0x00,
		Center  : 0x01,
		Inside  : 0x02,
		Outside : 0x03,
		Top     : 0x04
	};



	var c_oAscWatermarkType = {
		None       : 0,
		Text       : 1,
		Image      : 2
	};

	var c_oAscCalendarType = {
		Gregorian            : 0,
		GregorianArabic      : 1,
		GregorianMeFrench    : 2,
		GregorianUs          : 3,
		GregorianXlitEnglish : 4,
		GregorianXlitFrench  : 5,
		Hebrew               : 6,
		Hijri                : 7,
		Japan                : 8,
		Korea                : 9,
		None                 : 10,
		Saka                 : 11,
		Taiwan               : 12,
		Thai                 : 13
	};

	var c_oAscContentControlSpecificType = {
		None         : 0,
		CheckBox     : 1,
		Picture      : 2,
		ComboBox     : 3,
		DropDownList : 4,
		DateTime     : 5,

		TOC          : 10,
		Complex      : 11,
		
		toString : function(type)
		{
			switch (type)
			{
				case c_oAscContentControlSpecificType.None:
					return "text";
				case c_oAscContentControlSpecificType.CheckBox:
					return "checkBox";
				case c_oAscContentControlSpecificType.Picture:
					return "picture";
				case c_oAscContentControlSpecificType.ComboBox:
					return "comboBox";
				case c_oAscContentControlSpecificType.DropDownList:
					return "dropDownList";
				case c_oAscContentControlSpecificType.DateTime:
					return "dateTime";
				case c_oAscContentControlSpecificType.TOC:
					return "toc";
				case c_oAscContentControlSpecificType.Complex:
					return "complex";
			}
			
			return "unknown";
		},
		
		fromString : function(value)
		{
			switch (value)
			{
				case "checkBox":
				case "radio":
					return c_oAscContentControlSpecificType.CheckBox;
				case "picture":
					return c_oAscContentControlSpecificType.Picture;
				case "comboBox":
					return c_oAscContentControlSpecificType.ComboBox;
				case "dropDownList":
					return c_oAscContentControlSpecificType.DropDownList;
				case "dateTime":
					return c_oAscContentControlSpecificType.DateTime;
				case "toc":
					return c_oAscContentControlSpecificType.TOC;
				case "complex":
					return c_oAscContentControlSpecificType.Complex;
			}
			
			return c_oAscContentControlSpecificType.None;
		}
	};

	var c_oAscDefNameType = {
		none: 0,
		table: 1,
		slicer: 2
	};

	var c_oAscChangeTextCaseType = {
		SentenceCase    : 0,
		LowerCase       : 1,
		UpperCase       : 2,
		CapitalizeWords : 3,
		ToggleCase      : 4
	};

	var g_aLcidNameIdArray = [
		"ar", 0x0001 ,
		"bg", 0x0002 ,
		"ca", 0x0003 ,
		"zh-Hans", 0x0004 ,
		"cs", 0x0005 ,
		"da", 0x0006 ,
		"de", 0x0007 ,
		"el", 0x0008 ,
		"en", 0x0009 ,
		"es", 0x000a ,
		"fi", 0x000b ,
		"fr", 0x000c ,
		"he", 0x000d ,
		"hu", 0x000e ,
		"is", 0x000f ,
		"it", 0x0010 ,
		"ja", 0x0011 ,
		"ko", 0x0012 ,
		"nl", 0x0013 ,
		"no", 0x0014 ,
		"pl", 0x0015 ,
		"pt", 0x0016 ,
		"rm", 0x0017 ,
		"ro", 0x0018 ,
		"ru", 0x0019 ,
		"hr", 0x001a ,
		"sk", 0x001b ,
		"sq", 0x001c ,
		"sv", 0x001d ,
		"th", 0x001e ,
		"tr", 0x001f ,
		"ur", 0x0020 ,
		"id", 0x0021 ,
		"uk", 0x0022 ,
		"be", 0x0023 ,
		"sl", 0x0024 ,
		"et", 0x0025 ,
		"lv", 0x0026 ,
		"lt", 0x0027 ,
		"tg", 0x0028 ,
		"fa", 0x0029 ,
		"vi", 0x002a ,
		"hy", 0x002b ,
		"az", 0x002c ,
		"eu", 0x002d ,
		"hsb", 0x002e ,
		"mk", 0x002f ,
		"tn", 0x0032 ,
		"xh", 0x0034 ,
		"zu", 0x0035 ,
		"af", 0x0036 ,
		"ka", 0x0037 ,
		"fo", 0x0038 ,
		"hi", 0x0039 ,
		"mt", 0x003a ,
		"se", 0x003b ,
		"ga", 0x003c ,
		"ms", 0x003e ,
		"kk", 0x003f ,
		"ky", 0x0040 ,
		"sw", 0x0041 ,
		"tk", 0x0042 ,
		"uz", 0x0043 ,
		"tt", 0x0044 ,
		"bn", 0x0045 ,
		"pa", 0x0046 ,
		"gu", 0x0047 ,
		"or", 0x0048 ,
		"ta", 0x0049 ,
		"te", 0x004a ,
		"kn", 0x004b ,
		"ml", 0x004c ,
		"as", 0x004d ,
		"mr", 0x004e ,
		"sa", 0x004f ,
		"mn", 0x0050 ,
		"bo", 0x0051 ,
		"cy", 0x0052 ,
		"km", 0x0053 ,
		"lo", 0x0054 ,
		"gl", 0x0056 ,
		"kok", 0x0057 ,
		"syr", 0x005a ,
		"si", 0x005b ,
		"iu", 0x005d ,
		"am", 0x005e ,
		"tzm", 0x005f ,
		"ne", 0x0061 ,
		"fy", 0x0062 ,
		"ps", 0x0063 ,
		"fil", 0x0064 ,
		"dv", 0x0065 ,
		"ha", 0x0068 ,
		"yo", 0x006a ,
		"quz", 0x006b ,
		"nso", 0x006c ,
		"ba", 0x006d ,
		"lb", 0x006e ,
		"kl", 0x006f ,
		"ig", 0x0070 ,
		"ii", 0x0078 ,
		"arn", 0x007a ,
		"moh", 0x007c ,
		"br", 0x007e ,
		"ug", 0x0080 ,
		"mi", 0x0081 ,
		"oc", 0x0082 ,
		"co", 0x0083 ,
		"gsw", 0x0084 ,
		"sah", 0x0085 ,
		"qut", 0x0086 ,
		"rw", 0x0087 ,
		"wo", 0x0088 ,
		"prs", 0x008c ,
		"gd", 0x0091 ,
		"ar-SA", 0x0401 ,
		"bg-BG", 0x0402 ,
		"ca-ES", 0x0403 ,
		"zh-TW", 0x0404 ,
		"cs-CZ", 0x0405 ,
		"da-DK", 0x0406 ,
		"de-DE", 0x0407 ,
		"el-GR", 0x0408 ,
		"en-US", 0x0409 ,
		"es-ES_tradnl", 0x040a ,
		"fi-FI", 0x040b ,
		"fr-FR", 0x040c ,
		"he-IL", 0x040d ,
		"hu-HU", 0x040e ,
		"is-IS", 0x040f ,
		"it-IT", 0x0410 ,
		"ja-JP", 0x0411 ,
		"ko-KR", 0x0412 ,
		"nl-NL", 0x0413 ,
		"nb-NO", 0x0414 ,
		"pl-PL", 0x0415 ,
		"pt-BR", 0x0416 ,
		"rm-CH", 0x0417 ,
		"ro-RO", 0x0418 ,
		"ru-RU", 0x0419 ,
		"hr-HR", 0x041a ,
		"sk-SK", 0x041b ,
		"sq-AL", 0x041c ,
		"sv-SE", 0x041d ,
		"th-TH", 0x041e ,
		"tr-TR", 0x041f ,
		"ur-PK", 0x0420 ,
		"id-ID", 0x0421 ,
		"uk-UA", 0x0422 ,
		"be-BY", 0x0423 ,
		"sl-SI", 0x0424 ,
		"et-EE", 0x0425 ,
		"lv-LV", 0x0426 ,
		"lt-LT", 0x0427 ,
		"tg-Cyrl-TJ", 0x0428 ,
		"fa-IR", 0x0429 ,
		"vi-VN", 0x042a ,
		"hy-AM", 0x042b ,
		"az-Latn-AZ", 0x042c ,
		"eu-ES", 0x042d ,
		"wen-DE", 0x042e ,
		"mk-MK", 0x042f ,
		"st-ZA", 0x0430 ,
		"ts-ZA", 0x0431 ,
		"tn-ZA", 0x0432 ,
		"ven-ZA", 0x0433 ,
		"xh-ZA", 0x0434 ,
		"zu-ZA", 0x0435 ,
		"af-ZA", 0x0436 ,
		"ka-GE", 0x0437 ,
		"fo-FO", 0x0438 ,
		"hi-IN", 0x0439 ,
		"mt-MT", 0x043a ,
		"se-NO", 0x043b ,
		"ms-MY", 0x043e ,
		"kk-KZ", 0x043f ,
		"ky-KG", 0x0440 ,
		"sw-KE", 0x0441 ,
		"tk-TM", 0x0442 ,
		"uz-Latn-UZ", 0x0443 ,
		"tt-RU", 0x0444 ,
		"bn-IN", 0x0445 ,
		"pa-IN", 0x0446 ,
		"gu-IN", 0x0447 ,
		"or-IN", 0x0448 ,
		"ta-IN", 0x0449 ,
		"te-IN", 0x044a ,
		"kn-IN", 0x044b ,
		"ml-IN", 0x044c ,
		"as-IN", 0x044d ,
		"mr-IN", 0x044e ,
		"sa-IN", 0x044f ,
		"mn-MN", 0x0450 ,
		"bo-CN", 0x0451 ,
		"cy-GB", 0x0452 ,
		"km-KH", 0x0453 ,
		"lo-LA", 0x0454 ,
		"my-MM", 0x0455 ,
		"gl-ES", 0x0456 ,
		"kok-IN", 0x0457 ,
		"mni", 0x0458 ,
		"sd-IN", 0x0459 ,
		"syr-SY", 0x045a ,
		"si-LK", 0x045b ,
		"chr-US", 0x045c ,
		"iu-Cans-CA", 0x045d ,
		"am-ET", 0x045e ,
		"tmz", 0x045f ,
		"ne-NP", 0x0461 ,
		"fy-NL", 0x0462 ,
		"ps-AF", 0x0463 ,
		"fil-PH", 0x0464 ,
		"dv-MV", 0x0465 ,
		"bin-NG", 0x0466 ,
		"fuv-NG", 0x0467 ,
		"ha-Latn-NG", 0x0468 ,
		"ibb-NG", 0x0469 ,
		"yo-NG", 0x046a ,
		"quz-BO", 0x046b ,
		"nso-ZA", 0x046c ,
		"ba-RU", 0x046d ,
		"lb-LU", 0x046e ,
		"kl-GL", 0x046f ,
		"ig-NG", 0x0470 ,
		"kr-NG", 0x0471 ,
		"gaz-ET", 0x0472 ,
		"ti-ER", 0x0473 ,
		"gn-PY", 0x0474 ,
		"haw-US", 0x0475 ,
		"so-SO", 0x0477 ,
		"ii-CN", 0x0478 ,
		"pap-AN", 0x0479 ,
		"arn-CL", 0x047a ,
		"moh-CA", 0x047c ,
		"br-FR", 0x047e ,
		"ug-CN", 0x0480 ,
		"mi-NZ", 0x0481 ,
		"oc-FR", 0x0482 ,
		"co-FR", 0x0483 ,
		"gsw-FR", 0x0484 ,
		"sah-RU", 0x0485 ,
		"qut-GT", 0x0486 ,
		"rw-RW", 0x0487 ,
		"wo-SN", 0x0488 ,
		"prs-AF", 0x048c ,
		"plt-MG", 0x048d ,
		"gd-GB", 0x0491 ,
		"ar-IQ", 0x0801 ,
		"zh-CN", 0x0804 ,
		"de-CH", 0x0807 ,
		"en-GB", 0x0809 ,
		"es-MX", 0x080a ,
		"fr-BE", 0x080c ,
		"it-CH", 0x0810 ,
		"nl-BE", 0x0813 ,
		"nn-NO", 0x0814 ,
		"pt-PT", 0x0816 ,
		"ro-MO", 0x0818 ,
		"ru-MO", 0x0819 ,
		"sr-Latn-CS", 0x081a ,
		"sv-FI", 0x081d ,
		"ur-IN", 0x0820 ,
		"az-Cyrl-AZ", 0x082c ,
		"dsb-DE", 0x082e ,
		"se-SE", 0x083b ,
		"ga-IE", 0x083c ,
		"ms-BN", 0x083e ,
		"uz-Cyrl-UZ", 0x0843 ,
		"bn-BD", 0x0845 ,
		"pa-PK", 0x0846 ,
		"mn-Mong-CN", 0x0850 ,
		"bo-BT", 0x0851 ,
		"sd-PK", 0x0859 ,
		"iu-Latn-CA", 0x085d ,
		"tzm-Latn-DZ", 0x085f ,
		"ne-IN", 0x0861 ,
		"quz-EC", 0x086b ,
		"ti-ET", 0x0873 ,
		"ar-EG", 0x0c01 ,
		"zh-HK", 0x0c04 ,
		"de-AT", 0x0c07 ,
		"en-AU", 0x0c09 ,
		"es-ES", 0x0c0a ,
		"fr-CA", 0x0c0c ,
		"sr-Cyrl-CS", 0x0c1a ,
		"se-FI", 0x0c3b ,
		"tmz-MA", 0x0c5f ,
		"quz-PE", 0x0c6b ,
		"ar-LY", 0x1001 ,
		"zh-SG", 0x1004 ,
		"de-LU", 0x1007 ,
		"en-CA", 0x1009 ,
		"es-GT", 0x100a ,
		"fr-CH", 0x100c ,
		"hr-BA", 0x101a ,
		"smj-NO", 0x103b ,
		"ar-DZ", 0x1401 ,
		"zh-MO", 0x1404 ,
		"de-LI", 0x1407 ,
		"en-NZ", 0x1409 ,
		"es-CR", 0x140a ,
		"fr-LU", 0x140c ,
		"bs-Latn-BA", 0x141a ,
		"smj-SE", 0x143b ,
		"ar-MA", 0x1801 ,
		"en-IE", 0x1809 ,
		"es-PA", 0x180a ,
		"fr-MC", 0x180c ,
		"sr-Latn-BA", 0x181a ,
		"sma-NO", 0x183b ,
		"ar-TN", 0x1c01 ,
		"en-ZA", 0x1c09 ,
		"es-DO", 0x1c0a ,
		"fr-West", 0x1c0c ,
		"sr-Cyrl-BA", 0x1c1a ,
		"sma-SE", 0x1c3b ,
		"ar-OM", 0x2001 ,
		"en-JM", 0x2009 ,
		"es-VE", 0x200a ,
		"fr-RE", 0x200c ,
		"bs-Cyrl-BA", 0x201a ,
		"sms-FI", 0x203b ,
		"ar-YE", 0x2401 ,
		"en-CB", 0x2409 ,
		"es-CO", 0x240a ,
		"fr-CG", 0x240c ,
		"sr-Latn-RS", 0x241a ,
		"smn-FI", 0x243b ,
		"ar-SY", 0x2801 ,
		"en-BZ", 0x2809 ,
		"es-PE", 0x280a ,
		"fr-SN", 0x280c ,
		"sr-Cyrl-RS", 0x281a ,
		"ar-JO", 0x2c01 ,
		"en-TT", 0x2c09 ,
		"es-AR", 0x2c0a ,
		"fr-CM", 0x2c0c ,
		"sr-Latn-ME", 0x2c1a ,
		"ar-LB", 0x3001 ,
		"en-ZW", 0x3009 ,
		"es-EC", 0x300a ,
		"fr-CI", 0x300c ,
		"sr-Cyrl-ME", 0x301a ,
		"ar-KW", 0x3401 ,
		"en-PH", 0x3409 ,
		"es-CL", 0x340a ,
		"fr-ML", 0x340c ,
		"ar-AE", 0x3801 ,
		"en-ID", 0x3809 ,
		"es-UY", 0x380a ,
		"fr-MA", 0x380c ,
		"ar-BH", 0x3c01 ,
		"en-HK", 0x3c09 ,
		"es-PY", 0x3c0a ,
		"fr-HT", 0x3c0c ,
		"ar-QA", 0x4001 ,
		"en-IN", 0x4009 ,
		"es-BO", 0x400a ,
		"en-MY", 0x4409 ,
		"es-SV", 0x440a ,
		"en-SG", 0x4809 ,
		"es-HN", 0x480a ,
		"es-NI", 0x4c0a ,
		"es-PR", 0x500a ,
		"es-US", 0x540a ,
		"bs-Cyrl", 0x641a ,
		"bs-Latn", 0x681a ,
		"sr-Cyrl", 0x6c1a ,
		"sr-Latn", 0x701a ,
		"smn", 0x703b ,
		"az-Cyrl", 0x742c ,
		"sms", 0x743b ,
		"zh", 0x7804 ,
		"nn", 0x7814 ,
		"bs", 0x781a ,
		"az-Latn", 0x782c ,
		"sma", 0x783b ,
		"uz-Cyrl", 0x7843 ,
		"mn-Cyrl", 0x7850 ,
		"iu-Cans", 0x785d ,
		"zh-Hant", 0x7c04 ,
		"nb", 0x7c14 ,
		"sr", 0x7c1a ,
		"tg-Cyrl", 0x7c28 ,
		"dsb", 0x7c2e ,
		"smj", 0x7c3b ,
		"uz-Latn", 0x7c43 ,
		"mn-Mong", 0x7c50 ,
		"iu-Latn", 0x7c5d ,
		"tzm-Latn", 0x7c5f ,
		"ha-Latn", 0x7c68 ];
	var g_oLcidNameToIdMap = {};
	var g_oLcidIdToNameMap = {};
	for(var i = 0, length = g_aLcidNameIdArray.length; i + 1< length; i+=2)
	{
		var name = g_aLcidNameIdArray[i];
		var id = g_aLcidNameIdArray[i + 1];
		g_oLcidNameToIdMap[name] = id;
		g_oLcidIdToNameMap[id] = name;
	}
	var availableIdeographLanguages = ['zh-CN', 'vi-VN', 'ko-KR', 'ja-JP', 'zh-Hans', 'zh-TW', 'zh-CN', 'zh-HK', 'zh-SG',
		'zh-MO', 'zh-Hant', 'zh'];
	var availableBidiLanguages = [];
	var document_compatibility_mode_Word11 = 11;
	var document_compatibility_mode_Word12 = 12;
	var document_compatibility_mode_Word14 = 14;
	var document_compatibility_mode_Word15 = 15;

	var document_compatibility_mode_Current = document_compatibility_mode_Word12;

	var c_oAscCustomShortcutType = {
		Symbol : 1
	};

	var c_oAscLineNumberRestartType = {
		Continuous : 1,
		NewPage    : 2,
		NewSection : 3
	};

	var c_oAscSectionApplyType = {
		Current : 0,
		ToEnd   : 1,
		All     : 2
	};

	var c_oAscSlideSZType = {
			Sz35mm: 0,
			SzA3: 1,
			SzA4: 2,
			SzB4ISO: 3,
			SzB4JIS: 4,
			SzB5ISO: 5,
			SzB5JIS: 6,
			SzBanner: 7,
			SzCustom: 8,
			SzHagakiCard: 9,
			SzLedger: 10,
			SzLetter: 11,
			SzOverhead: 12,
			SzScreen16x10: 13,
			SzScreen16x9: 14,
			SzScreen4x3: 15,
			SzWidescreen: 16
	};

	var c_oAscPictureFormScaleFlag = {
		Always  : 0,
		Bigger  : 1,
		Smaller : 2,
		Never   : 3,
	};

	var c_oAscDisplayModeInReview = {
		Edit     : 0,
		Final    : 1,
		Original : 2,
		Simple   : 3
	};

	var c_oAscMathInputType = {
		Unicode : 0,
		LaTeX   : 1
	};

	const LigaturesFlags = {
		Standard     : 0x01, //AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_STANDARD,
		Contextual   : 0x02, //AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_CONTEXTUAL,
		Historical   : 0x04, //AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_HISTORICAL,
		Discretional : 0x08  //AscFonts.HB_FEATURE.HB_FEATURE_TYPE_LIGATURES_DISCRETIONARY
	};

	const LigaturesType = {
		None                             : 0x00,
		Standard                         : LigaturesFlags.Standard,
		Contextual                       : LigaturesFlags.Contextual,
		Historical                       : LigaturesFlags.Historical,
		Discretional                     : LigaturesFlags.Discretional,
		StandardContextual               : LigaturesFlags.Standard | LigaturesFlags.Contextual,
		StandardHistorical               : LigaturesFlags.Standard | LigaturesFlags.Historical,
		ContextualHistorical             : LigaturesFlags.Contextual | LigaturesFlags.Historical,
		StandardDiscretional             : LigaturesFlags.Standard | LigaturesFlags.Discretional,
		ContextualDiscretional           : LigaturesFlags.Contextual | LigaturesFlags.Discretional,
		HistoricalDiscretional           : LigaturesFlags.Historical | LigaturesFlags.Discretional,
		StandardContextualHistorical     : LigaturesFlags.Standard | LigaturesFlags.Contextual | LigaturesFlags.Historical,
		StandardContextualDiscretional   : LigaturesFlags.Standard | LigaturesFlags.Contextual | LigaturesFlags.Discretional,
		StandardHistoricalDiscretional   : LigaturesFlags.Standard | LigaturesFlags.Historical | LigaturesFlags.Discretional,
		ContextualHistoricalDiscretional : LigaturesFlags.Contextual | LigaturesFlags.Historical | LigaturesFlags.Discretional,
		All                              : LigaturesFlags.Standard | LigaturesFlags.Contextual | LigaturesFlags.Historical | LigaturesFlags.Discretional
	};

	const CombFormWidthRule = {
		AtLeast : 0,
		Auto    : 1,
		Exact   : 2
	};

	const UnderlineType = {
		Dash            : 0,
		DashDotDotHeavy : 1,
		DashDotHeavy    : 2,
		DashedHeavy     : 3,
		DashLong        : 4,
		DashLongHeavy   : 5,
		DotDash         : 6,
		DotDotDash      : 7,
		Dotted          : 8,
		DottedHeavy     : 9,
		Double          : 10,
		None            : 11,
		Single          : 12,
		Thick           : 13,
		Wave            : 14,
		WavyDouble      : 15,
		WavyHeavy       : 16,
		Words           : 17
	};

	const DocumentView = {
		MasterPages : 0,
		None        : 1,
		Normal      : 2,
		Outline     : 3,
		Print       : 4,
		Web         : 5
	};

	var c_oAscConfirm = {
		ConfirmReplaceRange: 0,
		ConfirmPutMergeRange: 1,
		ConfirmReplaceFormulaInTable: 2,
		ConfirmChangeProtectRange: 3,
		ConfirmMaxChangesSize: 4,
		ConfirmAddCellWatches: 5,
		ConfirmReplaceHeaderFooterPicture: 6
	};


	const c_oPluginContextMenuTypes = {
		None: "None",
		Target: "Target",
		Selection: "Selection",
		Image: "Image",
		Shape: "Shape",
		OleObject: "OleObject"
	};
	
	// 17.18.21 ST_EdGrp
	const ST_EdGrp = {
		administrators : 0,
		contributors   : 1,
		current        : 2,
		editors        : 3,
		everyone       : 4,
		none           : 5,
		owners         : 6
	};
	
	window['Asc']['ST_EdGrp'] = window['Asc'].ST_EdGrp = ST_EdGrp;
	ST_EdGrp['administrators'] = ST_EdGrp.administrators;
	ST_EdGrp['contributors']   = ST_EdGrp.contributors;
	ST_EdGrp['current']        = ST_EdGrp.current;
	ST_EdGrp['editors']        = ST_EdGrp.editors;
	ST_EdGrp['everyone']       = ST_EdGrp.everyone;
	ST_EdGrp['none']           = ST_EdGrp.none;

	// 17.18.13 ST_DisplacedByCustomXml
	const ST_DisplacedByCustomXml = {
		next : 0,
		prev : 1
	};
	window['Asc']['ST_DisplacedByCustomXml'] = window['Asc'].ST_DisplacedByCustomXml = ST_DisplacedByCustomXml;
	ST_DisplacedByCustomXml['next'] = ST_DisplacedByCustomXml.next;
	ST_DisplacedByCustomXml['prev'] = ST_DisplacedByCustomXml.prev;
	
	
	var c_oAscDateTimeFormat = {};
	c_oAscDateTimeFormat[lcid_arSA] = [
		"dd/MM/yyyy",
		"dddd, dd MMMM, yyyy",
		"dd MMMM, yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd.MM.yyyy",
		"dd MMM, yy",
		"d MMMM yyyy",
		"MMMM yy",
		"MMM-yy",
		"dd/MM/yyyy hh:mm am/pm",
		"dd/MM/yyyy hh:mm:ss am/pm",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_azLatnAZ] = [
		"dd.MM.yyyy",
		"d MMMM yyyy, dddd",
		"d MMMM yyyy",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd/MM/yyyy",
		"d MMM. yy",
		"dd/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"dd.MM.yyyy HH:mm",
		"dd.MM.yyyy HH:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_bgBG]     = [
		"d.M.yyyy 'г.'",
		"dddd, dd MMMM yyyy 'г.'",
		"dd MMMM yyyy 'г.'",
		"d.M.yy 'г.'",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/M/yyyy",
		"dd MMM. yy 'г.'",
		"d/M/yy",
		"MMMM yy 'г'",
		"MMM-yy",
		"d.M.yyyy 'г.' H:mm",
		"d.M.yyyy 'г.' H:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_csCZ]     = [
		"dd.MM.yyyy",
		"dddd d. MMMM yyyy",
		"d. MMMM yyyy",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd/MM/yyyy",
		"d. MMM. yy",
		"dd/MM/yy",
		"MMMM ’yy",
		"MMM-yy",
		"dd.MM.yyyy H:mm",
		"dd.MM.yyyy H:mm:ss",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_daDK]     = [
		"dd-MM-yyyy",
		"d. MMMM yyyy",
		"yy-MM-dd",
		"yyyy.MM.dd",
		"yyyy-MM-dd",
		"MMMM yyyy",
		"d.M.yy",
		"d/M yyyy",
		"dd.MM.yyyy",
		"d.M.yyyy",
		"dd/MM yyyy",
		"dd-MM-yyyy HH:mm",
		"dd-MM-yyyy HH:mm:ss",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_deAT]     = c_oAscDateTimeFormat[lcid_deDE] = [
		"dd.MM.yyyy",
		"dddd, d. MMMM yyyy",
		"d. MMMM yyyy",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"yy-MM-dd",
		"dd/MM/yyyy",
		"dd. MMM. yyyy",
		"dd/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"dd.MM.yyyy HH:mm",
		"dd.MM.yyyy HH:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_deCH]     = [
		"dd.MM.yyyy",
		"dddd, d. MMMM yyyy",
		"d. MMMM yyyy",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"yy-MM-dd",
		"dd/MM/yyyy",
		"dd. MMM. yyyy",
		"dd/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"dd.MM.yyyy HH:mm",
		"dd.MM.yyyy HH:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_elGR]     = [
		"d/M/yyyy",
		"dddd, d MMMM yyyy",
		"d MMMM yyyy",
		"d/M/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d.M.yyyy",
		"d MMM. yy",
		"d.M.yy",
		"MMMM yy",
		"MMM-yy",
		"d/M/yyyy h:mm am/pm",
		"d/M/yyyy h:mm:ss am/pm",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_enAU]     = [
		"d/MM/yyyy",
		"dddd, d MMMM yyyy",
		"d MMMM yyyy",
		"d/MM/yy",
		"yyyy-MM-dd",
		"d-MMM-yyyy",
		"d.MM.yyyy",
		"d MMM. yy",
		"MMMM yy",
		"MMM-yy",
		"d/MM/yyyy h:mm AM/PM",
		"d/MM/yyyy h:mm:ss AM/PM",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_enGB]     = [
		"dd/MM/yyyy",
		"dddd, dd MMMM yyyy",
		"dd MMMM yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd.MM.yyyy",
		"dd MMM. yy",
		"d MMMM yyyy",
		"MMMM yy",
		"MMM-yy",
		"dd/MM/yyyy HH:mm",
		"dd/MM/yyyy HH:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_esES]     = [
		"dd/MM/yyyy",
		"dddd, d 'de' MMMM 'de' yyyy",
		"d 'de' MMMM 'de' yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM.-yy",
		"dd.MM.yyyy",
		"d MMM. yy",
		"dd.MM.yy",
		"MMMM 'de' yyyy",
		"MMM.-yy",
		"dd/MM/yyyy H:mm",
		"dd/MM/yyyy H:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_esMX]     = [
		"dd/MM/yyyy",
		"dddd, d 'de' MMMM 'de' yyyy",
		"d 'de' MMMM 'de' yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM.-yy",
		"dd.MM.yyyy",
		"d MMM. yy",
		"dd.MM.yy",
		"MMMM 'de' yyyy",
		"MMM.-yy",
		"dd/MM/yyyy hh:mm am/pm",
		"dd/MM/yyyy hh:mm:ss am/pm",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_fiFI]     = [
		"d.M.yyyy",
		"dddd d. MMMM yyyy",
		"d. MMMM yyyy",
		"d.M.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/M/yyyy",
		"d. MMM. yy",
		"d/M/yy",
		"MMMM yy",
		"MMM-yy",
		"d.M.yyyy H.mm",
		"d.M.yyyy H.mm.ss",
		"h.mm am/pm",
		"h.mm.ss am/pm",
		"HH.mm",
		"HH.mm.ss"
	];
	c_oAscDateTimeFormat[lcid_frCH]     = [
		"dd/MM/yyyy",
		"dddd, d MMMM yyyy",
		"d MMMM yyyy",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"d-MMM.-yy",
		"dd/MM/yyyy",
		"d MMM. yy",
		"dd/MM/yy",
		"MMMM yy",
		"MMM.-yy",
		"dd/MM/yyyy HH:mm",
		"dd/MM/yyyy HH:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_frFR]     = [
		"dd/MM/yyyy",
		"dddd d MMMM yyyy",
		"d MMMM yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM.-yy",
		"dd.MM.yyyy",
		"d MMM. yy",
		"dd.MM.yy",
		"MMMM yy",
		"MMM.-yy",
		"dd/MM/yyyy HH:mm",
		"dd/MM/yyyy HH:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_enID]     = c_oAscDateTimeFormat[lcid_idID] = [
		"dd/MM/yyyy",
		"dddd, dd MMMM yyyy",
		"dd MMMM yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd.MM.yyyy",
		"dd MMM. yy",
		"dd.MM.yy",
		"MMMM yy",
		"MMM-yy",
		"dd/MM/yyyy HH:mm",
		"dd/MM/yyyy HH:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_itIT]     = c_oAscDateTimeFormat[lcid_itCH] = [
		"dd/MM/yyyy",
		"dddd d MMMM yyyy",
		"d MMMM yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MM.-yy",
		"dd.MM.yyyy",
		"d MMM. yy",
		"MMM. ’yy",
		"MMMM ’yy",
		"MMM-yy",
		"dd/MM/yyyy HH:mm",
		"dd/MM/yyyy HH:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_jaJP]     = [
		"yyyy/MM/dd",
		"yyyy年M月d日",
		"yyyy年M月d日(JDDD)",
		"yyyy年M月",
		"JYYYY年JM月JD日(JDDD)",
		"yyyy/M/d",
		"yy/M/d H時m分",
		"yy/M/d H時m分s秒",
		"AM/PMh時m分",
		"AM/PMh時m分s秒",
		"H時m分",
		"H時m分s秒",
		"yyyy-MM-dd"
	];
	c_oAscDateTimeFormat[lcid_koKR]     = [
		"yyyy-MM-dd",
		"yyyy년 M월 d일 dddd",
		"yyyy년 M월 d일",
		"yyyy/M/d",
		"yyMMdd",
		"yyyy년 M월",
		"yyyy년 M월 d일 AM/PM h시 m분",
		"y년 m월 d일 H시 m분 s초",
		"AM/PM h시 m분",
		"AM/PM h시 m분 s초",
		"H시 m분",
		"H시 m분 s초"
	];
	c_oAscDateTimeFormat[lcid_lvLV]     = [
		"dd.MM.yyyy",
		"dddd, yyyy. 'gada' d. MMMM",
		"yyyy. 'gada' d. MMMM",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"d-MMM.-yy",
		"dd/MM/yyyy",
		"yy. 'gada' d. MMM.",
		"dd/MM/yy",
		"yy MMMM",
		"MMM.-yy",
		"dd.MM.yyyy HH:mm",
		"dd.MM.yyyy HH:mm:ss",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_huHU]     = [
		"yyyy. MM. dd.",
		"yyyy. MMMM d., dddd",
		"yyyy. MMMM d.",
		"yy. MM. dd.",
		"yyyy-MM-dd",
		"yy-MMM.-d",
		"yyyy/MM/dd",
		"yy. MMM. d.",
		"’yy MMM.",
		"’yy MMMM",
		"MMM.-yy",
		"yyyy. MM. dd. H:mm",
		"yyyy. MM. dd. H:mm:ss",
		"am/pm h:mm",
		"am/pm h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_nlNL]     = [
		"d-M-yyyy",
		"dddd d MMMM yyyy",
		"d MMMM yyyy",
		"d-M-yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/M/yyyy",
		"d MMM. yy",
		"d/M/yy",
		"MMMM ’yy",
		"MMM-yy",
		"d-M-yyyy HH:mm",
		"d-M-yyyy HH:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_plPL]     = [
		"dd.MM.yyyy",
		"dddd, d MMMM yyyy",
		"d MMMM yyyy",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd/MM/yyyy",
		"dd MMM yy",
		"dd/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"dd.MM.yyyy HH:mm",
		"dd.MM.yyyy HH:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_ptBR]     = [
		"dd/MM/yyyy",
		"dddd, d 'de' MMMM 'de' yyyy",
		"d 'de' MMMM 'de' yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd.MM.yyyy",
		"d MMM. yy",
		"dd.MM.yy",
		"MMMM 'de' yy",
		"MMM-yy",
		"dd/MM/yyyy HH:mm",
		"dd/MM/yyyy HH:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_ptPT]     = [
		"dd/MM/yyyy",
		"dddd, d 'de' MMMM 'de' yyyy",
		"d 'de' MMMM 'de' yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd.MM.yyyy",
		"d MMM. yy",
		"dd.MM.yy",
		"MMMM 'de' yy",
		"MMM-yy",
		"dd/MM/yyyy HH:mm",
		"dd/MM/yyyy HH:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_ruRU]     = [
		"dd.MM.yyyy",
		"dddd, d MMMM yyyy 'г.'",
		"d MMMM yyyy 'г.'",
		"'«'d'»' MMMM yyyy 'года'",
		"d MMMM yyyy 'года'",
		"d MMMM yyyy",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd/MM/yyyy",
		"d MMM. yy 'г.'",
		"dd/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"dd.MM.yyyy H:mm",
		"dd.MM.yyyy H:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_srCyrlRS]     = [
		"dd.MM.yyyy.",
		"dddd, d. MMMM yyyy.",
		"d. MMMM yyyy.",
		"dd.MM.yy.",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd/MM/yyyy",
		"d. MMM. yy.",
		"dd/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"dd.MM.yyyy. H:mm",
		"dd.MM.yyyy. H:mm:ss",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_srLatnRS]     = [
		"d.M.yyyy.",
		"dddd, dd. MMMM yyyy.",
		"dd. MMMM yyyy.",
		"d.M.yy.",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/M/yyyy",
		"dd. MMM. yy.",
		"d/M/yy",
		"MMMM yy",
		"MMM-yy",
		"d.M.yyyy. HH:mm",
		"d.M.yyyy. HH:mm:ss",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_skSK]     = [
		"d. M. yyyy",
		"dddd d. MMMM yyyy",
		"d. MMMM yyyy",
		"d. M. yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/M/yyyy",
		"d. MMM. yy",
		"d/M/yy",
		"MMMM yy",
		"MMM-yy",
		"d. M. yyyy H:mm",
		"d. M. yyyy H:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_slSI]     = [
		"d. MM. yyyy",
		"dddd, dd. MMMM yyyy",
		"dd. MMMM yyyy",
		"d. MM. yy",
		"yyyy-MM-dd",
		"d-MMM.-yy",
		"d/MM/yyyy",
		"dd. MMM. yy",
		"d/MM/yy",
		"MMMM yy",
		"MMM.-yy",
		"d. MM. yyyy HH:mm",
		"d. MM. yyyy HH:mm:ss",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_srLatnCS]     = [
		"d.MM.yyyy.",
		"dddd, dd. MMMM yyyy.",
		"dd. MMMM yyyy.",
		"d.MM.yyyy.",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/MM/yyyy",
		"dd. MMM. yy.",
		"d/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"d.MM.yyyy. HH:mm",
		"d.MM.yyyy. HH:mm:ss",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_svFI]     = [
		"yyyy-MM-dd",
		"dddd 'den' d MMMM yyyy",
		"d MMMM yyyy",
		"yy-MM-dd",
		"yyMMdd",
		"d MMM. yy",
		"d/M yyyy",
		"d MMM. -yy",
		"M/d/yy",
		"MMMM yyyy",
		"'den' d MMMM yyyy",
		"yyyy-MM-dd HH:mm",
		"yy-MM-dd hh:mm",
		"h.mm am/pm",
		"h.mm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_svSE]     = [
		"yyyy-MM-dd",
		"dddd 'den' d MMMM yyyy",
		"d MMMM yyyy",
		"yy-MM-dd",
		"yyMMdd",
		"d MMM yy",
		"d/M yyyy",
		"d MMM -yy",
		"M/d/yy",
		"MMMM yyyy",
		"'den' d MMMM yyyy",
		"yyyy-MM-dd HH:mm",
		"yy-MM-dd hh:mm",
		"h.mm am/pm",
		"h.mm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_sqAL]     = [
		"d.M.yyyy",
		"dddd, d MMMM yyyy",
		"d MMMM yyyy",
		"d.M.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/M/yyyy",
		"d MMM. yy",
		"d/M/yy",
		"MMMM yy",
		"MMM-yy",
		"d.M.yyyy h:mm am/pm",
		"d.M.yyyy h:mm:ss am/pm",
		"h:mm am/pm",
		"h:mm:ss am/pm",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_trTR]     = [
		"d.MM.yyyy",
		"d MMMM yyyy dddd",
		"d MMMM yyyy",
		"d.MM.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d/MM/yyyy",
		"d MMM. yy",
		"d/MM/yy",
		"MMMM yy",
		"MMM-yy",
		"d.MM.yyyy HH:mm",
		"d.MM.yyyy HH:mm:ss",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_ukUA]     = [
		"dd.MM.yyyy",
		"dddd, d MMMM yyyy 'р.'",
		"d MMMM yyyy 'р.'",
		"dd.MM.yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"dd/MM/yyyy",
		"d MMM. yy 'р.'",
		"dd/MM/yy",
		"MMMM yy 'р'",
		"MMM-yy",
		"dd.MM.yyyy H:mm",
		"dd.MM.yyyy H:mm:ss",
		"h:mm",
		"h:mm:ss",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_viVN]     = [
		"dd/MM/yyyy",
		"dddd, dd MMMM yyyy",
		"dd MMMM yyyy",
		"dd/MM/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"d-MMMM-yy",
		"dd MMM. yy",
		"dd.MM.yy",
		"MMMM yy",
		"MMM-yy",
		"dd/MM/yyyy h:mm AM/PM",
		"dd/MM/yyyy h:mm:ss AM/PM",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];
	c_oAscDateTimeFormat[lcid_zhCN]     = c_oAscDateTimeFormat[lcid_zhTW] = [
		"yyyy/M/d",
		"yyyy'年'M'月'd'日'",
		"yyyy年M月d日星期W",
		"yy.M.d",
		"yyyy年M月",
		"h时m分s秒",
		"h时m分",
		"AM/PMh时m分",
		"YYYY年M月d日",
		"YYYY年M月d日星期W",
		"YYYY年M月"
	];
	c_oAscDateTimeFormat[lcid_enUS]     = [
		"M/d/yyyy",
		"dddd, MMMM d, yyyy",
		"MMMM d, yyyy",
		"M/d/yy",
		"yyyy-MM-dd",
		"d-MMM-yy",
		"M.d.yyyy",
		"MMM. d, yy",
		"d MMMM yyyy",
		"MMMM yy",
		"MMM-yy",
		"M/d/yyyy h:mm AM/PM",
		"M/d/yyyy h:mm:ss AM/PM",
		"h:mm AM/PM",
		"h:mm:ss AM/PM",
		"HH:mm",
		"HH:mm:ss"
	];


	//------------------------------------------------------------export--------------------------------------------------
	var prot;
	window['Asc']['FONT_THUMBNAIL_HEIGHT'] = FONT_THUMBNAIL_HEIGHT;
	window['Asc']['c_oAscMaxColumnWidth']  = window['Asc'].c_oAscMaxColumnWidth = c_oAscMaxColumnWidth;
	window['Asc']['c_oAscMaxRowHeight'] = window['Asc'].c_oAscMaxRowHeight = c_oAscMaxRowHeight;
    window['Asc']['c_nMaxConversionTime'] = window['Asc'].c_nMaxConversionTime = c_nMaxConversionTime;
	window['Asc']['c_nCommonRequestTime'] = window['Asc'].c_nCommonRequestTime = c_nCommonRequestTime;
	window['Asc']['c_nMaxDownloadTitleLen'] = window['Asc'].c_nMaxDownloadTitleLen = c_nMaxDownloadTitleLen;
	window['Asc']['c_nVersionNoBase64'] = window['Asc'].c_nVersionNoBase64 = c_nVersionNoBase64;
	window['Asc']['c_dMaxParaRunContentLength'] = window['Asc'].c_dMaxParaRunContentLength = c_dMaxParaRunContentLength;
	window['Asc']['c_nMaxHyperlinkLength'] = window['Asc'].c_nMaxHyperlinkLength = c_nMaxHyperlinkLength;
	window['Asc']['c_sNativeViewerFormats'] = window['Asc'].c_sNativeViewerFormats = c_sNativeViewerFormats;
	window['Asc']['c_sShardKeyName'] = window['Asc'].c_sShardKeyName = c_sShardKeyName;
	window['Asc']['c_sWopiSrcName'] = window['Asc'].c_sWopiSrcName = c_sWopiSrcName;
	window['Asc']['c_sUserSessionIdName'] = window['Asc'].c_sUserSessionIdName = c_sUserSessionIdName;
	window['Asc']['c_oAscFileType'] = window['Asc'].c_oAscFileType = c_oAscFileType;
	window['Asc'].g_oLcidNameToIdMap = g_oLcidNameToIdMap;
	window['Asc'].availableIdeographLanguages = availableIdeographLanguages;
	window['Asc'].availableBidiLanguages = availableBidiLanguages;
	window['Asc'].g_oLcidIdToNameMap = g_oLcidIdToNameMap;
	prot                         = c_oAscFileType;
	prot['UNKNOWN']              = prot.UNKNOWN;
	prot['PDF']                  = prot.PDF;
	prot['PDFA']                 = prot.PDFA;
	prot['DJVU']                 = prot.DJVU;
	prot['XPS']                  = prot.XPS;
	prot['HTML']                 = prot.HTML;
	prot['DOCX']                 = prot.DOCX;
	prot['DOC']                  = prot.DOC;
	prot['ODT']                  = prot.ODT;
	prot['RTF']                  = prot.RTF;
	prot['TXT']                  = prot.TXT;
	prot['MHT']                  = prot.MHT;
	prot['EPUB']                 = prot.EPUB;
	prot['FB2']                  = prot.FB2;
	prot['MOBI']                 = prot.MOBI;
	prot['DOCM']                 = prot.DOCM;
	prot['DOTX']                 = prot.DOTX;
	prot['DOTM']                 = prot.DOTM;
	prot['FODT']                 = prot.FODT;
	prot['OTT']                  = prot.OTT;
	prot['DOC_FLAT']             = prot.DOC_FLAT;
	prot['DOCX_FLAT']            = prot.DOCX_FLAT;
	prot['HTML_IN_CONTAINER']    = prot.HTML_IN_CONTAINER;
	prot['DOCX_PACKAGE']         = prot.DOCX_PACKAGE;
	prot['OFORM']                = prot.OFORM;
	prot['DOCXF']                = prot.DOCXF;
	prot['DOCY']                 = prot.DOCY;
	prot['JSON']                 = prot.JSON;
	prot['XLSX']                 = prot.XLSX;
	prot['XLS']                  = prot.XLS;
	prot['ODS']                  = prot.ODS;
	prot['CSV']                  = prot.CSV;
	prot['XLSM']                 = prot.XLSM;
	prot['XLTX']                 = prot.XLTX;
	prot['XLTM']                 = prot.XLTM;
	prot['XLSB']                 = prot.XLSB;
	prot['FODS']                 = prot.FODS;
	prot['OTS']                  = prot.OTS;
	prot['XLSX_FLAT']            = prot.XLSX_FLAT;
	prot['XLSX_PACKAGE']         = prot.XLSX_PACKAGE;
	prot['XLSY']                 = prot.XLSY;
	prot['CANVAS_SPREADSHEET']   = prot.CANVAS_SPREADSHEET;
	prot['PPTX']                 = prot.PPTX;
	prot['PPT']                  = prot.PPT;
	prot['ODP']                  = prot.ODP;
	prot['PPSX']                 = prot.PPSX;
	prot['PPTM']                 = prot.PPTM;
	prot['PPSM']                 = prot.PPSM;
	prot['POTX']                 = prot.POTX;
	prot['POTM']                 = prot.POTM;
	prot['FODP']                 = prot.FODP;
	prot['OTP']                  = prot.OTP;
	prot['PPTX_PACKAGE']         = prot.PPTX_PACKAGE;
	prot['CANVAS_PRESENTATION']  = prot.CANVAS_PRESENTATION;

	prot['VSDX']        		 = prot.VSDX;
	prot['VSSX']        		 = prot.VSSX;
	prot['VSTX']        		 = prot.VSTX;
	prot['VSDM']        		 = prot.VSDM;
	prot['VSTM']        		 = prot.VSTM;

	prot['JPG']                  = prot.JPG;
	prot['TIFF']                 = prot.TIFF;
	prot['TGA']                  = prot.TGA;
	prot['GIF']                  = prot.GIF;
	prot['PNG']                  = prot.PNG;
	prot['EMF']                  = prot.EMF;
	prot['WMF']                  = prot.WMF;
	prot['BMP']                  = prot.BMP;
	prot['CR2']                  = prot.CR2;
	prot['PCX']                  = prot.PCX;
	prot['RAS']                  = prot.RAS;
	prot['PSD']                  = prot.PSD;
	prot['ICO']                  = prot.ICO;

	window['Asc']['c_oAscTextAssociation'] = window['Asc'].c_oAscTextAssociation = c_oAscTextAssociation;
	prot = c_oAscTextAssociation;
	prot['BlockChar'] = prot.BlockChar;
	prot['BlockLine'] = prot.BlockLine;
	prot['PlainLine'] = prot.PlainLine;
	prot['PlainParagraph'] = prot.PlainParagraph;


	window['Asc']['c_oAscAsyncAction']       = window['Asc'].c_oAscAsyncAction = c_oAscAsyncAction;
	prot                                     = c_oAscAsyncAction;
	prot['Open']                             = prot.Open;
	prot['Save']                             = prot.Save;
	prot['LoadDocumentFonts']                = prot.LoadDocumentFonts;
	prot['LoadDocumentImages']               = prot.LoadDocumentImages;
	prot['LoadFont']                         = prot.LoadFont;
	prot['LoadImage']                        = prot.LoadImage;
	prot['DownloadAs']                       = prot.DownloadAs;
	prot['Print']                            = prot.Print;
	prot['UploadImage']                      = prot.UploadImage;
	prot['ApplyChanges']                     = prot.ApplyChanges;
	prot['SlowOperation']                    = prot.SlowOperation;
	prot['LoadTheme']                        = prot.LoadTheme;
	prot['MailMergeLoadFile']                = prot.MailMergeLoadFile;
	prot['DownloadMerge']                    = prot.DownloadMerge;
	prot['SendMailMerge']                    = prot.SendMailMerge;
	prot['ForceSaveButton']                  = prot.ForceSaveButton;
	prot['ForceSaveTimeout']                 = prot.ForceSaveTimeout;
	prot['Waiting']                          = prot.Waiting;
	prot['Submit']                           = prot.Submit;
	prot['Disconnect']                       = prot.Disconnect;
	prot['RefreshFile']                      = prot.RefreshFile;
	window['Asc']['c_oAscAdvancedOptionsID'] = window['Asc'].c_oAscAdvancedOptionsID = c_oAscAdvancedOptionsID;
	prot                                         = c_oAscAdvancedOptionsID;
	prot['CSV']                                  = prot.CSV;
	prot['TXT']                                  = prot.TXT;
	prot['DRM']                                  = prot.DRM;
	window['Asc']['c_oAscFontRenderingModeType'] = window['Asc'].c_oAscFontRenderingModeType = c_oAscFontRenderingModeType;
	prot                                   = c_oAscFontRenderingModeType;
	prot['noHinting']                      = prot.noHinting;
	prot['hinting']                        = prot.hinting;
	prot['hintingAndSubpixeling']          = prot.hintingAndSubpixeling;
	window['Asc']['c_oAscAsyncActionType'] = window['Asc'].c_oAscAsyncActionType = c_oAscAsyncActionType;
	prot                                 = c_oAscAsyncActionType;
	prot['Information']                  = prot.Information;
	prot['BlockInteraction']             = prot.BlockInteraction;
	window['Asc']['c_oAscNumFormatType'] = window['Asc'].c_oAscNumFormatType = c_oAscNumFormatType;
	prot                                     = c_oAscNumFormatType;
	prot['None']                             = prot.None;
	prot['General']                          = prot.General;
	prot['Number']                           = prot.Number;
	prot['Scientific']                       = prot.Scientific;
	prot['Accounting']                       = prot.Accounting;
	prot['Currency']                         = prot.Currency;
	prot['Date']                             = prot.Date;
	prot['Time']                             = prot.Time;
	prot['Percent']                          = prot.Percent;
	prot['Fraction']                         = prot.Fraction;
	prot['Text']                             = prot.Text;
	prot['Custom']                           = prot.Custom;
	window['Asc']['c_oAscDrawingLayerType']  = window['Asc'].c_oAscDrawingLayerType  = c_oAscDrawingLayerType;
	prot                                     = c_oAscDrawingLayerType;
	prot['BringToFront']                     = prot.BringToFront;
	prot['SendToBack']                       = prot.SendToBack;
	prot['BringForward']                     = prot.BringForward;
	prot['SendBackward']                     = prot.SendBackward;
	window['Asc']['c_oAscTypeSelectElement'] = window['Asc'].c_oAscTypeSelectElement = c_oAscTypeSelectElement;
	
	prot                      = c_oAscTypeSelectElement;
	prot['Paragraph']         = prot.Paragraph;
	prot['Table']             = prot.Table;
	prot['Image']             = prot.Image;
	prot['Header']            = prot.Header;
	prot['Hyperlink']         = prot.Hyperlink;
	prot['SpellCheck']        = prot.SpellCheck;
	prot['Shape']             = prot.Shape;
	prot['Slide']             = prot.Slide;
	prot['Chart']             = prot.Chart;
	prot['Math']              = prot.Math;
	prot['MailMerge']         = prot.MailMerge;
	prot['ContentControl']    = prot.ContentControl;
	prot['Animation']         = prot.Animation;
	prot['UnProtectedRegion'] = prot.UnProtectedRegion;
	prot['Text']              = prot.Text;
	prot['Annot']             = prot.Annot;
	prot['PdfPage']           = prot.PdfPage;
	
	window['Asc']['linerule_AtLeast'] = window['Asc'].linerule_AtLeast = linerule_AtLeast;
	window['Asc']['linerule_Auto'] = window['Asc'].linerule_Auto = linerule_Auto;
	window['Asc']['linerule_Exact'] = window['Asc'].linerule_Exact = linerule_Exact;
	window['Asc']['c_oAscShdClear'] = window['Asc'].c_oAscShdClear = c_oAscShdClear;
	window['Asc']['c_oAscShdNil'] = window['Asc'].c_oAscShdNil = c_oAscShdNil;
	window['Asc']['c_oAscShd']    = window['Asc'].c_oAscShd = c_oAscShd;
	prot                          = c_oAscShd;
	prot['Clear'] = prot.Clear;
	prot['Nil'] = prot.Nil;
	prot['DiagCross'] = prot.DiagCross;
	prot['DiagStripe'] = prot.DiagStripe;
	prot['HorzCross'] = prot.HorzCross;
	prot['HorzStripe'] = prot.HorzStripe;
	prot['Pct10'] = prot.Pct10;
	prot['Pct12'] = prot.Pct12;
	prot['Pct15'] = prot.Pct15;
	prot['Pct20'] = prot.Pct20;
	prot['Pct25'] = prot.Pct25;
	prot['Pct30'] = prot.Pct30;
	prot['Pct35'] = prot.Pct35;
	prot['Pct37'] = prot.Pct37;
	prot['Pct40'] = prot.Pct40;
	prot['Pct45'] = prot.Pct45;
	prot['Pct5'] = prot.Pct5;
	prot['Pct50'] = prot.Pct50;
	prot['Pct55'] = prot.Pct55;
	prot['Pct60'] = prot.Pct60;
	prot['Pct62'] = prot.Pct62;
	prot['Pct65'] = prot.Pct65;
	prot['Pct70'] = prot.Pct70;
	prot['Pct75'] = prot.Pct75;
	prot['Pct80'] = prot.Pct80;
	prot['Pct85'] = prot.Pct85;
	prot['Pct87'] = prot.Pct87;
	prot['Pct90'] = prot.Pct90;
	prot['Pct95'] = prot.Pct95;
	prot['ReverseDiagStripe'] = prot.ReverseDiagStripe;
	prot['Solid'] = prot.Solid;
	prot['ThinDiagCross'] = prot.ThinDiagCross;
	prot['ThinDiagStripe'] = prot.ThinDiagStripe;
	prot['ThinHorzCross'] = prot.ThinHorzCross;
	prot['ThinHorzStripe'] = prot.ThinHorzStripe;
	prot['ThinReverseDiagStripe'] = prot.ThinReverseDiagStripe;
	prot['ThinVertStripe'] = prot.ThinVertStripe;
	prot['VertStripe'] = prot.VertStripe;
	window['Asc']['c_oAscDropCap'] = window['Asc'].c_oAscDropCap = c_oAscDropCap;
	prot                                          = c_oAscDropCap;
	prot['None']                                  = prot.None;
	prot['Drop']                                  = prot.Drop;
	prot['Margin']                                = prot.Margin;
	window['Asc']['c_oAscChartTitleShowSettings'] = window['Asc'].c_oAscChartTitleShowSettings = c_oAscChartTitleShowSettings;
	prot                                                 = c_oAscChartTitleShowSettings;
	prot['none']                                         = prot.none;
	prot['overlay']                                      = prot.overlay;
	prot['noOverlay']                                    = prot.noOverlay;
	window['Asc']['c_oAscChartHorAxisLabelShowSettings'] = window['Asc'].c_oAscChartHorAxisLabelShowSettings = c_oAscChartHorAxisLabelShowSettings;
	prot                                                  = c_oAscChartHorAxisLabelShowSettings;
	prot['none']                                          = prot.none;
	prot['noOverlay']                                     = prot.noOverlay;
	window['Asc']['c_oAscChartVertAxisLabelShowSettings'] = window['Asc'].c_oAscChartVertAxisLabelShowSettings = c_oAscChartVertAxisLabelShowSettings;
	prot                                           = c_oAscChartVertAxisLabelShowSettings;
	prot['none']                                   = prot.none;
	prot['rotated']                                = prot.rotated;
	prot['vertical']                               = prot.vertical;
	prot['horizontal']                             = prot.horizontal;
	window['Asc']['c_oAscChartLegendShowSettings'] = window['Asc'].c_oAscChartLegendShowSettings = c_oAscChartLegendShowSettings;
	prot                                      = c_oAscChartLegendShowSettings;
	prot['none']                              = prot.none;
	prot['left']                              = prot.left;
	prot['top']                               = prot.top;
	prot['right']                             = prot.right;
	prot['bottom']                            = prot.bottom;
	prot['leftOverlay']                       = prot.leftOverlay;
	prot['rightOverlay']                      = prot.rightOverlay;
	prot['layout']                            = prot.layout;
	prot['topRight']                          = prot.topRight;
	window['Asc']['c_oAscChartDataLabelsPos'] = window['Asc'].c_oAscChartDataLabelsPos = c_oAscChartDataLabelsPos;
	prot                                     = c_oAscChartDataLabelsPos;
	prot['none']                             = prot.none;
	prot['b']                                = prot.b;
	prot['bestFit']                          = prot.bestFit;
	prot['ctr']                              = prot.ctr;
	prot['inBase']                           = prot.inBase;
	prot['inEnd']                            = prot.inEnd;
	prot['l']                                = prot.l;
	prot['outEnd']                           = prot.outEnd;
	prot['r']                                = prot.r;
	prot['t']                                = prot.t;
	prot['show']                             = prot.show;
	window['Asc']['c_oAscGridLinesSettings'] = window['Asc'].c_oAscGridLinesSettings = c_oAscGridLinesSettings;
	prot                                     = c_oAscGridLinesSettings;
	prot['none']                             = prot.none;
	prot['major']                            = prot.major;
	prot['minor']                            = prot.minor;
	prot['majorMinor']                       = prot.majorMinor;
	window['Asc']['c_oAscChartTypeSettings'] = window['Asc'].c_oAscChartTypeSettings = c_oAscChartTypeSettings;
	prot                               = c_oAscChartTypeSettings;
	prot['barNormal']                  = prot.barNormal;
	prot['barStacked']                 = prot.barStacked;
	prot['barStackedPer']              = prot.barStackedPer;
	prot['barNormal3d']                = prot.barNormal3d;
	prot['barStacked3d']               = prot.barStacked3d;
	prot['barStackedPer3d']            = prot.barStackedPer3d;
	prot['barNormal3dPerspective']     = prot.barNormal3dPerspective;
	prot['lineNormal']                 = prot.lineNormal;
	prot['lineStacked']                = prot.lineStacked;
	prot['lineStackedPer']             = prot.lineStackedPer;
	prot['lineNormalMarker']           = prot.lineNormalMarker;
	prot['lineStackedMarker']          = prot.lineStackedMarker;
	prot['lineStackedPerMarker']       = prot.lineStackedPerMarker;
	prot['line3d']                     = prot.line3d;
	prot['pie']                        = prot.pie;
	prot['pie3d']                      = prot.pie3d;
	prot['hBarNormal']                 = prot.hBarNormal;
	prot['hBarStacked']                = prot.hBarStacked;
	prot['hBarStackedPer']             = prot.hBarStackedPer;
	prot['hBarNormal3d']               = prot.hBarNormal3d;
	prot['hBarStacked3d']              = prot.hBarStacked3d;
	prot['hBarStackedPer3d']           = prot.hBarStackedPer3d;
	prot['areaNormal']                 = prot.areaNormal;
	prot['areaStacked']                = prot.areaStacked;
	prot['areaStackedPer']             = prot.areaStackedPer;
	prot['doughnut']                   = prot.doughnut;
	prot['stock']                      = prot.stock;
	prot['scatter']                    = prot.scatter;
	prot['scatterLine']                = prot.scatterLine;
	prot['scatterLineMarker']          = prot.scatterLineMarker;
	prot['scatterMarker']              = prot.scatterMarker;
	prot['scatterNone']                = prot.scatterNone;
	prot['scatterSmooth']              = prot.scatterSmooth;
	prot['scatterSmoothMarker']        = prot.scatterSmoothMarker;
	prot['surfaceNormal']              = prot.surfaceNormal;
	prot['surfaceWireframe']           = prot.surfaceWireframe;
	prot['contourNormal']              = prot.contourNormal;
	prot['contourWireframe']           = prot.contourWireframe;
	prot['comboCustom']                = prot.comboCustom;
	prot['comboBarLine']               = prot.comboBarLine;
	prot['comboBarLineSecondary']      = prot.comboBarLineSecondary;
	prot['comboAreaBar']               = prot.comboAreaBar;
	prot['radar']                      = prot.radar;
	prot['radarMarker']                = prot.radarMarker;
	prot['radarFilled']                = prot.radarFilled;
	prot['unknown']                    = prot.unknown;

	window['Asc']['c_oAscValAxisRule'] = window['Asc'].c_oAscValAxisRule = c_oAscValAxisRule;
	prot                              = c_oAscValAxisRule;
	prot['auto']                      = prot.auto;
	prot['fixed']                     = prot.fixed;
	window['Asc']['c_oAscValAxUnits'] = window['Asc'].c_oAscValAxUnits = c_oAscValAxUnits;
	prot                            = c_oAscValAxUnits;
	prot['none']                    = prot.none;
	prot['BILLIONS']                = prot.BILLIONS;
	prot['HUNDRED_MILLIONS']        = prot.HUNDRED_MILLIONS;
	prot['HUNDREDS']                = prot.HUNDREDS;
	prot['HUNDRED_THOUSANDS']       = prot.HUNDRED_THOUSANDS;
	prot['MILLIONS']                = prot.MILLIONS;
	prot['TEN_MILLIONS']            = prot.TEN_MILLIONS;
	prot['TEN_THOUSANDS']           = prot.TEN_THOUSANDS;
	prot['TRILLIONS']               = prot.TRILLIONS;
	prot['CUSTOM']                  = prot.CUSTOM;
	prot['THOUSANDS']               = prot.THOUSANDS;
	window['Asc']['c_oAscTickMark'] = window['Asc'].c_oAscTickMark = c_oAscTickMark;
	prot                                 = c_oAscTickMark;
	prot['TICK_MARK_CROSS']              = prot.TICK_MARK_CROSS;
	prot['TICK_MARK_IN']                 = prot.TICK_MARK_IN;
	prot['TICK_MARK_NONE']               = prot.TICK_MARK_NONE;
	prot['TICK_MARK_OUT']                = prot.TICK_MARK_OUT;
	window['Asc']['c_oAscTickLabelsPos'] = window['Asc'].c_oAscTickLabelsPos = c_oAscTickLabelsPos;
	prot                                = c_oAscTickLabelsPos;
	prot['TICK_LABEL_POSITION_HIGH']    = prot.TICK_LABEL_POSITION_HIGH;
	prot['TICK_LABEL_POSITION_LOW']     = prot.TICK_LABEL_POSITION_LOW;
	prot['TICK_LABEL_POSITION_NEXT_TO'] = prot.TICK_LABEL_POSITION_NEXT_TO;
	prot['TICK_LABEL_POSITION_NONE']    = prot.TICK_LABEL_POSITION_NONE;
	window['Asc']['c_oAscCrossesRule']  = window['Asc'].c_oAscCrossesRule = c_oAscCrossesRule;
	prot                                     = c_oAscCrossesRule;
	prot['auto']                             = prot.auto;
	prot['maxValue']                         = prot.maxValue;
	prot['value']                            = prot.value;
	prot['minValue']                         = prot.minValue;
	window['Asc']['c_oAscBetweenLabelsRule'] = window['Asc'].c_oAscBetweenLabelsRule = c_oAscBetweenLabelsRule;
	prot                                  = c_oAscBetweenLabelsRule;
	prot['auto']                          = prot.auto;
	prot['manual']                        = prot.manual;
	window['Asc']['c_oAscLabelsPosition'] = window['Asc'].c_oAscLabelsPosition = c_oAscLabelsPosition;
	prot                            = c_oAscLabelsPosition;
	prot['byDivisions']             = prot.byDivisions;
	prot['betweenDivisions']        = prot.betweenDivisions;
	window['Asc']['c_oAscAxisType'] = window['Asc'].c_oAscAxisType = c_oAscAxisType;
	prot                           = c_oAscAxisType;
	prot['auto']                   = prot.auto;
	prot['date']                   = prot.date;
	prot['text']                   = prot.text;
	prot['cat']                    = prot.cat;
	prot['val']                    = prot.val;
	prot['ser']                    = prot.ser;
	window['Asc']['c_oAscHAnchor'] = window['Asc'].c_oAscHAnchor = c_oAscHAnchor;
	prot                          = c_oAscHAnchor;
	prot['Margin']                = prot.Margin;
	prot['Page']                  = prot.Page;
	prot['Text']                  = prot.Text;
	window['Asc']['c_oAscXAlign'] = window['Asc'].c_oAscXAlign = c_oAscXAlign;
	prot                          = c_oAscXAlign;
	prot['Center']                = prot.Center;
	prot['Inside']                = prot.Inside;
	prot['Left']                  = prot.Left;
	prot['Outside']               = prot.Outside;
	prot['Right']                 = prot.Right;
	window['Asc']['c_oAscYAlign'] = window['Asc'].c_oAscYAlign = c_oAscYAlign;
	prot                           = c_oAscYAlign;
	prot['Bottom']                 = prot.Bottom;
	prot['Center']                 = prot.Center;
	prot['Inline']                 = prot.Inline;
	prot['Inside']                 = prot.Inside;
	prot['Outside']                = prot.Outside;
	prot['Top']                    = prot.Top;
	window['Asc']['c_oAscVAnchor'] = window['Asc'].c_oAscVAnchor = c_oAscVAnchor;
	prot                                 = c_oAscVAnchor;
	prot['Margin']                       = prot.Margin;
	prot['Page']                         = prot.Page;
	prot['Text']                         = prot.Text;
	window['Asc']['c_oAscRelativeFromH'] = window['Asc'].c_oAscRelativeFromH = c_oAscRelativeFromH;
	prot                                 = c_oAscRelativeFromH;
	prot['Character']                    = prot.Character;
	prot['Column']                       = prot.Column;
	prot['InsideMargin']                 = prot.InsideMargin;
	prot['LeftMargin']                   = prot.LeftMargin;
	prot['Margin']                       = prot.Margin;
	prot['OutsideMargin']                = prot.OutsideMargin;
	prot['Page']                         = prot.Page;
	prot['RightMargin']                  = prot.RightMargin;
	window['Asc']['c_oAscRelativeFromV'] = window['Asc'].c_oAscRelativeFromV = c_oAscRelativeFromV;
	prot                                   = c_oAscRelativeFromV;
	prot['BottomMargin']                   = prot.BottomMargin;
	prot['InsideMargin']                   = prot.InsideMargin;
	prot['Line']                           = prot.Line;
	prot['Margin']                         = prot.Margin;
	prot['OutsideMargin']                  = prot.OutsideMargin;
	prot['Page']                           = prot.Page;
	prot['Paragraph']                      = prot.Paragraph;
	prot['TopMargin']                      = prot.TopMargin;
	window['Asc']['c_oAscBorderStyles'] = window['Asc'].c_oAscBorderStyles = c_oAscBorderStyles;
	prot                         = c_oAscBorderStyles;
	prot['None']                 = prot.None;
	prot['Double']               = prot.Double;
	prot['Hair']                 = prot.Hair;
	prot['DashDotDot']           = prot.DashDotDot;
	prot['DashDot']              = prot.DashDot;
	prot['Dotted']               = prot.Dotted;
	prot['Dashed']               = prot.Dashed;
	prot['Thin']                 = prot.Thin;
	prot['MediumDashDotDot']     = prot.MediumDashDotDot;
	prot['SlantDashDot']         = prot.SlantDashDot;
	prot['MediumDashDot']        = prot.MediumDashDot;
	prot['MediumDashed']         = prot.MediumDashed;
	prot['Medium']               = prot.Medium;
	prot['Thick']                = prot.Thick;
	window['Asc']['c_oAscPageOrientation'] = window['Asc'].c_oAscPageOrientation = c_oAscPageOrientation;
	prot                         = c_oAscPageOrientation;
	prot['PagePortrait']         = prot.PagePortrait;
	prot['PageLandscape']        = prot.PageLandscape;
	window['Asc']['c_oAscColor'] = window['Asc'].c_oAscColor = c_oAscColor;
	prot                        = c_oAscColor;
	prot['COLOR_TYPE_NONE']     = prot.COLOR_TYPE_NONE;
	prot['COLOR_TYPE_SRGB']     = prot.COLOR_TYPE_SRGB;
	prot['COLOR_TYPE_PRST']     = prot.COLOR_TYPE_PRST;
	prot['COLOR_TYPE_SCHEME']   = prot.COLOR_TYPE_SCHEME;
	prot['COLOR_TYPE_SYS']      = prot.COLOR_TYPE_SYS;
	prot['COLOR_TYPE_STYLE']      = prot.COLOR_TYPE_STYLE;
	window['Asc']['c_oAscFill'] = window['Asc'].c_oAscFill = c_oAscFill;
	prot                                = c_oAscFill;
	prot['FILL_TYPE_NONE']              = prot.FILL_TYPE_NONE;
	prot['FILL_TYPE_BLIP']              = prot.FILL_TYPE_BLIP;
	prot['FILL_TYPE_NOFILL']            = prot.FILL_TYPE_NOFILL;
	prot['FILL_TYPE_SOLID']             = prot.FILL_TYPE_SOLID;
	prot['FILL_TYPE_GRAD']              = prot.FILL_TYPE_GRAD;
	prot['FILL_TYPE_PATT']              = prot.FILL_TYPE_PATT;
	prot['FILL_TYPE_GRP']               = prot.FILL_TYPE_GRP;
	window['Asc']['c_oAscFillGradType'] = window['Asc'].c_oAscFillGradType = c_oAscFillGradType;
	prot                                = c_oAscFillGradType;
	prot['GRAD_LINEAR']                 = prot.GRAD_LINEAR;
	prot['GRAD_PATH']                   = prot.GRAD_PATH;
	window['Asc']['c_oAscFillBlipType'] = window['Asc'].c_oAscFillBlipType = c_oAscFillBlipType;
	prot                              = c_oAscFillBlipType;
	prot['STRETCH']                   = prot.STRETCH;
	prot['TILE']                      = prot.TILE;
	window['Asc']['c_oAscStrokeType'] = window['Asc'].c_oAscStrokeType = c_oAscStrokeType;
	prot                                     = c_oAscStrokeType;
	prot['STROKE_NONE']                      = prot.STROKE_NONE;
	prot['STROKE_COLOR']                     = prot.STROKE_COLOR;
	window['Asc']['c_oAscVAlign'] = window['Asc'].c_oAscVAlign = c_oAscVAlign;
	prot                          = c_oAscVAlign;
	prot['Bottom']                = prot.Bottom;
	prot['Center']                = prot.Center;
	prot['Dist']                  = prot.Dist;
	prot['Just']                  = prot.Just;
	prot['Top']                   = prot.Top;
	window['Asc']['c_oAscVertDrawingText']   = c_oAscVertDrawingText;
	prot                                     = c_oAscVertDrawingText;
	prot['normal']                           = prot.normal;
	prot['vert']                             = prot.vert;
	prot['vert270']                          = prot.vert270;
	window['Asc']['c_oAscLineJoinType']      = c_oAscLineJoinType;
	prot                                     = c_oAscLineJoinType;
	prot['Round']                            = prot.Round;
	prot['Bevel']                            = prot.Bevel;
	prot['Miter']                            = prot.Miter;
	window['Asc']['c_oAscLineCapType']       = c_oAscLineCapType;
	prot                                     = c_oAscLineCapType;
	prot['Flat']                             = prot.Flat;
	prot['Round']                            = prot.Round;
	prot['Square']                           = prot.Square;
	window['Asc']['c_oAscLineBeginType']     = c_oAscLineBeginType;
	prot                                     = c_oAscLineBeginType;
	prot['None']                             = prot.None;
	prot['Arrow']                            = prot.Arrow;
	prot['Diamond']                          = prot.Diamond;
	prot['Oval']                             = prot.Oval;
	prot['Stealth']                          = prot.Stealth;
	prot['Triangle']                         = prot.Triangle;
	window['Asc']['c_oAscLineBeginSize']     = c_oAscLineBeginSize;
	prot                                     = c_oAscLineBeginSize;
	prot['small_small']                      = prot.small_small;
	prot['small_mid']                        = prot.small_mid;
	prot['small_large']                      = prot.small_large;
	prot['mid_small']                        = prot.mid_small;
	prot['mid_mid']                          = prot.mid_mid;
	prot['mid_large']                        = prot.mid_large;
	prot['large_small']                      = prot.large_small;
	prot['large_mid']                        = prot.large_mid;
	prot['large_large']                      = prot.large_large;
	window['Asc']['c_oAscCellTextDirection'] = window['Asc'].c_oAscCellTextDirection = c_oAscCellTextDirection;
	prot                                 = c_oAscCellTextDirection;
	prot['LRTB']                         = prot.LRTB;
	prot['TBRL']                         = prot.TBRL;
	prot['BTLR']                         = prot.BTLR;
	window['Asc']['c_oAscDocumentUnits'] = window['Asc'].c_oAscDocumentUnits = c_oAscDocumentUnits;
	prot                                    = c_oAscDocumentUnits;
	prot['Millimeter']                      = prot.Millimeter;
	prot['Inch']                            = prot.Inch;
	prot['Point']                           = prot.Point;

	window['Asc']['c_oAscMouseMoveDataTypes'] = window['Asc'].c_oAscMouseMoveDataTypes = c_oAscMouseMoveDataTypes;

	prot                 = c_oAscMouseMoveDataTypes;
	prot['Common']       = prot.Common;
	prot['Hyperlink']    = prot.Hyperlink;
	prot['LockedObject'] = prot.LockedObject;
	prot['Footnote']     = prot.Footnote;
	prot['Form']         = prot.Form;
	prot['Review']       = prot.Review;
	prot['Eyedropper']   = prot.Eyedropper;
	prot['Placeholder']   = prot.Placeholder;
	prot['EffectInfo']   = prot.EffectInfo;

	window['Asc']['c_oAscMaxTooltipLength'] = window['Asc'].c_oAscMaxTooltipLength = c_oAscMaxTooltipLength;
	window['Asc']['c_oAscMaxCellOrCommentLength'] = window['Asc'].c_oAscMaxCellOrCommentLength = c_oAscMaxCellOrCommentLength;
	window['Asc']['c_oAscMaxHeaderFooterLength']  = window['Asc'].c_oAscMaxHeaderFooterLength  = c_oAscMaxHeaderFooterLength;
	window['Asc']['c_oAscMaxFilterListLength']    = window['Asc'].c_oAscMaxFilterListLength  = c_oAscMaxFilterListLength;
	window['Asc']['c_oAscSelectionType'] = window['Asc'].c_oAscSelectionType = c_oAscSelectionType;
	prot                                 = c_oAscSelectionType;
	prot['RangeCells']                   = prot.RangeCells;
	prot['RangeCol']                     = prot.RangeCol;
	prot['RangeRow']                     = prot.RangeRow;
	prot['RangeMax']                     = prot.RangeMax;
	prot['RangeImage']                   = prot.RangeImage;
	prot['RangeChart']                   = prot.RangeChart;
	prot['RangeShape']                   = prot.RangeShape;
	prot['RangeSlicer']                  = prot.RangeSlicer;
	prot['RangeShapeText']               = prot.RangeShapeText;
	prot['RangeChartText']               = prot.RangeChartText;
	prot['RangeFrozen']                  = prot.RangeFrozen;
	window['Asc']['c_oAscInsertOptions'] = window['Asc'].c_oAscInsertOptions = c_oAscInsertOptions;
	prot                                 = c_oAscInsertOptions;
	prot['InsertCellsAndShiftRight']     = prot.InsertCellsAndShiftRight;
	prot['InsertCellsAndShiftDown']      = prot.InsertCellsAndShiftDown;
	prot['InsertColumns']                = prot.InsertColumns;
	prot['InsertRows']                   = prot.InsertRows;
	prot['InsertTableRowAbove']          = prot.InsertTableRowAbove;
	prot['InsertTableRowBelow']          = prot.InsertTableRowBelow;
	prot['InsertTableColLeft']           = prot.InsertTableColLeft;
	prot['InsertTableColRight']          = prot.InsertTableColRight;
	window['Asc']['c_oAscDeleteOptions'] = window['Asc'].c_oAscDeleteOptions = c_oAscDeleteOptions;
	prot                            = c_oAscDeleteOptions;
	prot['DeleteCellsAndShiftLeft'] = prot.DeleteCellsAndShiftLeft;
	prot['DeleteCellsAndShiftTop']  = prot.DeleteCellsAndShiftTop;
	prot['DeleteColumns']           = prot.DeleteColumns;
	prot['DeleteRows']              = prot.DeleteRows;
	prot['DeleteTable']             = prot.DeleteTable;

	window['Asc']['c_oAscPrintType'] = window['Asc'].c_oAscPrintType = c_oAscPrintType;
	prot = c_oAscPrintType;
	prot['ActiveSheets'] = prot.ActiveSheets;
	prot['EntireWorkbook'] = prot.EntireWorkbook;
	prot['Selection'] = prot.Selection;

	window['Asc']['c_oDashType'] = window['Asc'].c_oDashType = c_oDashType;
	prot                  = c_oDashType;
	prot['dash']          = prot.dash;
	prot['dashDot']       = prot.dashDot;
	prot['dot']           = prot.dot;
	prot['lgDash']        = prot.lgDash;
	prot['lgDashDot']     = prot.lgDashDot;
	prot['lgDashDotDot']  = prot.lgDashDotDot;
	prot['solid']         = prot.solid;
	prot['sysDash']       = prot.sysDash;
	prot['sysDashDot']    = prot.sysDashDot;
	prot['sysDashDotDot'] = prot.sysDashDotDot;
	prot['sysDot']        = prot.sysDot;


    window['Asc']['c_oAscMathInterfaceType'] = window['Asc'].c_oAscMathInterfaceType = c_oAscMathInterfaceType;
    prot                  = c_oAscMathInterfaceType;
    prot['Common'] = prot.Common;
    prot['Fraction'] = prot.Fraction;
    prot['Script'] = prot.Script;
    prot['Radical'] = prot.Radical;
    prot['LargeOperator'] = prot.LargeOperator;
    prot['Delimiter'] = prot.Delimiter;
    prot['Function'] = prot.Function;
    prot['Accent'] = prot.Accent;
    prot['BorderBox'] = prot.BorderBox;
    prot['Bar'] = prot.Bar;
    prot['Box'] = prot.Box;
    prot['Limit'] = prot.Limit;
    prot['GroupChar'] = prot.GroupChar;
    prot['Matrix'] = prot.Matrix;
    prot['EqArray'] = prot.EqArray;
    prot['Phantom'] = prot.Phantom;



	prot = window['Asc']['c_oAscMathInterfaceBarPos'] = window['Asc'].c_oAscMathInterfaceBarPos = c_oAscMathInterfaceBarPos;
	prot['Top']    = c_oAscMathInterfaceBarPos.Top;
	prot['Bottom'] = c_oAscMathInterfaceBarPos.Bottom;

	prot = window['Asc']['c_oAscMathInterfaceScript'] = window['Asc'].c_oAscMathInterfaceScript = c_oAscMathInterfaceScript;
	prot['None']      = c_oAscMathInterfaceScript.None;
	prot['Sup']       = c_oAscMathInterfaceScript.Sup;
	prot['Sub']       = c_oAscMathInterfaceScript.Sub;
	prot['SubSup']    = c_oAscMathInterfaceScript.SubSup;
	prot['PreSubSup'] = c_oAscMathInterfaceScript.PreSubSup;

	prot = window['Asc']['c_oAscMathInterfaceFraction'] = window['Asc'].c_oAscMathInterfaceFraction = c_oAscMathInterfaceFraction;
	prot['None']   = c_oAscMathInterfaceFraction.Bar;
	prot['Skewed'] = c_oAscMathInterfaceFraction.Skewed;
	prot['Linear'] = c_oAscMathInterfaceFraction.Linear;
	prot['NoBar']  = c_oAscMathInterfaceFraction.NoBar;

	prot = window['Asc']['c_oAscMathInterfaceLimitPos'] = window['Asc'].c_oAscMathInterfaceLimitPos = c_oAscMathInterfaceLimitPos;
	prot['None']   = c_oAscMathInterfaceLimitPos.None;
	prot['Top']    = c_oAscMathInterfaceLimitPos.Top;
	prot['Bottom'] = c_oAscMathInterfaceLimitPos.Bottom;

	prot = window['Asc']['c_oAscMathInterfaceMatrixMatrixAlign'] = window['Asc'].c_oAscMathInterfaceMatrixMatrixAlign = c_oAscMathInterfaceMatrixMatrixAlign;
	prot['Top']    = c_oAscMathInterfaceMatrixMatrixAlign.Top;
	prot['Center'] = c_oAscMathInterfaceMatrixMatrixAlign.Center;
	prot['Bottom'] = c_oAscMathInterfaceMatrixMatrixAlign.Bottom;

	prot = window['Asc']['c_oAscMathInterfaceMatrixColumnAlign'] = window['Asc'].c_oAscMathInterfaceMatrixColumnAlign = c_oAscMathInterfaceMatrixColumnAlign;
	prot['Left']   = c_oAscMathInterfaceMatrixColumnAlign.Left;
	prot['Center'] = c_oAscMathInterfaceMatrixColumnAlign.Center;
	prot['Right']  = c_oAscMathInterfaceMatrixColumnAlign.Right;

	prot = window['Asc']['c_oAscMathInterfaceEqArrayAlign'] = window['Asc'].c_oAscMathInterfaceEqArrayAlign = c_oAscMathInterfaceEqArrayAlign;
	prot['Top']    = c_oAscMathInterfaceEqArrayAlign.Top;
	prot['Center'] = c_oAscMathInterfaceEqArrayAlign.Center;
	prot['Bottom'] = c_oAscMathInterfaceEqArrayAlign.Bottom;

	prot = window['Asc']['c_oAscMathInterfaceNaryLimitLocation'] = window['Asc'].c_oAscMathInterfaceNaryLimitLocation = c_oAscMathInterfaceNaryLimitLocation;
	prot['UndOvr'] = c_oAscMathInterfaceNaryLimitLocation.UndOvr;
	prot['SubSup'] = c_oAscMathInterfaceNaryLimitLocation.SubSup;

	prot = window['Asc']['c_oAscMathInterfaceGroupCharPos'] = window['Asc'].c_oAscMathInterfaceGroupCharPos = c_oAscMathInterfaceGroupCharPos;
	prot['None']   = c_oAscMathInterfaceGroupCharPos.None;
	prot['Top']    = c_oAscMathInterfaceGroupCharPos.Top;
	prot['Bottom'] = c_oAscMathInterfaceGroupCharPos.Bottom;

	prot = window['Asc']['c_oAscTabLeader'] = window['Asc'].c_oAscTabLeader = c_oAscTabLeader;
	prot["None"]       = c_oAscTabLeader.None;
	prot["Heavy"]      = c_oAscTabLeader.Heavy;
	prot["Dot"]        = c_oAscTabLeader.Dot;
	prot["Hyphen"]     = c_oAscTabLeader.Hyphen;
	prot["MiddleDot"]  = c_oAscTabLeader.MiddleDot;
	prot["Underscore"] = c_oAscTabLeader.Underscore;

	prot = window['Asc']['c_oAscTabType'] = window['Asc'].c_oAscTabType = c_oAscTabType;
	prot["Bar"]     = c_oAscTabType.Bar;
	prot["Center"]  = c_oAscTabType.Center;
	prot["Clear"]   = c_oAscTabType.Clear;
	prot["Decimal"] = c_oAscTabType.Decimal;
	prot["Num"]     = c_oAscTabType.Num;
	prot["Right"]   = c_oAscTabType.Right;
	prot["Left"]    = c_oAscTabType.Left;


	prot = window['Asc']['c_oAscRestrictionType'] = window['Asc'].c_oAscRestrictionType = c_oAscRestrictionType;
	prot['None']           = c_oAscRestrictionType.None;
	prot['OnlyForms']      = c_oAscRestrictionType.OnlyForms;
	prot['OnlyComments']   = c_oAscRestrictionType.OnlyComments;
	prot['OnlySignatures'] = c_oAscRestrictionType.OnlySignatures;
	prot['View']           = c_oAscRestrictionType.View;

	prot = window['Asc']['c_oAscLocalRestrictionType'] = window['Asc'].c_oAscLocalRestrictionType = c_oAscLocalRestrictionType;
	prot['None']     = c_oAscLocalRestrictionType.None;
	prot['ReadOnly'] = c_oAscLocalRestrictionType.ReadOnly;
	prot['Locked']   = c_oAscLocalRestrictionType.Locked;
	prot['Nosafe']   = c_oAscLocalRestrictionType.Nosafe;


	prot =  window["AscCommon"]["c_oAscCellAnchorType"] = window["AscCommon"].c_oAscCellAnchorType = c_oAscCellAnchorType;
	prot["cellanchorAbsolute"] = prot.cellanchorAbsolute;
	prot["cellanchorOneCell"] = prot.cellanchorOneCell;
	prot["cellanchorTwoCell"] = prot.cellanchorTwoCell;

	prot = window['Asc'];
	prot["vertalign_Baseline"]    = vertalign_Baseline;
	prot["vertalign_SuperScript"] = vertalign_SuperScript;
	prot["vertalign_SubScript"]   = vertalign_SubScript;

	window['AscCommon']                             = window['AscCommon'] || {};
	window["AscCommon"].g_cCharDelimiter            = g_cCharDelimiter;
	window["AscCommon"].g_cGeneralFormat            = g_cGeneralFormat;
	window["AscCommon"].bDate1904                   = false;
	window["AscCommon"].c_oAscAdvancedOptionsAction = c_oAscAdvancedOptionsAction;
	window["AscCommon"].DownloadType                = DownloadType;
	window["AscCommon"].CellValueType               = CellValueType;
	window["AscCommon"].c_oAscChartDefines          = c_oAscChartDefines;
	window['Asc']['c_oAscStyleImage']               = window['Asc'].c_oAscStyleImage = window["AscCommon"].c_oAscStyleImage = c_oAscStyleImage;
	c_oAscStyleImage["Default"] = c_oAscStyleImage.Default;
	c_oAscStyleImage["Document"] = c_oAscStyleImage.Document;

	window["AscCommon"].c_oAscLineDrawingRule       = c_oAscLineDrawingRule;
	window["AscCommon"].vertalign_Baseline          = vertalign_Baseline;
	window["AscCommon"].vertalign_SuperScript       = vertalign_SuperScript;
	window["AscCommon"].vertalign_SubScript         = vertalign_SubScript;
	window["AscCommon"].hdrftr_Header               = hdrftr_Header;
	window["AscCommon"].hdrftr_Footer               = hdrftr_Footer;
	window["AscCommon"].vaKSize                     = vaKSize;
	window["AscCommon"].vaKSuper                    = vaKSuper;
	window["AscCommon"].vaKSub                      = vaKSub;
	window["AscCommon"].c_oAscSizeRelFromH          = c_oAscSizeRelFromH;
	window["AscCommon"].c_oAscSizeRelFromV          = c_oAscSizeRelFromV;
	window["AscCommon"].c_oAscWrapStyle             = c_oAscWrapStyle;
	window["AscCommon"].c_oAscBorderWidth           = c_oAscBorderWidth;
	window["AscCommon"].c_oAscBorderType            = c_oAscBorderType;
	window["AscCommon"].c_oAscLockTypes             = c_oAscLockTypes;
	window["AscCommon"].c_oAscFormatPainterState    = c_oAscFormatPainterState;
	window["AscCommon"].c_oAscSaveTypes             = c_oAscSaveTypes;
	window["AscCommon"].c_oAscChartType             = c_oAscChartType;
	window["AscCommon"].c_oAscChartSubType          = c_oAscChartSubType;
	window["AscCommon"].c_oAscCsvDelimiter          = c_oAscCsvDelimiter;
	window["AscCommon"].c_oAscUrlType               = c_oAscUrlType;
	window["AscCommon"].c_oAscPrintDefaultSettings  = c_oAscPrintDefaultSettings;
	window["AscCommon"].c_oZoomType                 = c_oZoomType;
	window["AscCommon"].c_oNotifyType               = c_oNotifyType;
	window["AscCommon"].c_oNotifyParentType         = c_oNotifyParentType;
	window["AscCommon"].c_oAscEncodings             = c_oAscEncodings;
	window["AscCommon"].c_oAscEncodingsMap          = c_oAscEncodingsMap;
	window["AscCommon"].c_oAscCodePageNone          = c_oAscCodePageNone;
	window["AscCommon"].c_oAscCodePageUtf7          = c_oAscCodePageUtf7;
	window["AscCommon"].c_oAscCodePageUtf8          = c_oAscCodePageUtf8;
	window["AscCommon"].c_oAscCodePageUtf16         = c_oAscCodePageUtf16;
	window["AscCommon"].c_oAscCodePageUtf16BE       = c_oAscCodePageUtf16BE;
	window["AscCommon"].c_oAscCodePageUtf32         = c_oAscCodePageUtf32;
	window["AscCommon"].c_oAscCodePageUtf32BE       = c_oAscCodePageUtf32BE;
	window["AscCommon"].c_oAscMaxFormulaLength      = c_oAscMaxFormulaLength;
	window["AscCommon"].c_oAscMaxFormulaReferenceLength = c_oAscMaxFormulaReferenceLength;
	window["AscCommon"].c_oAscMaxTableColumnTextLength = c_oAscMaxTableColumnTextLength;

	window["AscCommon"].c_oAscFrameDataType = c_oAscFrameDataType;

	prot =  window["AscCommon"]["c_oAscUrlType"] = window["AscCommon"].c_oAscUrlType = c_oAscUrlType;
	prot["Invalid"] = prot.Invalid;
	prot["Http"] = prot.Http;
	prot["Email"] = prot.Email;
	prot["Unsafe"] = prot.Unsafe;


	window["AscCommon"].changestype_None                      = changestype_None;
	window["AscCommon"].changestype_Paragraph_Content         = changestype_Paragraph_Content;
	window["AscCommon"].changestype_Paragraph_Properties      = changestype_Paragraph_Properties;
	window["AscCommon"].changestype_Paragraph_AddText         = changestype_Paragraph_AddText;
	window["AscCommon"].changestype_Paragraph_TextProperties  = changestype_Paragraph_TextProperties;
	window["AscCommon"].changestype_Document_Content          = changestype_Document_Content;
	window["AscCommon"].changestype_Document_Content_Add      = changestype_Document_Content_Add;
	window["AscCommon"].changestype_Document_SectPr           = changestype_Document_SectPr;
	window["AscCommon"].changestype_Document_Styles           = changestype_Document_Styles;
	window["AscCommon"].changestype_Table_Properties          = changestype_Table_Properties;
	window["AscCommon"].changestype_Table_RemoveCells         = changestype_Table_RemoveCells;
	window["AscCommon"].changestype_Image_Properties          = changestype_Image_Properties;
	window["AscCommon"].changestype_ContentControl_Remove     = changestype_ContentControl_Remove;
	window["AscCommon"].changestype_ContentControl_Properties = changestype_ContentControl_Properties;
	window["AscCommon"].changestype_ContentControl_Add        = changestype_ContentControl_Add;
	window["AscCommon"].changestype_HdrFtr                    = changestype_HdrFtr;
	window["AscCommon"].changestype_Remove                    = changestype_Remove;
	window["AscCommon"].changestype_Delete                    = changestype_Delete;
	window["AscCommon"].changestype_Drawing_Props             = changestype_Drawing_Props;
	window["AscCommon"].changestype_ColorScheme               = changestype_ColorScheme;
	window["AscCommon"].changestype_Text_Props                = changestype_Text_Props;
	window["AscCommon"].changestype_RemoveSlide               = changestype_RemoveSlide;
	window["AscCommon"].changestype_Theme                     = changestype_Theme;
	window["AscCommon"].changestype_SlideSize                 = changestype_SlideSize;
	window["AscCommon"].changestype_SlideBg                   = changestype_SlideBg;
	window["AscCommon"].changestype_SlideTransition           = changestype_SlideTransition;
	window["AscCommon"].changestype_MoveComment               = changestype_MoveComment;
	window["AscCommon"].changestype_AddComment                = changestype_AddComment;
	window["AscCommon"].changestype_Layout                    = changestype_Layout;
	window["AscCommon"].changestype_AddShape                  = changestype_AddShape;
	window["AscCommon"].changestype_AddShapes                 = changestype_AddShapes;
	window["AscCommon"].changestype_PresDefaultLang           = changestype_PresDefaultLang;
	window["AscCommon"].changestype_SlideHide                 = changestype_SlideHide;
	window["AscCommon"].changestype_CorePr                    = changestype_CorePr;
	window["AscCommon"].changestype_Document_Settings         = changestype_Document_Settings;
	window["AscCommon"].changestype_Timing                    = changestype_Timing;
	window["AscCommon"].changestype_ViewPr                    = changestype_ViewPr;
	window["AscCommon"].changestype_DocumentProtection        = changestype_DocumentProtection;
	window["AscCommon"].changestype_CustomPr                  = changestype_CustomPr;

	window["AscCommon"].changestype_2_InlineObjectMove        = changestype_2_InlineObjectMove;
	window["AscCommon"].changestype_2_HdrFtr                  = changestype_2_HdrFtr;
	window["AscCommon"].changestype_2_Comment                 = changestype_2_Comment;
	window["AscCommon"].changestype_2_Element_and_Type        = changestype_2_Element_and_Type;
	window["AscCommon"].changestype_2_ElementsArray_and_Type  = changestype_2_ElementsArray_and_Type;
	window["AscCommon"].changestype_2_AdditionalTypes         = changestype_2_AdditionalTypes;
	window["AscCommon"].changestype_2_Element_and_Type_Array  = changestype_2_Element_and_Type_Array;
	window["AscCommon"].contentchanges_Add                    = contentchanges_Add;
	window["AscCommon"].contentchanges_Remove                 = contentchanges_Remove;

	window["AscCommon"].PUNCTUATION_FLAG_BASE                 = PUNCTUATION_FLAG_BASE;
	window["AscCommon"].PUNCTUATION_FLAG_CANT_BE_AT_BEGIN     = PUNCTUATION_FLAG_CANT_BE_AT_BEGIN;
	window["AscCommon"].PUNCTUATION_FLAG_CANT_BE_AT_END       = PUNCTUATION_FLAG_CANT_BE_AT_END;
	window["AscCommon"].PUNCTUATION_FLAG_EAST_ASIAN           = PUNCTUATION_FLAG_EAST_ASIAN;
	window["AscCommon"].PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E   = PUNCTUATION_FLAG_CANT_BE_AT_BEGIN_E;
	window["AscCommon"].PUNCTUATION_FLAG_CANT_BE_AT_END_E     = PUNCTUATION_FLAG_CANT_BE_AT_END_E;
	window["AscCommon"].g_aPunctuation                        = g_aPunctuation;

	window["AscCommon"].offlineMode = offlineMode;
	window["AscCommon"].chartMode = chartMode;
	window["AscCommon"].dataMode = dataMode;

	window['AscCommon']['align_Right'] = window['AscCommon'].align_Right = align_Right;
	window['AscCommon']['align_Left'] = window['AscCommon'].align_Left = align_Left;
	window['AscCommon']['align_Center'] = window['AscCommon'].align_Center = align_Center;
	window['AscCommon']['align_Justify'] = window['AscCommon'].align_Justify = align_Justify;
	window['AscCommon']['align_Distributed'] = window['AscCommon'].align_Distributed = align_Distributed;
	window['AscCommon']['align_CenterContinuous'] = window['AscCommon'].align_CenterContinuous = align_CenterContinuous;


	window["AscCommon"]["c_oAscFormatPainterState"]    = c_oAscFormatPainterState;
	c_oAscFormatPainterState["kOff"] = c_oAscFormatPainterState.kOff;
	c_oAscFormatPainterState["kOn"] = c_oAscFormatPainterState.kOn;
	c_oAscFormatPainterState["kMultiple"] = c_oAscFormatPainterState.kMultiple;

	
	window['Asc']['c_oSpecialPasteProps'] = window['Asc'].c_oSpecialPasteProps = c_oSpecialPasteProps;
	prot = c_oSpecialPasteProps;
	prot['paste'] = prot.paste;
	prot['pasteOnlyFormula'] = prot.pasteOnlyFormula;
	prot['formulaNumberFormat'] = prot.formulaNumberFormat;
	prot['formulaAllFormatting'] = prot.formulaAllFormatting;
	prot['formulaWithoutBorders'] = prot.formulaWithoutBorders;
	prot['formulaColumnWidth'] = prot.formulaColumnWidth;
	prot['mergeConditionalFormating'] = prot.mergeConditionalFormating;
	prot['pasteOnlyValues'] = prot.pasteOnlyValues;
	prot['valueNumberFormat'] = prot.valueNumberFormat;
	prot['valueAllFormating'] = prot.valueAllFormating;
	prot['pasteOnlyFormating'] = prot.pasteOnlyFormating;
	prot['transpose'] = prot.transpose;
	prot['link'] = prot.link;
	prot['picture'] = prot.picture;
	prot['linkedPicture'] = prot.linkedPicture;
	prot['sourceformatting'] = prot.sourceformatting;
	prot['destinationFormatting'] = prot.destinationFormatting;
	prot['mergeFormatting'] = prot.mergeFormatting;
	prot['uniteList'] = prot.uniteList;
	prot['doNotUniteList'] = prot.doNotUniteList;
	prot['keepTextOnly'] = prot.keepTextOnly;
	prot['insertAsNestedTable'] = prot.insertAsNestedTable;
	prot['overwriteCells'] = prot.overwriteCells;
	prot['useTextImport'] = prot.useTextImport;
	prot['comments'] = prot.comments;
	prot['columnWidth'] = prot.columnWidth;

	window['Asc']['c_oSpecialPasteOperation'] = window['Asc'].c_oSpecialPasteOperation = c_oSpecialPasteOperation;
	prot = c_oSpecialPasteOperation;
	prot['none'] = prot.none;
	prot['add'] = prot.add;
	prot['subtract'] = prot.subtract;
	prot['multiply'] = prot.multiply;
	prot['divide'] = prot.divide;

	window['Asc']['c_oAbstractNumMultiLvlTypes'] = window['Asc'].c_oAbstractNumMultiLvlTypes = c_oAbstractNumMultiLvlTypes;
	prot = c_oAbstractNumMultiLvlTypes;
	prot['HybridMultiLevel'] = prot.HybridMultiLevel;
	prot['MultiLevel']       = prot.MultiLevel;
	prot['SingleLevel']      = prot.SingleLevel;

	window['Asc']['c_oAscSmartArtTypes'] = window['Asc'].c_oAscSmartArtTypes = c_oAscSmartArtTypes;
	prot = c_oAscSmartArtTypes;
	prot['AccentedPicture']                   = prot.AccentedPicture;
	prot['Balance']                           = prot.Balance;
	prot['TitledPictureBlocks']               = prot.TitledPictureBlocks;
	prot['PictureAccentBlocks']               = prot.PictureAccentBlocks;
	prot['BlockCycle']                        = prot.BlockCycle;
	prot['StackedVenn']                       = prot.StackedVenn;
	prot['VerticalEquation']                  = prot.VerticalEquation;
	prot['VerticalBlockList']                 = prot.VerticalBlockList;
	prot['VerticalBendingProcess']            = prot.VerticalBendingProcess;
	prot['VerticalBulletList']                = prot.VerticalBulletList;
	prot['VerticalCurvedList']                = prot.VerticalCurvedList;
	prot['VerticalProcess']                   = prot.VerticalProcess;
	prot['VerticalBoxList']                   = prot.VerticalBoxList;
	prot['VerticalPictureList']               = prot.VerticalPictureList;
	prot['VerticalCircleList']                = prot.VerticalCircleList;
	prot['VerticalPictureAccentList']         = prot.VerticalPictureAccentList;
	prot['VerticalArrowList']                 = prot.VerticalArrowList;
	prot['VerticalChevronList']               = prot.VerticalChevronList;
	prot['VerticalAccentList']                = prot.VerticalAccentList;
	prot['NestedTarget']                      = prot.NestedTarget;
	prot['Funnel']                            = prot.Funnel;
	prot['UpwardArrow']                       = prot.UpwardArrow;
	prot['IncreasingArrowsProcess']           = prot.IncreasingArrowsProcess;
	prot['StepUpProcess']                     = prot.StepUpProcess;
	prot['CircularPictureCallout']            = prot.CircularPictureCallout;
	prot['HorizontalHierarchy']               = prot.HorizontalHierarchy;
	prot['HorizontalLabeledHierarchy']        = prot.HorizontalLabeledHierarchy;
	prot['HorizontalMultiLevelHierarchy']     = prot.HorizontalMultiLevelHierarchy;
	prot['HorizontalOrganizationChart']       = prot.HorizontalOrganizationChart;
	prot['HorizontalBulletList']              = prot.HorizontalBulletList;
	prot['HorizontalPictureList']             = prot.HorizontalPictureList;
	prot['ClosedChevronProcess']              = prot.ClosedChevronProcess;
	prot['HierarchyList']                     = prot.HierarchyList;
	prot['Hierarchy']                         = prot.Hierarchy;
	prot['CirclePictureHierarchy']            = prot.CirclePictureHierarchy;
	prot['LabeledHierarchy']                  = prot.LabeledHierarchy;
	prot['InvertedPyramid']                   = prot.InvertedPyramid;
	prot['HexagonCluster']                    = prot.HexagonCluster;
	prot['CircleRelationship']                = prot.CircleRelationship;
	prot['CircleAccentTimeline']              = prot.CircleAccentTimeline;
	prot['CircularBendingProcess']            = prot.CircularBendingProcess;
	prot['ArrowRibbon']                       = prot.ArrowRibbon;
	prot['LinearVenn']                        = prot.LinearVenn;
	prot['PictureLineup']                     = prot.PictureLineup;
	prot['TitlePictureLineup']                = prot.TitlePictureLineup;
	prot['BendingPictureCaptionList']         = prot.BendingPictureCaptionList;
	prot['BendingPictureAccentList']          = prot.BendingPictureAccentList;
	prot['TitledMatrix']                      = prot.TitledMatrix;
	prot['IncreasingCircleProcess']           = prot.IncreasingCircleProcess;
	prot['BendingPictureBlocks']              = prot.BendingPictureBlocks;
	prot['BendingPictureCaption']             = prot.BendingPictureCaption;
	prot['BendingPictureSemiTransparentText'] = prot.BendingPictureSemiTransparentText;
	prot['NonDirectionalCycle']               = prot.NonDirectionalCycle;
	prot['ContinuousBlockProcess']            = prot.ContinuousBlockProcess;
	prot['ContinuousPictureList']             = prot.ContinuousPictureList;
	prot['ContinuousCycle']                   = prot.ContinuousCycle;
	prot['DescendingBlockList']               = prot.DescendingBlockList;
	prot['StepDownProcess']                   = prot.StepDownProcess;
	prot['ReverseList']                       = prot.ReverseList;
	prot['OrganizationChart']                 = prot.OrganizationChart;
	prot['NameAndTitleOrganizationChart']     = prot.NameAndTitleOrganizationChart;
	prot['AlternatingFlow']                   = prot.AlternatingFlow;
	prot['PyramidList']                       = prot.PyramidList;
	prot['PlusAndMinus']                      = prot.PlusAndMinus;
	prot['RepeatingBendingProcess']           = prot.RepeatingBendingProcess;
	prot['CaptionedPictures']                 = prot.CaptionedPictures;
	prot['DetailedProcess']                   = prot.DetailedProcess;
	prot['PictureStrips']                     = prot.PictureStrips;
	prot['HalfCircleOrganizationChart']       = prot.HalfCircleOrganizationChart;
	prot['PhasedProcess']                     = prot.PhasedProcess;
	prot['BasicVenn']                         = prot.BasicVenn;
	prot['BasicTimeline']                     = prot.BasicTimeline;
	prot['BasicPie']                          = prot.BasicPie;
	prot['BasicMatrix']                       = prot.BasicMatrix;
	prot['BasicPyramid']                      = prot.BasicPyramid;
	prot['BasicRadial']                       = prot.BasicRadial;
	prot['BasicTarget']                       = prot.BasicTarget;
	prot['BasicBlockList']                    = prot.BasicBlockList;
	prot['BasicBendingProcess']               = prot.BasicBendingProcess;
	prot['BasicProcess']                      = prot.BasicProcess;
	prot['BasicChevronProcess']               = prot.BasicChevronProcess;
	prot['BasicCycle']                        = prot.BasicCycle;
	prot['OpposingIdeas']                     = prot.OpposingIdeas;
	prot['OpposingArrows']                    = prot.OpposingArrows;
	prot['RandomToResultProcess']             = prot.RandomToResultProcess;
	prot['SubStepProcess']                    = prot.SubStepProcess;
	prot['PieProcess']                        = prot.PieProcess;
	prot['AccentProcess']                     = prot.AccentProcess;
	prot['AscendingPictureAccentProcess']     = prot.AscendingPictureAccentProcess;
	prot['PictureAccentProcess']              = prot.PictureAccentProcess;
	prot['RadialVenn']                        = prot.RadialVenn;
	prot['RadialCycle']                       = prot.RadialCycle;
	prot['RadialCluster']                     = prot.RadialCluster;
	prot['RadialList']                        = prot.RadialList;
	prot['MultiDirectionalCycle']             = prot.MultiDirectionalCycle;
	prot['DivergingRadial']                   = prot.DivergingRadial;
	prot['DivergingArrows']                   = prot.DivergingArrows;
	prot['FramedTextPicture']                 = prot.FramedTextPicture;
	prot['GroupedList']                       = prot.GroupedList;
	prot['SegmentedPyramid']                  = prot.SegmentedPyramid;
	prot['SegmentedProcess']                  = prot.SegmentedProcess;
	prot['SegmentedCycle']                    = prot.SegmentedCycle;
	prot['PictureGrid']                       = prot.PictureGrid;
	prot['GridMatrix']                        = prot.GridMatrix;
	prot['SpiralPicture']                     = prot.SpiralPicture;
	prot['StackedList']                       = prot.StackedList;
	prot['PictureCaptionList']                = prot.PictureCaptionList;
	prot['ProcessList']                       = prot.ProcessList;
	prot['BubblePictureList']                 = prot.BubblePictureList;
	prot['SquareAccentList']                  = prot.SquareAccentList;
	prot['LinedList']                         = prot.LinedList;
	prot['PictureAccentList']                 = prot.PictureAccentList;
	prot['TitledPictureAccentList']           = prot.TitledPictureAccentList;
	prot['SnapshotPictureList']               = prot.SnapshotPictureList;
	prot['ContinuousArrowProcess']            = prot.ContinuousArrowProcess;
	prot['CircleArrowProcess']                = prot.CircleArrowProcess;
	prot['ProcessArrows']                     = prot.ProcessArrows;
	prot['StaggeredProcess']                  = prot.StaggeredProcess;
	prot['ConvergingRadial']                  = prot.ConvergingRadial;
	prot['ConvergingArrows']                  = prot.ConvergingArrows;
	prot['TableHierarchy']                    = prot.TableHierarchy;
	prot['TableList']                         = prot.TableList;
	prot['TextCycle']                         = prot.TextCycle;
	prot['TrapezoidList']                     = prot.TrapezoidList;
	prot['DescendingProcess']                 = prot.DescendingProcess;
	prot['ChevronList']                       = prot.ChevronList;
	prot['Equation']                          = prot.Equation;
	prot['CounterbalanceArrows']              = prot.CounterbalanceArrows;
	prot['TargetList']                        = prot.TargetList;
	prot['CycleMatrix']                       = prot.CycleMatrix;
	prot['AlternatingPictureBlocks']          = prot.AlternatingPictureBlocks;
	prot['AlternatingPictureCircles']         = prot.AlternatingPictureCircles;
	prot['AlternatingHexagonList']            = prot.AlternatingHexagonList;
	prot['Gear']                              = prot.Gear;
	prot['ArchitectureLayout']                = prot.ArchitectureLayout;
	prot['ChevronAccentProcess']              = prot.ChevronAccentProcess;
	prot['CircleProcess']                     = prot.CircleProcess;
	prot['ConvergingText']                    = prot.ConvergingText;
	prot['HexagonRadial']                     = prot.HexagonRadial;
	prot['InterconnectedBlockProcess']        = prot.InterconnectedBlockProcess;
	prot['InterconnectedRings']               = prot.InterconnectedRings;
	prot['PictureFrame']                      = prot.PictureFrame;
	prot['PictureOrganizationChart']          = prot.PictureOrganizationChart;
	prot['RadialPictureList']                 = prot.RadialPictureList;
	prot['TabList']                           = prot.TabList;
	prot['TabbedArc']                         = prot.TabbedArc;
	prot['ThemePictureAccent']                = prot.ThemePictureAccent;
	prot['ThemePictureAlternatingAccent']     = prot.ThemePictureAlternatingAccent;
	prot['ThemePictureGrid']                  = prot.ThemePictureGrid;
	prot['VaryingWidthList']                  = prot.VaryingWidthList;
	prot['VerticalBracketList']               = prot.VerticalBracketList;

	window['Asc']['c_oAscSmartArtSectionNames'] = window['Asc'].c_oAscSmartArtSectionNames = c_oAscSmartArtSectionNames;
	prot = c_oAscSmartArtSectionNames;
	prot['List']         = prot.List;
	prot['Process']      = prot.Process;
	prot['Cycle']        = prot.Cycle;
	prot['Hierarchy']    = prot.Hierarchy;
	prot['Relationship'] = prot.Relationship;
	prot['Matrix']       = prot.Matrix;
	prot['Pyramid']      = prot.Pyramid;
	prot['Picture']      = prot.Picture;
	prot['OfficeCom']    = prot.OfficeCom;

	window['Asc']['c_oAscJSONNumberingType'] = window['Asc'].c_oAscJSONNumberingType = c_oAscJSONNumberingType;
	prot = c_oAscJSONNumberingType;
	prot["Remove"] = prot.Remove;
	prot["Bullet"] = prot.Bullet;
	prot["Number"] = prot.Number;
	prot["Hybrid"] = prot.Hybrid;

	window['Asc']['c_oAscNumberingFormat'] = window['Asc'].c_oAscNumberingFormat = c_oAscNumberingFormat;
	prot = c_oAscNumberingFormat;
	prot['Aiueo']                        = prot.Aiueo;
	prot['AiueoFullWidth']               = prot.AiueoFullWidth;
	prot['ArabicAbjad']                  = prot.ArabicAbjad;
	prot['ArabicAlpha']                  = prot.ArabicAlpha;
	prot['BahtText']                     = prot.BahtText;
	prot['Bullet']                       = prot.Bullet;
	prot['CardinalText']                 = prot.CardinalText;
	prot['Chicago']                      = prot.Chicago;
	prot['ChineseCounting']              = prot.ChineseCounting;
	prot['ChineseCountingThousand']      = prot.ChineseCountingThousand;
	prot['ChineseLegalSimplified']       = prot.ChineseLegalSimplified;
	prot['Chosung']                      = prot.Chosung;
	prot['Custom']                       = prot.Custom;
	prot['Decimal']                      = prot.Decimal;
	prot['DecimalEnclosedCircle']        = prot.DecimalEnclosedCircle;
	prot['DecimalEnclosedCircleChinese'] = prot.DecimalEnclosedCircleChinese;
	prot['DecimalEnclosedFullstop']      = prot.DecimalEnclosedFullstop;
	prot['DecimalEnclosedParen']         = prot.DecimalEnclosedParen;
	prot['DecimalFullWidth']             = prot.DecimalFullWidth;
	prot['DecimalFullWidth2']            = prot.DecimalFullWidth2;
	prot['DecimalHalfWidth']             = prot.DecimalHalfWidth;
	prot['DecimalZero']                  = prot.DecimalZero;
	prot['DollarText']                   = prot.DollarText;
	prot['Ganada']                       = prot.Ganada;
	prot['Hebrew1']                      = prot.Hebrew1;
	prot['Hebrew2']                      = prot.Hebrew2;
	prot['Hex']                          = prot.Hex;
	prot['HindiConsonants']              = prot.HindiConsonants;
	prot['HindiCounting']                = prot.HindiCounting;
	prot['HindiNumbers']                 = prot.HindiNumbers;
	prot['HindiVowels']                  = prot.HindiVowels;
	prot['IdeographDigital']             = prot.IdeographDigital;
	prot['IdeographEnclosedCircle']      = prot.IdeographEnclosedCircle;
	prot['IdeographLegalTraditional']    = prot.IdeographLegalTraditional;
	prot['IdeographTraditional']         = prot.IdeographTraditional;
	prot['IdeographZodiac']              = prot.IdeographZodiac;
	prot['IdeographZodiacTraditional']   = prot.IdeographZodiacTraditional;
	prot['Iroha']                        = prot.Iroha;
	prot['IrohaFullWidth']               = prot.IrohaFullWidth;
	prot['JapaneseCounting']             = prot.JapaneseCounting;
	prot['JapaneseDigitalTenThousand']   = prot.JapaneseDigitalTenThousand;
	prot['JapaneseLegal']                = prot.JapaneseLegal;
	prot['KoreanCounting']               = prot.KoreanCounting;
	prot['KoreanDigital']                = prot.KoreanDigital;
	prot['KoreanDigital2']               = prot.KoreanDigital2;
	prot['KoreanLegal']                  = prot.KoreanLegal;
	prot['LowerLetter']                  = prot.LowerLetter;
	prot['LowerRoman']                   = prot.LowerRoman;
	prot['None']                         = prot.None;
	prot['NumberInDash']                 = prot.NumberInDash;
	prot['Ordinal']                      = prot.Ordinal;
	prot['OrdinalText']                  = prot.OrdinalText;
	prot['RussianLower']                 = prot.RussianLower;
	prot['RussianUpper']                 = prot.RussianUpper;
	prot['TaiwaneseCounting']            = prot.TaiwaneseCounting;
	prot['TaiwaneseCountingThousand']    = prot.TaiwaneseCountingThousand;
	prot['TaiwaneseDigital']             = prot.TaiwaneseDigital;
	prot['ThaiCounting']                 = prot.ThaiCounting;
	prot['ThaiLetters']                  = prot.ThaiLetters;
	prot['ThaiNumbers']                  = prot.ThaiNumbers;
	prot['UpperLetter']                  = prot.UpperLetter;
	prot['UpperRoman']                   = prot.UpperRoman;
	prot['VietnameseCounting']           = prot.VietnameseCounting;

	window['Asc']['c_oAscAllNumberingTypes'] = window['Asc'].c_oAscAllNumberingTypes = c_oAscAllNumberingTypes;

	// new presentation types
	prot['Ea1JpnKor']                    = prot.Ea1JpnKor;
	prot['CircleNumWdBlack']             = prot.CircleNumWdBlack;
	prot['Ea1JpnChsDb']                  = prot.Ea1JpnChsDb;
	prot['Ea1Cht']                       = prot.Ea1Cht;
	prot['CircleNumWdWhitePlain']        = prot.CircleNumWdWhitePlain;

	//custom types
	prot['CustomGreece']                 = prot.CustomGreece;
	prot['CustomDecimalFourZero']        = prot.CustomDecimalFourZero;
	prot['CustomDecimalThreeZero']       = prot.CustomDecimalThreeZero;
	prot['CustomDecimalTwoZero']         = prot.CustomDecimalTwoZero;
	prot['CustomUpperTurkish']           = prot.CustomUpperTurkish;
	prot['CustomLowerTurkish']           = prot.CustomLowerTurkish;

	window['Asc']['c_oAscCustomNumberingFormatAssociation'] = window['Asc'].c_oAscCustomNumberingFormatAssociation = c_oAscCustomNumberingFormatAssociation;

	window['Asc']['c_oAscNumberingSuff'] = window['Asc'].c_oAscNumberingSuff = c_oAscNumberingSuff;
	prot = c_oAscNumberingSuff;
	prot['Tab']   = c_oAscNumberingSuff.Tab;
	prot['Space'] = c_oAscNumberingSuff.Space;
	prot['None']  = c_oAscNumberingSuff.None;

	window['Asc']['c_oAscNumberingLvlTextType'] = window['Asc'].c_oAscNumberingLvlTextType = c_oAscNumberingLvlTextType;
	prot = c_oAscNumberingLvlTextType;
	prot['Text'] = c_oAscNumberingLvlTextType.Text;
	prot['Num']  = c_oAscNumberingLvlTextType.Num;
	
	window['Asc']['c_oAscChapterSep'] = window['Asc'].c_oAscChapterSep = c_oAscChapterSep;
	prot = c_oAscChapterSep;
	prot['Colon'] = c_oAscChapterSep.Colon;
	prot['EmDash']  = c_oAscChapterSep.EmDash;
	prot['EnDash']  = c_oAscChapterSep.EnDash;
	prot['Hyphen']  = c_oAscChapterSep.Hyphen;
	prot['Period']  = c_oAscChapterSep.Period;

	prot = window['Asc']['c_oAscSdtAppearance'] = window['Asc'].c_oAscSdtAppearance = c_oAscSdtAppearance;
	prot['Frame']  = c_oAscSdtAppearance.Frame;
	prot['Hidden'] = c_oAscSdtAppearance.Hidden;


	prot = window['Asc']['c_oAscObjectsAlignType'] = window['Asc'].c_oAscObjectsAlignType = c_oAscObjectsAlignType;
	prot['Selected'] = c_oAscObjectsAlignType.Selected;
	prot['Slide'] = c_oAscObjectsAlignType.Slide;
	prot['Page'] = c_oAscObjectsAlignType.Page;
	prot['Margin'] = c_oAscObjectsAlignType.Margin;

	prot = window['Asc']['c_oAscItemType'] = window['Asc'].c_oAscItemType = c_oAscItemType;
	prot['Data'] = prot.Data;
	prot['Default'] = prot.Default;
	prot['Sum'] = prot.Sum;
	prot['CountA'] = prot.CountA;
	prot['Avg'] = prot.Avg;
	prot['Max'] = prot.Max;
	prot['Min'] = prot.Min;
	prot['Product'] = prot.Product;
	prot['Count'] = prot.Count;
	prot['StdDev'] = prot.StdDev;
	prot['StdDevP'] = prot.StdDevP;
	prot['Var'] = prot.Var;
	prot['VarP'] = prot.VarP;
	prot['Grand'] = prot.Grand;
	prot['Blank'] = prot.Blank;

	prot = window['Asc']['c_oAscRevisionsMove'] = window['Asc'].c_oAscRevisionsMove = c_oAscRevisionsMove;
	prot['NoMove']   = c_oAscRevisionsMove.NoMove;
	prot['MoveTo']   = c_oAscRevisionsMove.MoveTo;
	prot['MoveFrom'] = c_oAscRevisionsMove.MoveFrom;


	prot = window['Asc']['c_oAscRevisionsChangeType'] = window['Asc'].c_oAscRevisionsChangeType = c_oAscRevisionsChangeType;
	prot['Unknown']    = c_oAscRevisionsChangeType.Unknown;
	prot['TextAdd']    = c_oAscRevisionsChangeType.TextAdd;
	prot['TextRem']    = c_oAscRevisionsChangeType.TextRem;
	prot['ParaAdd']    = c_oAscRevisionsChangeType.ParaAdd;
	prot['ParaRem']    = c_oAscRevisionsChangeType.ParaRem;
	prot['TextPr']     = c_oAscRevisionsChangeType.TextPr;
	prot['ParaPr']     = c_oAscRevisionsChangeType.ParaPr;
	prot['TablePr']    = c_oAscRevisionsChangeType.TablePr;
	prot['RowsAdd']    = c_oAscRevisionsChangeType.RowsAdd;
	prot['RowsRem']    = c_oAscRevisionsChangeType.RowsRem;
	prot['TableRowPr'] = c_oAscRevisionsChangeType.TableRowPr;
	prot['MoveMark']   = c_oAscRevisionsChangeType.MoveMark;

	prot = window['Asc']['c_oAscSectionBreakType'] = window['Asc'].c_oAscSectionBreakType = c_oAscSectionBreakType;
	prot['NextPage']   = c_oAscSectionBreakType.NextPage;
	prot['OddPage']    = c_oAscSectionBreakType.OddPage;
	prot['EvenPage']   = c_oAscSectionBreakType.EvenPage;
	prot['Continuous'] = c_oAscSectionBreakType.Continuous;
	prot['Column']     = c_oAscSectionBreakType.Column;


	prot = window['Asc']['c_oAscSdtLockType'] = window['Asc'].c_oAscSdtLockType = c_oAscSdtLockType;
	prot['ContentLocked']    = c_oAscSdtLockType.ContentLocked;
	prot['SdtContentLocked'] = c_oAscSdtLockType.SdtContentLocked;
	prot['SdtLocked']        = c_oAscSdtLockType.SdtLocked;
	prot['Unlocked']         = c_oAscSdtLockType.Unlocked;


	prot = window['Asc']['c_oAscAlignH'] = window['Asc'].c_oAscAlignH = c_oAscAlignH;
	prot['Center']  = c_oAscAlignH.Center;
	prot['Inside']  = c_oAscAlignH.Inside;
	prot['Left']    = c_oAscAlignH.Left;
	prot['Outside'] = c_oAscAlignH.Outside;
	prot['Right']   = c_oAscAlignH.Right;


	prot = window['Asc']['c_oAscAlignV'] = window['Asc'].c_oAscAlignV = c_oAscAlignV;
	prot['Bottom']  = c_oAscAlignV.Bottom;
	prot['Center']  = c_oAscAlignV.Center;
	prot['Inside']  = c_oAscAlignV.Inside;
	prot['Outside'] = c_oAscAlignV.Outside;
	prot['Top']     = c_oAscAlignV.Top;

	prot = window['Asc']['c_oAscWatermarkType'] = window['Asc'].c_oAscWatermarkType = c_oAscWatermarkType;
	prot['None'] = prot.None;
	prot['Text'] = prot.Text;
	prot['Image'] = prot.Image;

	prot = window['Asc']['c_oAscCalendarType'] = window['Asc'].c_oAscCalendarType = c_oAscCalendarType;
	prot['Gregorian']            = c_oAscCalendarType.Gregorian;
	prot['GregorianArabic']      = c_oAscCalendarType.GregorianArabic;
	prot['GregorianMeFrench']    = c_oAscCalendarType.GregorianMeFrench;
	prot['GregorianUs']          = c_oAscCalendarType.GregorianUs;
	prot['GregorianXlitEnglish'] = c_oAscCalendarType.GregorianXlitEnglish;
	prot['GregorianXlitFrench']  = c_oAscCalendarType.GregorianXlitFrench;
	prot['Hebrew']               = c_oAscCalendarType.Hebrew;
	prot['Hijri']                = c_oAscCalendarType.Hijri;
	prot['Japan']                = c_oAscCalendarType.Japan;
	prot['Korea']                = c_oAscCalendarType.Korea;
	prot['None']                 = c_oAscCalendarType.None;
	prot['Saka']                 = c_oAscCalendarType.Saka;
	prot['Taiwan']               = c_oAscCalendarType.Taiwan;
	prot['Thai']                 = c_oAscCalendarType.Thai;

	prot = window['Asc']['c_oAscContentControlSpecificType'] = window['Asc'].c_oAscContentControlSpecificType = c_oAscContentControlSpecificType;
	prot['None']         = c_oAscContentControlSpecificType.None;
	prot['CheckBox']     = c_oAscContentControlSpecificType.CheckBox;
	prot['Picture']      = c_oAscContentControlSpecificType.Picture;
	prot['ComboBox']     = c_oAscContentControlSpecificType.ComboBox;
	prot['DropDownList'] = c_oAscContentControlSpecificType.DropDownList;
	prot['DateTime']     = c_oAscContentControlSpecificType.DateTime;
	prot['TOC']          = c_oAscContentControlSpecificType.TOC;
	prot['Complex']      = c_oAscContentControlSpecificType.Complex;

	window['Asc']['c_oAscDefNameType'] = window['Asc'].c_oAscDefNameType = c_oAscDefNameType;
	prot = c_oAscDefNameType;
	prot['table'] = prot.table;
	prot['slicer'] = prot.slicer;

	window['Asc']['c_oAscChangeTextCaseType'] = window['Asc'].c_oAscChangeTextCaseType = c_oAscChangeTextCaseType;
	prot = c_oAscChangeTextCaseType;
	prot['SentenceCase']    = prot.SentenceCase;
	prot['LowerCase']       = prot.LowerCase;
	prot['UpperCase']       = prot.UpperCase;
	prot['CapitalizeWords'] = prot.CapitalizeWords;
	prot['ToggleCase']      = prot.ToggleCase;

	window["AscCommon"].document_compatibility_mode_Word11  = document_compatibility_mode_Word11;
	window["AscCommon"].document_compatibility_mode_Word12  = document_compatibility_mode_Word12;
	window["AscCommon"].document_compatibility_mode_Word14  = document_compatibility_mode_Word14;
	window["AscCommon"].document_compatibility_mode_Word15  = document_compatibility_mode_Word15;
	window["AscCommon"].document_compatibility_mode_Current = document_compatibility_mode_Current;

	prot = window['AscCommon']['c_oAscCustomShortcutType'] = window['AscCommon'].c_oAscCustomShortcutType = c_oAscCustomShortcutType;
	prot['Symbol'] = c_oAscCustomShortcutType.Symbol;

	prot = window['Asc']['c_oAscLineNumberRestartType'] = window['Asc'].c_oAscLineNumberRestartType = c_oAscLineNumberRestartType;
	prot['Continuous'] = c_oAscLineNumberRestartType.Continuous;
	prot['NewPage']    = c_oAscLineNumberRestartType.NewPage;
	prot['NewSection'] = c_oAscLineNumberRestartType.NewSection;

	prot = window['Asc']['c_oAscSectionApplyType'] = window['Asc'].c_oAscSectionApplyType = c_oAscSectionApplyType;
	prot['Current'] = c_oAscSectionApplyType.Current;
	prot['ToEnd']   = c_oAscSectionApplyType.ToEnd;
	prot['All']     = c_oAscSectionApplyType.All;

	prot = window['Asc']['c_oAscSlideSZType'] = window['Asc'].c_oAscSlideSZType = c_oAscSlideSZType;
	prot['Sz35mm'] = prot.Sz35mm;
	prot['SzA3'] = prot.SzA3;
	prot['SzA4'] = prot.SzA4;
	prot['SzB4ISO'] = prot.SzB4ISO;
	prot['SzB4JIS'] = prot.SzB4JIS;
	prot['SzB5ISO'] = prot.SzB5ISO;
	prot['SzB5JIS'] = prot.SzB5JIS;
	prot['SzBanner'] = prot.SzBanner;
	prot['SzCustom'] = prot.SzCustom;
	prot['SzHagakiCard'] = prot.SzHagakiCard;
	prot['SzLedger'] = prot.SzLedger;
	prot['SzLetter'] = prot.SzLetter;
	prot['SzOverhead'] = prot.SzOverhead;
	prot['SzScreen16x10'] = prot.SzScreen16x10;
	prot['SzScreen16x9'] = prot.SzScreen16x9;
	prot['SzScreen4x3'] = prot.SzScreen4x3;
	prot['SzWidescreen'] = prot.SzWidescreen;

	prot = window['Asc']['c_oAscPictureFormScaleFlag'] = window['Asc'].c_oAscPictureFormScaleFlag = c_oAscPictureFormScaleFlag;
	prot['Always']  = prot.Always;
	prot['Bigger']  = prot.Bigger;
	prot['Smaller'] = prot.Smaller;
	prot['Never']   = prot.Never;

	prot = window['Asc']['c_oAscDisplayModeInReview'] = window['Asc'].c_oAscDisplayModeInReview = c_oAscDisplayModeInReview;
	prot['Edit']     = prot.Edit;
	prot['Final']    = prot.Final;
	prot['Original'] = prot.Original;
	prot['Simple']   = prot.Simple;

	prot = window['Asc']['c_oAscMathInputType'] = window['Asc'].c_oAscMathInputType = c_oAscMathInputType;
	prot['Unicode'] = prot.Unicode;
	prot['LaTeX']   = prot.LaTeX;

	window['AscFormat'] = window['AscFormat'] || {};

	window['AscFormat'].text_fit_No         = window['AscFormat']['text_fit_No']         = 0;
    window['AscFormat'].text_fit_Auto       = window['AscFormat']['text_fit_Auto']       = 1;
    window['AscFormat'].text_fit_NormAuto   = window['AscFormat']['text_fit_NormAuto']   = 2;

	//Vert Overflow Types
    window['AscFormat'].nVOTClip = window['AscFormat']['nVOTClip'] = 0;
    window['AscFormat'].nVOTEllipsis = window['AscFormat']['nVOTEllipsis'] = 1;
	window['AscFormat'].nVOTOverflow = window['AscFormat']['nVOTOverflow'] = 2;

	//Hor OverFlow Types
	window['AscFormat'].nHOTClip = window['AscFormat']['nHOTClip'] = 0;
	window['AscFormat'].nHOTOverflow = window['AscFormat']['nHOTOverflow'] = 1;

    window['AscFormat'].BULLET_TYPE_BULLET_NONE = window['AscFormat']['BULLET_TYPE_BULLET_NONE'] = 0;
    window['AscFormat'].BULLET_TYPE_BULLET_CHAR = window['AscFormat']['BULLET_TYPE_BULLET_CHAR'] = 1;
    window['AscFormat'].BULLET_TYPE_BULLET_AUTONUM = window['AscFormat']['BULLET_TYPE_BULLET_AUTONUM'] = 2;
    window['AscFormat'].BULLET_TYPE_BULLET_BLIP = window['AscFormat']['BULLET_TYPE_BULLET_BLIP'] = 3;
	window['AscCommon'].c_oAscRectAlignType = c_oAscRectAlignType;
	window['AscCommon'].ST_HueDir = ST_HueDir;
	window['AscCommon'].ST_ClrAppMethod = ST_ClrAppMethod;
	window['AscCommon'].ST_AnimLvlStr = ST_AnimLvlStr;
	window['AscCommon'].ST_AnimOneStr = ST_AnimOneStr;
	window['AscCommon'].ST_Direction = ST_Direction;
	window['AscCommon'].ST_HierBranchStyle = ST_HierBranchStyle;
	window['AscCommon'].ST_ResizeHandlesStr = ST_ResizeHandlesStr;
	window['AscCommon'].ST_PtType = ST_PtType;
	window['AscCommon'].ST_ChildOrderType = ST_ChildOrderType;
	window['AscCommon'].ST_AlgorithmType = ST_AlgorithmType;
	window['AscCommon'].ST_ConstraintRelationship = ST_ConstraintRelationship;
	window['AscCommon'].ST_BoolOperator = ST_BoolOperator;
	window['AscCommon'].ST_ElementType = ST_ElementType;
	window['AscCommon'].ST_ConstraintType = ST_ConstraintType;
	window['AscCommon'].ST_VariableType = ST_VariableType;
	window['AscCommon'].ST_AxisType = ST_AxisType;
	window['AscCommon'].ST_FunctionType = ST_FunctionType;
	window['AscCommon'].ST_FunctionOperator = ST_FunctionOperator;
	window['AscCommon'].ST_LayoutShapeType = ST_LayoutShapeType;
	window['AscCommon'].ST_ParameterId = ST_ParameterId;
	window['AscCommon'].ST_PresetCameraType = ST_PresetCameraType;
	window['AscCommon'].ST_LightRigDirection = ST_LightRigDirection;
	window['AscCommon'].ST_LightRigType = ST_LightRigType;
	window['AscCommon'].ST_BevelPresetType = ST_BevelPresetType;
	window['AscCommon'].ST_PresetMaterialType = ST_PresetMaterialType;
	window['AscCommon'].ST_TLAnimateEffectTransition = ST_TLAnimateEffectTransition;
	window['AscCommon'].c_oAscOleObjectTypes = c_oAscOleObjectTypes;
	window['AscCommon'].ST_CxnType = ST_CxnType;

	prot = window['Asc']['LigaturesType'] = window['Asc'].LigaturesType = LigaturesType;
	prot['None']                             = prot.None;
	prot['Standard']                         = prot.Standard;
	prot['Contextual']                       = prot.Contextual;
	prot['Historical']                       = prot.Historical;
	prot['Discretional']                     = prot.Discretional;
	prot['StandardContextual']               = prot.StandardContextual;
	prot['StandardHistorical']               = prot.StandardHistorical;
	prot['ContextualHistorical']             = prot.ContextualHistorical;
	prot['StandardDiscretional']             = prot.StandardDiscretional;
	prot['ContextualDiscretional']           = prot.ContextualDiscretional;
	prot['HistoricalDiscretional']           = prot.HistoricalDiscretional;
	prot['StandardContextualHistorical']     = prot.StandardContextualHistorical;
	prot['StandardContextualDiscretional']   = prot.StandardContextualDiscretional;
	prot['StandardHistoricalDiscretional']   = prot.StandardHistoricalDiscretional;
	prot['ContextualHistoricalDiscretional'] = prot.ContextualHistoricalDiscretional;
	prot['All']                              = prot.All;

	prot = window['Asc']['CombFormWidthRule'] = window['Asc'].CombFormWidthRule = CombFormWidthRule;
	prot['Auto']    = prot.Auto;
	prot['AtLeast'] = prot.AtLeast;
	prot['Exact']   = prot.Exact;

	prot = window['Asc']['UnderlineType'] = window['Asc'].UnderlineType = UnderlineType;
	prot['Dash']            = prot.Dash;
	prot['DashDotDotHeavy'] = prot.DashDotDotHeavy;
	prot['DashDotHeavy']    = prot.DashDotHeavy;
	prot['DashedHeavy']     = prot.DashedHeavy;
	prot['DashLong']        = prot.DashLong;
	prot['DashLongHeavy']   = prot.DashLongHeavy;
	prot['DotDash']         = prot.DotDash;
	prot['DotDotDash']      = prot.DotDotDash;
	prot['Dotted']          = prot.Dotted;
	prot['DottedHeavy']     = prot.DottedHeavy;
	prot['Double']          = prot.Double;
	prot['None']            = prot.None;
	prot['Single']          = prot.Single;
	prot['Thick']           = prot.Thick;
	prot['Wave']            = prot.Wave;
	prot['WavyDouble']      = prot.WavyDouble;
	prot['WavyHeavy']       = prot.WavyHeavy;
	prot['Words']           = prot.Words;

	prot = window['Asc']['DocumentView'] = window['Asc'].DocumentView = DocumentView;
	prot['MasterPages'] = prot.MasterPages;
	prot['None']        = prot.None;
	prot['Normal']      = prot.Normal;
	prot['Outline']     = prot.Outline;
	prot['Print']       = prot.Print;
	prot['Web']         = prot.Web;

	prot = window['Asc']['c_oAscConfirm'] = window['Asc'].c_oAscConfirm = c_oAscConfirm;
	prot['ConfirmReplaceRange'] = prot.ConfirmReplaceRange;
	prot['ConfirmPutMergeRange'] = prot.ConfirmPutMergeRange;
	prot['ConfirmChangeProtectRange'] = prot.ConfirmChangeProtectRange;
	prot['ConfirmMaxChangesSize'] = prot.ConfirmMaxChangesSize;
	prot['ConfirmAddCellWatches'] = prot.ConfirmAddCellWatches;
	prot['ConfirmReplaceHeaderFooterPicture'] = prot.ConfirmReplaceHeaderFooterPicture;

	window['Asc']['c_oAscSmartArtSections'] = window['Asc'].c_oAscSmartArtSections = c_oAscSmartArtSections;



	prot = window['Asc']['c_oPluginContextMenuTypes'] = window['Asc'].c_oPluginContextMenuTypes = c_oPluginContextMenuTypes;
	prot['None'] = prot.None;
	prot['Target'] = prot.Target;
	prot['Selection'] = prot.Selection;
	prot['Image'] = prot.Image;
	prot['Shape'] = prot.Shape;
	prot['OleObject'] = prot.OleObject;

	window['Asc']['c_oAscDateTimeFormat'] = window['Asc'].c_oAscDateTimeFormat = c_oAscDateTimeFormat;

})(window);
