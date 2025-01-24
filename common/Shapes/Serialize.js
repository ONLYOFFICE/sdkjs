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

(function(window, undefined){

var prot;
// Import
var g_nodeAttributeEnd = AscCommon.g_nodeAttributeEnd;

var c_oAscShdClear = Asc.c_oAscShdClear;
var c_oAscColor = Asc.c_oAscColor;
var c_oAscFill = Asc.c_oAscFill;

var c_dScalePPTXSizes = 36000;
function IsHiddenObj(object)
{
    if (!object)
        return false;
    var _uniProps = object.getUniNvProps && object.getUniNvProps();
    if (!_uniProps)
        return false;

    if (_uniProps.cNvPr && _uniProps.cNvPr.isHidden)
        return true;

    return false;
}

function CBuilderImages(blip_fill, full_url, image_shape, sp_pr, ln, text_pr, para_text_pr, run, paragraph, bullet)
{
    this.Url = full_url;
    this.BlipFill = blip_fill;
    this.ImageShape = image_shape;
    this.SpPr = sp_pr;
    this.Ln = ln;

    this.TextPr     = text_pr;
    this.ParaTextPr = para_text_pr;
    this.Run        = run;
    this.Paragraph  = paragraph;
    this.Bullet = bullet;
    this.AdditionalUrls = [];//для wmf, ole
}
CBuilderImages.prototype =
{
    SetUrl : function(url)
    {
        if(url !== "error")
        {
            var oCopyFill, oCopyBlipFill, oCopyLn;
            if(!this.Ln && this.SpPr && this.SpPr.Fill)
            {
                oCopyFill = this.SpPr.Fill.createDuplicate();
                if(oCopyFill.fill && oCopyFill.fill.type === c_oAscFill.FILL_TYPE_BLIP)
                {
                    oCopyFill.fill.setRasterImageId(url);
                    this.SpPr.setFill(oCopyFill);
                }
            }
            if(this.Ln && this.SpPr && this.SpPr === this.Ln && this.Ln.Fill && this.Ln.Fill.fill && this.Ln.Fill.fill.type === c_oAscFill.FILL_TYPE_BLIP)
            {
                oCopyLn = this.Ln.createDuplicate();
                oCopyLn.Fill.fill.setRasterImageId(url);
                this.SpPr.setLn(oCopyLn);
            }
            if(this.ImageShape && this.ImageShape.blipFill)
            {
                oCopyBlipFill = this.ImageShape.blipFill.createDuplicate();
                oCopyBlipFill.setRasterImageId(url);
                this.ImageShape.setBlipFill(oCopyBlipFill);
            }
            if(this.TextPr && !this.Ln)
            {
                if(this.Paragraph)
                {
                    var oPr = this.Paragraph.Pr;
                    if(oPr.DefaultRunPr && oPr.DefaultRunPr.Unifill && oPr.DefaultRunPr.Unifill.fill && oPr.DefaultRunPr.Unifill.fill.type === c_oAscFill.FILL_TYPE_BLIP)
                    {
                        var Pr = this.Paragraph.Pr.Copy();
                        Pr.DefaultRunPr.Unifill.fill.setRasterImageId(url);
                        this.Paragraph.Set_Pr(Pr);
                    }
                }
                else if(this.ParaTextPr || this.Run)
                {
                    if(this.ParaTextPr && this.ParaTextPr.Value && this.ParaTextPr.Value.Unifill && this.ParaTextPr.Value.Unifill.fill && this.ParaTextPr.Value.Unifill.fill.type === c_oAscFill.FILL_TYPE_BLIP)
                    {
                        oCopyFill = this.ParaTextPr.Value.Unifill.createDuplicate();
                        oCopyFill.fill.setRasterImageId(url);
                        this.ParaTextPr.Set_Unifill(oCopyFill);
                    }
                    if(this.Run && this.Run.Pr && this.Run.Pr.Unifill && this.Run.Pr.Unifill.fill && this.Run.Pr.Unifill.fill.type === c_oAscFill.FILL_TYPE_BLIP)
                    {
                        oCopyFill = this.Run.Pr.Unifill.createDuplicate();
                        oCopyFill.fill.setRasterImageId(url);
                        this.Run.Set_Unifill(oCopyFill);
                    }
                }
            }
            if(this.Bullet) 
            {
                if(this.Paragraph) 
                {
                    var oPr = this.Paragraph.Pr;
                    var oBullet = oPr.Bullet;
                    if(oBullet.getImageBulletURL()) 
                    {
                        var oNewPr = oPr.Copy();
                        oNewPr.Bullet.fillBulletImage(url);
                        this.Paragraph.Set_Pr(oNewPr);
                    }
                }
            }
            this.BlipFill.RasterImageId = url;
        }
    }
};

function BinaryPPTYLoader()
{
    this.stream = null;
    this.presentation = null;
    this.DrawingDocument = null;
    this.TempGroupObject = null;
    this.TempMainObject = null;

    this.IsFillingSmartArt = false;
    this.IsThemeLoader = false;
    this.Api = null;

    this.map_table_styles = {};

    this.ImageMapChecker = null;

    this.IsUseFullUrl = false;
	this.insertDocumentUrlsData = null;
    this.RebuildImages = [];

    this.aSlideLayouts = [];
    this.aThemes = [];

	this.oConnectedObjects = {};
	this.map_shapes_by_id = {};
	this.fields = [];
	this.smartarts = [];


	this.ClearConnectedObjects = function(){
        this.oConnectedObjects = {};
        this.map_shapes_by_id = {};
    };
	this.AddConnectedObject = function(oObject){
        this.oConnectedObjects[oObject.Id] = oObject;
    };
	this.AssignConnectedObjects = function () {
        for(var sId in this.oConnectedObjects) {
            if(this.oConnectedObjects.hasOwnProperty(sId)) {
                this.oConnectedObjects[sId].assignConnection(this.map_shapes_by_id);
            }
        }
        this.ClearConnectedObjects();
    };

    this.Start_UseFullUrl = function(insertDocumentUrlsData)
    {
        this.IsUseFullUrl = true;
		this.insertDocumentUrlsData = insertDocumentUrlsData;
    };

    this.End_UseFullUrl = function()
    {
        var _result = this.RebuildImages;

        this.IsUseFullUrl = false;
        this.RebuildImages = [];

        return _result;
    };

    this.Load = function(base64_ppty, presentation)
    {
        this.presentation = presentation;
        this.DrawingDocument = null;
        if(presentation)
        {
            this.DrawingDocument = presentation.DrawingDocument;
        }
        else
        {
            this.DrawingDocument = null;
        }
        this.ImageMapChecker = {};

		var isBase64 = typeof base64_ppty === 'string';
		var srcLen = isBase64 ? base64_ppty.length : base64_ppty.length;
        var nWritten = 0;

        var index = 0;
        var read_main_prop = "";
        while (true)
        {
            var _c = isBase64 ? base64_ppty.charCodeAt(index) : base64_ppty[index];
            if (_c == ";".charCodeAt(0))
                break;

            read_main_prop += String.fromCharCode(_c);
            index++;
        }
        index++;

        if ("PPTY" != read_main_prop)
            return false;

        read_main_prop = "";
        while (true)
        {
            var _c = isBase64 ? base64_ppty.charCodeAt(index) : base64_ppty[index];
            if (_c == ";".charCodeAt(0))
                break;

            read_main_prop += String.fromCharCode(_c);
            index++;
        }
        index++;

        var _version_num_str = read_main_prop.substring(1);
		var version = 1;
		if(_version_num_str.length > 0)
        {
            version = _version_num_str - 0;
        }
        read_main_prop = "";
        while (true)
        {
            var _c = isBase64 ? base64_ppty.charCodeAt(index) : base64_ppty[index];
            if (_c == ";".charCodeAt(0))
                break;

            read_main_prop += String.fromCharCode(_c);
            index++;
        }
        index++;

		if (Asc.c_nVersionNoBase64 !== version) {
			var dstLen_str = read_main_prop;

			var dstLen = parseInt(dstLen_str);
			var memoryData = AscCommon.Base64.decode(base64_ppty, false, dstLen, index);
			this.stream = new AscCommon.FileStream(memoryData, memoryData.length);
		} else {
			this.stream = new AscCommon.FileStream();
			this.stream.obj    = null;
			this.stream.data   = base64_ppty;
			this.stream.size   = base64_ppty.length;
			//skip header
			this.stream.EnterFrame(index);
			this.stream.Seek2(index);
		}

        this.presentation.ImageMap = {};
        this.presentation.Fonts = [];
        this.fields.length = 0;

        this.LoadDocument();
        if(AscFonts.IsCheckSymbols)
        {
            var bLoad = AscCommon.g_oIdCounter.m_bLoad;
            AscCommon.g_oIdCounter.Set_Load(false);
            for(var nField = 0; nField < this.fields.length; ++nField)
            {
                var oField = this.fields[nField];
                var sValue = oField.private_GetString();
                if(typeof sValue === "string" && sValue.length > 0)
                {
                    AscFonts.FontPickerByCharacter.getFontsByString(sValue);
                }
            }
            AscCommon.g_oIdCounter.Set_Load(bLoad);
        }
        this.fields.length = 0;
        AscFormat.checkPlaceholdersText();

        this.ImageMapChecker = null;
    };
		this.GenerateSmartArts = function () {
			while (this.smartarts.length) {
				const smartart = this.smartarts.pop();
				smartart.generateDrawingPart();
			}
		};
	this.ClearSmartArts = function () {
		this.smartarts.length = 0;
	};

    this.LoadDocument = function()
    {
        // чтение формата ppty
        var _main_tables = {};
        var s = this.stream;
        var err = 0;

        err = s.EnterFrame(5 * 30);
        if (err != 0)
            return err;

        for (var i = 0; i < 30; i++)
        {
            var _type = s.GetUChar();
            if (0 == _type)
                break;

            _main_tables["" + _type] = s.GetULong();
        }

        if (undefined != _main_tables["255"])
        {
            // signature
            s.Seek2(_main_tables["255"]);
            var _sign = s.GetString1(4);
            var _ver = s.GetULong();
        }

        if (!this.IsThemeLoader)
        {
            if (undefined != _main_tables["1"])
            {
                // app
                s.Seek2(_main_tables["1"]);

                this.presentation.App = new AscCommon.CApp();
                this.presentation.App.fromStream(s);
            }

            if (undefined != _main_tables["2"])
            {
                // core
                s.Seek2(_main_tables["2"]);

                this.presentation.Core = new AscCommon.CCore();
                this.presentation.Core.fromStream(s);
            }

            if (undefined != _main_tables["8"])
            {
                // Customs
                let nCustomPos = _main_tables["8"];
                s.Seek2(nCustomPos);

                //let nCustomType = s.GetUChar();
                let nCustomCount = s.GetULong();
                if(nCustomCount > 0) {
                    for(let nRecord = 0; nRecord < nCustomCount; ++nRecord) {

                        let nCustomType = s.GetUChar();
                        s.SkipRecord();
                    }
                }
                this.presentation.CustomXmlData = s.data.slice(nCustomPos, s.cur);
                s.Seek2(nCustomPos);
            }

            if (undefined != _main_tables["48"])
            {
                // CustomProperties
                s.Seek2(_main_tables["48"]);

                this.presentation.CustomProperties.fromStream(s);
            }
        }

        if (undefined != _main_tables["3"])
        {
            // core
            s.Seek2(_main_tables["3"]);

            this.presentation.pres = new AscCommon.CPres();
            var pres = this.presentation.pres;

            pres.fromStream(s, this);

            if(!this.IsThemeLoader)
            {
                if(pres.attrShowSpecialPlsOnTitleSld !== null)
                {
                    this.presentation.setShowSpecialPlsOnTitleSld(pres.attrShowSpecialPlsOnTitleSld);
                }

                if(pres.attrFirstSlideNum !== null)
                {
                    this.presentation.setFirstSlideNum(pres.attrFirstSlideNum);
                }
            }

            this.presentation.defaultTextStyle = pres.defaultTextStyle;
        }

        if (!this.IsThemeLoader)
        {
            if (undefined != _main_tables["4"])
            {
                // view props
                s.Seek2(_main_tables["4"]);
                this.presentation.setViewPr(this.ReadViewProps());
            }

            if (undefined != _main_tables["5"])
            {
                // vmldrawing
                s.Seek2(_main_tables["5"]);
                this.presentation.VmlDrawing = this.ReadVmlDrawing();
            }

            if (undefined != _main_tables["6"])
            {
                // tablestyles
                s.Seek2(_main_tables["6"]);
                this.presentation.TableStyles = this.ReadTableStyles();
            }
            if (undefined != _main_tables["7"])
            {
                // presprops
                s.Seek2(_main_tables["7"]);
                this.ReadPresProps(this.presentation);
            }
        }

        this.aThemes.length = 0;
        if (undefined != _main_tables["20"])
        {
            // themes
            s.Seek2(_main_tables["20"]);

            var _themes_count = s.GetULong();
            for (var i = 0; i < _themes_count; i++)
                this.aThemes[i] = this.ReadTheme();
        }

        if (undefined != _main_tables["22"])
        {
            // slide masters
            s.Seek2(_main_tables["22"]);

            var _sm_count = s.GetULong();
            for (var i = 0; i < _sm_count; i++)
            {
                this.presentation.slideMasters[i] = this.ReadSlideMaster();
                this.presentation.slideMasters[i].setSlideSize(this.presentation.GetWidthMM(), this.presentation.GetHeightMM());
            }
        }

        this.aSlideLayouts.length = 0;
        if (undefined != _main_tables["23"])
        {
            // slide masters
            s.Seek2(_main_tables["23"]);

            var _sl_count = s.GetULong();
            for (var i = 0; i < _sl_count; i++)
            {
                this.aSlideLayouts[i] = this.ReadSlideLayout();
                this.aSlideLayouts[i].setSlideSize(this.presentation.GetWidthMM(), this.presentation.GetHeightMM());
            }
        }

        if (!this.IsThemeLoader)
        {
            if (undefined != _main_tables["24"])
            {
                // slides
                s.Seek2(_main_tables["24"]);

                var _s_count = s.GetULong();
                var bOldVal;
                if(this.Api)
                {
                    bOldVal = this.Api.bNoSendComments;
                    this.Api.bNoSendComments = true;
                }
                for (var i = 0; i < _s_count; i++)
                {
                    this.presentation.insertSlide(i, this.ReadSlide(i)) ;
                }
                if(this.Api)
                {
                    this.Api.bNoSendComments = bOldVal;
                }
            }

            if (undefined != _main_tables["25"])
            {
                // slides
                s.Seek2(_main_tables["25"]);

                var _nm_count = s.GetULong();
                for (var i = 0; i < _nm_count; i++){
                    this.presentation.notesMasters[i] = this.ReadNoteMaster();
                    this.presentation.notesMasters[i].setTheme(this.aThemes[0]);//TODO: убрать после того как будут сделаны рельсы
                }
            }

            if (undefined != _main_tables["26"])
            {
                // slides
                s.Seek2(_main_tables["26"]);

                var _n_count = s.GetULong();
                for (var i = 0; i < _n_count; i++)
                    this.presentation.notes[i] = this.ReadNote();
            }
        }

        // теперь нужно прочитать используемые в презентации шрифты и картинки
        if (null == this.ImageMapChecker)
        {
            if (undefined != _main_tables["42"])
            {
                s.Seek2(_main_tables["42"]);

                var _type = s.GetUChar();
                var _len = s.GetULong();

                s.Skip2(1); // strat attr

                var _cur_ind = 0;

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    var image_id = s.GetString2();
                    if (this.IsThemeLoader)
                    {
                        image_id = "theme" + (this.Api.ThemeLoader.CurrentLoadThemeIndex + 1) + "/media/" + image_id;
                    }

                    this.presentation.ImageMap[_cur_ind++] = image_id;
                }
            }
        }
        else
        {
            var _cur_ind = 0;
            for (var k in this.ImageMapChecker)
            {
                if (this.IsThemeLoader)
                {
                    image_id = "theme" + (this.Api.ThemeLoader.CurrentLoadThemeIndex + 1) + "/media/" + k;
                }

                this.presentation.ImageMap[_cur_ind++] = k;
            }
        }

        if (undefined != _main_tables["43"])
        {
            s.Seek2(_main_tables["43"]);

            var _type = s.GetUChar();
            var _len = s.GetULong();

            s.Skip2(1); // strat attr

            var _cur_ind = 0;

            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                var f_name = s.GetString2();

                this.presentation.Fonts[this.presentation.Fonts.length] = new AscFonts.CFont(f_name);
            }
        }

        // все загружено, осталось расставить связи и загрузить картинки тем и шаблонов
        if (undefined != _main_tables["41"])
        {
            s.Seek2(_main_tables["41"]);

            s.Skip2(5); // type + len

            var _count = s.GetULong();

            for (var i = 0; i < _count; i++)
            {
                var _master_type = s.GetUChar(); // must be 0
                this.ReadMasterInfo(i);
            }
        }

        if (!this.IsThemeLoader)
        {
            if (undefined != _main_tables["40"])
            {
                s.Seek2(_main_tables["40"]);

                s.Skip2(6); // type + len + start attr

                var _slideNum = 0;
                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    var indexL = s.GetULong();
                    this.presentation.Slides[_slideNum].setLayout(this.aSlideLayouts[indexL]);
                    this.presentation.Slides[_slideNum].Master = this.aSlideLayouts[indexL].Master;
                    _slideNum++;
                }
            }
			if (undefined != _main_tables["45"])
			{
				s.Seek2(_main_tables["45"]);
				s.Skip2(6); // type + len + start attr

				var _slideNum = 0;
				while (true)
				{
					var _at = s.GetUChar();
					if (_at == g_nodeAttributeEnd)
						break;

					var indexL = s.GetLong();
					this.presentation.Slides[_slideNum].setNotes(this.presentation.notes[indexL]);
                    ++_slideNum;
				}
			}
			if (undefined != _main_tables["46"])
			{
				s.Seek2(_main_tables["46"]);
				s.Skip2(6); // type + len + start attr

				var _noteNum = 0;
				while (true)
				{
					var _at = s.GetUChar();
					if (_at == g_nodeAttributeEnd)
						break;

					var indexL = s.GetLong();
                    this.presentation.notes[_noteNum].setNotesMaster(this.presentation.notesMasters[indexL]);
					_noteNum++;
				}
			}
			if (undefined != _main_tables["47"])
			{
				s.Seek2(_main_tables["47"]);
				s.Skip2(6); // type + len + start attr

				var _noteMasterNum = 0;
				while (true)
				{
					var _at = s.GetUChar();
					if (_at == g_nodeAttributeEnd)
						break;

					var indexL = s.GetLong();
					var notesMaster = this.presentation.notesMasters[_noteMasterNum];
					var notesMasterTheme = this.aThemes[indexL];
					if (notesMaster && notesMasterTheme) {
						notesMaster.setTheme(notesMasterTheme);
					}
					_noteMasterNum++;
				}
			}
        }
				if (this.IsThemeLoader) {
					this.ClearSmartArts();
				} else {
					this.GenerateSmartArts();
				}

        if (this.Api != null && !this.IsThemeLoader)
        {
            this.presentation.createNecessaryObjectsIfNoPresent();
        }
        else if (this.Api != null && this.IsThemeLoader)
        {
            var theme_loader = this.Api.ThemeLoader;
            var _info = theme_loader.themes_info_editor[theme_loader.CurrentLoadThemeIndex];
            _info.ImageMap = this.presentation.ImageMap;
            _info.FontMap = this.presentation.Fonts;
        }
    };

    this.ReadMasterInfo = function(indexMaster)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        var master = this.presentation.slideMasters[indexMaster];

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    var indexTh = s.GetULong();
                    master.setTheme(this.aThemes[indexTh]);
                    break;
                }
                case 1:
                {
                    s.GetString2A();
                    break;
                }
                default:
                    break;
            }
        }

        var _lay_count = s.GetULong();
        for (var i = 0; i < _lay_count; i++)
        {
            s.Skip2(6); // type + len

            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                switch (_at)
                {
                    case 0:
                    {
                        var indexL = s.GetULong();
                        master.addToSldLayoutLstToPos(master.sldLayoutLst.length, this.aSlideLayouts[indexL]);
                        break;
                    }
                    case 1:
                    {
                        s.GetString2A();
                        break;
                    }
                    default:
                        break;
                }
            }
        }

        s.Seek2(_end_rec);

        if (this.Api != null && this.IsThemeLoader)
        {
            var theme_loader = this.Api.ThemeLoader;

            var theme_load_info = new CThemeLoadInfo();
            theme_load_info.Master = master;
            theme_load_info.Theme = master.Theme;

            var _lay_cnt = master.sldLayoutLst.length;
            for (var i = 0; i < _lay_cnt; i++)
                theme_load_info.Layouts[i] = master.sldLayoutLst[i];

            theme_loader.themes_info_editor[theme_loader.CurrentLoadThemeIndex] = theme_load_info;
        }
    };

    this.ReadViewProps = function()
    {
       //this.stream.SkipRecord();
       //return null;
        let oViewPr = new AscFormat.CViewPr();
        this.stream.GetUChar();
        oViewPr.fromPPTY(this);
        return oViewPr;
    };
    this.ReadVmlDrawing = function()
    {
        return null;
    };
    this.ReadPresProps = function(presentation)
    {
        var s = this.stream;

        var _type = s.GetUChar();

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    s.SkipRecord();
                    break;
                }
                case 1:
                {
                    presentation.showPr = this.ReadShowPr();
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };
    this.ReadShowPr = function()
    {
        var showPr = new CShowPr();
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    showPr.loop = s.GetBool();
                    break;
                }
                case 1:
                {
                    showPr.showAnimation = s.GetBool();
                    break;
                }
                case 2:
                {
                    showPr.showNarration = s.GetBool();
                    break;
                }
                case 3:
                {
                    showPr.useTimings = s.GetBool();
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    //todo browseShowScrollbar
                    showPr.browse = true;
                    s.SkipRecord();
                    break;
                }
                case 1:
                {
                    this.ReadShowPrCustShow(showPr);
                    break;
                }
                case 2:
                {
                    this.ReadShowPrKiosk(showPr);
                    break;
                }
                case 3:
                {
                    showPr.penClr = this.ReadUniColor();
                    break;
                }
                case 4:
                {
                    showPr.present = true;
                    s.SkipRecord();
                    break;
                }
                case 5:
                {
                    if (!showPr.show){
                        showPr.show = {showAll: null, range: null, custShow: null};
                    }
                    showPr.show.showAll = true;
                    s.SkipRecord();
                    break;
                }
                case 6:
                {
                    this.ReadShowPrSldRg(showPr);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return showPr;
    };
    this.ReadShowPrCustShow = function(showPr)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    if (!showPr.show){
                        showPr.show = {showAll: null, range: null, custShow: null};
                    }
                    showPr.show.custShow = s.GetLong();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
    };
    this.ReadShowPrKiosk = function(showPr)
    {
        showPr.kiosk = {restart: null};
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    showPr.kiosk.restart = s.GetLong();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
    };
    this.ReadShowPrSldRg = function(showPr)
    {
        if (!showPr.show){
            showPr.show = {showAll: null, range: null, custShow: null};
        }
        showPr.show.range = {start: null, end: null};
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    showPr.show.range.start = s.GetLong();
                    break;
                }
                case 1:
                {
                    showPr.show.range.end = s.GetLong();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
    };
    this.ReadTableStyles = function()
    {
        //var _styles = this.presentation.globalTableStyles;
        var s = this.stream;

        var _type = s.GetUChar();

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        let sNewDefaultStyleStyleId;
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    sNewDefaultStyleStyleId = s.GetString2();
                    break;
                }
                default:
                    break;
            }
        }

        var _type = s.GetUChar(); // 0!!!
        s.Skip2(4); // len


        while (s.cur < _end_rec)
        {
            s.Skip2(1);
            this.ReadTableStyle();
        }
        if(typeof sNewDefaultStyleStyleId === 'string' && this.presentation && this.presentation.globalTableStyles)
        {
            const oDefaultStyle = this.presentation.globalTableStyles.GetStyleByStyleId(sNewDefaultStyleStyleId);
            if (oDefaultStyle)
            {
                this.presentation.DefaultTableStyleId = oDefaultStyle.Id;
            }
        }
        s.Seek2(_end_rec);
    };

    this.ReadTableStyle = function(bNotAddStyle)
    {
        var s = this.stream;

        var _style = new CStyle("", null, null, styletype_Table);

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    var _id = s.GetString2();
                    _style.SetStyleId(_id);
                   // _style.Id = _id;
					if(AscCommon.isRealObject(this.presentation.TableStylesIdMap) && !bNotAddStyle)
						this.presentation.TableStylesIdMap[_style.Id] = true;

                    if (this.presentation && this.presentation.globalTableStyles)
                    {
                        const oOldStyle = this.presentation.globalTableStyles.GetStyleByStyleId(_id);
                        if (oOldStyle)
                        {
                            this.presentation.globalTableStyles.Remove(oOldStyle.GetId());
                            this.presentation.globalTableStyles.Add(_style);
                        }
                        else
                        {
                            this.map_table_styles[_id] = _style;
                        }
                    }
                    break;
                }
                case 1:
                {
                    _style.Name = s.GetString2();
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var _end_rec2 = s.cur + s.GetLong() + 4;

                    while (s.cur < _end_rec2)
                    {
                        var _at2 = s.GetUChar();
                        switch (_at2)
                        {
                            case 0:
                            {
                                var _end_rec3 = s.cur + s.GetLong() + 4;
                                while (s.cur < _end_rec3)
                                {
                                    var _at3 = s.GetUChar();
                                    switch (_at3)
                                    {
                                        case 0:
                                        {
                                            var _unifill = this.ReadUniFill();
                                            if (_unifill && _unifill.fill !== undefined && _unifill.fill != null)
                                            {
                                                if (undefined === _style.TablePr.Shd || null == _style.TablePr.Shd)
                                                {
                                                    _style.TablePr.Shd = new CDocumentShd();
                                                    _style.TablePr.Shd.Value = c_oAscShdClear;
                                                }
                                                _style.TablePr.Shd.Unifill = _unifill;
                                            }
                                        }
                                        default:
                                        {
                                            s.SkipRecord();
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                            case 1:
                            {
                                if (undefined === _style.TablePr.Shd || null == _style.TablePr.Shd)
                                {
                                    _style.TablePr.Shd = new CDocumentShd();
                                    _style.TablePr.Shd.Value = c_oAscShdClear;
                                }
                                _style.TablePr.Shd.FillRef = this.ReadStyleRef();
                                break;
                            }
                            default:
                            {
                                s.SkipRecord();
                                break;
                            }
                        }
                    }

                    s.Seek2(_end_rec2);
                    break;
                }
                case 1:
                {
                    _style.TableWholeTable = this.ReadTableStylePart();
                    break;
                }
                case 2:
                {
                    _style.TableBand1Horz = this.ReadTableStylePart();
                    break;
                }
                case 3:
                {
                    _style.TableBand2Horz = this.ReadTableStylePart();
                    break;
                }
                case 4:
                {
                    _style.TableBand1Vert = this.ReadTableStylePart();
                    break;
                }
                case 5:
                {
                    _style.TableBand2Vert = this.ReadTableStylePart();
                    break;
                }
                case 6:
                {
                    _style.TableLastCol = this.ReadTableStylePart();
                    break;
                }
                case 7:
                {
                    _style.TableFirstCol = this.ReadTableStylePart();
                    break;
                }
                case 8:
                {
                    _style.TableFirstRow = this.ReadTableStylePart();
                    break;
                }
                case 9:
                {
                    _style.TableLastRow = this.ReadTableStylePart();
                    break;
                }
                case 10:
                {
                    _style.TableBRCell = this.ReadTableStylePart();
                    break;
                }
                case 11:
                {
                    _style.TableBLCell = this.ReadTableStylePart();
                    break;
                }
                case 12:
                {
                    _style.TableTRCell = this.ReadTableStylePart();
                    break;
                }
                case 13:
                {
                    _style.TableTLCell = this.ReadTableStylePart();
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        _style.wholeToTablePr();
		if(bNotAddStyle)
		{
			return _style;
		}
		else
		{
			if(this.presentation && this.presentation.globalTableStyles)
				this.presentation.globalTableStyles.Add(_style);
		}
    };

    this.ReadTableStylePart = function()
    {
        var s = this.stream;

        var _part = new CTableStylePr();

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var _end_rec2 = s.cur + s.GetLong() + 4;

                    s.Skip2(1); // start attributes

                    var _i, _b;
                    while (true)
                    {
                        var _at2 = s.GetUChar();
                        if (_at2 == g_nodeAttributeEnd)
                            break;

                        switch (_at2)
                        {
                            case 0:
                            {
                                _i = s.GetUChar();
                                break;
                            }
                            case 1:
                            {
                                _b = s.GetUChar();
                                break;
                            }
                            default:
                                break;
                        }
                    }

                    if(_i === 0)
                    {
                        _part.TextPr.Italic = true;
                    }
                    else if(_i === 1)
                    {
                        _part.TextPr.Italic = false;
                    }

                    if(_b === 0)
                    {
                        _part.TextPr.Bold = true;
                    }
                    else if(_b === 1)
                    {
                        _part.TextPr.Bold = false;
                    }

                    while (s.cur < _end_rec2)
                    {
                        var _at3 = s.GetUChar();
                        switch (_at3)
                        {
                            case 0:
                            {
                                _part.TextPr.FontRef = this.ReadFontRef();
                                break;
                            }
                            case 1:
                            {
                                var _Unicolor = this.ReadUniColor();
                                if(_Unicolor && _Unicolor.color)
                                {
                                    _part.TextPr.Unifill =  AscFormat.CreateUniFillByUniColor(_Unicolor);
                                }
                                break;
                            }
                            default:
                            {
                                s.SkipRecord();
                                break;
                            }
                        }
                    }

                    s.Seek2(_end_rec2);
                    break;
                }
                case 1:
                {
                    var _end_rec2 = s.cur + s.GetLong() + 4;

                    while (s.cur < _end_rec2)
                    {
                        var _at2 = s.GetUChar();
                        switch (_at2)
                        {
                            case 0:
                            {
                                this.ReadTcBdr(_part);
                                break;
                            }
                            case 1:
                            {
                                if (undefined === _part.TableCellPr.Shd || null == _part.TableCellPr.Shd)
                                {
                                    _part.TableCellPr.Shd = new CDocumentShd();
                                    _part.TableCellPr.Shd.Value = c_oAscShdClear;
                                }
                                _part.TableCellPr.Shd.FillRef = this.ReadStyleRef();
                                break;
                            }
                            case 2:
                            {
                                var _end_rec3 = s.cur + s.GetLong() + 4;
                                while (s.cur < _end_rec3)
                                {
                                    var _at3 = s.GetUChar();
                                    switch (_at3)
                                    {
                                        case 0:
                                        {
                                            var _unifill = this.ReadUniFill();
                                            if (_unifill && _unifill.fill !== undefined && _unifill.fill != null)
                                            {
                                                if (undefined === _part.TableCellPr.Shd || null == _part.TableCellPr.Shd)
                                                {
                                                    _part.TableCellPr.Shd = new CDocumentShd();
                                                    _part.TableCellPr.Shd.Value = c_oAscShdClear;
                                                }
                                                _part.TableCellPr.Shd.Unifill = _unifill;
                                            }
                                            break;
                                        }
                                        default:
                                        {
                                            s.SkipRecord();
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                            case 3:
                            {
                                s.SkipRecord();
                                break;
                            }
                            default:
                            {
                                s.SkipRecord();
                                break;
                            }
                        }
                    }

                    s.Seek2(_end_rec2);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return _part;
    };

    this.ReadTcBdr = function(_part)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    _part.TableCellPr.TableCellBorders.Left = new CDocumentBorder();
                    this.ReadTableBorderLineStyle(_part.TableCellPr.TableCellBorders.Left);
                    break;
                }
                case 1:
                {
                    _part.TableCellPr.TableCellBorders.Right = new CDocumentBorder();
                    this.ReadTableBorderLineStyle(_part.TableCellPr.TableCellBorders.Right);
                    break;
                }
                case 2:
                {
                    _part.TableCellPr.TableCellBorders.Top = new CDocumentBorder();
                    this.ReadTableBorderLineStyle(_part.TableCellPr.TableCellBorders.Top);
                    break;
                }
                case 3:
                {
                    _part.TableCellPr.TableCellBorders.Bottom = new CDocumentBorder();
                    this.ReadTableBorderLineStyle(_part.TableCellPr.TableCellBorders.Bottom);
                    break;
                }
                case 4:
                {
                    _part.TablePr.TableBorders.InsideH = new CDocumentBorder();
                    this.ReadTableBorderLineStyle(_part.TablePr.TableBorders.InsideH);
                    break;
                }
                case 5:
                {
                    _part.TablePr.TableBorders.InsideV = new CDocumentBorder();
                    this.ReadTableBorderLineStyle(_part.TablePr.TableBorders.InsideV);
                    break;
                }
                case 6:
                case 7:
                {
                    s.SkipRecord();
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return _part;
    };

    this.ReadTableBorderLineStyle = function(_border)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var ln = this.ReadLn();
                    ln.fillDocumentBorder(_border);
                    break;
                }
                case 1:
                {
                    _border.LineRef = this.ReadStyleRef();
                    _border.Value = border_Single;
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    // UNICOLOR ---------------------------------

    this.ReadUniColor = function()
    {
        var s = this.stream;
        var _len = s.GetULong();
        var read_start = s.cur;
        var read_end = read_start + _len;

        var uni_color = new AscFormat.CUniColor();

        if (s.cur < read_end)
        {
            var _type = s.GetUChar();

            var _e = s.cur + s.GetULong() + 4;

            switch (_type)
            {
                case c_oAscColor.COLOR_TYPE_PRST:
                {
                    s.Skip2(2);
                    uni_color.setColor(new AscFormat.CPrstColor());
                    uni_color.color.setId(s.GetString2());
                    s.Skip2(1);

                    if (s.cur < _e)
                    {
                        if (0 == s.GetUChar())
                        {
                            uni_color.setMods(this.ReadColorMods());
                        }
                    }

                    break;
                }
                case c_oAscColor.COLOR_TYPE_SCHEME:
                {
                    s.Skip2(2);
                    uni_color.setColor(new AscFormat.CSchemeColor());
                    uni_color.color.setId(s.GetUChar());
                    s.Skip2(1);

                    if (s.cur < _e)
                    {
                        if (0 == s.GetUChar())
                        {
                            uni_color.setMods(this.ReadColorMods());
                        }
                    }

                    break;
                }
                case c_oAscColor.COLOR_TYPE_SRGB:
                {
                    var r, g, b;
                    s.Skip2(1);
                    uni_color.setColor(new AscFormat.CRGBColor());
                    s.Skip2(1);
                    r = s.GetUChar();
                    s.Skip2(1);
                    g = s.GetUChar();
                    s.Skip2(1);
                    b = s.GetUChar();
                    s.Skip2(1);
                    uni_color.color.setColor(r, g, b);
                    if (s.cur < _e)
                    {
                        if (0 == s.GetUChar())
                        {
                            uni_color.setMods(this.ReadColorMods());
                        }
                    }

                    break;
                }
                case c_oAscColor.COLOR_TYPE_SYS:
                {
                    s.Skip2(1);
                    uni_color.setColor(new AscFormat.CSysColor());

                    while (true)
                    {
                        var _at = s.GetUChar();
                        if (_at == g_nodeAttributeEnd)
                            break;

                        switch (_at)
                        {
                            case 0:
                            {
                                uni_color.color.setId(s.GetString2());
                                break;
                            }
                            case 1:
                            {
                                uni_color.color.setR(s.GetUChar());
                                break;
                            }
                            case 2:
                            {
                                uni_color.color.setG(s.GetUChar());
                                break;
                            }
                            case 3:
                            {
                                uni_color.color.setB(s.GetUChar());
                                break;
                            }
                            default:
                                break;
                        }
                    }

                    if (s.cur < _e)
                    {
                        if (0 == s.GetUChar())
                        {
                            uni_color.setMods(this.ReadColorMods());
                        }
                    }

                    break;
                }
                case c_oAscColor.COLOR_TYPE_STYLE:
                {
                    var oColor = new AscFormat.CStyleColor();
                    s.Skip2(1);
                    while(true)
                    {
                        var _at2 = s.GetUChar();
                        if (_at2 == g_nodeAttributeEnd)
                            break;
                        switch (_at2) {
                            case 0:
                            {
                                oColor.val = s.GetULong();
                                break;
                            }
                            case 1:
                            {
                                oColor.bAuto = s.GetBool();
                                break;
                            }
                            default:
                            {
                                break;
                            }
                        }
                    }
                    uni_color.setColor(oColor);
                    break;
                }
            }
        }

        if(!uni_color.color){
            return null;
        }

        s.Seek2(read_end);
        return uni_color;
    };

    this.ReadColorMods = function()
    {
        var ret = new AscFormat.CColorModifiers();
        var _mods = this.ReadColorModifiers();
        if(_mods)
        {
            for(var i = 0; i < _mods.length; ++i)
                ret.addMod(_mods[i]);
        }
        return ret;
    };

    this.ReadColorMod = function()
    {
        var s = this.stream;
        var _s1 = s.cur;
        var _e1 = _s1 + s.GetULong() + 4;
        var _mod = null;
        if (_s1 < _e1)
        {
            s.Skip2(1);
            _mod = new AscFormat.CColorMod();
            while (true)
            {
                var _type = s.GetUChar();

                if (0 == _type)
                {
                    _mod.setName(s.GetString2());
                    var _find = _mod.name.indexOf(":");
                    if (_find >= 0 && _find < (_mod.name.length - 1))
                        _mod.setName(_mod.name.substring(_find + 1));
                }
                else if (1 == _type)
                    _mod.setVal(s.GetLong());
                else if (g_nodeAttributeEnd == _type)
                    break;
                else
                    break;
            }
        }
        s.Seek2(_e1);
        return _mod;
    };

    this.ReadColorModifiers = function()
    {
        var s = this.stream;
        var _start = s.cur;
        var _end = _start + s.GetULong() + 4;

        var _ret = null;

        var _count = s.GetULong();
        for (var i = 0; i < _count; i++)
        {
            if (s.cur > _end)
                break;

            s.Skip2(1);

            var _mod = this.ReadColorMod();
            if(_mod)
            {
                if (null == _ret)
                    _ret = [];
                _ret[_ret.length] = _mod;
            }
        }

        s.Seek2(_end);
        return _ret;
    };

    // ------------------------------------------

    // UNIFILL ----------------------------------

    this.ReadRect = function(bIsMain)
    {
        var _ret = {};

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    _ret.l = s.GetPercentage();
                    break;
                }
                case 1:
                {
                    _ret.t = s.GetPercentage();
                    break;
                }
                case 2:
                {
                    _ret.r = s.GetPercentage();
                    break;
                }
                case 3:
                {
                    _ret.b = s.GetPercentage();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);

        if (null == _ret.l && null == _ret.t && null == _ret.r && null == _ret.b)
            return null;

        if (_ret.l == null)
            _ret.l = 0;
        if (_ret.t == null)
            _ret.t = 0;
        if (_ret.r == null)
            _ret.r = 0;
        if (_ret.b == null)
            _ret.b = 0;

        if (!bIsMain)
        {
            var _absW = Math.abs(_ret.l) + Math.abs(_ret.r) + 100;
            var _absH = Math.abs(_ret.t) + Math.abs(_ret.b) + 100;

            _ret.l = -100 * _ret.l / _absW;
            _ret.t = -100 * _ret.t / _absH;
            _ret.r = -100 * _ret.r / _absW;
            _ret.b = -100 * _ret.b / _absH;
        }

        _ret.r = 100 - _ret.r;
        _ret.b = 100 - _ret.b;

        if (_ret.l > _ret.r)
        {
            var tmp = _ret.l;
            _ret.l = _ret.r;
            _ret.r = tmp;
        }
        if (_ret.t > _ret.b)
        {
            var tmp = _ret.t;
            _ret.t = _ret.b;
            _ret.b = tmp;
        }
        var ret = new AscFormat.CSrcRect();
        ret.setLTRB(_ret.l, _ret.t, _ret.r, _ret.b);
        return ret;
    };

    this.ReadGradLin = function()
    {
        var _lin = new AscFormat.GradLin();
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    _lin.setAngle(s.GetLong());
                    break;
                }
                case 1:
                {
                    _lin.setScale(s.GetBool());
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
        return _lin;
    };

    this.ReadGradPath = function()
    {
        var _path = new AscFormat.GradPath();
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    _path.setPath(s.GetUChar());
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
        return _path;
    };


    this.ReadBlur = function()
    {
        var nRecStart, nRecLen, nRecEnd;
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        nRecStart = s.cur;
        nRecLen = s.GetLong();
        nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.CBlur();
        s.Skip2(1);

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:	oEffect.rad = s.GetULong(); break;
                case 1:	oEffect.grow = s.GetBool(); break;
            }
        }
        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadFillOverlay = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var nRecStart, nRecLen, nRecEnd;
        nRecStart = s.cur;
        nRecLen = s.GetLong();
        nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.CFillOverlay();
        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (_at == 0)
                oEffect.blend = s.GetUChar();
            else break;
        }

        while (s.cur < nRecEnd)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    oEffect.fill = this.ReadUniFill();
                    break;
                }
                default:
                    break;
            }
        }
        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadGlow = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var nRecStart = s.cur;
        var nRecLen = s.GetLong();
        var nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.CGlow();
        s.Skip2(1);

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (_at == 0)
                oEffect.rad = s.GetLong();
            else break;
        }
        while (s.cur < nRecEnd)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    oEffect.color = this.ReadUniColor();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadInnerShdw = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var nRecStart = s.cur;
        var nRecLen = s.GetLong();
        var nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.CInnerShdw();
        s.Skip2(1);

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:	oEffect.dir = s.GetLong(); break;
                case 1:	oEffect.dist = s.GetLong(); break;
                case 2:	oEffect.blurRad = s.GetLong(); break;
            }
        }
        while (s.cur < nRecEnd)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    oEffect.color = this.ReadUniColor();
                    break;
                }
                default:
                    break;
            }
        }


        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadOuterShdw = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();

        var nRecStart = s.cur;
        var nRecLen = s.GetLong();
        var nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.COuterShdw();
        s.Skip2(1);

        while (true)
        {
            var  _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0: oEffect.algn = s.GetUChar(); break;
                case 1:	oEffect.blurRad = s.GetLong(); break;
                case 2:	oEffect.dir		= s.GetLong(); break;
                case 3:	oEffect.dist	= s.GetLong(); break;
                case 4:	oEffect.kx		= s.GetLong(); break;
                case 5:	oEffect.ky		= s.GetLong(); break;
                case 6:	oEffect.sx		= s.GetLong(); break;
                case 7:	oEffect.sy		= s.GetLong(); break;
                case 8:	oEffect.rotWithShape = s.GetBool(); break;
            }
        }
        while (s.cur < nRecEnd)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    oEffect.color = this.ReadUniColor();
                    break;
                }
                default:
                    break;
            }
        }
        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadPrstShdw = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var nRecStart = s.cur;
        var nRecLen = s.GetLong();
        var nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.CPrstShdw();
        s.Skip2(1);

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:	oEffect.dir		= s.GetLong(); break;
                case 1:	oEffect.dist	= s.GetLong(); break;
                case 2:	oEffect.prst = s.GetUChar(); break;
            }

        }
        while (s.cur < nRecEnd)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    oEffect.color = this.ReadUniColor();
                    break;
                }
                default:
                    break;
            }
        }
        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadReflection = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var nRecStart = s.cur;
        var nRecLen = s.GetLong();
        var nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.CReflection();
        s.Skip2(1);

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    oEffect.algn = ( s.GetUChar());
                }break;
                case 1:	oEffect.blurRad = s.GetLong(); break;
                case 2:	oEffect.stA		= s.GetLong(); break;
                case 3:	oEffect.endA	= s.GetLong(); break;
                case 4:	oEffect.stPos	= s.GetLong(); break;
                case 5:	oEffect.endPos	= s.GetLong(); break;
                case 6:	oEffect.dir		= s.GetLong(); break;
                case 7:	oEffect.fadeDir	= s.GetLong(); break;
                case 8:	oEffect.dist	= s.GetLong(); break;
                case 9:	oEffect.kx		= s.GetLong(); break;
                case 10:oEffect.ky		= s.GetLong(); break;
                case 11:oEffect.sx		= s.GetLong(); break;
                case 12:oEffect.sy		= s.GetLong(); break;
                case 13:oEffect.rotWithShape = s.GetBool(); break;
            }
        }

        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadSoftEdge = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var nRecStart = s.cur;
        var nRecLen = s.GetLong();
        var nRecEnd = nRecStart + nRecLen + 4;
        var oEffect = new AscFormat.CSoftEdge();
        s.Skip2(1);

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (_at == 0) oEffect.rad = s.GetULong();
            else break;
        }

        s.Seek2(nRecEnd);
        return oEffect;
    };

    this.ReadEffect = function()
    {

        var s = this.stream;
        var pos = s.cur;
        var nUniEffectLength = s.GetLong(); // len
        if(nUniEffectLength === 0)
        {
            return null;
        }
        var  nEffectType = s.GetUChar();
        s.Seek2(pos);
        var nRecStart, nRecLen, nRecEnd;
        var oEffect = null;
        switch(nEffectType)
        {
            case 0:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_NONE			=
            }
            case 1:
            {
                oEffect = this.ReadOuterShdw();
                break;//var  EFFECT_TYPE_OUTERSHDW		=
            }
            case 2:
            {
                oEffect = this.ReadGlow();
                break;//var  EFFECT_TYPE_GLOW			=
            }
            case 3:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CDuotone();
                var count = s.GetULong();
                for (var i = 0; i < count; ++i)
                {
                    s.Skip2(1); // type

                    var oUniColor = this.ReadUniColor();
                    if(oUniColor.color)
                    {
                        oEffect.colors.push(oUniColor);
                    }
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_DUOTONE		=
            }
            case 4:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CXfrmEffect();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    switch (_at)
                    {
                        case 0:	oEffect.kx	= s.GetLong(); break;
                        case 1: oEffect.ky	= s.GetLong(); break;
                        case 2: oEffect.sx	= s.GetLong(); break;
                        case 3: oEffect.sy	= s.GetLong(); break;
                        case 4: oEffect.tx	= s.GetULong(); break;
                        case 5: oEffect.ty	= s.GetULong(); break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_XFRM			=
            }
            case 5:
            {
                oEffect = this.ReadBlur();
                break;//var  EFFECT_TYPE_BLUR			=
            }
            case 6:
            {
                oEffect = this.ReadPrstShdw();
                break;//var  EFFECT_TYPE_PRSTSHDW		=
            }
            case 7:
            {
                oEffect = this.ReadInnerShdw();
                break;//var  EFFECT_TYPE_INNERSHDW		=
            }
            case 8:
            {
                oEffect = this.ReadReflection();
                break;//var  EFFECT_TYPE_REFLECTION		=
            }
            case 9:
            {
                oEffect = this.ReadSoftEdge();
                break;//var  EFFECT_TYPE_SOFTEDGE		=
            }
            case 10:
            {
                oEffect = this.ReadFillOverlay();
                break;//var  EFFECT_TYPE_FILLOVERLAY	=
            }
            case 11:
            {
                s.GetLong();
                s.GetUChar();

                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaCeiling();

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_ALPHACEILING	=
            }
            case 12:
            {
                s.GetLong();
                s.GetUChar();

                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaFloor();

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_ALPHAFLOOR		=
            }
            case 13:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CTintEffect();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    switch (_at)
                    {
                        case 0:	oEffect.amt = s.GetLong(); break;
                        case 1:	oEffect.hue = s.GetLong(); break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_TINTEFFECT		=
            }
            case 14:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CRelOff();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    switch (_at)
                    {
                        case 0:	oEffect.tx	= s.GetLong(); break;
                        case 1:	oEffect.ty	= s.GetLong(); break;
                    }

                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_RELOFF			=
            }
            case 15:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CLumEffect();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    switch (_at)
                    {
                        case 0:	oEffect.bright = s.GetLong(); break;
                        case 1:	oEffect.contrast = s.GetLong(); break;
                    }
                }


                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_LUM			=
            }
            case 16:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CHslEffect();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    switch (_at)
                    {
                        case 0:
                            oEffect.hue = s.GetLong(); break;
                        case 1:
                            oEffect.lum = s.GetLong(); break;
                        case 2:
                            oEffect.sat = s.GetLong(); break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_HSL			=
            }
            case 17:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CGrayscl();

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_GRAYSCL		=
            }
            case 18:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CEffectElement();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (_at == 0)
                        oEffect.ref = s.GetString2();
                    else break;
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_ELEMENT		=
            }
            case 19:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaRepl();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (_at == 0)
                        oEffect.a = s.GetLong();
                    else break;
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_ALPHAREPL		=
            }
            case 20:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaOutset();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (_at == 0)
                        oEffect.rad = s.GetULong();
                    else break;
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_ALPHAOUTSET	=
            }
            case 21:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaModFix();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (_at == 0)
                        oEffect.amt = s.GetLong();
                    else break;
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_ALPHAMODFIX	=
            }
            case 22:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaBiLevel();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (_at == 0)
                        oEffect.thresh = s.GetLong();
                    else break;
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_ALPHABILEVEL	=
            }
            case 23:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CBiLevel();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (_at == 0)
                        oEffect.thresh = s.GetLong();
                    else break;
                }

                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_BILEVEL		=
            }
            case 24:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CEffectContainer();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    switch (_at)
                    {
                        case 0:
                            oEffect.name = s.GetString2(); break;
                        case 1:
                        {
                            oEffect.type = (s.GetUChar());
                        }break;
                    }
                }
                while (s.cur < nRecEnd)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            var count_effects2 = s.GetULong();
                            for (var _eff2 = 0; _eff2 < count_effects2; ++_eff2)
                            {
                                s.Skip2(1); // type
                                var eff2 = this.ReadEffect();
                                if(!eff2)
                                {
                                    oEffect.effectList.push(eff2);
                                }
                            }
                        }break;
                        default:
                            break;
                    }
                }


                s.Seek2(nRecEnd);
                break;//var  EFFECT_TYPE_DAG			=
            }
            case 25:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CFillEffect();
                s.Skip2(1); // start attributes

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;
                }

                while (s.cur < nRecEnd)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            oEffect.fill = this.ReadUniFill();
                            break;
                        }
                        default:
                            break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var EFFECT_TYPE_FILL			=
            }
            case 26:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CClrRepl();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;
                }
                while (s.cur < nRecEnd)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            oEffect.color = this.ReadUniColor();
                            break;
                        }
                        default:
                            break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var EFFECT_TYPE_CLRREPL		=
            }
            case 27:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CClrChange();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    switch (_at)
                    {
                        case 0:
                            oEffect.useA = s.GetBool(); break;
                    }
                }
                while (s.cur < nRecEnd)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            oEffect.clrFrom = this.ReadUniColor();
                        }break;
                        case 1:
                        {
                            oEffect.clrTo = this.ReadUniColor();
                        }break;
                        default:
                            break;
                    }
                }


                s.Seek2(nRecEnd);
                break;//var EFFECT_TYPE_CLRCHANGE		=
            }
            case 28:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaInv();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;
                }
                while (s.cur < nRecEnd)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            oEffect.color = this.ReadUniColor();
                            break;
                        }
                        default:
                            break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var EFFECT_TYPE_ALPHAINV		=
            }
            case 29:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CAlphaMod();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;
                }
                while (s.cur < nRecEnd)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            oEffect.cont = this.ReadEffectDag();
                            break;
                        }
                        default:
                            break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var EFFECT_TYPE_ALPHAMOD		=
            }
            case 30:
            {
                s.GetLong();
                s.GetUChar();
                nRecStart = s.cur;
                nRecLen = s.GetLong();
                nRecEnd = nRecStart + nRecLen + 4;
                oEffect = new AscFormat.CBlend();
                s.Skip2(1);

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (_at == 0)
                        oEffect.blend = (s.GetUChar());
                    else break;
                }
                while (s.cur < nRecEnd)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            oEffect.cont = this.ReadEffectDag();
                            break;
                        }
                        default:
                            break;
                    }
                }

                s.Seek2(nRecEnd);
                break;//var EFFECT_TYPE_BLEND			=
            }
            default:
            {
                s.SkipRecord();
                break;//var
            }
        }

        return oEffect;
    };

    this.ReadEffectDag = function ()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var _start_pos = s.cur;
        var _end_rec = _start_pos + s.GetLong() + 4;
        s.Skip(1);

        var ret = new AscFormat.CEffectContainer();
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    ret.name = s.GetString2(); break;
                }
                case 1:
                {
                    ret.type = (s.GetUChar()); break;
                }
            }
        }
        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var count_effects = s.GetULong();
                    for (var _eff = 0; _eff < count_effects; ++_eff)
                    {
                        s.Skip2(1); // type
                        var effect = this.ReadEffect();
                        if(effect)
                        {
                            ret.effectList.push(effect);
                        }
                    }
                }break;
                default:
                    break;
            }
        }

        s.Seek(_end_rec);
        return ret;
    };

    this.ReadBlip = function (uni_fill, oSpPr, oImageShape, oLn, oParagraph, oBullet) {
        var s = this.stream;
        var _s2 = s.cur;
        var _e2 = _s2 + s.GetLong() + 4;

        s.Skip2(1);

        while (true)
        {
            var _at = s.GetUChar();
            if (g_nodeAttributeEnd == _at)
                break;

            if (_at == 0)
                s.Skip2(1);
        }

        while (s.cur < _e2)
        {
            var _t = s.GetUChar();

            switch (_t)
            {
                case 0:
                case 1:
                {
                    // id. embed / link
                    s.Skip2(4);
                    break;
                }
                case 10:
                case 11:
                {
                    // id. embed / link
                    s.GetString2();
                    break;
                }
                case 2:
                {
                    var len2 = s.GetLong();

                    var  _end_rec_effect = s.cur + len2;

                    var count_effects = s.GetULong();
                    for (var _eff = 0; _eff < count_effects; ++_eff)
                    {

                        s.Skip2(1); // type
                        var oEffect = this.ReadEffect();
                        if(oEffect)
                        {
                            uni_fill.fill.Effects.push(oEffect);
                        }
                    }
                    s.Seek2(_end_rec_effect);
                    break;
                }
                case 3:
                {
                    s.Skip2(6); // len + start attributes + type

                    var sReadPath = s.GetString2();
                    if (this.IsUseFullUrl && this.insertDocumentUrlsData && this.insertDocumentUrlsData.imageMap) {
                        var sReadPathNew = this.insertDocumentUrlsData.imageMap[AscCommon.g_oDocumentUrls.mediaPrefix + sReadPath];
                        if(sReadPathNew){
                            sReadPath = sReadPathNew;
                        }
                    }
                    if(this.IsUseFullUrl) {
                        if(window["native"] && window["native"]["CopyTmpToMedia"]){
                            if(!(window.documentInfo && window.documentInfo["iscoauthoring"])){
                                var sMedia = window["native"]["CopyTmpToMedia"](sReadPath);
                                if(typeof sMedia === "string" && sMedia.length > 0){
                                    sReadPath = sMedia;
                                }
                            }
                        }
                    }
                    uni_fill.fill.setRasterImageId(sReadPath);

                    // TEST version ---------------
                    var _s = sReadPath;
                    var indS = _s.lastIndexOf("emf");
                    if (indS == -1)
                        indS = _s.lastIndexOf("wmf");

                    if (indS != -1 && (indS == (_s.length - 3)))
                    {
                        _s = _s.substring(0, indS);
                        _s += "svg";
                        sReadPath = _s;
                        uni_fill.fill.setRasterImageId(_s);
                    }
                    // ----------------------------

                    if (this.IsThemeLoader)
                    {
                        sReadPath = "theme" + (this.Api.ThemeLoader.CurrentLoadThemeIndex + 1) + "/media/" + sReadPath;
                        uni_fill.fill.setRasterImageId(sReadPath);
                    }

                    if (this.ImageMapChecker != null)
                    {
                        let bAddToMap = true;
                        if(oImageShape && oImageShape instanceof AscFormat.COleObject)
                        {
                            if(sReadPath.indexOf(".") === -1)
                            {
                                bAddToMap = false;
                            }
                        }
                        if(bAddToMap)
                        {
                            this.ImageMapChecker[sReadPath] = true;
                        }
                    }

                    if (this.IsUseFullUrl)
                        this.RebuildImages.push(new CBuilderImages(uni_fill.fill, sReadPath, oImageShape, oSpPr, oLn, undefined, undefined, undefined, oParagraph, oBullet));

                    s.Skip2(1); // end attribute
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_e2);
    }


    this.ReadUniFill = function(oSpPr, oImageShape, oLn, oParagraph, oBullet)
    {
        var s = this.stream;
        var read_start = s.cur;
        var read_end = read_start + s.GetULong() + 4;

        var uni_fill = new AscFormat.CUniFill();

        if (s.cur < read_end)
        {
            var _type = s.GetUChar();
            var _e = s.cur + s.GetULong() + 4;

            switch (_type)
            {
                case c_oAscFill.FILL_TYPE_BLIP:
                {
                    s.Skip2(1);

                    uni_fill.setFill(new AscFormat.CBlipFill());

                    while (true)
                    {
                        var _at = s.GetUChar();
                        if (_at == g_nodeAttributeEnd)
                            break;

                        switch (_at)
                        {
                            case 0:
                                s.Skip2(4); // dpi
                                break;
                            case 1:
                                s.Skip2(1); // rotWithShape
                                break;
                            default:
                                break;
                        }
                    }

                    while (s.cur < _e)
                    {
                        var rec = s.GetUChar();

                        switch (rec)
                        {
                            case 0:
                            {
                                this.ReadBlip(uni_fill, oSpPr, oImageShape, oLn, oParagraph, oBullet);
                               break;
                            }
                            case 1:
                            {
                                uni_fill.fill.setSrcRect(this.ReadRect(true));
                                break;
                            }
                            case 2:
                            {
                                var oBlipTile = new AscFormat.CBlipFillTile();


                                var s = this.stream;

                                var _rec_start = s.cur;
                                var _end_rec = _rec_start + s.GetLong() + 4;

                                s.Skip2(1); // start attributes

                                while (true)
                                {
                                    var _at = s.GetUChar();
                                    if (_at == g_nodeAttributeEnd)
                                break;

                                    switch (_at)
                                    {
                                        case 0:
                                        {
                                            oBlipTile.sx = s.GetLong();
                                            break;
                            }
                                        case 1:
                                        {
                                            oBlipTile.sy = s.GetLong();
                                            break;
                                        }
                                        case 2:
                                        {
                                            oBlipTile.tx = s.GetLong();
                                            break;
                                        }
                            case 3:
                            {
                                            oBlipTile.ty = s.GetLong();
                                            break;
                                        }
                                        case 4:
                                        {
                                            oBlipTile.algn = s.GetUChar();
                                            break;
                                        }
                                        case 5:
                                        {
                                            oBlipTile.flip = s.GetUChar();
                                            break;
                                        }
                                        default:
                                        {
                                            break;
                                        }
                                    }
                                }
                                s.Seek2(_end_rec);
                                uni_fill.fill.setTile(oBlipTile);
                                break;
                            }
                            case 3:
                            {
                                var _e2 = s.cur + s.GetLong() + 4;

                                while (s.cur < _e2)
                                {
                                    var _t = s.GetUChar();

                                    switch (_t)
                                    {
                                        case 0:
                                        {
                                            var _srcRect = this.ReadRect(false);
                                            if (_srcRect != null)
                                                uni_fill.fill.setSrcRect(_srcRect);
                                            break;
                                        }
                                        default:
                                        {
                                            s.SkipRecord();
                                            break;
                                        }
                                    }
                                }

                                s.Seek2(_e2);
                                break;
                            }
                            case 101:
                            {
                              var oBuilderImages = this.RebuildImages[this.RebuildImages.length - 1];
                              if (this.IsUseFullUrl && oBuilderImages) {
                                s.Skip2(4);
                                var urlsCount = s.GetUChar();
                                for (var i = 0; i < urlsCount; ++i) {
                                  oBuilderImages.AdditionalUrls.push(s.GetString2());
                                }
                              } else {
                                s.SkipRecord();
                              }
                              break;
                            }
                            default:
                            {
                                // пока никаких настроек градиента нет
                                var _len = s.GetULong();
                                s.Skip2(_len);
                            }
                        }
                    }

                    break;
                }
                case c_oAscFill.FILL_TYPE_GRAD:
                {
                    s.Skip2(1);

                    uni_fill.setFill(new AscFormat.CGradFill());

                    while (true)
                    {
                        var _at = s.GetUChar();
                        if (_at == g_nodeAttributeEnd)
                            break;

                        switch (_at)
                        {
                            case 0:
                                s.Skip2(1);
                                break;
                            case 1:
                                uni_fill.fill.rotateWithShape = s.GetBool();
                                break;
                            default:
                                break;
                        }
                    }

                    while (s.cur < _e)
                    {
                        var rec = s.GetUChar();

                        switch (rec)
                        {
                            case 0:
                            {
                                var _s1 = s.cur;
                                var _e1 = _s1 + s.GetULong() + 4;

                                var _count = s.GetULong();
                                var colors_ = [];
                                for (var i = 0; i < _count; i++)
                                {
                                    if (s.cur >= _e1)
                                        break;

                                    s.Skip2(1); // type
                                    s.Skip2(4); // len

                                    var _gs = new AscFormat.CGs();

                                    s.Skip2(1); // start attr
                                    s.Skip2(1); // pos type
                                    _gs.pos = s.GetLong();
                                    s.Skip2(1); // end attr

                                    s.Skip2(1);
                                    _gs.color = this.ReadUniColor();

                                    colors_[colors_.length] = _gs;
                                }

                                s.Seek2(_e1);
                                colors_.sort(function(a,b){return a.pos- b.pos;});

                                for(var z = 0; z < colors_.length; ++z)
                                {
                                    uni_fill.fill.addColor(colors_[z]);
                                }
                                break;
                            }
                            case 1:
                            {
                                uni_fill.fill.setLin(this.ReadGradLin());
                                break;
                            }
                            case 2:
                            {
                                uni_fill.fill.setPath(this.ReadGradPath());
                                break;
                            }
                            case 3:
                            {
                                s.SkipRecord();
                                break;
                            }
                            default:
                            {
                                // пока никаких настроек градиента нет
                                var _len = s.GetULong();
                                s.Skip2(_len);
                            }
                        }


                    }
                    if (null != uni_fill.fill.lin && null != uni_fill.fill.path)
                    {
                        // ms office не открывает такие файлы.
                        uni_fill.fill.setPath(null);
                    }

                    if(uni_fill.fill.colors.length < 2)
                    {
                        if(uni_fill.fill.colors.length === 1)
                        {
                            var oUniColor = uni_fill.fill.colors[0].color;
                            uni_fill.fill = new AscFormat.CSolidFill();
                            uni_fill.fill.color = oUniColor;
                        }
                        else
                        {
                            uni_fill.fill = new AscFormat.CSolidFill();
                            uni_fill.fill.color =  AscFormat.CreateUniColorRGB(0, 0, 0);
                        }
                    }

                    break;
                }
                case c_oAscFill.FILL_TYPE_PATT:
                {
                    uni_fill.setFill(new AscFormat.CPattFill());

                    s.Skip2(1);
                    while (true)
                    {
                        var _atPF = s.GetUChar();
                        if (_atPF == g_nodeAttributeEnd)
                            break;

                        switch (_atPF)
                        {
                            case 0:
                            {
                                uni_fill.fill.setFType(s.GetUChar());
                                break;
                            }
                            default:
                                break;
                        }
                    }

                    while (s.cur < _e)
                    {
                        var rec = s.GetUChar();

                        switch (rec)
                        {
                            case 0:
                            {
                                uni_fill.fill.setFgColor(this.ReadUniColor());
                                break;
                            }
                            case 1:
                            {
                                uni_fill.fill.setBgColor(this.ReadUniColor());
                                break;
                            }
                            default:
                            {
                                // пока никаких настроек градиента нет
                                s.SkipRecord();
                            }
                        }
                    }

                    break;
                }
                case c_oAscFill.FILL_TYPE_SOLID:
                {
                    s.Skip2(1); // type + len

                    uni_fill.setFill(new AscFormat.CSolidFill());
                    uni_fill.fill.setColor(this.ReadUniColor());
                    if(uni_fill.fill.color && !uni_fill.fill.color.Mods){
                        uni_fill.fill.color.setMods(new AscFormat.CColorModifiers());
                    }
                    break;
                }
                case c_oAscFill.FILL_TYPE_NOFILL:
                {
                    uni_fill.setFill(new AscFormat.CNoFill());
                    break;
                }
                case c_oAscFill.FILL_TYPE_GRP:
                {
                    uni_fill.setFill(new AscFormat.CGrpFill());
                    break;
                }
            }
        }

        s.Seek2(read_end);
        if(!uni_fill.fill){
            return null;
        }
        uni_fill.checkTransparent();
        return uni_fill;
    };

    this.initAfterBlipFill = function(sReadPath, blipFill)
    {
        if (this.IsUseFullUrl && this.insertDocumentUrlsData && this.insertDocumentUrlsData.imageMap) {
            var sReadPathNew = this.insertDocumentUrlsData.imageMap[AscCommon.g_oDocumentUrls.mediaPrefix + sReadPath];
            if(sReadPathNew){
                sReadPath = sReadPathNew;
            }
        }
        if(this.IsUseFullUrl) {
            if(window["native"] && window["native"]["CopyTmpToMedia"]){
                if(!(window.documentInfo && window.documentInfo["iscoauthoring"])){
                    var sMedia = window["native"]["CopyTmpToMedia"](sReadPath);
                    if(typeof sMedia === "string" && sMedia.length > 0){
                        sReadPath = sMedia;
                    }
                }
            }
        }
        blipFill.setRasterImageId(sReadPath);

        let oApi = Asc.editor || editor;
        if(!oApi.isOpenOOXInBrowser) {
            // TEST version ---------------
            var _s = sReadPath;
            var indS = _s.lastIndexOf("emf");
            if (indS == -1)
                indS = _s.lastIndexOf("wmf");

            if (indS != -1 && (indS == (_s.length - 3)))
            {
                _s = _s.substring(0, indS);
                _s += "svg";
                sReadPath = _s;
                blipFill.setRasterImageId(_s);
            }
            // ----------------------------
        }

        if (this.IsThemeLoader)
        {
            sReadPath = "theme" + (this.Api.ThemeLoader.CurrentLoadThemeIndex + 1) + "/media/" + sReadPath;
            blipFill.setRasterImageId(sReadPath);
        }

        if (this.ImageMapChecker != null)
            this.ImageMapChecker[sReadPath] = true;

        if (this.IsUseFullUrl)
            this.RebuildImages.push(new CBuilderImages(blipFill, sReadPath));
    }
    // ------------------------------------------

    // COLOR SCHEME -----------------------------

    this.ReadExtraColorScheme = function()
    {
        var extra = new AscFormat.ExtraClrScheme();

        var s = this.stream;
        var _e = s.cur + s.GetULong() + 4;

        while (s.cur < _e)
        {
            var _rec = s.GetUChar();

            switch (_rec)
            {
                case 0:
                {
                    extra.setClrScheme(new AscFormat.ClrScheme());
                    this.ReadClrScheme(extra.clrScheme);
                    break;
                }
                case 1:
                {
                    extra.setClrMap(new AscFormat.ClrMap());
                    this.ReadClrMap(extra.clrMap);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_e);
        return extra;
    };

    this.ReadClrScheme = function(clrscheme)
    {
        var s = this.stream;
        var _e = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start attribute

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                clrscheme.setName(s.GetString2());
        }

        while (s.cur < _e)
        {
            var _rec = s.GetUChar();

            clrscheme.addColor(_rec,this.ReadUniColor());
        }

        s.Seek2(_e);
    };

    this.ReadClrMap = function(clrmap)
    {
        var s = this.stream;
        var _e = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start sttribute

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            clrmap.setClr(_at, s.GetUChar());
        }

        s.Seek2(_e);
    };

    this.ReadClrOverride = function()
    {
        var s = this.stream;
        var _e = s.cur + s.GetULong() + 4;

        var clr_map = null;
        if (s.cur < _e)
        {
            clr_map = new AscFormat.ClrMap();
            s.Skip2(1); // "0"-rectype
            this.ReadClrMap(clr_map);
        }

        s.Seek2(_e);
        return clr_map;
    };

    // ------------------------------------------

    // LINE PROPERTIES --------------------------

    this.ReadLn = function(spPr)
    {
        var ln = new AscFormat.CLn();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;


        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    ln.setAlgn(s.GetUChar());
                    break;
                }
                case 1:
                {
                    ln.setCap(s.GetUChar());
                    break;
                }
                case 2:
                {
                    ln.setCmpd(s.GetUChar());
                    break;
                }
                case 3:
                {
                    ln.setW(s.GetLong());
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    // themeElements
                    ln.setFill(this.ReadUniFill(spPr, null, ln));
                    break;
                }
                case 1:
                {
                    ln.setPrstDash(this.ReadLineDash());
                    break;
                }
                case 2:
                {
                    ln.setJoin(this.ReadLineJoin());
                    break;
                }
                case 3:
                {
                    ln.setHeadEnd(this.ReadLineEnd());
                    break;
                }
                case 4:
                {
                    ln.setTailEnd(this.ReadLineEnd());
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return ln;
    };

    this.ReadLineEnd = function()
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        var endL = new AscFormat.EndArrow();

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    endL.setType(s.GetUChar());
                    break;
                }
                case 1:
                {
                    endL.setW(s.GetUChar());
                    break;
                }
                case 2:
                {
                    endL.setLen(s.GetUChar());
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
        return endL;
    };

    this.ReadLineDash = function()
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        var _dash = 6; // solid

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    _dash = s.GetUChar();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
        return _dash;
    };

    this.ReadLineJoin = function()
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        var join = new AscFormat.LineJoin();

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    join.setType(s.GetLong());
                    break;
                }
                case 1:
                {
                    join.setLimit(s.GetLong());
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
        if( AscFormat.LineJoinType.Empty === join.type ||
            AscFormat.LineJoinType.Round === join.type ||
            AscFormat.LineJoinType.Bevel === join.type ||
            AscFormat.LineJoinType.Miter === join.type ||
            null === join.type) {
            return join;
        }
        return null;
    };

    // ------------------------------------------

    // SLIDE MASTER -----------------------------

    this.ReadSlideMaster = function()
    {
        var master = new MasterSlide(this.presentation, null);
        this.ClearConnectedObjects();
        this.TempMainObject = master;

        var s = this.stream;

        s.Skip2(1); // type
        var end = s.cur + s.GetULong() + 4;

        s.Skip2(1); // attribute start
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    master.preserve = s.GetBool();
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < end)
        {
            var _rec = s.GetUChar();

            switch (_rec)
            {
                case 0:
                {
                    var cSld = new AscFormat.CSld(master);
                    this.ReadCSld(cSld);
                    AscCommonSlide.fFillFromCSld(master, cSld);
                    break;
                }
                case 1:
                {
                    var clrMap = new AscFormat.ClrMap();
                    this.ReadClrMap(clrMap);
                    master.setClMapOverride(clrMap);
                    break;
                }
                case 2:
                {
                    var _transition = this.ReadTransition();
                    master.applyTransition(_transition);
                    break;
                }
                case 3:
                {
                    var oTiming = new AscFormat.CTiming();
                    oTiming.fromPPTY(this);
                    master.setTiming(oTiming);
                    break;
                }
                case 4:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                case 5:
                {
                    master.setHF(this.ReadHF());
                    break;
                }
                case 6:
                {
                    master.setTxStyles(this.ReadTxStyles());
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(end);
        this.TempMainObject = null;
        this.AssignConnectedObjects();
        return master;
    };

    this.ReadTxStyles = function()
    {
        var txStyles = new AscFormat.CTextStyles();

        var s = this.stream;
        var end = s.cur + s.GetULong() + 4;

        while (s.cur < end)
        {
            var _rec = s.GetUChar();

            switch (_rec)
            {
                case 0:
                {
                    txStyles.titleStyle = this.ReadTextListStyle();
                    break;
                }
                case 1:
                {
                    txStyles.bodyStyle = this.ReadTextListStyle();
                    break;
                }
                case 2:
                {
                    txStyles.otherStyle = this.ReadTextListStyle();
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(end);
        return txStyles;
    };

    // ------------------------------------------

    // SLIDE LAYOUT -----------------------------

    this.ReadSlideLayout = function()
    {
        var layout = new SlideLayout(null);
        this.ClearConnectedObjects();
        this.TempMainObject = layout;

        var s = this.stream;

        s.Skip2(1); // type
        var end = s.cur + s.GetULong() + 4;

        s.Skip2(1); // attribute start
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    layout.setMatchingName(s.GetString2());
                    break;
                }
                case 1:
                {
                    layout.preserve = s.GetBool();
                    break;
                }
                case 2:
                {
                    layout.setShowPhAnim(s.GetBool());
                    break;
                }
                case 3:
                {
                    layout.setShowMasterSp(s.GetBool());
                    break;
                }
                case 4:
                {
                    layout.userDrawn = s.GetBool();
                    break;
                }
                case 5:
                {
                    layout.setType(s.GetUChar());
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < end)
        {
            var _rec = s.GetUChar();

            switch (_rec)
            {
                case 0:
                {
                    var cSld = new AscFormat.CSld(layout);
                    this.ReadCSld(cSld);
                    AscCommonSlide.fFillFromCSld(layout, cSld);
                    break;
                }
                case 1:
                {
                    layout.setClMapOverride(this.ReadClrOverride());
                    break;
                }

                case 2:
                {
                    var _transition = this.ReadTransition();
                    layout.applyTransition(_transition);
                    break;
                }
                case 3:
                {
                    var oTiming = new AscFormat.CTiming();
                    oTiming.fromPPTY(this);
                    layout.setTiming(oTiming);
                    break;
                }
                case 4:
                {
                    layout.setHF(this.ReadHF());
                    break;
                }
                default:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
            }
        }

        s.Seek2(end);
        this.TempMainObject = null;
        this.AssignConnectedObjects();
        return layout;
    };

    // ------------------------------------------

    // SLIDE ------------------------------------

    this.ReadSlide = function(sldIndex)
    {
        var slide = new Slide(this.presentation, null, sldIndex);
        this.ClearConnectedObjects();
        this.TempMainObject = slide;

        var s = this.stream;
        s.Skip2(1); // type
        var end = s.cur + s.GetULong() + 4;

        s.Skip2(1); // attribute start
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                slide.setShow(s.GetBool());
            else if (1 == _at)
                slide.setShowPhAnim(s.GetBool());
            else if (2 == _at)
                slide.setShowMasterSp(s.GetBool());
        }

        while (s.cur < end)
        {
            var _rec = s.GetUChar();

            switch (_rec)
            {
                case 0:
                {
                    var cSld = new AscFormat.CSld(slide);
                    this.ReadCSld(cSld);
                    AscCommonSlide.fFillFromCSld(slide, cSld);
                    break;
                }
                case 1:
                {
                    slide.setClMapOverride(this.ReadClrOverride());
                    break;
                }
                case 2:
                {
                    var _transition = this.ReadTransition();
                    slide.applyTransition(_transition);
                    break;
                }
                case 3:
                {
                    var oTiming = new AscFormat.CTiming();
                    oTiming.fromPPTY(this);
                    slide.setTiming(oTiming);
                    break;
                }
                case 4:
                {
                    this.ReadComments(slide.writecomments);
                    break;
                }
                default:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
            }
        }

        slide.Load_Comments(this.presentation.CommentAuthors);

        s.Seek2(end);
        this.TempMainObject = null;
        this.AssignConnectedObjects();
        return slide;
    };

    this.ReadComments = function(writecomments)
    {
        var s = this.stream;
        var end2 = s.cur + s.GetLong() + 4;
        while (s.cur < end2)
        {
            var _rec2 = s.GetUChar();
            switch (_rec2)
            {
                case 0:
                {
                    s.Skip2(4); // len
                    var lCount = s.GetULong();

                    for (var i = 0; i < lCount; i++)
                    {
                        s.Skip2(1);

                        var _comment = new AscCommon.CWriteCommentData();

                        var _end_rec3 = s.cur + s.GetLong() + 4;

                        s.Skip2(1); // start attributes
                        while (true)
                        {
                            var _at3 = s.GetUChar();
                            if (_at3 == g_nodeAttributeEnd)
                                break;

                            switch (_at3)
                            {
                                case 0:
                                    _comment.WriteAuthorId = s.GetLong();
                                    break;
                                case 1:
                                    _comment.WriteTime = s.GetString2();
                                    break;
                                case 2:
                                    _comment.WriteCommentId = s.GetLong();
                                    break;
                                case 3:
                                    _comment.x = s.GetLong();
                                    break;
                                case 4:
                                    _comment.y = s.GetLong();
                                    break;
                                case 5:
                                    _comment.WriteText = s.GetString2();
                                    break;
                                case 6:
                                    _comment.WriteParentAuthorId = s.GetLong();
                                    break;
                                case 7:
                                    _comment.WriteParentCommentId = s.GetLong();
                                    break;
                                case 8:
                                    _comment.AdditionalData = s.GetString2();
                                    break;
                                default:
                                    break;
                            }
                        }

                        while (s.cur < _end_rec3)
                        {
                            var _rec3 = s.GetUChar();
                            switch (_rec3)
                            {
                                case 0:
                                {
                                    var _end_rec4 = s.cur + s.GetLong() + 4;
                                    s.Skip2(1); // start attributes
                                    while (true)
                                    {
                                        var _at = s.GetUChar();
                                        if (_at == g_nodeAttributeEnd)
                                            break;

                                        switch (_at)
                                        {
                                            case 9: { _comment.timeZoneBias = s.GetLong(); break; }
                                            default:
                                                return;
                                        }
                                    }
                                    s.Seek2(_end_rec4);
                                    break;
                                }
                                default:
                                {
                                    s.SkipRecord();
                                    break;
                                }
                            }
                        }

                        s.Seek2(_end_rec3);

                        _comment.Calculate2();
                        writecomments.push(_comment);
                    }

                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(end2);
    };

    this.ReadTransition = function()
    {
        var _transition = new Asc.CAscSlideTransition();
        _transition.setDefaultParams();

        var s = this.stream;
        var end = s.cur + s.GetULong() + 4;

        if (s.cur == end)
            return _transition;

        s.Skip2(1); // attribute start
        var _presentDuration = false;
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
            {
                _transition.SlideAdvanceOnMouseClick = s.GetBool();
            }
            else if (1 == _at)
            {
                _transition.SlideAdvanceAfter = true;
                _transition.SlideAdvanceDuration = s.GetULong();
            }
            else if (2 == _at)
            {
                _transition.TransitionDuration = s.GetULong();
                _presentDuration = true;
            }
            else if (3 == _at)
            {
                var _spd = s.GetUChar();
                if (!_presentDuration)
                {
                    _transition.TransitionDuration = 250;
                    if (_spd == 1)
                        _transition.TransitionDuration = 500;
                    else if (_spd == 2)
                        _transition.TransitionDuration = 750;
                }
            }
        }

        while (s.cur < end)
        {
            var _rec = s.GetUChar();

            switch (_rec)
            {
                case 0:
                {
                    var _type = "";
                    var _paramNames = [];
                    var _paramValues = [];

                    var _end_rec2 = s.cur + s.GetULong() + 4;

                    s.Skip2(1); // start attributes
                    while (true)
                    {
                        var _at2 = s.GetUChar();
                        if (_at2 == g_nodeAttributeEnd)
                            break;

                        switch (_at2)
                        {
                            case 0:
                            {
                                _type = s.GetString2();
                                break;
                            }
                            case 1:
                            {
                                _paramNames.push(s.GetString2());
                                break;
                            }
                            case 2:
                            {
                                _paramValues.push(s.GetString2());
                                break;
                            }
                            default:
                                break;
                        }
                    }
                    _transition.parseXmlParameters(_type, _paramNames, _paramValues);
                    s.Seek2(_end_rec2);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(end);
        return _transition;
    };

    this.ReadHF = function()
    {
        var hf = new AscFormat.HF();

        var s = this.stream;
        var _e = s.cur + s.GetULong() + 4;

        s.Skip2(1); // attribute start
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                hf.setDt(s.GetBool());
            else if (1 == _at)
                hf.setFtr(s.GetBool());
            else if (2 == _at)
                hf.setHdr(s.GetBool());
            else if (3 == _at)
                hf.setSldNum(s.GetBool());
        }

        s.Seek2(_e);
        return hf;
    };

    // ------------------------------------------

    this.ReadNoteMaster = function()
    {
        var oNotesMaster = new AscCommonSlide.CNotesMaster();
        this.ClearConnectedObjects();
        this.TempMainObject = oNotesMaster;
        this.stream.Skip2(1); // type
        var end = this.stream.cur + this.stream.GetLong() + 4;
        while(this.stream.cur < end){
            var at = this.stream.GetUChar();
            switch (at)
            {
                case 0:
                {
                    var cSld = new AscFormat.CSld(oNotesMaster);
                    this.ReadCSld(cSld);
                    for(var i = 0; i < cSld.spTree.length; ++i){
                        oNotesMaster.addToSpTreeToPos(i, cSld.spTree[i]);
                    }
                    if(cSld.Bg)
                    {
                        oNotesMaster.changeBackground(cSld.Bg);
                    }
                    oNotesMaster.setCSldName(cSld.name);
                    break;
                }
                case 1:
                {
                    let clrMap = new AscFormat.ClrMap();
                    this.ReadClrMap(clrMap);
                    oNotesMaster.setClrMap(clrMap);
                    break;
                }
                case 2:
                {
                    oNotesMaster.setHF(this.ReadHF());
                    break;
                }
                case 3:
                {
                    oNotesMaster.setNotesStyle(this.ReadTextListStyle());
                    break;
                }
                default:
                {
                    this.stream.SkipRecord();
                    break;
                }
            }
        }

        this.stream.Seek2(end);
        this.TempMainObject = null;
        this.AssignConnectedObjects()
        return oNotesMaster;
    };

    this.ReadNote = function()
    {
        var oNotes = new AscCommonSlide.CNotes();
        this.ClearConnectedObjects()
        this.TempMainObject = oNotes;
        var _s = this.stream;
        _s.Skip2(1); // type
        var _end = _s.cur + _s.GetLong() + 4;

        _s.Skip2(1); // attribute start
        while (true)
        {
            var _at = _s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                oNotes.setShowMasterPhAnim(_s.GetBool());
            else if (1 == _at)
                oNotes.setShowMasterSp(_s.GetBool());
        }

        while (_s.cur < _end)
        {
            var _rec = _s.GetUChar();

            switch (_rec)
            {
                case 0:
                {
                    var cSld = new AscFormat.CSld(oNotes);
                    this.ReadCSld(cSld);
                    for(var i = 0; i < cSld.spTree.length; ++i){
                        oNotes.addToSpTreeToPos(i, cSld.spTree[i]);
                    }
                    if(cSld.Bg)
                    {
                        oNotes.changeBackground(cSld.Bg);
                    }
                    oNotes.setCSldName(cSld.name);
                    break;
                }
                case 1:
                {
                    oNotes.setClMapOverride(this.ReadClrOverride());
                    break;
                }
                default:
                {
                    _s.SkipRecord();
                    break;
                }
            }
        }
        this.TempMainObject = null;
        _s.Seek2(_end);
        this.AssignConnectedObjects();
        return oNotes;
    };

    this.ReadCSld = function(csld)
    {
        var s = this.stream;
        var _end_rec = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                csld.name = s.GetString2();
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    // themeElements
                    csld.Bg = this.ReadBg();
                    break;
                }
                case 1:
                {
                    csld.spTree = this.ReadGroupShapeMain();
                    break;
                }
                default:
                {
                    s.Seek2(_end_rec);
                    return;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadBg = function()
    {
        var bg = new AscFormat.CBg();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                bg.bwMode = s.GetUChar();
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    // themeElements
                    bg.bgPr = this.ReadBgPr();
                    break;
                }
                case 1:
                {
                    bg.bgRef = this.ReadStyleRef();
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return bg;
    };

    this.ReadBgPr = function()
    {
        var bgpr = new AscFormat.CBgPr();

        var s = this.stream;
        var _end_rec = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                bgpr.shadeToTitle = s.GetBool();
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    // themeElements
                    bgpr.Fill = this.ReadUniFill();
                    break;
                }
                case 1:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return bgpr;
    };

    this.ReadStyleRef = function()
    {
        var ref = new AscFormat.StyleRef();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                ref.setIdx(s.GetLong());
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    // themeElements
                    ref.setColor(this.ReadUniColor());
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return ref;
    };

    this.ReadFontRef = function()
    {
        var ref = new AscFormat.FontRef();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                ref.setIdx(s.GetUChar());
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    // themeElements
                    ref.setColor(this.ReadUniColor());
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return ref;
    };

    // THEME ------------------------------------

    this.ReadTheme = function()
    {
        var theme = new AscFormat.CTheme();
        theme.presentation = this.presentation;

        var s = this.stream;
        var type = s.GetUChar();

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;


        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                theme.setName(s.GetString2());
            else if (1 == _at)
                theme.setIsThemeOverride(s.GetBool());
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    // themeElements
                    var themeElements = new AscFormat.ThemeElements(theme);
                    this.ReadThemeElements(themeElements);
                    theme.setFontScheme(themeElements.fontScheme);
                    theme.setFormatScheme(themeElements.fmtScheme);
                    theme.setColorScheme(themeElements.clrScheme);

                    break;
                }
                case 1:
                {
                    theme.setSpDef(this.ReadDefaultShapeProperties());
                    break;
                }
                case 2:
                {
                    theme.setLnDef(this.ReadDefaultShapeProperties());
                    break;
                }
                case 3:
                {
                    theme.setTxDef(this.ReadDefaultShapeProperties());
                    break;
                }
                case 4:
                {
                    s.Skip2(4); // len
                    var _len = s.GetULong();
                    for (var i = 0; i < _len; i++)
                    {
                        s.Skip2(1); // type
                        theme.extraClrSchemeLst[i] = this.ReadExtraColorScheme();
                    }
                }
            }
        }

        s.Seek2(_end_rec);
        return theme;
    };

    this.ReadThemeElements = function(thelems)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    this.ReadClrScheme(thelems.clrScheme);
                    break;
                }
                case 1:
                {
                    this.ReadFontScheme(thelems.fontScheme);
                    break;
                }
                case 2:
                {
                    this.ReadFmtScheme(thelems.fmtScheme);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadFontScheme = function(fontscheme)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                fontscheme.setName(s.GetString2());
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    this.ReadFontCollection(fontscheme.majorFont);
                    break;
                }
                case 1:
                {
                    this.ReadFontCollection(fontscheme.minorFont);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadFontCollection = function(fontcolls)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    fontcolls.setLatin(this.ReadTextFontTypeface());
                    break;
                }
                case 1:
                {
                    fontcolls.setEA(this.ReadTextFontTypeface());
                    break;
                }
                case 2:
                {
                    fontcolls.setCS(this.ReadTextFontTypeface());
                    break;
                }
                case 3:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadTextFontTypeface = function()
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var charset = "";
        var panose = "";
        var pitchFamily = "";
        var typeface = "";

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    charset = s.GetString2();
                    break;
                }
                case 1:
                {
                    panose = s.GetString2();
                    break;
                }
                case 2:
                {
                    pitchFamily = s.GetString2();
                    break;
                }
                case 3:
                {
					typeface = s.GetString2();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);

        return typeface;
    };

    this.ReadFmtScheme = function(fmt)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                fmt.setName(s.GetString2());
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    s.Skip2(4); // len
                    var _c = s.GetULong();

                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1); // type
                        fmt.fillStyleLst[i] = this.ReadUniFill();
                    }

                    break;
                }
                case 1:
                {
                    s.Skip2(4); // len
                    var _c = s.GetULong();

                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1); // type1
                        fmt.lnStyleLst[i] = this.ReadLn();
                    }
                    break;
                }
                case 2:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                case 3:
                {
                    s.Skip2(4); // len
                    var _c = s.GetULong();

                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1); // type
                        fmt.bgFillStyleLst[i] = this.ReadUniFill();
                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadDefaultShapeProperties = function()
    {
        var def = new AscFormat.DefaultShapeDefinition();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    this.ReadSpPr(def.spPr);
                    break;
                }
                case 1:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);

                    // bodyPr
                    break;
                }
                case 2:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);

                    // textstyles
                    break;
                }
                case 3:
                {
                    def.style = this.ReadShapeStyle();
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return def;
    };

    this.ReadEffectLst = function()
    {
        var s = this.stream;
        s.GetULong();
        s.GetUChar();
        var nRecStart = s.cur;
        var nRecLen = s.GetLong();
        var nRecEnd = nRecStart + nRecLen + 4;
        var oEffectLst = new AscFormat.CEffectLst();

        while (s.cur < nRecEnd)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    oEffectLst.blur = this.ReadBlur();
                    break;
                }
                case 1:
                {
                    oEffectLst.fillOverlay = this.ReadFillOverlay();
                    break;
                }
                case 2:
                {
                    oEffectLst.glow = this.ReadGlow();
                    break;
                }
                case 3:
                {
                    oEffectLst.innerShdw = this.ReadInnerShdw();
                    break;
                }
                case 4:
                {
                    oEffectLst.outerShdw = this.ReadOuterShdw();
                    break;
                }
                case 5:
                {
                    oEffectLst.prstShdw = this.ReadPrstShdw();
                    break;
                }
                case 6:
                {
                    oEffectLst.reflection = this.ReadReflection();
                    break;
                }
                case 7:
                {
                    oEffectLst.softEdge = this.ReadSoftEdge();
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }
        s.Seek2(nRecEnd);
        return oEffectLst;
    };

    this.ReadEffectProperties = function()
    {
        var s = this.stream;
        var pos = s.cur;
        var nLength = s.GetLong();
        if(nLength === 0)
        {
            return null;
        }
        var type = s.GetUChar();
        s.Seek2(pos);
        var oEffectProperties = new AscFormat.CEffectProperties();
        if(type === 1)
        {
            oEffectProperties.EffectLst = this.ReadEffectLst();
        }
        else
        {

            oEffectProperties.EffectDag = this.ReadEffectDag();
        }
        return oEffectProperties;
    };

    this.ReadSpPr = function(spPr)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                spPr.setBwMode(s.GetUChar());
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    spPr.setXfrm(this.ReadXfrm());
                    spPr.xfrm.setParent(spPr);
                    break;
                }
                case 1:
                {
                    spPr.setGeometry(this.ReadGeometry(spPr.xfrm));
                    break;
                }
                case 2:
                {
                    spPr.setFill(this.ReadUniFill(spPr, null, null));
                    break;
                }
                case 3:
                {
                    spPr.setLn(this.ReadLn(spPr));
                    break;
                }
                case 4:
                {
                    spPr.setEffectPr(this.ReadEffectProperties());
                    break;
                }
                case 5:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                case 6:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadGrSpPr = function(spPr)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            if (0 == _at)
                spPr.setBwMode(s.GetUChar());
            else
                break;
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    spPr.setXfrm(this.ReadXfrm());
                    spPr.xfrm.setParent(spPr);
                    break;
                }
                case 1:
                {
                    spPr.setFill(this.ReadUniFill(spPr, null, null));
                    break;
                }
                case 2:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                case 3:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadXfrm = function()
    {
        var ret = new AscFormat.CXfrm();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    ret.setOffX(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 1:
                {
                    ret.setOffY(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 2:
                {
                    ret.setExtX(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 3:
                {
                    ret.setExtY(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 4:
                {
                    ret.setChOffX(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 5:
                {
                    ret.setChOffY(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 6:
                {
                    ret.setChExtX(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 7:
                {
                    ret.setChExtY(s.GetLong() / c_dScalePPTXSizes);
                    break;
                }
                case 8:
                {
                    ret.setFlipH(s.GetBool());
                    break;
                }
                case 9:
                {
                    ret.setFlipV(s.GetBool());
                    break;
                }
                case 10:
                {
                    ret.setRot((s.GetLong()/60000)*Math.PI/180);
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
        return ret;
    };

	this.ReadSignatureLine = function()
	{
		var ret = new AscFormat.CSignatureLine();

		var s = this.stream;

		var _rec_start = s.cur;
		var _end_rec = _rec_start + s.GetULong() + 4;

		s.Skip2(1); // start attributes

		while (true)
		{
			var _at = s.GetUChar();
			if (_at == g_nodeAttributeEnd)
				break;
			switch (_at)
			{
				case 0:
				{
					s.GetString2();
					break;
				}
				case 1:
				{
					s.GetBool();
					break;
				}
				case 2:
				{
					s.GetUChar();
					break;
				}
				case 3:
				{
					ret.id = s.GetString2();
					break;
				}
				case 4:
				{
					s.GetBool();
					break;
				}
				case 5:
				{
					s.GetString2();
					break;
				}
				case 6:
				{
                    ret.showDate = s.GetBool();
					break;
				}
				case 7:
				{
					ret.instructions = s.GetString2();
					break;
				}
				case 8:
				{
					s.GetBool();
					break;
				}
				case 9:
				{
					s.GetString2();
					break;
				}
				case 10:
				{
					ret.signer = s.GetString2();
					break;
				}
				case 11:
				{
					ret.signer2 = s.GetString2();
					break;
				}
				case 12:
				{
					ret.email = s.GetString2();
					break;
				}
				default:
					break;
			}
		}

		s.Seek2(_end_rec);
		return ret;
	};

    this.ReadShapeStyle = function()
    {
        var def = new AscFormat.CShapeStyle();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    def.setLnRef(this.ReadStyleRef());
                    break;
                }
                case 1:
                {
                    def.setFillRef(this.ReadStyleRef());
                    break;
                }
                case 2:
                {
                    def.setEffectRef(this.ReadStyleRef());
                    break;
                }
                case 3:
                {
                    def.setFontRef(this.ReadFontRef());
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return def;
    };

    this.ReadOleInfo = function(ole)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetLong() + 4;

        s.Skip2(1); // start attributes
        var dxaOrig = 0;
        var dyaOrig = 0;
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    ole.setApplicationId(s.GetString2());
                    break;
                }
                case 1:
                {
                    ole.setData(s.GetString2());
                    break;
                }
                case 2:
                {
                    dxaOrig = s.GetULong();
                    break;
                }
                case 3:
                {
                    dyaOrig = s.GetULong();
                    break;
                }
                case 4:
                {
                    const drawAspect = s.GetUChar();
                    ole.setDrawAspect(drawAspect);
                    break;
                }
                case 5:
                {
                    s.GetUChar();
                    break;
                }
                case 6:
                {
                    s.GetUChar();
                    break;
                }
                case 7:
                {
                    ole.setObjectFile(s.GetString2());
                    break;
                }
                default:
                {
                    break;
                }
            }
        }
        var oleType = null;
        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 1:
                {
                    s.GetLong();//length
                    oleType = s.GetUChar();
                    ole.setOleType(oleType);
                    break;
                }
                case 2:
                {

                    var binary_length;
                    switch(oleType)
                    {
                        case 0:
                        {
                            binary_length = s.GetULong();
                            ole.setBinaryData(s.data.slice(s.cur, s.cur + binary_length));
                            s.Seek2(s.cur + binary_length);
                            break;
                        }
                        case 1:
                        {
                            ole.setObjectFile("maskFile.docx");
                            binary_length = s.GetULong();
                            ole.setBinaryData(s.data.slice(s.cur, s.cur + binary_length));
                            s.Seek2(s.cur + binary_length);
                            break;
                        }
                        case 2:
                        {
                            ole.setObjectFile("maskFile.xlsx");
                            binary_length = s.GetULong();
                            ole.setBinaryData(s.data.slice(s.cur, s.cur + binary_length));
                            s.Seek2(s.cur + binary_length);
                            break;
                        }
                        case 4://Equation
                        {
                            binary_length = s.GetLong();//length
                            var type2 = s.GetUChar();
                            if (AscCommonWord.c_oSer_OMathContentType.OMath === type2)
                            {
                                var oReadResult = new AscCommonWord.DocReadResult(null);
                                var oMathPara = this.ReadMathObject(s, oReadResult, new AscWord.Paragraph(null, true));
                                ole.setMathObject(oMathPara);
                            }
                            else
                            {
                                s.SkipRecord();
                            }
                            break;
                        }
                        default:
                        {
                            s.SkipRecord();
                            break;
                        }
                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }


        if (dxaOrig > 0 && dyaOrig > 0) {
			var ratio = 4 / 3 / 20;//twips to px
			ole.setPixSizes(ratio * dxaOrig, ratio * dyaOrig);
		}
        s.Seek2(_end_rec);
    };

    this.ReadMathObject = function(oStream, oReadResult, oParagraph)
    {
        var length2 = oStream.GetLong();
        var _stream = new AscCommon.FT_Stream2();
        _stream.data = oStream.data;
        _stream.pos = oStream.pos;
        _stream.cur = oStream.cur;
        _stream.size = oStream.size;
        var boMathr = new Binary_oMathReader(_stream, oReadResult, null);
        var oMathPara = new ParaMath();
        boMathr.bcr.Read1(length2, function(t, l){
            return boMathr.ReadMathArg(t,l,oMathPara.Root,oParagraph);
        });
        oMathPara.Root.Correct_Content(true);
        return oMathPara;
    };

    this.ReadGeometry = function(_xfrm)
    {
        var geom = null;

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        if (s.cur < _end_rec)
        {
            var _t = s.GetUChar();

            if (1 == _t)
            {
                // preset shape

                var _len = s.GetULong();
                var _s = s.cur;
                var _e = _s + _len;

                s.Skip2(1); // start attributes

                while (true)
                {
                    var _at = s.GetUChar();
                    if (_at == g_nodeAttributeEnd)
                        break;

                    if (0 == _at)
                    {
                        var tmpStr = s.GetString2();
                        geom = AscFormat.CreateGeometry(tmpStr);
                        geom.isLine = tmpStr == "line";
                        geom.setPreset(tmpStr);
                    }

                    else
                        break;
                }

                while (s.cur < _e)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            this.ReadGeomAdj(geom);
                            break;
                        }
                        default:
                        {
                            s.SkipRecord();
                            break;
                        }
                    }
                }
            }
            else if (2 == _t)
            {
                var _len = s.GetULong();
                var _s = s.cur;
                var _e = _s + _len;

                geom = AscFormat.CreateGeometry("");
                geom.preset = null;
                while (s.cur < _e)
                {
                    var _at = s.GetUChar();
                    switch (_at)
                    {
                        case 0:
                        {
                            this.ReadGeomAdj(geom);
                            break;
                        }
                        case 1:
                        {
                            this.ReadGeomGd(geom);
                            break;
                        }
                        case 2:
                        {
                            this.ReadGeomAh(geom);
                            break;
                        }
                        case 3:
                        {
                            this.ReadGeomCxn(geom);
                            break;
                        }
                        case 4:
                        {
                            this.ReadGeomPathLst(geom, _xfrm);
                            break;
                        }
                        case 5:
                        {
                            this.ReadGeomRect(geom);
                            break;
                        }
                        default:
                        {
                            s.SkipRecord();
                            break;
                        }
                    }
                }
            }
        }

        s.Seek2(_end_rec);
        return geom;
    };

    this.ReadGeomAdj = function(geom)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var _c = s.GetULong();

        for (var i = 0; i < _c; i++)
        {
            s.Skip2(6); // type + len + start attributes

            var arr = [];
            var cp = 0;

            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                if (cp == 1)
                    arr[cp] = s.GetLong();
                else
                    arr[cp] = s.GetString2();
                cp++;
            }

            if (arr.length >= 3)
                geom.AddAdj(arr[0], arr[1], arr[2]);
        }

        s.Seek2(_end_rec);
    };

    this.ReadGeomGd = function(geom)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var _c = s.GetULong();

        for (var i = 0; i < _c; i++)
        {
            s.Skip2(6); // type + len + start attributes

            var arr = [];
            var cp = 0;

            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                if (cp == 1)
                    arr[cp] = s.GetLong();
                else
                    arr[cp] = s.GetString2();
                cp++;
            }

            geom.AddGuide(arr[0], arr[1], arr[2], arr[3], arr[4]);
        }

        s.Seek2(_end_rec);
    };

    this.ReadGeomAh = function(geom)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var _c = s.GetULong();

        for (var i = 0; i < _c; i++)
        {
            var _type1 = s.GetUChar();
            s.Skip2(4); // len
            var _type = s.GetUChar();
            s.Skip2(5); // len + start attributes

            var arr = [];
            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                arr[_at] = s.GetString2();
            }

            if (1 == _type)
                geom.AddHandlePolar(arr[2], arr[6], arr[4], arr[3], arr[7], arr[5], arr[0], arr[1]);
            else
                geom.AddHandleXY(arr[2], arr[6], arr[4], arr[3], arr[7], arr[5], arr[0], arr[1]);
        }

        s.Seek2(_end_rec);
    };

    this.ReadGeomCxn = function(geom)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var _c = s.GetULong();

        for (var i = 0; i < _c; i++)
        {
            var _type = s.GetUChar();
            s.Skip2(5); // len + start attributes

            var arr = [];
            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                arr[_at] = s.GetString2();
            }

            geom.AddCnx(arr[2], arr[0], arr[1]);
        }

        s.Seek2(_end_rec);
    };

    this.ReadGeomRect = function(geom)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        var arr = [];
        arr[0] = "l";
        arr[1] = "t";
        arr[2] = "r";
        arr[3] = "b";
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            arr[_at] = s.GetString2();
        }

        geom.AddRect(arr[0], arr[1], arr[2], arr[3]);

        s.Seek2(_end_rec);
    };

    this.ReadGeomPathLst = function(geom, _xfrm)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var _c = s.GetULong();

        for (var i = 0; i < _c; i++)
        {
            var _type = s.GetUChar();
            var _len = s.GetULong();

            var _s = s.cur;
            var _e = _s + _len;
            s.Skip2(1); // start attributes

            var extrusionOk = false;
            var fill = 5;
            var stroke = true;
            var w = undefined;
            var h = undefined;

            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                switch (_at)
                {
                    case 0:
                    {
                        extrusionOk = s.GetBool();
                        break;
                    }
                    case 1:
                    {
                        fill = s.GetUChar();
                        break;
                    }
                    case 2:
                    {
                        h = s.GetLong();
                        break;
                    }
                    case 3:
                    {
                        stroke = s.GetBool();
                        break;
                    }
                    case 4:
                    {
                        w = s.GetLong();
                        break;
                    }
                    default:
                        break;
                }
            }

            geom.AddPathCommand(0, extrusionOk, (fill == 4) ? "none" : "norm", stroke, w, h);
            var isKoords = false;

            while (s.cur < _e)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 0:
                    {
                        s.Skip2(4); // len

                        var _cc = s.GetULong();

                        for (var j = 0; j < _cc; j++)
                        {
                            s.Skip2(5); // type + len
                            isKoords |= this.ReadUniPath2D(geom);
                        }

                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }

            s.Seek2(_e);
        }

        var _path = geom.pathLst[geom.pathLst.length - 1];
        // if (isKoords && undefined === _path.pathW && undefined === _path.pathH)
        // {
        //     if(_xfrm)
        //     {
        //         _path.pathW = _xfrm.extX * c_dScalePPTXSizes;
        //         _path.pathH = _xfrm.extY * c_dScalePPTXSizes;
        //     }
        // }

        s.Seek2(_end_rec);
    };

    this.ReadUniPath2D = function(geom)
    {
        var s = this.stream;

        var _type = s.GetUChar();
        var _len = s.GetULong();

        var _s = s.cur;
        var _e = _s + _len;

        if (3 == _type)
        {
            geom.AddPathCommand(6);
            s.Seek2(_e);
            return;
        }

        s.Skip2(1);

        var isKoord = false;

        var arr = [];
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            arr[_at] = s.GetString2();

            if (!isKoord && !isNaN(parseInt(arr[_at])))
                isKoord = true;
        }

        switch (_type)
        {
            case 1:
            {
                geom.AddPathCommand(1, arr[0], arr[1]);
                break;
            }
            case 2:
            {
                geom.AddPathCommand(2, arr[0], arr[1]);
                break;
            }
            case 3:
            {
                geom.AddPathCommand(6);
                break;
            }
            case 4:
            {
                geom.AddPathCommand(5, arr[0], arr[1], arr[2], arr[3], arr[4], arr[5]);
                break;
            }
            case 5:
            {
                geom.AddPathCommand(3, arr[0], arr[1], arr[2], arr[3]);
                break;
            }
            case 6:
            {
                geom.AddPathCommand(4, arr[0], arr[1], arr[2], arr[3]);
                break;
            }
            default:
            {
                s.SkipRecord();
                break;
            }
        }

        s.Seek2(_e);

        return isKoord;
    };

    // ------------------------------------------

    this.ReadGraphicObject = function()
    {
        var s = this.stream;
        var _type = s.GetUChar();
        var _object = null;

        switch (_type)
        {
            case 1:
            {
                _object = this.ReadShape();
                break;
            }
            case 2://pic
            case 6://ole
            case 7://video
            case 8://audio
            {
                _object = this.ReadPic(_type);
                break;
            }
            case 3:
            {
                _object = this.ReadCxn();
                break;
            }
            case 4:
            {
                _object = this.ReadGroupShape();
                break;
            }
            case 5:
            {
                _object = this.ReadGrFrame();
                break;
            }
            default:
            {
                s.SkipRecord();
                break;
            }
        }

        return _object;
    };

    // SHAPE PROPERTIES -------------------------

    this.ReadShape = function()
    {
        var s = this.stream;

        var shape = Asc.editor.isPdfEditor() ? new AscPDF.CPdfShape(this.TempMainObject) : new AscFormat.CShape(this.TempMainObject);

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;
        shape.setBDeleted(false);
        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    shape.setUseBgFill(s.GetBool());
                    break;
                }
                default:
                    break;
            }
        }

        var txXfrm = null;
        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var pr = this.ReadNvUniProp(shape);
                    shape.setNvSpPr(pr);
                    if(AscFormat.isRealNumber(pr.locks))
                    {
                        shape.setLocks(pr.locks);
                    }
                    break;
                }
                case 1:
                {
                    var sp_pr = new AscFormat.CSpPr();
                    this.ReadSpPr(sp_pr);
                    shape.setSpPr(sp_pr);
                    sp_pr.setParent(shape);
                    break;
                }
                case 2:
                {
                    shape.setStyle(this.ReadShapeStyle());
                    break;
                }
                case 3:
                {
                    shape.setTxBody(this.ReadTextBody(shape));
                    shape.txBody.setParent(shape);
                    break;
                }
				case 6:
				{
                    txXfrm = this.ReadXfrm();
                    shape.setTxXfrm(txXfrm);
					break;
				}
                case 7:
                {
                    shape.setSignature(this.ReadSignatureLine());
                    break;
                }
                case 8:
                {
                    if(typeof AscFormat.SmartArt !== undefined)
                    {
                        s.GetULong();//length
                        s.GetUChar();//attr type - 0
                        shape.setModelId(s.GetString2());
                    }
                    else
                    {
                        s.SkipRecord();
                    }
                    break;
                }
                case 9:
                {
                    s.GetULong();//length
                    s.GetUChar();//attr type - 0
                    shape.setFLocksText(s.GetBool());
                    break;
                }
                case 0xA1:
                {
                    shape.readMacro(s);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }
        if(txXfrm ){
            shape.convertFromSmartArt();
        }

        s.Seek2(_end_rec);
        return shape;
    };

    this.CheckGroupXfrm = function(oGroup){
        if(!oGroup){
            return;
        }
        oGroup.checkXfrm();
    };

    this.ReadGroupShape = function(type)
    {
        var s = this.stream;

        var shape;
        if(type === 9){
            shape = new AscFormat.CLockedCanvas();
        }
        else {
            shape = new AscFormat.CGroupShape();
        }
        shape.setBDeleted(false);
        this.TempGroupObject = shape;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var pr = this.ReadNvUniProp(shape);
                    shape.setNvSpPr(pr);
                    if(AscFormat.isRealNumber(pr.locks))
                    {
                        shape.setLocks(pr.locks);
                    }
                    break;
                }
                case 1:
                {
                    var spPr = new AscFormat.CSpPr();
                    this.ReadGrSpPr(spPr);
                    shape.setSpPr(spPr);
                    spPr.setParent(shape);
                    break;
                }
                case 2:
                {
                    s.Skip2(4); // len
                    var _c = s.GetULong();
                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1);
                        var __len = s.GetULong();
                        if (__len == 0)
                            continue;

                        var _type = s.GetUChar();

                        var _object = null;

                        switch (_type)
                        {
                            case 1:
                            {
                                _object = this.ReadShape();
                                if (!IsHiddenObj(_object) && _object.spPr && _object.spPr.xfrm)
                                {
                                    shape.addToSpTree(shape.spTree.length,_object);
                                    shape.spTree[shape.spTree.length-1].setGroup(shape);
                                }
                                break;
                            }
                            case 6:
                            case 2:
                            case 7:
                            case 8:
                            {
                                _object = this.ReadPic(_type);
                                if (!IsHiddenObj(_object) && _object.spPr && _object.spPr.xfrm)
                                {
                                    shape.addToSpTree(shape.spTree.length,_object);
                                    shape.spTree[shape.spTree.length-1].setGroup(shape);
                                }
                                break;
                            }
                            case 3:
                            {
                                _object = this.ReadCxn();
                                if (!IsHiddenObj(_object) && _object.spPr && _object.spPr.xfrm)
                                {
                                    shape.addToSpTree(shape.spTree.length,_object);
                                    shape.spTree[shape.spTree.length-1].setGroup(shape);
                                }
                                break;
                            }
                            case 4:
                            {
                                _object = this.ReadGroupShape();
                                if (!IsHiddenObj(_object) && _object.spPr && _object.spPr.xfrm && _object.spTree.length > 0)
                                {
                                    shape.addToSpTree(shape.spTree.length,_object);
                                    shape.spTree[shape.spTree.length-1].setGroup(shape);
                                    this.TempGroupObject = shape;
                                }
                                break;
                            }
                            case 5:
                            {
                                var _ret = null;
                                if ("undefined" != typeof(AscFormat.CGraphicFrame))
                                    _ret = this.ReadGrFrame();
                                else
                                    _ret = this.ReadChartDataInGroup(shape);
                                if (null != _ret)
                                {
                                    shape.addToSpTree(shape.spTree.length, _ret);
                                    shape.spTree[shape.spTree.length-1].setGroup(shape);
                                }
                                break;
                            }
                            default:
                            {
                                s.SkipRecord();
                                break;
                            }
                        }
                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }
        this.CheckGroupXfrm(shape);
        s.Seek2(_end_rec);
        this.TempGroupObject = null;
        return shape;
    };

    this.IsNoSlideSpTree = function ()
    {
        return this.TempMainObject && AscFormat.isRealNumber(this.TempMainObject.kind) && (this.TempMainObject.kind !== AscFormat.TYPE_KIND.SLIDE);
    };
    this.ReadSpTreeElement = function()
    {
        let s = this.stream;

        let bIsNoSlideSpTree = this.IsNoSlideSpTree();
        let _type = s.GetUChar();
        let _object;
        switch (_type)
        {
            case 1:
            {
                _object = this.ReadShape();
                if (!IsHiddenObj(_object) || bIsNoSlideSpTree)
                {
                    _object.setParent2(this.TempMainObject);
                }
                else
                {
                    _object = null;
                }
                break;
            }
            case 6:
            case 2:
            case 7:
            case 8:
            {
                _object = this.ReadPic(_type);
                if(_type === 6 && !_object.checkCorrect())
                {
                    _object = null;
                }
                break;
            }
            case 3:
            {
                _object = this.ReadCxn();
                break;
            }
            case 4:
            {
                _object = this.ReadGroupShape();
                break;
            }
            case 5:
            {
                _object = this.ReadGrFrame();
                break;
            }
            case 0x99:
            {
                _object = this.ReadSpTreeElement();
                break;
            }
            default:
            {
                s.SkipRecord();
                break;
            }
        }
        if(_object)
        {
            if (!IsHiddenObj(_object) || bIsNoSlideSpTree)
            {
                _object.setParent2(this.TempMainObject);
            }
            else
            {
                _object = null
            }
        }
        let nCurPos = s.cur;
        let nType = s.GetUChar();
        let oAltDrawing = null;
        if(nType === 0x99)
        {
            let nAltContentLen = s.GetULong();
            let nEndPos = s.cur + nAltContentLen;
            if(!_object || !_object.isSupported())
            oAltDrawing = this.ReadSpTreeElement();
            if(oAltDrawing)
            {
                _object = oAltDrawing;
            }
            s.Seek2(nEndPos);
        }
        else
        {
            s.Seek2(nCurPos);
        }
        return _object;
    };
    this.ReadGroupShapeMain = function()
    {
        let s = this.stream;

        let shapes = [];

        let _rec_start = s.cur;
        let _end_rec = _rec_start + s.GetULong() + 4;
        let _object;

        s.Skip2(5); // type SPTREE + len

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                case 1:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                case 2:
                {
                    s.Skip2(4); // len
                    var _c = s.GetULong();
                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1);
                        var __len = s.GetULong();
                        if (__len == 0)
                            continue;

                        _object = this.ReadSpTreeElement();
                        if(_object)
                        {
                            shapes[shapes.length] = _object;
                        }
                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return shapes;
    };


    this.ReadPic = function(type)
    {
        var s = this.stream;

        var isOle = (type === 6);
        var pic;
        if (isOle)
        {
            pic = new AscFormat.COleObject(this.TempMainObject)
        }
        else
        {
            if (Asc.editor.isPdfEditor())
                pic = new AscPDF.CPdfImage();
            else
                pic = new AscFormat.CImageShape(this.TempMainObject);
        }

        pic.setBDeleted(false);

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var sMaskFileName;
        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var pr = this.ReadNvUniProp(pic);
                    pic.setNvSpPr(pr);
                    if(AscFormat.isRealNumber(pr.locks)){
                        pic.setLocks(pr.locks);
                    }
                    break;
                }
                case 1:
                {
                    pic.setBlipFill(this.ReadUniFill(null, pic, null).fill);
                    break;
                }
                case 2:
                {

                    var spPr = new AscFormat.CSpPr();
                    spPr.setParent(pic);
                    this.ReadSpPr(spPr);
                    pic.setSpPr(spPr);
                    break;
                }
                case 3:
                {
                    pic.setStyle(this.ReadShapeStyle());
                    break;
                }
                case 4:
                {
                    if(isOle) {
                        this.ReadOleInfo(pic);
                    } else {
                        s.SkipRecord();
                    }
                    break;
                }
                case 5:
                {
                    if(type === 7 || type === 8){//video or audio
                        s.GetLong();
                        s.GetUChar();//start attributes

                        while(true)
                        {
                            var _at2 = s.GetUChar();
                            if (_at2 == g_nodeAttributeEnd)
                                break;
                            switch (_at2) {
                                case 0:
                                {
                                    sMaskFileName = s.GetString2();
                                    break;
                                }
                                case 1:
                                {
                                    s.GetBool();
                                    break;
                                }
                                default:
                                {
                                    break;
                                }
                            }
                        }

                    }
                    else{
                        s.SkipRecord();
                    }
                    break;
                }
                case 0xA1:
                {
                    pic.readMacro(s);
                    break;
                }
                default:
                {
                    this.stream.SkipRecord();
                    break;
                }
            }
        }

        if(type === 7 || type === 8){//video or audio
            if(typeof sMaskFileName === "string" && sMaskFileName.length > 0 &&
                pic.nvPicPr && pic.nvPicPr.nvPr /*&& pic.nvPicPr.nvPr.unimedia*/){
                var oUniMedia = new AscFormat.UniMedia();
                oUniMedia.type = type;
                oUniMedia.media = sMaskFileName;
                pic.nvPicPr.nvPr.setUniMedia(oUniMedia);
            }
        }
        s.Seek2(_end_rec);
        return pic;
    };
    this.ReadCxn = function()
    {
        var s = this.stream;

        var shape = false == Asc.editor.isPdfEditor() ? new AscFormat.CConnectionShape() : new AscPDF.CPdfConnectionShape();
        shape.setBDeleted(false);

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var pr = this.ReadNvUniProp(shape);
                    shape.setNvSpPr(pr);
                    if(AscFormat.isRealNumber(pr.locks)){
                        shape.setLocks(pr.locks);
                    }
                    break;
                }
                case 1:
                {
                    var spPr = new AscFormat.CSpPr();
                    spPr.setParent(shape);
                    this.ReadSpPr(spPr);
                    shape.setSpPr(spPr);

                    break;
                }
                case 2:
                {
                    shape.setStyle(this.ReadShapeStyle());
                    break;
                }
                case 0xA1:
                {
                    shape.readMacro(s);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }
        this.AddConnectedObject(shape);
        s.Seek2(_end_rec);
        return shape;
    };

    this.ReadChartDataInGroup = function(group)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        this.TempGroupObject = group;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    var spid = s.GetString2();
                    break;
                }
                default:
                    break;
            }
        }

        var _nvGraphicFramePr = null;
        var _xfrm = null;
        var _chart = null;
        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    _nvGraphicFramePr = this.ReadNvUniProp(AscFormat.ExecuteNoHistory(function () {
                        return new AscFormat.CGraphicFrame();
                    }, this, []));
                    break;
                }
                case 1:
                {
                    _xfrm = this.ReadXfrm();
                    break;
                }
                case 2:
                {
                    s.SkipRecord();
                    break;
                }
                case 3:
                {
                    var _length = s.GetLong();
                    var _pos = s.cur;

                    var _stream = new AscCommon.FT_Stream2();
                    _stream.data = s.data;
                    _stream.pos = s.pos;
                    _stream.cur = s.cur;
                    _stream.size = s.size;

                    _chart = new AscFormat.CChartSpace();
                    _chart.setBDeleted(false);
                    var oBinaryChartReader = new AscCommon.BinaryChartReader(_stream);
                    oBinaryChartReader.ExternalReadCT_ChartSpace(_length, _chart, this.presentation);
                    _chart.setBDeleted(false);
                    if(AscCommon.isRealObject(_nvGraphicFramePr) && AscFormat.isRealNumber(_nvGraphicFramePr.locks))
                    {
                        _chart.setLocks(_nvGraphicFramePr.locks);
                    }
                    if(_xfrm)
                    {
                        if(!_chart.spPr)
                        {
                            _chart.setSpPr(new AscFormat.CSpPr());
                            _chart.spPr.setParent(_chart);
                        }

                        _chart.spPr.setXfrm(_xfrm);
                        _xfrm.setParent(_chart.spPr);
                    }

                    s.Seek2(_pos + _length);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);

        this.TempGroupObject = null;
        if (_chart == null || !_chart.hasCharts())
            return null;

        return _chart;
    };

    this.ReadGrFrame = function()
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var _graphic_frame = Asc.editor.isPdfEditor() ? new AscPDF.CPdfGraphicFrame() : new AscFormat.CGraphicFrame();
        _graphic_frame.setParent2(this.TempMainObject);
        this.TempGroupObject = _graphic_frame;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    var spid = s.GetString2();
                    break;
                }
                default:
                    break;
            }
        }

        var _nvGraphicFramePr = null;
        var _xfrm = null;
        var _table = null;
        var _chart = null;
        var _slicer = null;
        var _timeslicer = null;
        var _smartArt = null;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    _nvGraphicFramePr = this.ReadNvUniProp(_graphic_frame);
                    break;
                }
                case 1:
                {
                    _xfrm = this.ReadXfrm();
                    break;
                }
                case 2:
                {
                    _table = this.ReadTable(_xfrm, _graphic_frame);
                    break;
                }
                case 3:
                {
                    var _length = s.GetLong();
                    var _pos = s.cur;

                    if(typeof AscFormat.CChartSpace !== "undefined" && _length)
                    {
                        var _stream = new AscCommon.FT_Stream2();
                        _stream.data = s.data;
                        _stream.pos = s.pos;
                        _stream.cur = s.cur;
                        _stream.size = s.size;
                        _chart = new AscFormat.CChartSpace();
                        _chart.setBDeleted(false);
                        AscCommon.pptx_content_loader.ImageMapChecker = this.ImageMapChecker;
                        AscCommon.pptx_content_loader.Reader.ImageMapChecker = this.ImageMapChecker;
                        var oBinaryChartReader = new AscCommon.BinaryChartReader(_stream);
                        oBinaryChartReader.ExternalReadCT_ChartSpace(_length, _chart, this.presentation);
                        if(!_chart.hasCharts())
                        {
                            _chart = null;
                        }
                    }

                    s.Seek2(_pos + _length);
                    break;
                }
                case 5:
                case 6:
                {
                    if (typeof AscFormat.CSlicer !== "undefined")
                    {
                        _slicer = new AscFormat.CSlicer();
                        _slicer.fromStream(s);
                    }
                    else
                    {
                        s.SkipRecord();
                    }
                    break;
                }
                case 7: // chartEx
                {
                    var _length = s.GetLong();
                    var _pos = s.cur;

                    if(typeof AscFormat.CChartSpace !== "undefined" && _length)
                    {
                        var _stream = new AscCommon.FT_Stream2();
                        _stream.data = s.data;
                        _stream.pos = s.pos;
                        _stream.cur = s.cur;
                        _stream.size = s.size;
                        _chart = new AscFormat.CChartSpace();
                        _chart.setBDeleted(false);
                        AscCommon.pptx_content_loader.ImageMapChecker = this.ImageMapChecker;
                        AscCommon.pptx_content_loader.Reader.ImageMapChecker = this.ImageMapChecker;
                        var oBinaryChartReader = new AscCommon.BinaryChartReader(_stream);
                        oBinaryChartReader.ExternalReadCT_ChartExSpace(_length, _chart, this.presentation);
                        if(!_chart.hasCharts())
                        {
                            _chart = null;
                        }
                    }

                    s.Seek2(_pos + _length);
                    break;
                }
                case 8://smartArt
                {
                    _smartArt = this.ReadSmartArt();
										this.smartarts.push(_smartArt);
                    break;
                }
                case 9:
                {
                    if (typeof AscFormat.CTimeslicer !== "undefined")
                    {
                        _timeslicer = new AscFormat.CTimeslicer();
                        _timeslicer.fromStream(s);
                    }
                    else
                    {
                        s.SkipRecord();
                    }
                    break;
                }
                case 0xA1:
                {
                    _graphic_frame.readMacro(s);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);

        this.TempGroupObject = null;
        if (_table == null && _chart == null && _slicer == null && _smartArt == null && _timeslicer == null)
            return null;

        if (_table != null)
        {
            _graphic_frame.checkEmptySpPrAndXfrm(_xfrm);
            _graphic_frame.setNvSpPr(_nvGraphicFramePr);
            if(AscCommon.isRealObject(_nvGraphicFramePr) && AscFormat.isRealNumber(_nvGraphicFramePr.locks))
            {
                _graphic_frame.setLocks(_nvGraphicFramePr.locks);
            }
            _graphic_frame.setGraphicObject(_table);
            _graphic_frame.setBDeleted(false);
        }
        else if (_chart != null)
        {
            _chart.checkEmptySpPrAndXfrm(_xfrm);
            if(AscCommon.isRealObject(_nvGraphicFramePr) )
            {
                _chart.setNvSpPr(_nvGraphicFramePr);
                if(AscFormat.isRealNumber(_nvGraphicFramePr.locks))
                {
                    _chart.setLocks(_nvGraphicFramePr.locks);
                }
                if(_nvGraphicFramePr.cNvPr) 
                {
                    this.map_shapes_by_id[_nvGraphicFramePr.cNvPr.id] = _chart;
                }
            }
            return _chart;
        }
        else if(_slicer != null)
        {
            _slicer.setBDeleted(false);
            _slicer.checkEmptySpPrAndXfrm(_xfrm);
            if(AscCommon.isRealObject(_nvGraphicFramePr) )
            {
                _slicer.setNvSpPr(_nvGraphicFramePr);
                if(AscFormat.isRealNumber(_nvGraphicFramePr.locks))
                {
                    _slicer.setLocks(_nvGraphicFramePr.locks);
                }
            }
            return _slicer;
        }
        else if(_timeslicer != null)
        {
            _timeslicer.setBDeleted(false);
            _timeslicer.checkEmptySpPrAndXfrm(_xfrm);
            if(AscCommon.isRealObject(_nvGraphicFramePr) )
            {
                _timeslicer.setNvSpPr(_nvGraphicFramePr);
                if(AscFormat.isRealNumber(_nvGraphicFramePr.locks))
                {
                    _timeslicer.setLocks(_nvGraphicFramePr.locks);
                }
            }
            return _timeslicer;
        }
        else if(_smartArt != null)
        {
            _smartArt.checkEmptySpPrAndXfrm(_xfrm);
            if(AscCommon.isRealObject(_nvGraphicFramePr) )
            {
                _smartArt.setNvSpPr(_nvGraphicFramePr);
                if(AscFormat.isRealNumber(_nvGraphicFramePr.locks))
                {
                    _smartArt.setLocks(_nvGraphicFramePr.locks);
                }
                if(_nvGraphicFramePr.cNvPr) 
                {
                    this.map_shapes_by_id[_nvGraphicFramePr.cNvPr.id] = _smartArt;
                }
            }
            return _smartArt;
        }

        return _graphic_frame;
    };

    this.ReadSmartArt = function(CDrawing)
    {
        var s = this.stream;
        var _smartArt;
        if(typeof AscFormat.SmartArt !== "undefined" && !CDrawing)
        {
            _smartArt = new AscFormat.SmartArt();
            _smartArt.fromPPTY(this);
            _smartArt.setBDeleted(false);
						_smartArt.generateDefaultStructures();
            _smartArt.checkDataModel();
            _smartArt.checkNodePointsAfterRead();
						_smartArt.correctUngeneratedSmartArtContent();
        }
        else
        {
            //read drawing part of smartArt as a group
            var _rec_start = s.cur;
            var _end_rec = _rec_start + s.GetULong() + 4;
            //no attributes
            while (s.cur < _end_rec)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 0:
                    {
                        //DiagramDrawing
                        _smartArt = this.ReadSmartArtGroup(CDrawing);
                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }
        }
        s.Seek2(_end_rec);
        _smartArt.setType(_smartArt.getTypeOfSmartArt());
        return _smartArt;
    };

    this.ReadSmartArtGroup = function(CDrawing)
    {
        var _smartArt;
        var s = this.stream;
        var _rec_start2 = s.cur;
        var _end_rec2 = _rec_start2 + s.GetULong() + 4;
        while (s.cur < _end_rec2)
        {
            var _at2 = s.GetUChar();
            switch (_at2)
            {
                case 0:
                {
                    //_smartArt = this.ReadGroupShape();
                    var shapes = this.ReadGroupShapeMain();
                    if (CDrawing) {
                        _smartArt = CDrawing;
                    } else {
                        _smartArt = new AscFormat.CGroupShape();
                    }
                    _smartArt.setBDeleted(false);
                    for(var nSp = 0; nSp < shapes.length; ++nSp)
                    {
                        var oSp = shapes[nSp];
                        _smartArt.addToSpTree(undefined, oSp);
                        oSp.setGroup(_smartArt);
                    }

                    _smartArt.setParent(this.TempMainObject == null ? this.ParaDrawing : null);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }
        s.Seek2(_end_rec2);
        return _smartArt;
    };

    this.ReadNvUniProp = function(drawing)
    {
        var prop = new AscFormat.UniNvPr();

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    this.ReadCNvPr(prop.cNvPr);
                    if(AscCommon.isRealObject(drawing))
                    {
                        this.map_shapes_by_id[prop.cNvPr.id] = drawing;
                    }
                    break;
                }
                case 1:
                {

                    var end = s.cur + s.GetULong() + 4;
                    var locks = 0;
                    if(AscCommon.isRealObject(drawing))
                    {
                        var drawingType = drawing.getObjectType();
                        switch(drawingType)
                        {
                            case AscDFH.historyitem_type_Shape:
                            {
                                s.Skip2(1); // attribute start
                                while (true)
                                {
                                    var _at2 = s.GetUChar();
                                    if (_at2 == g_nodeAttributeEnd)
                                        break;

                                    var value;
                                    switch(_at2)
                                    {
                                        case 0:{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.txBox, value);
                                            break;
                                        }
                                        case 1 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noAdjustHandles, value);
                                            break;
                                        }
                                        case 2 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noChangeArrowheads, value);
                                            break;
                                        }
                                        case 3 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noChangeAspect, value);
                                            break;
                                        }
                                        case 4 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noChangeShapeType, value);
                                            break;
                                        }
                                        case 5 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noEditPoints, value);
                                            break;
                                        }
                                        case 6 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noGrp, value);
                                            break;
                                        }
                                        case 7 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noMove, value);
                                            break;
                                        }
                                        case 8 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noResize, value);
                                            break;
                                        }
                                        case 9 :{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noRot, value);
                                            break;
                                        }
                                        case 10:{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noSelect, value);
                                            break;
                                        }
                                        case 11:{
                                            value = s.GetBool();
                                            locks = AscFormat.fUpdateLocksValue(locks, AscFormat.LOCKS_MASKS.noTextEdit, value);
                                            break;
                                        }
                                    }
                                }
                                prop.locks = locks;
                                break;
                            }
                            case AscDFH.historyitem_type_GroupShape:
                            {
                                s.Skip2(1); // attribute start
                                while (true)
                                {
                                    var _at2 = s.GetUChar();
                                    if (_at2 == g_nodeAttributeEnd)
                                        break;

                                    var value;
                                    switch(_at2)
                                    {
                                        case 0:
                                        {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeAspect | (value ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
                                            break;
                                        }
                                        case 1:
                                        {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noGrp | (value ? AscFormat.LOCKS_MASKS.noGrp << 1 : 0));
                                            break;
                                        }
                                        case 2:
                                        {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noMove | (value ? AscFormat.LOCKS_MASKS.noMove << 1 : 0));
                                            break;
                                        }
                                        case 3:
                                        {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noResize | (value ? AscFormat.LOCKS_MASKS.noResize << 1 : 0));
                                            break;
                                        }
                                        case 4:
                                        {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noRot | (value ? AscFormat.LOCKS_MASKS.noRot << 1 : 0));
                                            break;
                                        }
                                        case 5:
                                        {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noSelect | (value ? AscFormat.LOCKS_MASKS.noSelect << 1 : 0));
                                            break;
                                        }
                                        case 6:
                                        {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noUngrp | (value ? AscFormat.LOCKS_MASKS.noUngrp << 1 : 0));
                                            break;
                                        }
                                    }
                                }
                                prop.locks = locks;
                                break;
                            }
                            case AscDFH.historyitem_type_ImageShape:
                            {
                                s.Skip2(1); // attribute start
                                while (true)
                                {
                                    var _at2 = s.GetUChar();
                                    if (_at2 == g_nodeAttributeEnd)
                                        break;

                                    var value;
                                    switch(_at2)
                                    {
                                        case 0 :{
                                            value = s.GetBool();
                                            break;
                                        }
                                        case 1 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noAdjustHandles | (value ? AscFormat.LOCKS_MASKS.noAdjustHandles << 1 : 0));
                                            break;
                                            }
                                        case 2 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeArrowheads | (value ? AscFormat.LOCKS_MASKS.noChangeArrowheads << 1 : 0));
                                            break;
                                            }
                                        case 3 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeAspect | (value ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
                                            break;
                                            }
                                        case 4 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeShapeType | (value ? AscFormat.LOCKS_MASKS.noChangeShapeType << 1 : 0));
                                            break;
                                            }
                                        case 5 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noCrop | (value ? AscFormat.LOCKS_MASKS.noCrop << 1 : 0));
                                            break;
                                            }
                                        case 6 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noEditPoints | (value ? AscFormat.LOCKS_MASKS.noEditPoints << 1 : 0));
                                            break;
                                            }
                                        case 7 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noGrp | (value ? AscFormat.LOCKS_MASKS.noGrp << 1 : 0));
                                            break;
                                            }
                                        case 8 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noMove | (value ? AscFormat.LOCKS_MASKS.noMove << 1 : 0));
                                            break;
                                            }
                                        case 9 :{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noResize | (value ? AscFormat.LOCKS_MASKS.noResize << 1 : 0));
                                            break;
                                            }
                                        case 10:{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noRot | (value ? AscFormat.LOCKS_MASKS.noRot << 1 : 0));
                                            break;
                                            }
                                        case 11:{
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noSelect | (value ? AscFormat.LOCKS_MASKS.noSelect << 1 : 0));
                                            break;
                                            }
                                    }
                                }
                                prop.locks = locks;
                                break;
                            }
                            case AscDFH.historyitem_type_GraphicFrame:
                            case AscDFH.historyitem_type_ChartSpace:
                            case AscDFH.historyitem_type_SmartArt:
                            {
                                s.Skip2(1); // attribute start
                                while (true)
                                {
                                    var _at2 = s.GetUChar();
                                    if (_at2 == g_nodeAttributeEnd)
                                        break;

                                    var value;
                                    switch(_at2)
                                    {
                                        case 0: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeAspect | (value ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
                                            break;
                                        }
                                        case 1: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noDrilldown | (value ? AscFormat.LOCKS_MASKS.noDrilldown << 1 : 0));
                                            break;
                                        }
                                        case 2: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noGrp | (value ? AscFormat.LOCKS_MASKS.noGrp << 1 : 0));
                                            break;
                                        }
                                        case 3: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noMove | (value ? AscFormat.LOCKS_MASKS.noMove << 1 : 0));
                                            break;
                                        }
                                        case 4: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noResize | (value ? AscFormat.LOCKS_MASKS.noResize << 1 : 0));
                                            break;
                                        }
                                        case 5: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noSelect | (value ? AscFormat.LOCKS_MASKS.noSelect << 1 : 0));
                                            break;
                                        }
                                    }
                                }
                                prop.locks = locks;
                                break;
                            }
                            case AscDFH.historyitem_type_Cnx:{

                                s.Skip2(1); // attribute start
                                while (true)
                                {
                                    var _at2 = s.GetUChar();
                                    if (_at2 == g_nodeAttributeEnd)
                                        break;

                                    var value;
                                    switch(_at2)
                                    {
                                        case 0: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noAdjustHandles | (value ? AscFormat.LOCKS_MASKS.noAdjustHandles << 1 : 0));
                                            break;
                                        }
                                        case 1: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeArrowheads | (value ? AscFormat.LOCKS_MASKS.noChangeArrowheads << 1 : 0));
                                            break;
                                        }
                                        case 2: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeAspect | (value ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
                                            break;
                                        }
                                        case 3: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noChangeShapeType | (value ? AscFormat.LOCKS_MASKS.noChangeShapeType << 1 : 0));
                                            break;
                                        }
                                        case 4: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noEditPoints | (value ? AscFormat.LOCKS_MASKS.noEditPoints << 1 : 0));
                                            break;
                                        }
                                        case 5: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noGrp | (value ? AscFormat.LOCKS_MASKS.noGrp << 1 : 0));
                                            break;
                                        }
                                        case 6: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noMove | (value ? AscFormat.LOCKS_MASKS.noMove << 1 : 0));
                                            break;
                                        }
                                        case 7: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noResize | (value ? AscFormat.LOCKS_MASKS.noResize << 1 : 0));
                                            break;
                                        }
                                        case 8: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noRot | (value ? AscFormat.LOCKS_MASKS.noRot << 1 : 0));
                                            break;
                                        }
                                        case 9: {
                                            value = s.GetBool();
                                            locks |= (AscFormat.LOCKS_MASKS.noSelect | (value ? AscFormat.LOCKS_MASKS.noSelect << 1 : 0));
                                            break;
                                        }
                                        case 10:{
                                            prop.nvUniSpPr.stCnxId = s.GetULong();
                                            break;
                                        }
                                        case 11:{
                                            prop.nvUniSpPr.stCnxIdx = s.GetULong();
                                            break;
                                        }
                                        case 12:{
                                            prop.nvUniSpPr.endCnxId = s.GetULong();
                                            break;
                                        }
                                        case 13:{
                                            prop.nvUniSpPr.endCnxIdx = s.GetULong();
                                            break;
                                        }
                                    }
                                }
                                prop.locks = locks;
                                prop.setUniSpPr(prop.nvUniSpPr.copy());
                                break;
                            }
                        }
                    }
                    s.Seek2(end);
                    break;
                }
                case 2:
                {
                    this.ReadNvPr(prop.nvPr);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return prop;
    };

    this.ReadCNvPr = function(cNvPr)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    cNvPr.setId(s.GetLong());
                    break;
                }
                case 1:
                {
                    cNvPr.setName(s.GetString2());
                    break;
                }
                case 2:
                {
                    cNvPr.setIsHidden((1 == s.GetUChar()) ? true : false);
                    break;
                }
                case 3:
                {
                    cNvPr.setTitle(s.GetString2());
                    break;
                }
                case 4:
                {
                    cNvPr.setDescr(s.GetString2());
                    break;
                }
                case 5:
                {
                    cNvPr.form =  s.GetBool();
                    break;
                }
                default:{
                    break;
                }
            }
        }

        while(s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch(_at){
                case 0:{
                    cNvPr.setHlinkClick(this.ReadHyperlink());
                    break;
                }
                case 1:{
                    cNvPr.setHlinkHover(this.ReadHyperlink());
                    break;
                }
                default:{
                    this.stream.SkipRecord();
                    break;
                }
            }
        }
        s.Seek2(_end_rec);
    };

    this.ReadTable = function(_xfrm, _graphic_frame)
    {
        if (_xfrm == null)
        {
            this.stream.SkipRecord();
            return null;
        }

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        var cols = null;
        var rows = null;
        var _return_to_rows = 0;
        var props = null;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    props = this.ReadTablePr();
                    break;
                }
                case 1:
                {
                    s.Skip2(4);
                    var _len = s.GetULong();
                    cols = new Array(_len);
                    for (var i = 0; i < _len; i++)
                    {
                        s.Skip2(7); // type, len + startAttr + 0 (attrType)
                        cols[i] = s.GetULong() / 36000;
                        s.Skip2(1); // endAttr
                    }
                    break;
                }
                case 2:
                {
                    var _end_rec2 = s.cur + s.GetULong() + 4;
                    rows = s.GetULong();
                    _return_to_rows = s.cur;
                    s.Seek2(_end_rec2);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        if(cols.length === 0)
        {
            cols.push(_xfrm.extX);
        }
        var _table = new CTable(this.presentation.DrawingDocument, _graphic_frame, false, rows, cols.length, cols, true);
        _table.Reset(0, 0, _xfrm.extX, 100000, 0, 0, 1);
        if (null != props)
        {
            var style;
            if(this.map_table_styles[props.style])
            {
                _table.Set_TableStyle(this.map_table_styles[props.style].Id);
            }
            else if (this.presentation && this.presentation.globalTableStyles && this.presentation.globalTableStyles.GetStyleByStyleId(props.style))
            {
                style = this.presentation.globalTableStyles.GetStyleByStyleId(props.style);
                _table.Set_TableStyle(style.GetId());
            }
            _table.Set_Pr(props.props);
            _table.Set_TableLook(props.look);
        }
        _table.SetTableLayout(tbllayout_Fixed);

        s.Seek2(_return_to_rows);

        for (var i = 0; i < rows; i++)
        {
            s.Skip2(1); // 0!
            this.ReadTableRow(_table.Content[i]);
        }

        s.Seek2(_end_rec);

        return _table;
    };

    this.ReadTableRow = function(row)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

		var fRowHeight = 5;
        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
					fRowHeight = s.GetULong() / 36000;

                    break;
                }
                default:
                    break;
            }
        }

        s.Skip2(5); // type + len
        var _count = s.GetULong();
        _count = Math.min(_count, row.Content.length);
		for (var i = 0; i < _count; i++)
		{
			s.Skip2(1);
			var bIsNoHMerge = this.ReadCell(row.Content[i]);
			if (bIsNoHMerge === false)
			{
				row.Remove_Cell(i);
				i--;
				_count--;
			}
			var _gridCol = 1;
			if ("number" == typeof (row.Content[i].Pr.GridSpan))
			{
				_gridCol = row.Content[i].Pr.GridSpan;
			}

			if (_gridCol > (_count - i))
			{
				_gridCol = _count - i;
				row.Content[i].Pr.GridSpan = _gridCol;
				if (1 == row.Content[i].Pr.GridSpan)
					row.Content[i].Pr.GridSpan = undefined;
			}

			_gridCol--;
			while (_gridCol > 0)
			{
				i++;
				if (i >= _count)
					break;

				s.Skip2(1);
				this.ReadCell(row.Content[i]);

				// удаляем
				row.Remove_Cell(i);
				i--;
				_count--;

				--_gridCol;
			}
		}

		if(this.presentation && Array.isArray(this.presentation.Slides)){
            AscFormat.updateRowHeightAfterOpen(row, fRowHeight);
        }
        s.Seek2(_end_rec);
    };

    this.ReadCell = function(cell)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    var _id = s.GetString2();
                    break;
                }
                case 1:
                {
                    var rowSpan = s.GetULong();
                    if (1 < rowSpan)
                    {
                        cell.SetVMerge(vmerge_Restart);
                    }
                    break;
                }
                case 2:
                {
                    cell.Set_GridSpan(s.GetULong());
                    break;
                }
                case 3:
                {
                    var bIsHMerge = s.GetBool();
                    if (bIsHMerge && cell.Index > 0)
                    {
                        s.Seek2(_end_rec);
                        return false;
                    }
                    break;
                }
                case 4:
                {
                    var bIsVMerge = s.GetBool();
                    if (bIsVMerge && cell.Pr.VMerge != vmerge_Restart)
                    {
                        cell.SetVMerge(vmerge_Continue);
                    }
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var props = new CTableCellPr();
                    this.ReadCellProps(props);
                    props.Merge(cell.Pr);
                    cell.Set_Pr(props);
                    break;
                }
                case 1:
                {
                    this.ReadTextBody2(cell.Content);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return true;
    };

    this.ReadCellProps = function(props)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        //props.TableCellMar = {};
        //props.TableCellMar.Top    = new CTableMeasurement(tblwidth_Mm, 1.27);
        //props.TableCellMar.Left   = new CTableMeasurement(tblwidth_Mm, 2.54);
        //props.TableCellMar.Bottom = new CTableMeasurement(tblwidth_Mm, 1.27);
        //props.TableCellMar.Right  = new CTableMeasurement(tblwidth_Mm, 2.54);

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    if(props.TableCellMar == null)
                        props.TableCellMar = {};
                    props.TableCellMar.Left   = new CTableMeasurement(tblwidth_Mm, s.GetULong() / 36000);
                    //props.TableCellMar.Left.W = s.GetULong() / 36000;
                    break;
                }
                case 1:
                {
                    if(props.TableCellMar == null)
                        props.TableCellMar = {};
                    props.TableCellMar.Top = new CTableMeasurement(tblwidth_Mm, s.GetULong() / 36000);

                    //  props.TableCellMar.Top.W = s.GetULong() / 36000;
                    break;
                }
                case 2:
                {
                    if(props.TableCellMar == null)
                        props.TableCellMar = {};
                    props.TableCellMar.Right   = new CTableMeasurement(tblwidth_Mm, s.GetULong() / 36000);
                    // props.TableCellMar.Right.W = s.GetULong() / 36000;
                    break;
                }
                case 3:
                {
                    if(props.TableCellMar == null)
                        props.TableCellMar = {};
                    props.TableCellMar.Bottom   = new CTableMeasurement(tblwidth_Mm, s.GetULong() / 36000);

                    //props.TableCellMar.Bottom.W = s.GetULong() / 36000;
                    break;
                }
                case 4:
                {
                    s.Skip2(1);
                    break;
                }
                case 5:
                {
                    var nVert = s.GetUChar();
                    switch (nVert)
                    {
                        case 0: props.TextDirection = Asc.c_oAscCellTextDirection.TBRL; break;
                        case 1: props.TextDirection = Asc.c_oAscCellTextDirection.LRTB;/*_T("horz"); */break;
                        case 2: props.TextDirection = Asc.c_oAscCellTextDirection.TBRL; break;
                        case 3: props.TextDirection = Asc.c_oAscCellTextDirection.TBRL; break;
                        case 4: props.TextDirection = Asc.c_oAscCellTextDirection.BTLR; break;
                        case 5: props.TextDirection = Asc.c_oAscCellTextDirection.BTLR; break;
                        case 6: props.TextDirection = Asc.c_oAscCellTextDirection.TBRL; break;
                        default:
                            props.TextDirection = Asc.c_oAscCellTextDirection.LRTB;
                            break;
                    }
                    break;
                }
                case 6:
                {
                    var nVertAlign = s.GetUChar();
                    switch (nVertAlign)
                    {
                        case 0://bottom
                        {
                            props.VAlign = vertalignjc_Bottom;
                            break;
                        }
                        case 1://ctr
                        case 2://dist
                        case 3: //just
                        {
                            props.VAlign = vertalignjc_Center;
                            break;
                        }
                        case 4://top
                        {
                            props.VAlign = vertalignjc_Top;
                            break;
                        }
                    }
                    //s.Skip2(1);
                    break;
                }
                case 7:
                {
                    s.Skip2(1);
                    break;
                }
                default:
                    break;
            }
        }


        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    if(!props.TableCellBorders)
                    {
                        props.TableCellBorders = {};
                    }
                    props.TableCellBorders.Left = this.ReadTableBorderLn();
                    break;
                }
                case 1:
                {
                    if(!props.TableCellBorders)
                    {
                        props.TableCellBorders = {};
                    }
                    props.TableCellBorders.Top = this.ReadTableBorderLn();
                    break;
                }
                case 2:
                {
                    if(!props.TableCellBorders)
                    {
                        props.TableCellBorders = {};
                    }
                    props.TableCellBorders.Right = this.ReadTableBorderLn();
                    break;
                }
                case 3:
                {
                    if(!props.TableCellBorders)
                    {
                        props.TableCellBorders = {};
                    }
                    props.TableCellBorders.Bottom = this.ReadTableBorderLn();
                    break;
                }
                case 4:
                {
                    s.SkipRecord();
                    break;
                }
                case 5:
                {
                    s.SkipRecord();
                    break;
                }
                case 6:
                {
                    var _unifill = this.ReadUniFill();

                    if (_unifill && _unifill.fill !== undefined && _unifill.fill != null)
                    {
                        props.Shd = new CDocumentShd();
                        props.Shd.Value = c_oAscShdClear;
                        props.Shd.Unifill = _unifill;
                    }
                    break;
                }
                case 7:
                {
                    s.SkipRecord();
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadTableBorderLn = function()
    {
        var ln = this.ReadLn();

        var border = new CDocumentBorder();
        ln.fillDocumentBorder(border);

        return border;
    };

    this.ReadTablePr = function()
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        var obj = {};
        obj.props = new CTablePr();
        obj.look = new AscCommon.CTableLook(false, false, false, false, false, false);
        obj.style = -1;

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                   //var ind = this.map_table_styles[s.GetString2()];
                   //if (undefined !== ind)
                    obj.style = s.GetString2();
                    break;
                }
                case 1:
                {
                    s.Skip2(1);// rtl
                    break;
                }
                case 2:
                {
                    obj.look.SetFirstRow(s.GetBool());
                    break;
                }
                case 3:
                {
                    obj.look.SetFirstCol(s.GetBool());
                    break;
                }
                case 4:
                {
                    obj.look.SetLastRow(s.GetBool());
                    break;
                }
                case 5:
                {
                    obj.look.SetLastCol(s.GetBool());
                    break;
                }
                case 6:
                {
                    obj.look.SetBandHor(s.GetBool());
                    break;
                }
                case 7:
                {
                    obj.look.SetBandVer(s.GetBool());
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    var _unifill = this.ReadUniFill();
                    if (_unifill && _unifill.fill !== undefined && _unifill.fill != null)
                    {
                        obj.props.Shd = new CDocumentShd();
                        obj.props.Shd.Value = c_oAscShdClear;
                        obj.props.Shd.Unifill = _unifill;
                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return obj;
    };

    this.ReadNvPr = function(nvPr)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    nvPr.setIsPhoto(s.GetBool());
                    break;
                }
                case 1:
                {
                    nvPr.setUserDrawn(s.GetBool());
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    nvPr.setPh(this.ReadPH());
                    break;
                }
                case 1:
                {
                    nvPr.setUniMedia(new AscFormat.UniMedia());
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
                default:
                {
                    var _len = s.GetULong();
                    s.Skip2(_len);
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadPH = function()
    {
        var ph = new AscFormat.Ph();
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    ph.setHasCustomPrompt(s.GetBool());
                    break;
                }
                case 1:
                {
                    ph.setIdx(s.GetString2());
                    break;
                }
                case 2:
                {
                    ph.setOrient(s.GetUChar());
                    break;
                }
                case 3:
                {
                    ph.setSz(s.GetUChar());
                    break;
                }
                case 4:
                {
                    ph.setType(s.GetUChar());
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);
        return ph;
    };

    // ------------------------------------------

    // TEXT PROPERTIES --------------------------

    this.ReadRunProperties = function()
    {
        var rPr = new CTextPr();

        var s = this.stream;
        var _end_rec = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    var altLang = s.GetString2();
                    break;
                }
                case 1:
                {
                    rPr.Bold = s.GetBool();
                    break;
                }
                case 2:
                {
                    var baseline = s.GetLong();

                    if (baseline < 0)
                        rPr.VertAlign = AscCommon.vertalign_SubScript;
                    else if (baseline > 0)
                        rPr.VertAlign = AscCommon.vertalign_SuperScript;

                    break;
                }
                case 3:
                {
                    var bmk = s.GetString2();
                    break;
                }
                case 4:
                {
                    var _cap = s.GetUChar();
                    if (_cap == 0)
                    {
                        rPr.Caps = true;
                        rPr.SmallCaps = false;
                    }
                    else if (_cap == 1)
                    {
                        rPr.Caps = false;
                        rPr.SmallCaps = true;
                    }
                    else if (_cap == 2)
                    {
                        rPr.SmallCaps = false;
                        rPr.Caps = false;
                    }
                    break;
                }
                case 5:
                {
                    s.Skip2(1); // dirty
                    break;
                }
                case 6:
                {
                    s.Skip2(1); // error
                    break;
                }
                case 7:
                {
                    rPr.Italic = s.GetBool();
                    break;
                }
                case 8:
                {
                    s.Skip2(4); // kern
                    break;
                }
                case 9:
                {
                    s.Skip2(1); // kumimoji
                    break;
                }
                case 10:
                {
                    var lang = s.GetString2();
                    if(!(this.IsThemeLoader || this.IsFillingSmartArt))
                    {
                        var nLcid = Asc.g_oLcidNameToIdMap[lang];
                        if(nLcid)
                            rPr.Lang.Val = nLcid;
                    }
                    break;
                }
                case 11:
                {
                    s.Skip2(1); // noproof
                    break;
                }
                case 12:
                {
                    s.Skip2(1); // normalizeH
                    break;
                }
                case 13:
                {
                    s.Skip2(1); // smtClean
                    break;
                }
                case 14:
                {
                    s.Skip2(4); // smtId
                    break;
                }
                case 15:
                {
                    //s.Skip2(4); // spc
                    rPr.Spacing = s.GetLong() * 25.4 / 7200;
                    break;
                }
                case 16:
                {
                    var _strike = s.GetUChar();
                    if (0 == _strike)
                    {
                        rPr.Strikeout = false;
                        rPr.DStrikeout = true;
                    }
                    else if (2 == _strike)
                    {
                        rPr.Strikeout = true;
                        rPr.DStrikeout = false;
                    }
                    else
                    {
                        rPr.Strikeout = false;
                        rPr.DStrikeout = false;
                    }
                    break;
                }
                case 17:
                {
                    var _size = s.GetLong() / 100;
                    _size = ((_size * 2) + 0.5) >> 0;
                    _size /= 2;
                    rPr.FontSize = _size;
                    rPr.FontSizeCS = _size;
                    break;
                }
                case 18:
                {
                    rPr.Underline = (s.GetUChar() != 12);
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    rPr.TextOutline = this.ReadLn();
                    break;
                }
                case 1:
                {
                    var oUniFill = this.ReadUniFill();
                    if(oUniFill && oUniFill.fill){
                        rPr.Unifill = oUniFill;
                    }
                    break;
                }
                case 2:
                {
                    s.SkipRecord();
                    break;
                }
                case 3:
                {
                    //latin
                    rPr.RFonts.Ascii = { Name: this.ReadTextFontTypeface(), Index : -1 };
                    rPr.RFonts.HAnsi = { Name: rPr.RFonts.Ascii.Name, Index : -1 };
                    break;
                }
                case 4:
                {
                    //ea
                    rPr.RFonts.EastAsia = { Name: this.ReadTextFontTypeface(), Index : -1 };
                    break;
                }
                case 5:
                {
                    //cs
                    rPr.RFonts.CS = { Name: this.ReadTextFontTypeface(), Index : -1 };
                    break;
                }
                case 6:
                {
                    //sym

                    s.SkipRecord();
                    //rPr.RFonts.HAnsi = { Name: this.ReadTextFontTypeface(), Index : -1 };
                    break;
                }
                case 7:
                {
                    rPr.hlink = this.ReadHyperlink();
                    if (null == rPr.hlink)
                        delete rPr.hlink;
                    break;
                }
                case 8:
                {
                    s.SkipRecord();
                    break;
                }
                case 12:
                {
                    //highlight
                    var end_rec__ = s.cur + s.GetULong() + 4;

                    s.Skip2(1); // start attributes
                    var  at__;
                    while (true)
                    {
                        at__ = s.GetUChar();
                        if (at__ === g_nodeAttributeEnd)
                            break;
                    }
                    while (s.cur < end_rec__)
                    {
                        at__ = s.GetUChar();
                        switch (at__)
                        {
                            case 0:
                            {
                                rPr.HighlightColor = this.ReadUniColor();
                                break;
                            }
                            default:
                            {
                                break;
                            }
                        }
                    }

                    s.Seek2(end_rec__);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                }
            }
        }

        s.Seek2(_end_rec);
        //checkTextPr(rPr);
        return rPr;
    };

    this.ReadHyperlink = function()
    {
        var hyper = new AscFormat.CT_Hyperlink();
        var s = this.stream;
        var _end_rec = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    hyper.id = s.GetString2();
                    break;
                }
                case 1:
                {
                    hyper.invalidUrl = s.GetString2();
                    break;
                }
                case 2:
                {
                    hyper.action = s.GetString2();
                    break;
                }
                case 3:
                {
                    hyper.tgtFrame = s.GetString2();
                    break;
                }
                case 4:
                {
                    hyper.tooltip = s.GetString2();
                    break;
                }
                case 5:
                {
                    hyper.history = s.GetBool();
                    break;
                }
                case 6:
                {
                    hyper.highlightClick = s.GetBool();
                    break;
                }
                case 7:
                {
                    hyper.endSnd = s.GetBool();
                    break;
                }
                default:
                    break;
            }
        }

        s.Seek2(_end_rec);

        // correct hyperlink
        if (hyper.action != null && hyper.action != "")
        {
            if (hyper.action == "ppaction://hlinkshowjump?jump=firstslide")
                hyper.id = "ppaction://hlinkshowjump?jump=firstslide";
            else if (hyper.action == "ppaction://hlinkshowjump?jump=lastslide")
                hyper.id = "ppaction://hlinkshowjump?jump=lastslide";
            else if (hyper.action == "ppaction://hlinkshowjump?jump=nextslide")
                hyper.id = "ppaction://hlinkshowjump?jump=nextslide";
            else if (hyper.action == "ppaction://hlinkshowjump?jump=previousslide")
                hyper.id = "ppaction://hlinkshowjump?jump=previousslide";
            else if (hyper.action == "ppaction://hlinksldjump")
            {
                if (hyper.id != null && hyper.id.indexOf("slide") == 0)
                {
                    var _url = hyper.id.substring(5);
                    var _indexXml = _url.indexOf(".");
                    if (-1 != _indexXml)
                        _url = _url.substring(0, _indexXml);

                    var _slideNum = parseInt(_url);
                    if (isNaN(_slideNum))
                        _slideNum = 1;

                    --_slideNum;

                    hyper.id = hyper.action + "slide" + _slideNum;
                }
                else
                {
                    hyper.id = null;
                }
            }
            else
            {
                hyper.id = null;
            }
        }

        if (hyper.id == null)
            return null;

        return hyper;
    };

    this.CorrectBodyPr = function(bodyPr)
    {

        //TODO: сделать через методы
        var s = this.stream;
        var _end_rec = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    bodyPr.flatTx = s.GetLong();
                    break;
                }
                case 1:
                {
                    bodyPr.anchor = s.GetUChar();
                    break;
                }
                case 2:
                {
                    bodyPr.anchorCtr = s.GetBool();
                    break;
                }
                case 3:
                {
                    bodyPr.bIns = s.GetLong()/36000;
                    break;
                }
                case 4:
                {
                    bodyPr.compatLnSpc = s.GetBool();
                    break;
                }
                case 5:
                {
                    bodyPr.forceAA = s.GetBool();
                    break;
                }
                case 6:
                {
                    bodyPr.fromWordArt = s.GetBool();
                    break;
                }
                case 7:
                {
                    bodyPr.horzOverflow = s.GetUChar();
                    break;
                }
                case 8:
                {
                    bodyPr.lIns = s.GetLong()/36000;
                    break;
                }
                case 9:
                {
                    bodyPr.numCol = s.GetLong();
                    break;
                }
                case 10:
                {
                    bodyPr.rIns = s.GetLong()/36000;
                    break;
                }
                case 11:
                {
                    bodyPr.rot = s.GetLong();
                    break;
                }
                case 12:
                {
                    bodyPr.rtlCol = s.GetBool();
                    break;
                }
                case 13:
                {
                    bodyPr.spcCol = s.GetLong()/36000;
                    break;
                }
                case 14:
                {
                    bodyPr.spcFirstLastPara = s.GetBool();
                    break;
                }
                case 15:
                {
                    bodyPr.tIns = s.GetLong()/36000;
                    break;
                }
                case 16:
                {
                    bodyPr.upright = s.GetBool();
                    break;
                }
                case 17:
                {
                    bodyPr.setVertOpen(s.GetUChar());
                    break;
                }
                case 18:
                {
                    bodyPr.vertOverflow = s.GetUChar();
                    break;
                }
                case 19:
                {
                    bodyPr.wrap = s.GetUChar();
                    break;
                }
                default:
                    break;
            }
        }

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0://prstTxWarp
                {
                    var _end_rec3 = s.cur + s.GetULong() + 4;
                    s.Skip2(1);// start attributes
                    while(true)
                    {
                        var _at2 = s.GetUChar();
                        if (_at2 == g_nodeAttributeEnd)
                            break;
                        switch (_at2) {
                            case 0:
                            {
                                var sPrst = s.GetUChar();
                                bodyPr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry(AscFormat.getPrstByNumber(sPrst));
                                break;
                            }
                        }
                    }
                    while (s.cur < _end_rec3)
                    {
                       var _at = s.GetUChar();
                       switch (_at)
                       {
                           case 0:
                           {
                               this.ReadGeomAdj(bodyPr.prstTxWarp );
                               break;
                           }
                           default:
                           {
                               s.SkipRecord();
                               break;
                           }
                       }
                    }
                    s.Seek2(_end_rec3);
                    break;
                }
                case 1:
                {
                    var _end_rec2 = s.cur + s.GetULong() + 4;

                    s.Skip2(1); // start attributes

                    var txFit = new AscFormat.CTextFit();
                    txFit.type = -1;

                    while (true)
                    {
                        var _at2 = s.GetUChar();
                        if (_at2 == g_nodeAttributeEnd)
                            break;

                        switch (_at2)
                        {
                            case 0:
                            {
                                txFit.type = s.GetLong() - 1;
                                break;
                            }
                            case 1:
                            {
                                txFit.fontScale = s.GetLong();
                                break;
                            }
                            case 2:
                            {
                                txFit.lnSpcReduction = s.GetLong();
                                break;
                            }
                            default:
                                break;
                        }
                    }
                    if (txFit.type != -1)
                    {
                        bodyPr.textFit = txFit;
                    }

                    s.Seek2(_end_rec2);
                    break;
                }
                default:
                {
                    s.SkipRecord();
                }
            }
        }
        s.Seek2(_end_rec);
    };

    this.ReadBodyPr = function()
    {
        var bodyPr = new AscFormat.CBodyPr();
        this.CorrectBodyPr(bodyPr);
        return bodyPr;
    };

    this.ReadTextParagraphPr = function(par)
    {

        var para_pr = new CParaPr();
        var s = this.stream;
        var _end_rec = s.cur + s.GetULong() + 4;

        s.Skip2(1); // start attributes

        while (true)
        {
            var _at = s.GetUChar();
            if (_at == g_nodeAttributeEnd)
                break;

            switch (_at)
            {
                case 0:
                {
                    var _align = s.GetUChar();
                    switch (_align)
                    {
                        case 0: { para_pr.Jc = AscCommon.align_Center; break; }
                        case 1: { para_pr.Jc = AscCommon.align_Justify; break; }
                        case 2: { para_pr.Jc = AscCommon.align_Justify; break; }
                        case 3: { para_pr.Jc = AscCommon.align_Justify; break; }
                        case 4: { para_pr.Jc = AscCommon.align_Left; break; }
                        case 5: { para_pr.Jc = AscCommon.align_Right; break; }
                        case 6: { para_pr.Jc = AscCommon.align_Justify; break; }
                        default:
                            para_pr.Jc = AscCommon.align_Center;
                            break;
                    }
                    break;
                }
                case 1:
                {
                    para_pr.DefaultTab = s.GetLong()/36000;
                    break;
                }
                case 2:
                {
                    s.Skip2(1); // eaLnBrk
                    break;
                }
                case 3:
                {
                    s.Skip2(1); // font align
                    break;
                }
                case 4:
                {
                    s.Skip2(1); // hangingPunct
                    break;
                }
                case 5:
                {
                    para_pr.Ind.FirstLine = s.GetLong()/36000;
                    break;
                }
                case 6:
                {
                    s.Skip2(1); // latinLnBrk
                    break;
                }
                case 7:
                {
                    para_pr.Lvl = s.GetLong();
                    break;
                }
                case 8:
                {
                    para_pr.Ind.Left = s.GetLong()/36000;
                    break;
                }
                case 9:
                {
                    para_pr.Ind.Right = s.GetLong()/36000;
                    break;
                }
                case 10:
                {
                    s.Skip2(1); // rtl
                    break;
                }
                default:
                    break;
            }
        }

        var bullet = new AscFormat.CBullet();
        var b_bullet = false;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    s.Skip2(5); // len start attr

                    var Pts = null;
                    var Pct = null;
                    while (true)
                    {
                        var _at = s.GetUChar();
                        if (_at == g_nodeAttributeEnd)
                            break;

                        switch (_at)
                        {
                            case 0:
                            {
                                Pct = s.GetLong();
                                para_pr.Spacing.Line = Pct/100000;
                                para_pr.Spacing.LineRule = Asc.linerule_Auto;
                                break;
                            }
                            case 1:
                            {
                                Pts = s.GetLong();
                                para_pr.Spacing.Line = Pts*0.00352777778;
                                para_pr.Spacing.LineRule = Asc.linerule_Exact;
                                break;
                            }
                            default:
                                break;
                        }
                    }


                    // lnSpc
                    // TODO:
                    break;
                }
                case 1:
                {
                    s.Skip2(5); // len + start attr

                    var Pts = null;
                    var Pct = null;
                    while (true)
                    {
                        var _at = s.GetUChar();
                        if (_at == g_nodeAttributeEnd)
                            break;

                        switch (_at)
                        {
                            case 0:
                            {
                                Pct = s.GetLong();
                                para_pr.Spacing.After = 0;
                                para_pr.Spacing.AfterPct = Pct;
                                break;
                            }
                            case 1:
                            {
                                Pts = s.GetLong();
                                para_pr.Spacing.After = Pts*0.00352777778;
                                break;
                            }
                            default:
                                break;
                        }
                    }
                    // spcAft
                    // TODO:
                    break;
                }
                case 2:
                {
                    s.Skip2(5); // len + start attr

                    var Pts = null;
                    var Pct = null;
                    while (true)
                    {
                        var _at = s.GetUChar();
                        if (_at == g_nodeAttributeEnd)
                            break;

                        switch (_at)
                        {
                            case 0:
                            {
                                Pct = s.GetLong();
                                para_pr.Spacing.Before = 0;
                                para_pr.Spacing.BeforePct = Pct;
                                break;
                            }
                            case 1:
                            {
                                Pts = s.GetLong();
                                para_pr.Spacing.Before = Pts*0.00352777778;
                                break;
                            }
                            default:
                                break;
                        }
                    }

                    // spcBef
                    // TODO:
                    break;
                }
                case 3:
                {

                    var cur_pos = s.cur;
                    var _len = s.GetULong();
                    if (0 != _len)
                    {
                        b_bullet = true;
                        bullet.bulletColor = new AscFormat.CBulletColor();
                        bullet.bulletColor.type = s.GetUChar();

                        if (bullet.bulletColor.type == AscFormat.BULLET_TYPE_COLOR_CLRTX)
                        {
                            s.SkipRecord();
                        }
                        else
                        {
                            var _l = s.GetULong();
                            if (0 !== _l) {
                                s.Skip2(1);
                                bullet.bulletColor.UniColor = this.ReadUniColor();
                            }
                        }
                    }
                    s.Seek2(cur_pos + _len + 4);
                    break;
                }
                case 4:
                {
                    var cur_pos = s.cur;
                    var _len = s.GetULong();
                    if (0 != _len)
                    {
                        b_bullet = true;
                        bullet.bulletSize = new AscFormat.CBulletSize();

                        bullet.bulletSize.type = s.GetUChar();

                        if (bullet.bulletSize.type == AscFormat.BULLET_TYPE_SIZE_TX)
                        {
                            s.SkipRecord();
                        }
                        else
                        {
                            var _l = s.GetULong();
                            s.Skip2(2); // start attributes + type
                            bullet.bulletSize.val = s.GetLong();
                            s.Skip2(1); // end attributes
                        }
                    }
                    s.Seek2(cur_pos + _len + 4);
                    break;
                }
                case 5:
                {

                    var cur_pos = s.cur;
                    var _len = s.GetULong();
                    if (0 != _len)
                    {
                        b_bullet = true;
                        bullet.bulletTypeface = new AscFormat.CBulletTypeface();
                        bullet.bulletTypeface.type = s.GetUChar();

                        if (bullet.bulletTypeface.type == AscFormat.BULLET_TYPE_TYPEFACE_BUFONT)
                        {
                            bullet.bulletTypeface.typeface = this.ReadTextFontTypeface();
                        }
                        else
                        {
                            s.SkipRecord();
                        }
                    }
                    s.Seek2(cur_pos + _len + 4);
                    break;
                }
                case 6:
                {

                    var cur_pos = s.cur;
                    var _len = s.GetULong();
                    if (0 != _len)
                    {
                        b_bullet = true;
                        bullet.bulletType = new AscFormat.CBulletType();
                        bullet.bulletType.type = s.GetUChar();

                        if (bullet.bulletType.type == AscFormat.BULLET_TYPE_BULLET_NONE)
                        {
                            s.SkipRecord();
                        }
                        else if (bullet.bulletType.type == AscFormat.BULLET_TYPE_BULLET_BLIP)
                        {
                            s.Skip2(5);
                            var buBlip = new AscFormat.CBuBlip();
                            buBlip.fromPPTY(this, par, bullet);
                            bullet.bulletType.setBlip(buBlip);
                        }
                        else if (bullet.bulletType.type == AscFormat.BULLET_TYPE_BULLET_AUTONUM)
                        {
                            s.Skip2(5); // len + type + start attr

                            while (true)
                            {
                                var _at = s.GetUChar();
                                if (_at == g_nodeAttributeEnd)
                                    break;

                                switch (_at)
                                {
                                    case 0:
                                    {
                                        bullet.bulletType.AutoNumType = s.GetUChar();
                                        if(AscFonts.IsCheckSymbols) {
                                            var adaptedVal = window['AscCommonWord'].getAdaptedNumberingFormat(bullet.bulletType.AutoNumType);
                                            switch (adaptedVal) {
                                                case Asc.c_oAscNumberingFormat.LowerLetter:
                                                case Asc.c_oAscNumberingFormat.UpperLetter:
                                                case Asc.c_oAscNumberingFormat.LowerRoman:
                                                case Asc.c_oAscNumberingFormat.UpperRoman:
                                                    for (var nValue = 0; nValue < 26; ++nValue)
                                                    {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(97 + nValue);
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(65 + nValue);
                                                    }
                                                    break;

                                                case Asc.c_oAscNumberingFormat.ArabicAlpha:
                                                case Asc.c_oAscNumberingFormat.ArabicAbjad:
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0623);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0628);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x062A);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x062B);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x062C);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x062D);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x062E);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x062F);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0630);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0631);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0632);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0633);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0634);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0635);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0636);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0637);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0638);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0639);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x063A);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0641);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0642);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0643);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0644);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0645);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0646);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0647);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0648);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x064A);
                                                    break;

                                                case Asc.c_oAscNumberingFormat.DecimalFullWidth:
                                                    for (var i = 0; i < 10; i += 1) {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(0xFF10 + i);
                                                    }
                                                    break;

                                                case Asc.c_oAscNumberingFormat.Decimal:
                                                    for (var i = 0; i < 10; i += 1) {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(0x30 + i);
                                                    }
                                                    break;

                                                case Asc.c_oAscNumberingFormat.DecimalEnclosedCircle:
                                                    for (var nValue = 0; nValue < 20; ++nValue)
                                                    {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(0x2460 + nValue);
                                                    }
                                                    for (var i = 0; i < 10; i += 1) {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(0x30 + i);
                                                    }
                                                    break;

                                                case Asc.c_oAscNumberingFormat.ChineseCounting:
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x25CB);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E00);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E8C);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E09);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x56DB);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E94);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x516D);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E03);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x516B);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x4E5D);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x5341);
                                                    break;

                                                case Asc.c_oAscNumberingFormat.Hebrew2:
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D0);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D1);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D2);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D3);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D4);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D5);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D6);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D7);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D8);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05D9);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DB);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DC);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05DE);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E0);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E1);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E2);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E4);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E6);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E7);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E8);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05E9);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x05EA);
                                                    break;

                                                case Asc.c_oAscNumberingFormat.HindiConsonants:
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(2306);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(2307);
                                                    for (var i = 0x0905; i <= 0x0914; i += 1) {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(i);
                                                    }
                                                    break;

                                                case Asc.c_oAscNumberingFormat.HindiVowels:
                                                    for (var i = 0x0915; i <= 0x0939; i += 1) {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(i);
                                                    }
                                                    break;

                                                case Asc.c_oAscNumberingFormat.HindiNumbers:
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0967);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0968);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0969);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x096A);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x096B);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x096C);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x096D);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x096E);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x096F);
                                                    break;
                                                case Asc.c_oAscNumberingFormat.ThaiLetters:
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E01);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E02);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E04);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E25);
                                                    for (var i = 0x0E07; i <= 0x0E23; i += 1) {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(i);
                                                    }
                                                    for (var i = 0x0E27; i <= 0x0E2E; i += 1) {
                                                        AscFonts.FontPickerByCharacter.getFontBySymbol(i);
                                                    }
                                                    break;

                                                case Asc.c_oAscNumberingFormat.ThaiNumbers:
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E50);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E51);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E52);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E53);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E54);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E55);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E56);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E57);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E58);
                                                    AscFonts.FontPickerByCharacter.getFontBySymbol(0x0E59);
                                                    break;

                                                case Asc.c_oAscNumberingFormat.None:
                                                default:
                                                    break;
                                            }
                                        }

                                        break;
                                    }
                                    case 1:
                                    {
                                        bullet.bulletType.startAt = s.GetLong();
                                        break;
                                    }
                                    default:
                                        break;
                                }
                            }
                        }
                        else if (bullet.bulletType.type == AscFormat.BULLET_TYPE_BULLET_CHAR)
                        {
                            s.Skip2(6);
                            bullet.bulletType.Char = s.GetString2();
                            AscFonts.FontPickerByCharacter.getFontsByString(bullet.bulletType.Char);
                            s.Skip2(1);
                        }
                    }
                    s.Seek2(cur_pos + _len + 4);
                    break;
                }
                case 7:
                {
                    s.Skip2(4);
                    var _c = s.GetULong();

                    if (0 != _c)
                    {
                        para_pr.Tabs = new CParaTabs();
                        var _value, _pos;
                        for (var i = 0; i < _c; i++)
                        {
                            s.Skip2(6); // type, len, start attr
                            _value = null;
                            _pos = null;
                            while (true)
                            {
                                var _at = s.GetUChar();
                                if (_at == g_nodeAttributeEnd)
                                    break;

                                switch (_at)
                                {
                                    case 0:
                                    {
                                        _value = s.GetUChar();

                                        if (_value == 0)
                                            _value = tab_Center;
                                        else if (_value == 3)
                                            _value = tab_Right;
                                        else
                                            _value = tab_Left;

                                        break;
                                    }
                                    case 1:
                                    {
                                        _pos = s.GetLong() / 36000;
                                        break;
                                    }
                                    default:
                                        break;
                                }
                            }
                            para_pr.Tabs.Add(new CParaTab(_value, _pos))
                        }
                    }
                    break;
                }
                case 8:
                {
                    var OldBlipCount = 0;

                    if (this.IsUseFullUrl && par)
                        OldBlipCount = this.RebuildImages.length;

                    var r_pr = this.ReadRunProperties();
                    if(r_pr)
                    {
                        para_pr.DefaultRunPr = new CTextPr();

                        if(r_pr.Unifill && !r_pr.Unifill.fill)
                        {
                            r_pr.Unifill = undefined;
                        }
                        para_pr.DefaultRunPr.Set_FromObject(r_pr);


                        if (this.IsUseFullUrl && par)
                        {

                            if(this.RebuildImages.length > OldBlipCount)
                            {
                                for(var _t = OldBlipCount; _t < this.RebuildImages.length; ++_t)
                                {
                                    var oTextPr = new CTextPr();
                                    oTextPr.Set_FromObject(r_pr);
                                    this.RebuildImages[_t].TextPr = oTextPr;
                                    this.RebuildImages[_t].Paragraph = par;
                                }
                            }

                        }

                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                }
            }
        }

        if(b_bullet)
            para_pr.Bullet = bullet;
        // пока записи не поддерживаем
        s.Seek2(_end_rec);
        return para_pr;
    };

    this.ReadTextListStyle = function()
    {
        var styles = new AscFormat.TextListStyle();
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            styles.levels[_at] = this.ReadTextParagraphPr();
        }

        s.Seek2(_end_rec);
        return styles;
    };

    this.ReadTextBody = function(shape)
    {
        var txbody;

        if(shape)
        {
            if(shape.txBody)
            {
                txbody = shape.txBody;
            }
            else
            {
                txbody = new AscFormat.CTextBody();
                txbody.setParent(shape);
            }
        }
        else
        {
            txbody = new AscFormat.CTextBody();
        }
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    txbody.setBodyPr(this.ReadBodyPr());
                    break;
                }
                case 1:
                {
                    txbody.setLstStyle(this.ReadTextListStyle());
                    break;
                }
                case 2:
                {
                    s.Skip2(4);
                    var _c = s.GetULong();
                    txbody.setContent(new AscFormat.CDrawingDocContent(txbody, this.DrawingDocument, 0, 0, 0, 0, 0, 0, true));
                    if(_c>0)
                    {
                        txbody.content.Internal_Content_RemoveAll();
                    }
                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1); // type
                        var _paragraph = this.ReadParagraph(txbody.content);
                        txbody.content.Internal_Content_Add(txbody.content.Content.length, _paragraph);

                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return txbody;
    };

    this.ReadTextBodyTxPr = function(shape)
    {
        var txbody;

        if(shape.txPr)
            txbody = shape.txPr;
        else
        {
            shape.txPr = new AscFormat.CTextBody();
            txbody = shape.txPr;
        }
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    shape.setBodyPr(this.ReadBodyPr());
                    break;
                }
                case 1:
                {
                    txbody.setLstStyle(this.ReadTextListStyle());
                    break;
                }
                case 2:
                {
                    s.Skip2(4);
                    var _c = s.GetULong();
                    /*if(History != null)
                     {
                     History.TurnOff();
                     }*/
                    if(!txbody.content)
                        txbody.content = new AscFormat.CDrawingDocContent(shape, this.DrawingDocument, 0, 0, 0, 0, 0, 0, true);
                    if(_c>0)
                    {
                        txbody.content.Internal_Content_RemoveAll();
                    }

                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1); // type
                        var _paragraph = this.ReadParagraph(txbody.content);
                        _paragraph.Set_Parent(txbody.content);
                        txbody.content.Internal_Content_Add(txbody.content.Content.length, _paragraph);

                    }


                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
        return txbody;
    };

    this.ReadTextBody2 = function(content)
    {
        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;
        var oBodyPr;
        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    s.SkipRecord();
                    break;
                }
                case 1:
                {
                    s.SkipRecord();
                    break;
                }
                case 2:
                {
                    s.Skip2(4);
                    var _c = s.GetULong();
                    if(_c > 0)
                    {
                        content.Internal_Content_RemoveAll();
                    }
                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(1); // type
                        var _paragraph = this.ReadParagraph(content);
                        content.Internal_Content_Add(content.Content.length, _paragraph);
                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }

        s.Seek2(_end_rec);
    };

    this.ReadParagraph = function(DocumentContent)
    {
        var par = new AscWord.Paragraph(DocumentContent, true);

        var EndPos = 0;

        var s = this.stream;

        var _rec_start = s.cur;
        var _end_rec = _rec_start + s.GetULong() + 4;

        while (s.cur < _end_rec)
        {
            var _at = s.GetUChar();
            switch (_at)
            {
                case 0:
                {
                    par.Set_Pr(this.ReadTextParagraphPr(par));
                    break;
                }
                case 1:
                {
                    var OldImgCount = 0;
                    if(this.IsUseFullUrl)
                    {
                        OldImgCount = this.RebuildImages.length;
                    }
                    var endRunPr =  this.ReadRunProperties();
                    var _value_text_pr = new CTextPr();
                    if(endRunPr.Unifill && !endRunPr.Unifill.fill)
                    {
                        endRunPr.Unifill = undefined;
                    }
                    _value_text_pr.Set_FromObject(endRunPr);
                    par.TextPr.Apply_TextPr(_value_text_pr);//endRunProperties
                    var oTextPrEnd = new CTextPr();
                    oTextPrEnd.Set_FromObject(endRunPr);
                    par.Content[0].Set_Pr(oTextPrEnd);
                    if(this.IsUseFullUrl)
                    {
                        if(this.RebuildImages.length > OldImgCount)
                        {
                            for(var _t = OldImgCount; _t < this.RebuildImages.length; ++_t)
                            {
                                var _text_pr = new CTextPr();
                                _text_pr.Set_FromObject(endRunPr);
                                this.RebuildImages[_t].TextPr = _text_pr;
                                this.RebuildImages[_t].ParaTextPr = par.TextPr;
                                this.RebuildImages[_t].Run = par.Content[0];

                            }
                        }
                    }
                    break;
                }
                case 2:
                {
                    s.Skip2(4);

                    var _c = s.GetULong();
                    for (var i = 0; i < _c; i++)
                    {
                        s.Skip2(5); // type (0) + len
                        var _type = s.GetUChar();

                        switch (_type)
                        {
                            case AscFormat.PARRUN_TYPE_RUN:
                            {
                                var _end = s.cur + s.GetULong() + 4;

                                s.Skip2(1); // start attr

                                var _text = "";
                                while (true)
                                {
                                    var _at = s.GetUChar();
                                    if (_at == g_nodeAttributeEnd)
                                        break;

                                    if (0 == _at)
                                        _text = s.GetString2();
                                }

                                var OldImgCount = 0;
                                if(this.IsUseFullUrl)
                                {
                                    OldImgCount = this.RebuildImages.length;
                                }

                                var _run = null;
                                while (s.cur < _end)
                                {
                                    var _rec = s.GetUChar();

                                    if (0 == _rec)
                                        _run = this.ReadRunProperties();
                                    else
                                        s.SkipRecord();
                                }

                                s.Seek2(_end);


                                var new_run = new ParaRun(par, false), hyperlink = null;
                                if (null != _run)
                                {

                                    var text_pr = new CTextPr();
                                    if(_run.Unifill && !_run.Unifill.fill)
                                    {
                                        _run.Unifill = undefined;
                                    }
                                    if (_run.hlink !== undefined)
                                    {
                                        hyperlink = new ParaHyperlink();
                                        hyperlink.SetValue(_run.hlink.id);
                                        if (_run.hlink.tooltip) {
                                          hyperlink.SetToolTip(_run.hlink.tooltip);
                                        }
                                        // if(!_run.Unifill)
                                        // {
                                        //     _run.Unifill = AscFormat.CreateUniFillSchemeColorWidthTint(11, 0);
                                        // }
                                        _run.Underline = true;
                                    }
                                    text_pr.Set_FromObject(_run);
                                    new_run.Set_Pr(text_pr);
                                    if(this.IsUseFullUrl)
                                    {
                                        if(this.RebuildImages.length > OldImgCount)
                                        {
                                            for(var _t = OldImgCount; _t < this.RebuildImages.length; ++_t)
                                            {
                                                var _text_pr = new CTextPr();
                                                _text_pr.Set_FromObject(text_pr);
                                                this.RebuildImages[_t].TextPr = _text_pr;
                                                this.RebuildImages[_t].Run = new_run;

                                            }
                                        }
                                    }
                                }

                                new_run.AddText(_text);

                                if (hyperlink !== null)
                                {
                                    hyperlink.Add_ToContent(0, new_run, false);
                                    par.Internal_Content_Add(EndPos++, hyperlink);
                                }
                                else
                                {
                                    par.Internal_Content_Add(EndPos++, new_run);
                                }

                                break;
                            }
                            case AscFormat.PARRUN_TYPE_FLD:
                            {
                                var _end = s.cur + s.GetULong() + 4;

                                s.Skip2(1); // start attr

                                while (true)
                                {
                                    var _at = s.GetUChar();
                                    if (_at == g_nodeAttributeEnd)
                                        break;

                                    if (0 == _at)
                                        var f_id = s.GetString2();
                                    else if (1 == _at)
                                        var f_type = s.GetString2();
                                    else
                                        var f_text = s.GetString2();
                                }

                                var _rPr = null, _pPr = null;
                                while (s.cur < _end)
                                {
                                    var _at2 = s.GetUChar();
                                    switch (_at2)
                                    {
                                        case 0:
                                        {
                                            _rPr = this.ReadRunProperties();
                                            break;
                                        }
                                        case 1:
                                        {
                                            _pPr = this.ReadTextParagraphPr();
                                            break;
                                        }
                                        default:
                                        {
                                            s.SkipRecord();
                                            break;
                                        }
                                    }
                                }

                                var Fld = new AscCommonWord.CPresentationField(par);
                                if(f_id)
                                {
                                    Fld.SetGuid(f_id);
                                }
                                if(f_type)
                                {
                                    Fld.SetFieldType(f_type);
                                }
                                if(f_text)
                                {
                                    Fld.CanAddToContent = true;
                                    Fld.AddText(f_text);
                                    Fld.CanAddToContent = false;
                                }
                                if(_rPr)
                                {
                                    Fld.SetPr(_rPr);
                                }
                                if(_pPr)
                                {
                                    Fld.SetPPr(_pPr);
                                }

                                par.Internal_Content_Add(EndPos++, new ParaRun(par, false));
                                par.Internal_Content_Add(EndPos++, Fld);
                                par.Internal_Content_Add(EndPos++, new ParaRun(par, false));
                                s.Seek2(_end);
                                this.fields.push(Fld);
                                break;
                            }
                            case AscFormat.PARRUN_TYPE_BR:
                            {
                                var _end = s.cur + s.GetULong() + 4;

                                var _run = null;
                                while (s.cur < _end)
                                {
                                    var _rec = s.GetUChar();

                                    if (0 == _rec)
                                        _run = this.ReadRunProperties();
                                    else
                                        s.SkipRecord();
                                }

                                s.Seek2(_end);

                                var new_run = new ParaRun(par, false), hyperlink = null;
                                if (null != _run)
                                {
                                    if (_run.hlink !== undefined)
                                    {
                                        hyperlink = new ParaHyperlink();
                                        hyperlink.SetValue(_run.hlink.id);
                                        if (_run.hlink.tooltip) {
                                          hyperlink.SetToolTip(_run.hlink.tooltip);
                                        }
                                    }
                                    var text_pr = new CTextPr();
                                    if(_run.Unifill && !_run.Unifill.fill)
                                    {
                                        _run.Unifill = undefined;
                                    }
                                    text_pr.Set_FromObject(_run);
                                    new_run.Set_Pr(text_pr);
                                }
                                new_run.Add_ToContent( 0, new AscWord.CRunBreak(AscWord.break_Line));
                                if (hyperlink !== null)
                                {
                                    hyperlink.Add_ToContent(0, new_run, false);
                                    par.Internal_Content_Add(EndPos++, hyperlink);
                                }
                                else
                                {
                                    par.Internal_Content_Add(EndPos++, new_run);
                                }
                                break;
                            }
							case AscFormat.PARRUN_TYPE_MATHPARA:
							case AscFormat.PARRUN_TYPE_MATH:
							{
								var _end = s.cur + s.GetULong() + 4;

								var _stream = new AscCommon.FT_Stream2();
								_stream.data = s.data;
								_stream.pos = s.pos;
								_stream.cur = s.cur;
								_stream.size = s.size;
								var parContentOld = par.Content.length;

								var oReadResult = new AscCommonWord.DocReadResult(null);
								var boMathr = new Binary_oMathReader(_stream, oReadResult, null);
								var nDocLength = _stream.GetULongLE();
								if (AscFormat.PARRUN_TYPE_MATHPARA == _type) {
									var props = {};
									boMathr.bcr.Read1(nDocLength, function(t, l){
										return boMathr.ReadMathOMathPara(t,l,par, props);
									});
								} else {
									var oMath = new ParaMath();
                                    par.AddToContentToEnd(oMath);
									boMathr.bcr.Read1(nDocLength, function(t, l){
										return boMathr.ReadMathArg(t,l,oMath.Root,par);
									});
									oMath.Root.Correct_Content(true);
								}
								s.Seek2(_end);

								EndPos += par.Content.length - parContentOld;
								break;
							}
                            default:
                            {
                                s.SkipRecord();
                                break;
                            }
                        }
                    }
                    break;
                }
                default:
                {
                    s.SkipRecord();
                    break;
                }
            }
        }
        s.Seek2(_end_rec);
        par.Correct_Content();
        return par;
    };

    // ------------------------------------------
}








    function CPPTXContentLoader()
    {
        this.Reader = new AscCommon.BinaryPPTYLoader();
        this.Writer = null;

        this.stream = null;
        this.TempMainObject = null;
        this.ParaDrawing = null;
        this.LogicDocument = null;
        this.BaseReader = null;

        this.ImageMapChecker = {};

        this.Start_UseFullUrl = function(insertDocumentUrlsData)
        {
            this.Reader.Start_UseFullUrl(insertDocumentUrlsData);
        };
        this.End_UseFullUrl = function()
        {
            return this.Reader.End_UseFullUrl();
        };
        this.CheckStreamStart = function(oOtherStream) {
            if (!this.Reader) {
                this.Reader = new AscCommon.BinaryPPTYLoader();
            }
            this.Reader.ImageMapChecker = this.ImageMapChecker;
            if (!this.stream ||
                this.stream.data !== oOtherStream.data ||
                this.stream.size !== oOtherStream.size) {
                this.stream = new AscCommon.FileStream();
                this.stream.obj    = oOtherStream.obj;
                this.stream.data   = oOtherStream.data;
                this.stream.size   = oOtherStream.size;
            }
            this.stream.pos    = oOtherStream.pos;
            this.stream.cur    = oOtherStream.cur;
            this.Reader.stream = this.stream;
        };
        this.CheckStreamEnd = function(oOtherStream) {
            oOtherStream.pos = this.stream.pos;
            oOtherStream.cur = this.stream.cur;
        };
        this.ReadPPTXElement = function(reader, stream, fReadFunction) {
            var oOldReader = this.BaseReader;
            var oOldLogicDocument = this.LogicDocument;
            var oOldStream = this.stream;
            if(reader) {
                this.BaseReader = reader;
            }
            this.CheckStreamStart(stream);
            var oStream = this.stream;
            var oResult = fReadFunction();
            this.CheckStreamEnd(stream);


            this.BaseReader = oOldReader;
            this.LogicDocument = oOldLogicDocument;
            this.stream = oOldStream;
            return oResult;
        };
        this.ReadBodyPr = function(reader, stream) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                oThis.stream.GetUChar();
                return oThis.Reader.ReadBodyPr();
            });
        };
        this.ReadFontRef = function(reader, stream) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                oThis.stream.GetUChar();
                return oThis.Reader.ReadFontRef();
            });
        };
        this.ReadStyleRef = function(reader, stream) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                oThis.stream.GetUChar();
                return oThis.Reader.ReadStyleRef();
            });
        };
        this.ReadUniColor = function(reader, stream) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                oThis.stream.GetUChar();
                return oThis.Reader.ReadUniColor();
            });
        };
        this.ReadColorMod = function(reader, stream) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                oThis.stream.GetUChar();
                return oThis.Reader.ReadColorMod();
            });
        };


        this.ReadTheme = function(reader, stream) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                return oThis.Reader.ReadTheme();
            });
        };

        this.ReadDrawing = function(reader, stream, logicDocument, paraDrawing) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function () {
                if (null != paraDrawing)
                {
                    oThis.ParaDrawing = paraDrawing;
                    oThis.TempMainObject = null;
                }
                oThis.LogicDocument = logicDocument;
                oThis.Reader.presentation = logicDocument;
                if(logicDocument)
                {
                    oThis.Reader.DrawingDocument = logicDocument.DrawingDocument;
                }
                else
                {
                    oThis.Reader.DrawingDocument = null;
                }

                var GrObject = null;

                var s = oThis.stream;
                var _main_type = s.GetUChar(); // 0!!!

                var _rec_start = s.cur;
                var _end_rec = _rec_start + s.GetULong() + 4;
                if (s.cur < _end_rec){
                    s.Skip2(5); // 1 + 4 byte - len

                    var _type = s.GetUChar();
                    switch (_type)
                    {
                        case 1:
                        {
                            GrObject = oThis.ReadShape();
                            break;
                        }
                        case 6:
                        case 2:
                        case 7:
                        case 8:
                        {
                            GrObject = oThis.ReadPic(_type);
                            break;
                        }
                        case 3:
                        {
                            GrObject = oThis.ReadCxn();
                            break;
                        }
                        case 4:
                        {
                            GrObject = oThis.ReadGroupShape();
                            break;
                        }
                        case 5:
                        {

                            var oGrFrDrawing = oThis.Reader.ReadGrFrame();
                            if(oGrFrDrawing) {
                                if(oGrFrDrawing.getObjectType() === AscDFH.historyitem_type_GroupShape) {
                                    GrObject = oGrFrDrawing.convertToWord(oThis.LogicDocument);
                                    if(paraDrawing){
                                        GrObject.setParent(paraDrawing);
                                    }
                                }
                                else if(oGrFrDrawing.getObjectType() === AscDFH.historyitem_type_SmartArt) {
                                    GrObject = oGrFrDrawing;
                                    if(paraDrawing){
                                        GrObject.setParent(paraDrawing);
                                    }
                                }
                            }
                            break;
                        }
                        case 9:
                        {
                            GrObject = oThis.Reader.ReadGroupShape(9);
                            if(paraDrawing){
                                GrObject.setParent(paraDrawing);
                            }
                            break;
                        }
                        default:
                        {
                            s.SkipRecord();
                            break;
                        }
                    }
                }

                s.Seek2(_end_rec);
                return GrObject;
            });
        };

        this.ReadGraphicObject = function(stream, presentation, drawingDocument) {
            var oThis = this;
            return this.ReadPPTXElement(undefined, stream, function() {
                if(presentation)
                {
                    oThis.Reader.presentation = presentation;
                }
                if(drawingDocument)
                {
                    oThis.Reader.DrawingDocument = drawingDocument;
                }
                oThis.LogicDocument = null;
                var s = oThis.stream;
                var _main_type = s.GetUChar(); // 0!!!

                var _rec_start = s.cur;
                var _end_rec = _rec_start + s.GetULong() + 4;

                s.Skip2(5); // 1 + 4 byte - len

                var GrObject = oThis.Reader.ReadGraphicObject();

                s.Seek2(_end_rec);
                return GrObject;
            });
        };
        this.ReadGraphicObject2 = function(stream, presentation, drawingDocument) {
            var oThis = this;
            return this.ReadPPTXElement(undefined, stream, function() {
                if(presentation)
                {
                    oThis.Reader.presentation = presentation;
                }
                if(drawingDocument)
                {
                    oThis.Reader.DrawingDocument = drawingDocument;
                }
                oThis.LogicDocument = null;
                var s = oThis.stream;
                var GrObject = oThis.Reader.ReadGraphicObject();
                return GrObject;
            });
        };

        this.ReadTextBody = function(reader, stream, shape, presentation, drawingDocument) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                if(presentation)
                    oThis.Reader.presentation = presentation;
                if(drawingDocument)
                    oThis.Reader.DrawingDocument = drawingDocument;
                oThis.LogicDocument = null;
                var s = oThis.stream;
                var _main_type = s.GetUChar(); // 0!!!
                return oThis.Reader.ReadTextBody(shape);
            });
        };

        this.ReadTextBodyTxPr = function(reader, stream, shape) {
            var oThis = this;
            return this.ReadPPTXElement(reader, stream, function() {
                oThis.LogicDocument = null;
                var s = oThis.stream;
                var _main_type = s.GetUChar(); // 0!!!
                return oThis.Reader.ReadTextBodyTxPr(shape);
            });

        };

        this.ReadShapeProperty = function(stream, type) {
            var oThis = this;
            return this.ReadPPTXElement(undefined, stream, function() {
                oThis.LogicDocument = null;
                var s = oThis.stream;
                var _main_type = s.GetUChar(); // 0!!!

                var oNewSpPr;
                if(0 == type){
                    oNewSpPr = oThis.Reader.ReadLn()
                }
                else if(1 == type){
                    oNewSpPr = oThis.Reader.ReadUniFill();
                }
                else{
                    oNewSpPr = new AscFormat.CSpPr();
                    oThis.Reader.ReadSpPr(oNewSpPr);
                }
                return oNewSpPr;
            });
        };

		this.ReadRunProperties = function(stream, type) {
            var oThis = this;
            return this.ReadPPTXElement(undefined, stream, function() {
                oThis.LogicDocument = null;
                var s = oThis.stream;
                var _main_type = s.GetUChar(); // 0!!!
                return oThis.Reader.ReadRunProperties();
            });
		};

        this.ReadShape = function()
        {
            var s = this.stream;

            var shape = new AscFormat.CShape();
            shape.setWordShape(true);
            shape.setBDeleted(false);
            shape.setParent(this.TempMainObject == null ? this.ParaDrawing : null);
            var _rec_start = s.cur;
            var _end_rec = _rec_start + s.GetULong() + 4;

            s.Skip2(1); // start attributes

            while (true)
            {
                var _at = s.GetUChar();
                if (_at == AscCommon.g_nodeAttributeEnd)
                    break;

                switch (_at)
                {
                    case 0:
                    {
                        shape.setUseBgFill(s.GetBool());
                        break;
                    }
                    default:
                        break;
                }
            }

            var oXFRM = null;
            while (s.cur < _end_rec)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 0:
                    {
                        var pr = this.Reader.ReadNvUniProp(shape);
                        shape.setNvSpPr(pr);
                        if(AscFormat.isRealNumber(pr.locks))
                        {
                            shape.setLocks(pr.locks);
                        }
                        break;
                    }
                    case 1:
                    {
                        var spPr = new AscFormat.CSpPr();
                        this.ReadSpPr(spPr);
                        shape.setSpPr(spPr);
                        shape.spPr.setParent(shape);
                        break;
                    }
                    case 2:
                    {
                        shape.setStyle(this.Reader.ReadShapeStyle());
                        break;
                    }
                    case 3:
                    {
                        s.SkipRecord();
                        break;
                    }
                    case 4:
                    {
                        var oThis = this.BaseReader;

                        shape.setTextBoxContent(new CDocumentContent(shape, this.LogicDocument.DrawingDocument, 0, 0, 0, 0, false, false));

                        var _old_cont = shape.textBoxContent.Content[0];

                        shape.textBoxContent.Internal_Content_RemoveAll();

                        s.Skip2(4); // rec len

                        oThis.stream.pos = s.pos;
                        oThis.stream.cur = s.cur;

                        var oBinary_DocumentTableReader = new Binary_DocumentTableReader(shape.textBoxContent, oThis.oReadResult, oThis.openParams, oThis.stream, false, oThis.oComments);
                        var nDocLength = oThis.stream.GetULongLE();
                        var content_arr = [];
                        var oCurParaDrawing = this.ParaDrawing;
                        oThis.bcr.Read1(nDocLength, function(t,l){
                            return oBinary_DocumentTableReader.ReadDocumentContent(t,l, content_arr);
                        });
                        this.ParaDrawing = oCurParaDrawing;
                        for(var i = 0, length = content_arr.length; i < length; ++i){
                            if(i == length - 1)
                                shape.textBoxContent.Internal_Content_Add(i, content_arr[i], true);
                            else
                                shape.textBoxContent.Internal_Content_Add(i, content_arr[i], false);
                        }

                        s.pos = oThis.stream.pos;
                        s.cur = oThis.stream.cur;

                        if (shape.textBoxContent.Content.length == 0)
                            shape.textBoxContent.Internal_Content_Add(0, _old_cont);

                        break;
                    }
                    case 5:
                    {
                        var bodyPr = new AscFormat.CBodyPr();
                        this.Reader.CorrectBodyPr(bodyPr);
                        shape.setBodyPr(bodyPr);
                        break;
                    }
                    case 6:
                    {
                        oXFRM = this.Reader.ReadXfrm();
                        break;
                    }
                    case 7:
                    {
                        shape.setSignature(this.Reader.ReadSignatureLine());
                        break;
                    }
                    case 8:
                    {
                        if(typeof AscFormat.SmartArt !== undefined)
                        {
                            s.GetULong();//length
                            s.GetUChar();//attr type - 0
                            shape.setModelId(s.GetString2());
                        }
                        else
                        {
                            s.SkipRecord();
                        }
                        break;
                    }
                    case 0xA1:
                    {
                        shape.readMacro(s);
                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }

            if(oXFRM)
            {
                var oRet = new AscFormat.CGroupShape();
                shape.setParent(null);
                oRet.setParent(this.TempMainObject == null ? this.ParaDrawing : null);
                oRet.setBDeleted(false);
                var oSpPr = new AscFormat.CSpPr();
                var oXfrm = new AscFormat.CXfrm();
                oXfrm.setOffX(shape.spPr.xfrm.offX);
                oXfrm.setOffY(shape.spPr.xfrm.offY);
                oXfrm.setExtX(shape.spPr.xfrm.extX);
                oXfrm.setExtY(shape.spPr.xfrm.extY);
                oXfrm.setChExtX(shape.spPr.xfrm.extX);
                oXfrm.setChExtY(shape.spPr.xfrm.extY);
                oXfrm.setChOffX(0);
                oXfrm.setChOffY(0);
                oSpPr.setXfrm(oXfrm);
                oXfrm.setParent(oSpPr);
                shape.spPr.xfrm.setOffX(0);
                shape.spPr.xfrm.setOffY(0);
                oRet.setSpPr(oSpPr);
                oSpPr.setParent(oRet);
                oRet.addToSpTree(0, shape);
                var oShape2 = new AscFormat.CShape();
                var oSpPr2 = new AscFormat.CSpPr();
                oShape2.setSpPr(oSpPr2);
                oSpPr2.setParent(oShape2);
                var oXfrm2 = oXFRM;
                oXfrm2.setParent(oSpPr2);
                oSpPr2.setXfrm(oXfrm2);
                oXfrm2.setOffX(oXfrm2.offX - oXfrm.offX);
                oXfrm2.setOffY(oXfrm2.offY - oXfrm.offY);
                oSpPr2.setFill(AscFormat.CreateNoFillUniFill());
                oSpPr2.setLn(AscFormat.CreateNoFillLine());
                oShape2.setTxBody(shape.txBody);
                shape.setTxBody(null);
                shape.setGroup(oRet);
                oShape2.setBDeleted(false);
                oShape2.setWordShape(true);
                if(shape.spPr.xfrm && AscFormat.isRealNumber(shape.spPr.xfrm.rot))
                {
                    oXfrm2.setRot((AscFormat.isRealNumber(oXfrm2.rot) ? oXfrm2.rot : 0) + shape.spPr.xfrm.rot);
                }
                if(oShape2.txBody)
                {
                    oShape2.txBody.setParent(oShape2);
                }
                if(shape.textBoxContent)
                {
                    oShape2.setTextBoxContent(shape.textBoxContent.Copy(oShape2, shape.textBoxContent.DrawingDocument));
                    shape.setTextBoxContent(null);
                }
                if(shape.bodyPr)
                {
                    oShape2.setBodyPr(shape.bodyPr);
                    shape.setBodyPr(null);
                }
                oRet.addToSpTree(1, oShape2);
                oShape2.setGroup(oRet);
                s.Seek2(_end_rec);
                return oRet;
            }
            s.Seek2(_end_rec);
            return shape;

        };
        this.ReadCxn = function()
        {
            var s = this.stream;

            var shape = new AscFormat.CConnectionShape( );
            shape.setWordShape(true);
            shape.setParent(this.TempMainObject == null ? this.ParaDrawing : null);
            var _rec_start = s.cur;
            var _end_rec = _rec_start + s.GetULong() + 4;

            while (s.cur < _end_rec)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 0:
                    {
                        s.SkipRecord();
                        break;
                    }
                    case 1:
                    {
                        var spPr = new AscFormat.CSpPr();
                        this.ReadSpPr(spPr);
                        shape.setSpPr(spPr);
                        break;
                    }
                    case 2:
                    {
                        shape.setStyle(this.Reader.ReadShapeStyle());
                        break;
                    }
                    case 0xA1:
                    {
                        shape.readMacro(s);
                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }

            s.Seek2(_end_rec);
            return shape;
        };
        this.ReadPic = function(type)
        {
            var s = this.stream;

            var isOle = (type === 6);
            var pic;
            if (isOle)
            {
                pic = new AscFormat.COleObject(this.TempMainObject)
            }
            else
            {
                if (Asc.editor.isPdfEditor())
                    pic = new AscPDF.CPdfImage();
                else
                    pic = new AscFormat.CImageShape(this.TempMainObject);
            }

            pic.setBDeleted(false);
            pic.setParent(this.TempMainObject == null ? this.ParaDrawing : null);

            var _rec_start = s.cur;
            var _end_rec = _rec_start + s.GetULong() + 4;

            var sMaskFileName = "";
            while (s.cur < _end_rec)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 0:
                    {
                        var pr = this.Reader.ReadNvUniProp(pic);
                        pic.setNvSpPr(pr);
                        if(AscFormat.isRealNumber(pr.locks))
                        {
                            pic.setLocks(pr.locks);
                        }
                        break;
                    }
                    case 1:
                    {
                        var unifill = this.Reader.ReadUniFill(null, pic, null);
                        pic.setBlipFill(unifill.fill);//this.Reader.ReadUniFill();

                        //pic.spPr.Fill = new AscFormat.CUniFill();
                        //pic.spPr.Fill.fill = pic.blipFill;
                        //pic.brush = pic.spPr.Fill;

                        break;
                    }
                    case 2:
                    {
                        var spPr = new AscFormat.CSpPr();
                        this.ReadSpPr(spPr);
                        pic.setSpPr(spPr);
                        pic.spPr.setParent(pic);
                        break;
                    }
                    case 3:
                    {
                        pic.setStyle(this.Reader.ReadShapeStyle());
                        break;
                    }
                    case 4:
                    {
                        if(isOle) {
                            this.ReadOleInfo(pic);
                        } else {
                            s.SkipRecord();
                        }
                        break;
                    }
                    case 5:
                    {
                        if(type === 7 || type === 8){//video or audio
                            s.GetLong();
                            s.GetUChar();//start attributes
                            while(true){
                                var _at2 = s.GetUChar();
                                if (_at2 == g_nodeAttributeEnd)
                                    break;
                                switch (_at2) {
                                    case 0:
                                    {
                                        sMaskFileName = s.GetString2();
                                        break;
                                    }
                                }
                            }
                        }
                        else{
                            s.SkipRecord();
                        }
                        break;
                    }
                    case 0xA1:
                    {
                        pic.readMacro(s);
                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }

            if(type === 7 || type === 8){//video or audio
                if(typeof sMaskFileName === "string" && sMaskFileName.length > 0 &&
                    pic.nvPicPr && pic.nvPicPr.nvPr /*&& pic.nvPicPr.nvPr.unimedia*/){
                    var oUniMedia = new AscFormat.UniMedia();
                    oUniMedia.type = type;
                    oUniMedia.media = sMaskFileName;
                    pic.nvPicPr.nvPr.setUniMedia(oUniMedia);
                }
            }

            s.Seek2(_end_rec);
            return pic;
        };
        this.ReadOleInfo = function(ole)
        {
            var s = this.stream;

            var _rec_start = s.cur;
            var _end_rec = _rec_start + s.GetLong() + 4;

            s.Skip2(1); // start attributes
            var dxaOrig = 0;
            var dyaOrig = 0;
            while (true)
            {
                var _at = s.GetUChar();
                if (_at == g_nodeAttributeEnd)
                    break;

                switch (_at)
                {
                    case 0:
                    {
                        ole.setApplicationId(s.GetString2());
                        break;
                    }
                    case 1:
                    {
                        ole.setData(s.GetString2());
                        break;
                    }
                    case 2:
                    {
                        dxaOrig = s.GetULong();
                        break;
                    }
                    case 3:
                    {
                        dyaOrig = s.GetULong();
                        break;
                    }
                    case 4:
                    {
                        const drawAspect = s.GetUChar();
                        ole.setDrawAspect(drawAspect);
                        break;
                    }
                    case 5:
                    {
                        s.GetUChar();
                        break;
                    }
                    case 6:
                    {
                        s.GetUChar();
                        break;
                    }
                    case 7:
                    {
                        ole.setObjectFile(s.GetString2());
                        break;
                    }
                    default:
                        break;
                }
            }




            var oleType = null;
            while (s.cur < _end_rec)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 1:
                    {
                        s.GetLong();//length
                        oleType = s.GetUChar();
                        ole.setOleType(oleType);
                        break;
                    }
                    case 2:
                    {

                        var binary_length;
                        switch(oleType)
                        {
                            case 0:
                            {
                                binary_length = s.GetULong();
                                ole.setBinaryData(s.data.slice(s.cur, s.cur + binary_length));
                                s.Seek2(s.cur + binary_length);
                                break;
                            }
                            case 1:
                            {
                                ole.setObjectFile("maskFile.docx");
                                binary_length = s.GetULong();
                                ole.setBinaryData(s.data.slice(s.cur, s.cur + binary_length));
                                s.Seek2(s.cur + binary_length);
                                break;
                            }
                            case 2:
                            {
                                ole.setObjectFile("maskFile.xlsx");
                                binary_length = s.GetULong();
                                ole.setBinaryData(s.data.slice(s.cur, s.cur + binary_length));
                                s.Seek2(s.cur + binary_length);
                                break;
                            }
                            case 4:
                            {
                                s.GetLong();//length

                                var type2 = s.GetUChar();
                                if (AscCommonWord.c_oSer_OMathContentType.OMath === type2 && ole.parent && ole.parent.Parent)
                                {
                                    var oReadResult = this.BaseReader ? this.BaseReader.oReadResult : new AscCommonWord.DocReadResult(null);
                                    ole.parent.ParaMath = this.Reader.ReadMathObject(s, oReadResult, ole.parent.Parent);
                                }
                                else
                                {
                                    s.SkipRecord();
                                }
                                break;
                            }
                            default:
                            {
                                s.SkipRecord();
                                break;
                            }
                        }
                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }




			if (dxaOrig > 0 && dyaOrig > 0) {
				var ratio = 4 / 3 / 20;//twips to px
				ole.setPixSizes(ratio * dxaOrig, ratio * dyaOrig);
			}
            s.Seek2(_end_rec);
        };
        this.ReadGroupShape = function()
        {
            var s = this.stream;

            var shape = new AscFormat.CGroupShape();

            shape.setBDeleted(false);
            shape.setParent(this.TempMainObject == null ? this.ParaDrawing : null);
            this.TempGroupObject = shape;

            var oldParaDrawing = this.ParaDrawing;
            this.ParaDrawing = null;

            var _rec_start = s.cur;
            var _end_rec = _rec_start + s.GetULong() + 4;

            while (s.cur < _end_rec)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 0:
                    {
                        s.SkipRecord();
                        break;
                    }
                    case 1:
                    {
                        var spPr = new AscFormat.CSpPr();
                        this.Reader.ReadGrSpPr(spPr);
                        shape.setSpPr(spPr);
                        shape.spPr.setParent(shape);
                        break;
                    }
                    case 2:
                    {
                        s.Skip2(4); // len
                        var _c = s.GetULong();
                        for (var i = 0; i < _c; i++)
                        {
                            s.Skip2(1);
                            var __len = s.GetULong();
                            if (__len == 0)
                                continue;

                            var _type = s.GetUChar();

                            var sp;
                            switch (_type)
                            {
                                case 1:
                                {
                                    sp = this.ReadShape();
                                    if(!IsHiddenObj(sp) && sp.spPr && sp.spPr.xfrm){
                                        sp.setGroup(shape);
                                        shape.addToSpTree(shape.spTree.length, sp);
                                    }
                                    break;
                                }
                                case 6:
                                case 2:
                                case 7:
                                case 8:
                                {
                                    sp = this.ReadPic(_type);
                                    if(!IsHiddenObj(sp) && sp.spPr && sp.spPr.xfrm){
                                        sp.setGroup(shape);
                                        shape.addToSpTree(shape.spTree.length, sp);
                                    }
                                    break;
                                }
                                case 3:
                                {
                                    sp = this.ReadCxn();
                                    if(!IsHiddenObj(sp) && sp.spPr && sp.spPr.xfrm) {
                                        sp.setGroup(shape);
                                        shape.addToSpTree(shape.spTree.length, sp);
                                    }
                                    break;
                                }
                                case 4:
                                {
                                    sp = this.ReadGroupShape();
                                    if(sp && !IsHiddenObj(sp) && sp.spPr && sp.spPr.xfrm && sp.spTree.length > 0 ) {
                                        sp.setGroup(shape);
                                        shape.addToSpTree(shape.spTree.length, sp);
                                    }
                                    break;
                                }
                                case 5:
                                {
                                    var _chart = null;
                                    if ("undefined" != typeof(AscFormat.CGraphicFrame))
                                        _chart = this.Reader.ReadGrFrame();
                                    else
                                        _chart = this.Reader.ReadChartDataInGroup(shape);
                                    if (null != _chart)
                                    {
                                        _chart.setGroup(shape);
                                        shape.addToSpTree(shape.spTree.length, _chart);
                                    }
                                    break;
                                }
                                default:
                                {
                                    s.SkipRecord();
                                    break;
                                }
                            }
                        }
                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }

            if(oldParaDrawing && shape.spPr && !shape.spPr.xfrm){
                shape.bEmptyTransform = true;
            }
            if(!oldParaDrawing){
                this.Reader.CheckGroupXfrm(shape);
            }
            this.ParaDrawing = oldParaDrawing;
            s.Seek2(_end_rec);
            this.TempGroupObject = null;
            // if(shape.spTree.length === 0){
            //     return null;
            // }
            return shape;
        };

        this.ReadSpPr = function(spPr)
        {
            var s = this.stream;

            var _rec_start = s.cur;
            var _end_rec = _rec_start + s.GetULong() + 4;

            s.Skip2(1); // start attributes

            while (true)
            {
                var _at = s.GetUChar();
                if (_at == AscCommon.g_nodeAttributeEnd)
                    break;

                if (0 == _at)
                    spPr.bwMode = s.GetUChar();
                else
                    break;
            }

            while (s.cur < _end_rec)
            {
                var _at = s.GetUChar();
                switch (_at)
                {
                    case 0:
                    {
                        spPr.setXfrm(this.Reader.ReadXfrm());
                        spPr.xfrm.setParent(spPr);
                        //this.CorrectXfrm(spPr.xfrm);
                        break;
                    }
                    case 1:
                    {
                        spPr.setGeometry(this.Reader.ReadGeometry(spPr.xfrm));
                        break;
                    }
                    case 2:
                    {
                        spPr.setFill(this.Reader.ReadUniFill(spPr, null, null));
                        break;
                    }
                    case 3:
                    {
                        spPr.setLn(this.Reader.ReadLn());
                        break;
                    }
                    case 4:
                    {
                        spPr.setEffectPr(this.Reader.ReadEffectProperties());
                        break;
                    }
                    case 5:
                    {
                        var _len = s.GetULong();
                        s.Skip2(_len);
                        break;
                    }
                    case 6:
                    {
                        var _len = s.GetULong();
                        s.Skip2(_len);
                        break;
                    }
                    default:
                    {
                        s.SkipRecord();
                        break;
                    }
                }
            }

            s.Seek2(_end_rec);
        };

        this.CorrectXfrm = function(_xfrm)
        {
            if (!_xfrm)
                return;

            if (null == _xfrm.rot)
                return;

            var nInvertRotate = 0;
            if (true === _xfrm.flipH)
                nInvertRotate += 1;
            if (true === _xfrm.flipV)
                nInvertRotate += 1;

            var _rot = _xfrm.rot;
            var _del = 2 * Math.PI;

            if (nInvertRotate)
                _rot = -_rot;

            if (_rot >= _del)
            {
                var _intD = (_rot / _del) >> 0;
                _rot = _rot - _intD * _del;
            }
            else if (_rot < 0)
            {
                var _intD = (-_rot / _del) >> 0;
                _intD = 1 + _intD;
                _rot = _rot + _intD * _del;
            }

            _xfrm.rot = _rot;
        };


        this.CheckImagesNeeds = function(logicDoc)
        {
            var index = 0;
            logicDoc.ImageMap = {};
            for (var i in this.ImageMapChecker)
            {
                logicDoc.ImageMap[index++] = i;
            }
        };

        this.Clear = function(bClearStreamOnly)
        {
            //вызывается пока только перед вставкой
            this.Reader.stream = null;
            this.stream = null;
            this.BaseReader = null;
            if(!bClearStreamOnly)
                this.ImageMapChecker = {};
	        this.Reader.ClearSmartArts();
        };
    }

    //----------------------------------------------------------export----------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].c_dScalePPTXSizes = c_dScalePPTXSizes;
    window['AscCommon'].CBuilderImages = CBuilderImages;
    window['AscCommon'].BinaryPPTYLoader = BinaryPPTYLoader;
    window['AscCommon'].IsHiddenObj = IsHiddenObj;
    window['AscCommon'].pptx_content_loader = new CPPTXContentLoader();




})(window);
