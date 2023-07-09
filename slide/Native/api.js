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

var _api = null;

var sdkCheck = true;
var spellCheck = true;
var _internalStorage = {};

function asc_menu_WriteSlidePr(_slidePr, _stream){
    if(_slidePr.Background) {
        _slidePr.Background.write(0, _stream);
    }
    asc_menu_WriteTransition(1, _slidePr.Transition, _stream);
    if(AscFormat.isRealNumber(_slidePr.LayoutIndex)){
        _stream["WriteByte"](2);
        _stream["WriteLong"](_slidePr.LayoutIndex);
    }
    if(AscFormat.isRealBool(_slidePr.isHidden)){
        _stream["WriteByte"](3);
        _stream["WriteBool"](_slidePr.isHidden);
    }
    if(AscFormat.isRealBool(_slidePr.lockBackground)){
        _stream["WriteByte"](4);
        _stream["WriteBool"](_slidePr.lockBackground);
    }
    if(AscFormat.isRealBool(_slidePr.lockDelete)){
        _stream["WriteByte"](5);
        _stream["WriteBool"](_slidePr.lockDelete);
    }
    if(AscFormat.isRealBool(_slidePr.lockLayout)){
        _stream["WriteByte"](6);
        _stream["WriteBool"](_slidePr.lockLayout);
    }
    if(AscFormat.isRealBool(_slidePr.lockRemove)){
        _stream["WriteByte"](7);
        _stream["WriteBool"](_slidePr.lockRemove);
    }
    if(AscFormat.isRealBool(_slidePr.lockTiming)){
        _stream["WriteByte"](8);
        _stream["WriteBool"](_slidePr.lockTiming);
    }
    if(AscFormat.isRealBool(_slidePr.lockTransition)){
        _stream["WriteByte"](9);
        _stream["WriteBool"](_slidePr.lockTransition);
    }
    _stream["WriteByte"](255);
}

function asc_menu_ReadSlidePr(_params, _cursor){
    var _settings = new Asc.CAscSlideProps();

    var _continue = true;
    while (_continue)
    {
        var _attr = _params[_cursor.pos++];
        switch (_attr)
        {
            case 0:
            {
                _settings.Background = new Asc.asc_CShapeFill();
                _settings.Background.read(_params, _cursor);
                break;
            }
            case 1:
            {
                _settings.Transition = asc_menu_ReadTransition(_params, _cursor);
                break;
            }
            case 2:
            {
                _settings.LayoutIndex = _params[_cursor.pos++];
                break;
            }
            case 3:
            {
                _settings.isHidden = _params[_cursor.pos++];
                break;
            }
            case 4:
            {
                _settings.lockBackground = _params[_cursor.pos++];
                break;
            }
            case 5:
            {
                _settings.lockDelete = _params[_cursor.pos++];
                break;
            }
            case 6:
            {
                _settings.lockLayout = _params[_cursor.pos++];
                break;
            }
            case 7:
            {
                _settings.lockRemove = _params[_cursor.pos++];
                break;
            }
            case 8:
            {
                _settings.lockTiming = _params[_cursor.pos++];
                break;
            }
            case 9:
            {
                _settings.lockTransition = _params[_cursor.pos++];
                break;
            }
            case 255:
            default:
            {
                _continue = false;
                break;
            }
        }
    }
    return _settings;
}

function asc_menu_WriteTransition(type, _transition, _stream){

    _stream["WriteByte"](type);
    if(AscFormat.isRealNumber(_transition.TransitionType)){
        _stream["WriteByte"](0);
        _stream["WriteLong"](_transition.TransitionType);
    }
    if(AscFormat.isRealNumber(_transition.TransitionOption)){
        _stream["WriteByte"](1);
        _stream["WriteLong"](_transition.TransitionOption);
    }
    if(AscFormat.isRealNumber(_transition.TransitionDuration)){
        _stream["WriteByte"](2);
        _stream["WriteLong"](_transition.TransitionDuration);
    }
    if(AscFormat.isRealBool(_transition.SlideAdvanceOnMouseClick)){
        _stream["WriteByte"](3);
        _stream["WriteBool"](_transition.SlideAdvanceOnMouseClick);
    }
    if(AscFormat.isRealBool(_transition.SlideAdvanceAfter)){
        _stream["WriteByte"](4);
        _stream["WriteBool"](_transition.SlideAdvanceAfter);
    }
    if(AscFormat.isRealBool(_transition.ShowLoop)){
        _stream["WriteByte"](5);
        _stream["WriteBool"](_transition.ShowLoop);
    }
    if(AscFormat.isRealNumber(_transition.SlideAdvanceDuration)){
        _stream["WriteByte"](6);
        _stream["WriteLong"](_transition.SlideAdvanceDuration);
    }

    _stream["WriteByte"](255);
}

function asc_menu_ReadTransition(_params, _cursor)
{
    var _settings = new Asc.CAscSlideTransition();

    var _continue = true;
    while (_continue)
    {
        var _attr = _params[_cursor.pos++];
        switch (_attr)
        {
            case 0:
            {
                _settings.TransitionType = _params[_cursor.pos++];
                break;
            }
            case 1:
            {
                _settings.TransitionOption = _params[_cursor.pos++];
                break;
            }
            case 2:
            {
                _settings.TransitionDuration = _params[_cursor.pos++];
                break;
            }
            case 3:
            {
                _settings.SlideAdvanceOnMouseClick = _params[_cursor.pos++];
                break;
            }
            case 4:
            {
                _settings.SlideAdvanceAfter = _params[_cursor.pos++];
                break;
            }
            case 5:
            {
                _settings.ShowLoop = _params[_cursor.pos++];
                break;
            }
            case 6:
            {
                _settings.SlideAdvanceDuration = _params[_cursor.pos++];
            }
            case 255:
            default:
            {
                _continue = false;
                break;
            }
        }
    }

    return _settings;
}

function initSpellCheckApi() {
    
    _api.SpellCheckApi = new AscCommon.CSpellCheckApi();
    _api.isSpellCheckEnable = true;

    _api.SpellCheckApi.spellCheck = function (spellData) {
        window["native"]["SpellCheck"](JSON.stringify(spellData));
    };
    
    _api.SpellCheckApi.disconnect = function () {};

    _api.sendEvent('asc_onSpellCheckInit', [
        "1026",
        "1027",
        "1029",
        "1030",
        "1031",
        "1032",
        "1033",
        "1036",
        "1038",
        "1040",
        "1042",
        "1043",
        "1044",
        "1045",
        "1046",
        "1048",
        "1049",
        "1050",
        "1051",
        "1053",
        "1055",
        "1057",
        "1058",
        "1060",
        "1062",
        "1063",
        "1066",
        "1068",
        "1069",
        "1087",
        "1104",
        "1110",
        "1134",
        "2051",
        "2055",
        "2057",
        "2068",
        "2070",
        "3079",
        "3081",
        "3082",
        "4105",
        "7177",
        "9242",
        "10266"
    ]);

    _api.SpellCheckApi.onInit = function (e) {
        _api.sendEvent('asc_onSpellCheckInit', e);
    };

    _api.SpellCheckApi.onSpellCheck = function (e) {
        _api.SpellCheck_CallBack(e);
    };

    _api.SpellCheckApi.init(_api.documentId);

    _api.asc_setSpellCheck(spellCheck);

}

function NativeOpenFileP(_params, documentInfo){
    window.NATIVE_DOCUMENT_TYPE = window["native"]["GetEditorType"]();
    if ("presentation" !== window.NATIVE_DOCUMENT_TYPE){
        return;
    }

    sdkCheck = documentInfo["sdkCheck"];
    spellCheck = documentInfo["spellCheck"];

    var translations = documentInfo["translations"];
    if (undefined != translations && null != translations && translations.length > 0) {
        translations = JSON.parse(translations)
    } else {
        translations = "";
    }

    window["_api"] = window["API"] = _api = new window["Asc"]["asc_docs_api"](translations);
    window["_editor"] = window.editor;

    AscCommon.g_clipboardBase.Init(_api);
    _api["Native_Editor_Initialize_Settings"](_params);
    window.documentInfo = documentInfo;
    var userInfo = new Asc.asc_CUserInfo();
    userInfo.asc_putId(window.documentInfo["docUserId"]);
    userInfo.asc_putFullName(window.documentInfo["docUserName"]);
    userInfo.asc_putFirstName(window.documentInfo["docUserFirstName"]);
    userInfo.asc_putLastName(window.documentInfo["docUserLastName"]);

    var docInfo = new Asc.asc_CDocInfo();
    docInfo.put_Id(window.documentInfo["docKey"]);
    docInfo.put_Url(window.documentInfo["docURL"]);
    docInfo.put_Format("pptx");
    docInfo.put_UserInfo(userInfo);
    docInfo.put_Token(window.documentInfo["token"]);

    _internalStorage.externalUserInfo = userInfo;
    _internalStorage.externalDocInfo = docInfo;

    var permissions = window.documentInfo["permissions"];
    if (undefined != permissions && null != permissions && permissions.length > 0) {
        docInfo.put_Permissions(JSON.parse(permissions));
    }
    _api.asc_setDocInfo(docInfo);
    
    _api.asc_registerCallback("asc_onAdvancedOptions", function(type, options) {
        var stream = global_memory_stream_menu;
        if (options === undefined) {
            options = {};
        }
        options["optionId"] = type;
        stream["ClearNoAttack"]();
        stream["WriteString2"](JSON.stringify(options));
        window["native"]["OnCallMenuEvent"](22000, stream); // ASC_MENU_EVENT_TYPE_ADVANCED_OPTIONS
    });
    
    _api.asc_registerCallback("asc_onSendThemeColorSchemes", function(schemes) {
        var stream = global_memory_stream_menu;
        stream["ClearNoAttack"]();
        asc_WriteColorSchemes(schemes, stream);
        window["native"]["OnCallMenuEvent"](2404, stream); // ASC_SPREADSHEETS_EVENT_TYPE_COLOR_SCHEMES
    });

    _api.asc_registerCallback("asc_onSendThemeColors", onApiSendThemeColors);

    _api.asc_registerCallback("asc_onUpdateThemeIndex", function(nIndex) {
        var stream = global_memory_stream_menu;
        stream["ClearNoAttack"]();
        stream["WriteLong"](nIndex);
        window["native"]["OnCallMenuEvent"](8093, stream); // ASC_PRESENTATIONS_EVENT_TYPE_THEME_INDEX
    });
    _api.asc_registerCallback('asc_onError', onApiError);

    // Edit

    _api.asc_registerCallback('asc_canIncreaseIndent', onApiCanIncreaseIndent);
    _api.asc_registerCallback('asc_canDecreaseIndent', onApiCanDecreaseIndent);

    // Comments

    _api.asc_registerCallback("asc_onAddComment", onApiAddComment);
    _api.asc_registerCallback("asc_onAddComments", onApiAddComments);
    _api.asc_registerCallback("asc_onRemoveComment", onApiRemoveComment);
    _api.asc_registerCallback("asc_onChangeComments", onApiChangeComments);
    _api.asc_registerCallback("asc_onRemoveComments", onApiRemoveComments);
    _api.asc_registerCallback("asc_onChangeCommentData", onApiChangeCommentData);
    _api.asc_registerCallback("asc_onLockComment", onApiLockComment);
    _api.asc_registerCallback("asc_onUnLockComment", onApiUnLockComment);
    _api.asc_registerCallback("asc_onShowComment", onApiShowComment);
    _api.asc_registerCallback("asc_onHideComment", onApiHideComment);
    _api.asc_registerCallback("asc_onUpdateCommentPosition", onApiUpdateCommentPosition);

    if (window.documentInfo["iscoauthoring"]) {
        _api.isSpellCheckEnable = false;
        _api.asc_setAutoSaveGap(1);
        _api._coAuthoringInit();
        _api.asc_SetFastCollaborative(true);
        _api.SetCollaborativeMarksShowType(Asc.c_oAscCollaborativeMarksShowType.None);
        window["native"]["onTokenJWT"](_api.CoAuthoringApi.get_jwt());

        _api.asc_registerCallback("asc_onAuthParticipantsChanged", onApiAuthParticipantsChanged);
        _api.asc_registerCallback("asc_onParticipantsChanged", onApiParticipantsChanged);

        _api.asc_registerCallback("asc_onGetEditorPermissions", function(state) {

            var rData = {
                "c"             : "open",
                "id"            : window.documentInfo["docKey"],
                "userid"        : window.documentInfo["docUserId"],
                "format"        : "pptx",
                "vkey"          : undefined,
                "url"           : window.documentInfo["docURL"],
                "title"         : this.documentTitle,
                "nobase64"      : true};

            _api.CoAuthoringApi.auth(window.documentInfo["viewmode"], rData);
        });

        _api.asc_registerCallback("asc_onDocumentUpdateVersion", function(callback) {
            var me = this;
            me.needToUpdateVersion = true;
            if (callback) callback.call(me);
        });
    } else {
        var doc_bin = window["native"]["GetFileString"]("native_open_file");
        _api["asc_nativeOpenFile"](doc_bin);
        _api.documentId = "1";
        _api.WordControl.m_oDrawingDocument.AfterLoad();
        window["_api"] = window["API"] = Api = _api;
        window["_editor"] = window.editor;
        if (window.documentInfo["viewmode"]) {
            _api.ShowParaMarks = false;
            AscCommon.CollaborativeEditing.Set_GlobalLock(true);
            _api.isViewMode = true;
            _api.WordControl.m_oDrawingDocument.IsViewMode = true;
          }
        var _presentation = _api.WordControl.m_oLogicDocument;

        var nSlidesCount = _presentation.Slides.length;
        var dPresentationWidth = _presentation.GetWidthMM();
        var dPresentationHeight = _presentation.GetHeightMM();

        var aTransitions = [];
        var slides = _presentation.Slides;
        // for(var i = 0; i < slides.length; ++i){
        //     aTransitions.push(slides[i].transition.ToArray());
        // }

        _api.asc_GetDefaultTableStyles();
	    _presentation.Recalculate({Drawings:{All:true, Map:{}}});
	    _presentation.CurPage = Math.min(0, _presentation.Slides.length - 1);
        _presentation.Document_UpdateInterfaceState();
        _presentation.DrawingDocument.CheckThemes();
        _presentation.DrawingDocument.CheckGuiControlColors();
        _api.WordControl.CheckLayouts();
        initSpellCheckApi();

        if (!_api.bNoSendComments) {
            var _slides = _presentation.Slides;
            var _slidesCount = _slides.length;
            for (var i = 0; i < _slidesCount; i++) {
                var slideComments = _slides[i].slideComments;
                if (slideComments) {
                    var _comments = slideComments.comments;
                    var _commentsCount = _comments.length;
                    for (var j = 0; j < _commentsCount; j++) {
                        _api.sync_AddComment(_comments[j].Get_Id(), _comments[j].Data);
                    }
                }
            }
        }

        return [nSlidesCount, dPresentationWidth, dPresentationHeight, aTransitions];
    }
}

function onApiCanIncreaseIndent(value) {
    var data = { "result": value };
    postDataAsJSONString(data, 8127); // ASC_PRESENTATIONS_EVENT_CANINCREASEINDENT
}

function onApiCanDecreaseIndent(value) {
    var data = { "result": value };
    postDataAsJSONString(data, 8128); // ASC_PRESENTATIONS_EVENT_CANDECREASEINDENT
}

Asc['asc_docs_api'].prototype.UpdateTextPr = function(TextPr)
{
    if (!TextPr)
        return;

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    if (TextPr.Bold !== undefined)
    {
        _stream["WriteByte"](0);
        _stream["WriteBool"](TextPr.Bold);
    }
    if (TextPr.Italic !== undefined)
    {
        _stream["WriteByte"](1);
        _stream["WriteBool"](TextPr.Italic);
    }
    if (TextPr.Underline !== undefined)
    {
        _stream["WriteByte"](2);
        _stream["WriteBool"](TextPr.Underline);
    }
    if (TextPr.Strikeout !== undefined)
    {
        _stream["WriteByte"](3);
        _stream["WriteBool"](TextPr.Strikeout);
    }

    asc_menu_WriteFontFamily(4, TextPr.FontFamily, _stream);

    if (TextPr.FontSize !== undefined)
    {
        _stream["WriteByte"](5);
        _stream["WriteDouble2"](TextPr.FontSize);
    }

    if(TextPr.Unifill && TextPr.Unifill.fill && TextPr.Unifill.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID && TextPr.Unifill.fill.color)
    {
        var _color = AscCommon.CreateAscColor(TextPr.Unifill.fill.color);
        asc_menu_WriteColor(6, AscCommon.CreateAscColorCustom(_color.r, _color.g, _color.b, false), _stream);
    }
    else if (TextPr.Color !== undefined)
    {
        asc_menu_WriteColor(6, AscCommon.CreateAscColorCustom(TextPr.Color.r, TextPr.Color.g, TextPr.Color.b, TextPr.Color.Auto), _stream);
    }

    if (TextPr.VertAlign !== undefined)
    {
        _stream["WriteByte"](7);
        _stream["WriteLong"](TextPr.VertAlign);
    }

    if (TextPr.HighLight !== undefined)
    {
        if (TextPr.HighLight === AscCommonWord.highlight_None)
        {
            _stream["WriteByte"](12);
        }
        else
        {
            asc_menu_WriteColor(8, AscCommon.CreateAscColorCustom(TextPr.HighLight.r, TextPr.HighLight.g, TextPr.HighLight.b), _stream);
        }
    }

    if (TextPr.DStrikeout !== undefined)
    {
        _stream["WriteByte"](9);
        _stream["WriteBool"](TextPr.DStrikeout);
    }
    if (TextPr.Caps !== undefined)
    {
        _stream["WriteByte"](10);
        _stream["WriteBool"](TextPr.Caps);
    }
    if (TextPr.SmallCaps !== undefined)
    {
        _stream["WriteByte"](11);
        _stream["WriteBool"](TextPr.SmallCaps);
    }
    if (TextPr.Spacing !== undefined)
    {
        _stream["WriteByte"](13);
        _stream["WriteDouble2"](TextPr.Spacing);
    }

    _stream["WriteByte"](255);

    window["native"]["OnCallMenuEvent"](1, _stream);
};

Asc['asc_docs_api'].prototype["Native_Editor_Initialize_Settings"] = function(_params)
{
    window["NativeSupportTimeouts"] = true;

    if (!_params)
        return;

    var _current = { pos : 0 };
    var _continue = true;
    while (_continue)
    {
        var _attr = _params[_current.pos++];
        switch (_attr)
        {
            case 0:
            {
				AscCommon.GlobalSkin.STYLE_THUMBNAIL_WIDTH = _params[_current.pos++];
                break;
            }
            case 1:
            {
				AscCommon.GlobalSkin.STYLE_THUMBNAIL_HEIGHT = _params[_current.pos++];
                break;
            }
            case 2:
            {
                TABLE_STYLE_WIDTH_PIX = _params[_current.pos++];
                break;
            }
            case 3:
            {
                TABLE_STYLE_HEIGHT_PIX = _params[_current.pos++];
                break;
            }
            case 4:
            {
                this.chartPreviewManager.CHART_PREVIEW_WIDTH_PIX = _params[_current.pos++];
                break;
            }
            case 5:
            {
                this.chartPreviewManager.CHART_PREVIEW_HEIGHT_PIX = _params[_current.pos++];
                break;
            }
            case 6:
            {
                var _val = _params[_current.pos++];
                if (_val === true)
                {
                    this.ShowParaMarks = false;
                    AscCommon.CollaborativeEditing.Set_GlobalLock(true);

                    this.isViewMode = true;
                    this.WordControl.m_oDrawingDocument.IsViewMode = true;
                }
                break;
            }
            case 100:
            {
                this.WordControl.m_oDrawingDocument.IsRetina = _params[_current.pos++];
                break;
            }
            case 101:
            {
                this.WordControl.m_oDrawingDocument.IsMobile = _params[_current.pos++];
                window.AscAlwaysSaveAspectOnResizeTrack = true;
                break;
            }
            case 255:
            default:
            {
                _continue = false;
                break;
            }
        }
    }

    AscCommon.AscBrowser.isRetina = this.WordControl.m_oDrawingDocument.IsRetina;
};


Asc['asc_docs_api'].prototype["CheckSlideBounds"] = function(nSlideIndex){
    var oBoundsChecker = new AscFormat.CSlideBoundsChecker();
    this.WordControl.m_oLogicDocument.Draw(nSlideIndex, oBoundsChecker);
    var oBounds = oBoundsChecker.Bounds;
    return [
        oBounds.min_x, oBounds.max_x, oBounds.min_y, oBounds.max_y
    ]
}

Asc['asc_docs_api'].prototype["GetNativePageMeta"] = function(pageIndex, bTh, bIsPlayMode)
{
    this.WordControl.m_oDrawingDocument.RenderPage(pageIndex, bTh, bIsPlayMode);
};

Asc["asc_docs_api"].prototype["asc_nativeOpenFile2"] = function(base64File, version)
{
    this.SpellCheckUrl = '';

    this.WordControl.m_bIsRuler = false;
    this.WordControl.Init();

    this.InitEditor();

    this.DocumentType   = 2;

    AscCommon.g_oIdCounter.Set_Load(true);

    var _loader = new AscCommon.BinaryPPTYLoader();
    _loader.Api = this;

    _loader.Load(base64File, this.WordControl.m_oLogicDocument);
    this.LoadedObject = 1;
    AscCommon.g_oIdCounter.Set_Load(false);
};

Asc['asc_docs_api'].prototype.openDocument = function(file)
{
    _api.asc_nativeOpenFile2(file.data);


    var _presentation = _api.WordControl.m_oLogicDocument;

    var nSlidesCount = _presentation.Slides.length;
    var dPresentationWidth = _presentation.GetWidthMM();
    var dPresentationHeight = _presentation.GetHeightMM();

    var aTransitions = [];
    var slides = _presentation.Slides;
    // for(var i = 0; i < slides.length; ++i){
    //     aTransitions.push(slides[i].transition.ToArray());
    // }
    var _result =  [nSlidesCount, dPresentationWidth, dPresentationHeight, aTransitions];
    var oTheme = null;

    if (null != slides[0])
    {
        oTheme = slides[0].getTheme();
    }
    if (false) {

        this.WordControl.m_oDrawingDocument.AfterLoad();

        
        this.ImageLoader.bIsLoadDocumentFirst = true;

        if (oTheme)
        {
            _api.sendColorThemes(oTheme);
        }

        window["native"]["onEndLoadingFile"](_result);
        this.asc_nativeCalculateFile();

        return;
    }

    var version;
    if (file.changes && this.VersionHistory)
    {
        this.VersionHistory.changes = file.changes;
        this.VersionHistory.applyChanges(this);
    }

    this.WordControl.m_oDrawingDocument.AfterLoad();

    //console.log("ImageMap : " + JSON.stringify(this.WordControl.m_oLogicDocument));

    this.ImageLoader.bIsLoadDocumentFirst = true;
    this.ImageLoader.LoadDocumentImages(this.WordControl.m_oLogicDocument.ImageMap);

    this.WordControl.m_oLogicDocument.Continue_FastCollaborativeEditing();

    //this.asyncFontsDocumentEndLoaded();
    //
    // if (oTheme)
    // {
    //     _api.sendColorThemes(oTheme);
    // }

    if (null != this.WordControl.m_oLogicDocument)
    {
        this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
    }
    window["native"]["onTokenJWT"](_api.CoAuthoringApi.get_jwt());
    window["native"]["onEndLoadingFile"](_result);
    
    this.asc_nativeCalculateFile();

    this.WordControl.m_oDrawingDocument.Collaborative_TargetsUpdate(true);

    _api.asc_GetDefaultTableStyles();

    initSpellCheckApi();

    var t = this;
    setInterval(function() {
        t._autoSave();
    }, 40);
};

Asc['asc_docs_api'].prototype.Update_ParaInd = function( Ind )
{
   // this.WordControl.m_oDrawingDocument.Update_ParaInd(Ind);
};

Asc['asc_docs_api'].prototype.Internal_Update_Ind_Left = function(Left)
{
};

Asc['asc_docs_api'].prototype.Internal_Update_Ind_Right = function(Right)
{
};

Asc['asc_docs_api'].prototype.IsAsyncOpenDocumentImages = function()
{
    return true;
};

Asc['asc_docs_api'].prototype.asyncImageEndLoadedBackground = function(_image)
{
};

/***************************** COPY|PASTE *******************************/

Asc['asc_docs_api'].prototype.Call_Menu_Context_Copy = function()
{
    var dataBuffer = {};

    var clipboard = {};
    clipboard.pushData = function(type, data) {

        if (AscCommon.c_oAscClipboardDataFormat.Text === type) {

            dataBuffer.text = data;

        } else if (AscCommon.c_oAscClipboardDataFormat.Internal === type) {

            if (null != data.drawingUrls && data.drawingUrls.length > 0) {
                dataBuffer.drawingUrls = data.drawingUrls[0];
            }

            dataBuffer.sBase64 = data;
        }
    };

    this.asc_CheckCopy(clipboard, AscCommon.c_oAscClipboardDataFormat.Internal|AscCommon.c_oAscClipboardDataFormat.Text);

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    if (dataBuffer.text) {
        _stream["WriteByte"](0);
        _stream["WriteString2"](dataBuffer.text);
    }

    if (dataBuffer.drawingUrls) {
        _stream["WriteByte"](1);
        _stream["WriteStringA"](dataBuffer.drawingUrls);
    }

    if (dataBuffer.sBase64) {
        _stream["WriteByte"](2);
        _stream["WriteStringA"](dataBuffer.sBase64);
    }

    _stream["WriteByte"](255);

    return _stream;
};
Asc['asc_docs_api'].prototype.Call_Menu_Context_Cut = function()
{
    var dataBuffer = {};

    var clipboard = {};
    clipboard.pushData = function(type, data) {

        if (AscCommon.c_oAscClipboardDataFormat.Text === type) {

            dataBuffer.text = data;

        } else if (AscCommon.c_oAscClipboardDataFormat.Internal === type) {

            if (null != data.drawingUrls && data.drawingUrls.length > 0) {
                dataBuffer.drawingUrls = data.drawingUrls[0];
            }

            dataBuffer.sBase64 = data;
        }
    }

    this.asc_CheckCopy(clipboard, AscCommon.c_oAscClipboardDataFormat.Internal|AscCommon.c_oAscClipboardDataFormat.Text);

    this.asc_SelectionCut();

    var _stream = global_memory_stream_menu;
    _stream["ClearNoAttack"]();

    if (dataBuffer.text) {
        _stream["WriteByte"](0);
        _stream["WriteString2"](dataBuffer.text);
    }

    if (dataBuffer.drawingUrls) {
        _stream["WriteByte"](1);
        _stream["WriteStringA"](dataBuffer.drawingUrls);
    }

    if (dataBuffer.sBase64) {
        _stream["WriteByte"](2);
        _stream["WriteStringA"](dataBuffer.sBase64);
    }

    _stream["WriteByte"](255);

    return _stream;
};

Asc['asc_docs_api'].prototype.Call_Menu_Context_Paste = function(type, param)
{
    if (0 == type)
    {
        this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, param);
    }
    else if (1 == type)
    {
        this.AddImageUrlNative(param, 200, 200);
    }
    else if (2 == type)
    {
        this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, param);
    }
};

Asc['asc_docs_api'].prototype.Call_Menu_Context_Select = function()
{
    this.WordControl.m_oLogicDocument.MoveCursorLeft(false, true);
    this.WordControl.m_oLogicDocument.MoveCursorRight(true, true);
    this.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
};

Asc['asc_docs_api'].prototype.Call_Menu_Context_Delete = function()
{
    this.WordControl.m_oLogicDocument.Remove(-1);
};

Asc['asc_docs_api'].prototype.Call_Menu_Context_SelectAll = function()
{
    this.WordControl.m_oLogicDocument.SelectAll();
};

Asc['asc_docs_api'].prototype.asc_setDocumentPassword = function(password)
{
    var v = {
        "id": this.documentId,
        "userid": this.documentUserId,
        "format": this.documentFormat,
        "c": "reopen",
        "title": this.documentTitle,
        "password": password
    };

    AscCommon.sendCommand(this, null, v);
};

Asc["asc_docs_api"].prototype["asc_nativeGetFileData"] = function()
{
    var oBinaryFileWriter = new AscCommon.CBinaryFileWriter();
    this.WordControl.m_oLogicDocument.CalculateComments();
    oBinaryFileWriter.WriteDocument3(this.WordControl.m_oLogicDocument);

    window["native"]["GetFileData"](oBinaryFileWriter.ImData.data, oBinaryFileWriter.GetCurPosition());

    return true;
};

Asc['asc_docs_api'].prototype.asc_setSpellCheck = function(isOn)
{
    if (editor.WordControl.m_oLogicDocument)
    {
        var _presentation = editor.WordControl.m_oLogicDocument;
        _presentation.Spelling.Use = isOn;
        var _drawing_document = editor.WordControl.m_oDrawingDocument;
        if(isOn)
        {
            this.spellCheckTimerId = setInterval(function(){_presentation.ContinueSpellCheck();}, 500);
        }
        else
        {
            if(this.spellCheckTimerId)
            {
               clearInterval(this.spellCheckTimerId);
            }
        }
        var oCurSlide = _presentation.Slides[_presentation.CurPage];

        if(oCurSlide)
        {
            _drawing_document.OnStartRecalculate(_presentation.Slides.length);
            _drawing_document.OnRecalculatePage(_presentation.CurPage, oCurSlide);
            _drawing_document.OnEndRecalculate();
        }
    }
};

window["native"]["Call_CheckSlideBounds"] = function(nIndex){
    if (window.editor) {
        return window.editor["CheckSlideBounds"](nIndex);
    }
};

window["native"]["Call_GetPageMeta"] = function(nIndex, bTh, bIsPlayMode){
    if (window.editor) {
        return window.editor["GetNativePageMeta"](nIndex, bTh, bIsPlayMode);
    }
};

window["native"]["Call_OnMouseDown"] = function(e) {
    if (window.editor) {
      return window.editor.WordControl.m_oDrawingDocument.OnMouseDown(e);
    }
    return -1;
  };

window["native"]["Call_OnMouseUp"] = function(e) {
    if(window.editor) {
        return window.editor.WordControl.m_oDrawingDocument.OnMouseUp(e);
    }

    return [];
};

window["native"]["Call_OnMouseMove"] = function(e) {
    if(window.editor) {
        window.editor.WordControl.m_oDrawingDocument.OnMouseMove(e);
    }
};

window["native"]["Call_OnKeyboardEvent"] = function(e) {
    return window.editor.WordControl.m_oDrawingDocument.OnKeyboardEvent(e);
};

window["native"]["Call_OnCheckMouseDown"] = function(e) {
    return window.editor.WordControl.m_oDrawingDocument.OnCheckMouseDown(e);
};

window["native"]["Call_ResetSelection"] = function() {
    window.editor.WordControl.m_oLogicDocument.RemoveSelection(false);
    window.editor.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
    window.editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
};

window["native"]["Call_OnUpdateOverlay"] = function(param) {
    if (window.editor) {
        window.editor.WordControl.OnUpdateOverlay(param);
    }
};
window["native"]["Call_SetCurrentPage"] = function(param){
    if (window.editor) {
        var oWC = window.editor.WordControl;
        oWC.m_oLogicDocument.Set_CurPage(param);
        if(oWC.m_oDrawingDocument)
        {
            oWC.m_oDrawingDocument.SlidesCount = oWC.m_oLogicDocument.Slides.length;
            oWC.m_oDrawingDocument.SlideCurrent = oWC.m_oLogicDocument.CurPage;
        }
        oWC.CheckLayouts(false);
    }
};

window["native"]["Call_Menu_Event"] = function (type, _params)
{
    return _api["Call_Menu_Event"](type, _params);
};

window["AscCommon"].sendImgUrls = function(api, images, callback)
{
	var _data = [];
	callback(_data);
};

window["native"]["offline_of"] = function(_params, documentInfo) { return NativeOpenFileP(_params, documentInfo); };

