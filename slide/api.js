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

(function(window, document)
{

	// Import
	var locktype_None = AscCommon.c_oAscLockTypes.kLockTypeNone;
	var locktype_Mine = AscCommon.c_oAscLockTypes.kLockTypeMine;
	var locktype_Other = AscCommon.c_oAscLockTypes.kLockTypeOther;
	var locktype_Other2 = AscCommon.c_oAscLockTypes.kLockTypeOther2;
	var locktype_Other3 = AscCommon.c_oAscLockTypes.kLockTypeOther3;
	var changestype_Drawing_Props   = AscCommon.changestype_Drawing_Props;
	var asc_CSelectedObject         = AscCommon.asc_CSelectedObject;
	var g_oDocumentUrls             = AscCommon.g_oDocumentUrls;
	var sendCommand                 = AscCommon.sendCommand;
	var g_oIdCounter                = AscCommon.g_oIdCounter;
	var g_oTableId                  = AscCommon.g_oTableId;
	var PasteElementsId             = null;
	var global_mouseEvent           = null;
	var History                     = null;

	var c_oAscError             = Asc.c_oAscError;
	var c_oAscFileType          = Asc.c_oAscFileType;
	var c_oAscAsyncAction       = Asc.c_oAscAsyncAction;
	var c_oAscAdvancedOptionsID = Asc.c_oAscAdvancedOptionsID;
	var c_oAscAsyncActionType   = Asc.c_oAscAsyncActionType;
	var c_oAscTypeSelectElement = Asc.c_oAscTypeSelectElement;
	var c_oAscFill              = Asc.c_oAscFill;
	var asc_CShapeFill          = Asc.asc_CShapeFill;
	var asc_CFillBlip           = Asc.asc_CFillBlip;
    var c_oAscFontRenderingModeType = Asc.c_oAscFontRenderingModeType;

	function CAscSlideProps()
	{
		this.Background               = null;
		this.Transition               = null;
		this.LayoutIndex              = null;
		this.lockDelete               = false;
		this.lockLayout               = false;
		this.lockTiming               = false;
		this.lockBackground           = false;
		this.lockTransition           = false;
		this.lockRemove               = false;
		this.lockApplyBackgroundToAll = false;
		this.lockResetBackground      = false;
		this.isHidden                 = false;
		this.resetBackground          = false;
		this.applyBackgroundToAll     = false;
		this.showMasterSp             = null;
		this.IsMaster                 = false;
		this.layoutName               = undefined;
		this.masterName               = undefined;
	}

	CAscSlideProps.prototype.get_background     = function()
	{
		return this.Background;
	};
	CAscSlideProps.prototype.put_background     = function(v)
	{
		this.Background = v;
	};
	CAscSlideProps.prototype.get_LayoutName     = function()
	{
		return this.layoutName;
	};
	CAscSlideProps.prototype.put_LayoutName     = function(v)
	{
		this.layoutName = v;
	};
	CAscSlideProps.prototype.get_MasterName     = function()
	{
		return this.masterName;
	};
	CAscSlideProps.prototype.put_MasterName     = function(v)
	{
		this.masterName = v;
	};
	CAscSlideProps.prototype.get_LayoutIndex     = function()
	{
		return this.LayoutIndex;
	};
	CAscSlideProps.prototype.put_LayoutIndex     = function(v)
	{
		this.LayoutIndex = v;
	};
	CAscSlideProps.prototype.get_transition         = function()
	{
		return this.Transition;
	};
	CAscSlideProps.prototype.put_transition         = function(v)
	{
		this.Transition = v;
	};
	CAscSlideProps.prototype.get_LockDelete     = function()
	{
		return this.lockDelete;
	};
	CAscSlideProps.prototype.put_LockDelete     = function(v)
	{
		this.lockDelete = v;
	};
	CAscSlideProps.prototype.get_LockLayout     = function()
	{
		return this.lockLayout;
	};
	CAscSlideProps.prototype.put_LockLayout     = function(v)
	{
		this.lockLayout = v;
	};
	CAscSlideProps.prototype.get_LockTiming     = function()
	{
		return this.lockTiming;
	};
	CAscSlideProps.prototype.put_LockTiming     = function(v)
	{
		this.lockTiming = v;
	};
	CAscSlideProps.prototype.get_LockBackground = function()
	{
		return this.lockBackground;
	};
	CAscSlideProps.prototype.put_LockBackground = function(v)
	{
		this.lockBackground = v;
	};
	CAscSlideProps.prototype.get_LockTransition = function()
	{
		return this.lockTransition;
	};
	CAscSlideProps.prototype.put_LockTransition = function(v)
	{
		this.lockTransition = v;
	};
	CAscSlideProps.prototype.get_LockRemove     = function()
	{
		return this.lockRemove;
	};
	CAscSlideProps.prototype.put_LockRemove     = function(v)
	{
		this.lockRemove = v;
	};
	CAscSlideProps.prototype.get_IsHidden     = function()
	{
		return this.isHidden;
	};
	CAscSlideProps.prototype.get_LockApplyBackgroundToAll     = function()
	{
		return this.lockApplyBackgroundToAll;
	};
	CAscSlideProps.prototype.put_LockApplyBackgroundToAll     = function(v)
	{
		this.lockApplyBackgroundToAll = v;
	};
	CAscSlideProps.prototype.get_LockResetBackground     = function()
	{
		return this.lockResetBackground;
	};
	CAscSlideProps.prototype.put_LockResetBackground     = function(v)
	{
		this.lockResetBackground = v;
	};
	CAscSlideProps.prototype.get_ResetBackground     = function()
	{
		return this.resetBackground;
	};
	CAscSlideProps.prototype.put_ResetBackground     = function(v)
	{
		this.resetBackground = v;
	};
	CAscSlideProps.prototype.get_ApplyBackgroundToAll     = function()
	{
		return this.applyBackgroundToAll;
	};
	CAscSlideProps.prototype.put_ApplyBackgroundToAll     = function(v)
	{
		this.applyBackgroundToAll = v;
	};
	CAscSlideProps.prototype.get_ShowMasterSp     = function()
	{
		return this.showMasterSp;
	};
	CAscSlideProps.prototype.put_ShowMasterSp     = function(v)
	{
		this.showMasterSp = v;
	};
	CAscSlideProps.prototype.get_IsMasterSelected     = function()
	{
		return this.IsMaster;
	};

	function CAscChartProp(obj)
	{
		if (obj)
		{

			this.Width    = (undefined != obj.w) ? obj.w : undefined;
			this.Height   = (undefined != obj.h) ? obj.h : undefined;
			this.Position = new Asc.CPosition({X : obj.x, Y : obj.y});

			this.Locked          = (undefined != obj.locked) ? obj.locked : false;
			this.lockAspect      = (undefined != obj.lockAspect) ? obj.lockAspect : false;
			this.ChartProperties = (undefined != obj.chartProps) ? obj.chartProps : null;

			this.severalCharts      = obj.severalCharts != undefined ? obj.severalCharts : false;
			this.severalChartTypes  = obj.severalChartTypes != undefined ? obj.severalChartTypes : undefined;
			this.severalChartStyles = obj.severalChartStyles != undefined ? obj.severalChartStyles : undefined;

			this.title = obj.title != undefined ? obj.title : undefined;
			this.description = obj.description != undefined ? obj.description : undefined;
			this.name = obj.name != undefined ? obj.name : undefined;
		}
		else
		{
			this.Width           = undefined;
			this.Height          = undefined;
			this.Position        = undefined;
			this.Locked          = false;
			this.lockAspect      = undefined;
			this.ChartProperties = new Asc.asc_ChartSettings();

			this.severalCharts      = false;
			this.severalChartTypes  = undefined;
			this.severalChartStyles = undefined;
            this.title = undefined;
            this.description = undefined;
		}
	}

	CAscChartProp.prototype.get_ChangeLevel = function()
	{
		return this.ChangeLevel;
	};
	CAscChartProp.prototype.put_ChangeLevel = function(v)
	{
		this.ChangeLevel = v;
	};

	CAscChartProp.prototype.get_CanBeFlow     = function()
	{
		return this.CanBeFlow;
	};
	CAscChartProp.prototype.get_Width         = function()
	{
		return this.Width;
	};
	CAscChartProp.prototype.put_Width         = function(v)
	{
		this.Width = v;
	};
	CAscChartProp.prototype.get_Height        = function()
	{
		return this.Height;
	};
	CAscChartProp.prototype.put_Height        = function(v)
	{
		this.Height = v;
	};
	CAscChartProp.prototype.get_WrappingStyle = function()
	{
		return this.WrappingStyle;
	};
	CAscChartProp.prototype.put_WrappingStyle = function(v)
	{
		this.WrappingStyle = v;
	};
	// Возвращается объект класса Asc.asc_CPaddings
	CAscChartProp.prototype.get_Paddings      = function()
	{
		return this.Paddings;
	};
	// Аргумент объект класса Asc.asc_CPaddings
	CAscChartProp.prototype.put_Paddings      = function(v)
	{
		this.Paddings = v;
	};
	CAscChartProp.prototype.get_AllowOverlap  = function()
	{
		return this.AllowOverlap;
	};
	CAscChartProp.prototype.put_AllowOverlap  = function(v)
	{
		this.AllowOverlap = v;
	};
	// Возвращается объект класса CPosition
	CAscChartProp.prototype.get_Position      = function()
	{
		return this.Position;
	};
	// Аргумент объект класса CPosition
	CAscChartProp.prototype.put_Position      = function(v)
	{
		this.Position = v;
	};
	CAscChartProp.prototype.get_PositionH     = function()
	{
		return this.PositionH;
	};
	CAscChartProp.prototype.put_PositionH     = function(v)
	{
		this.PositionH = v;
	};
	CAscChartProp.prototype.get_PositionV     = function()
	{
		return this.PositionV;
	};
	CAscChartProp.prototype.put_PositionV     = function(v)
	{
		this.PositionV = v;
	};
	CAscChartProp.prototype.get_Value_X       = function(RelativeFrom)
	{
		if (null != this.Internal_Position) return this.Internal_Position.Calculate_X_Value(RelativeFrom);
		return 0;
	};
	CAscChartProp.prototype.get_Value_Y       = function(RelativeFrom)
	{
		if (null != this.Internal_Position) return this.Internal_Position.Calculate_Y_Value(RelativeFrom);
		return 0;
	};

	CAscChartProp.prototype.get_ImageUrl     = function()
	{
		return this.ImageUrl;
	};
	CAscChartProp.prototype.put_ImageUrl     = function(v)
	{
		this.ImageUrl = v;
	};
	CAscChartProp.prototype.get_Group        = function()
	{
		return this.Group;
	};
	CAscChartProp.prototype.put_Group        = function(v)
	{
		this.Group = v;
	};
	CAscChartProp.prototype.asc_getFromGroup = function()
	{
		return this.fromGroup;
	};
	CAscChartProp.prototype.asc_putFromGroup = function(v)
	{
		this.fromGroup = v;
	};

	CAscChartProp.prototype.get_isChartProps = function()
	{
		return this.isChartProps;
	};
	CAscChartProp.prototype.put_isChartPross = function(v)
	{
		this.isChartProps = v;
	};

	CAscChartProp.prototype.get_SeveralCharts     = function()
	{
		return this.severalCharts;
	};
	CAscChartProp.prototype.put_SeveralCharts     = function(v)
	{
		this.severalCharts = v;
	};
	CAscChartProp.prototype.get_SeveralChartTypes = function()
	{
		return this.severalChartTypes;
	};
	CAscChartProp.prototype.put_SeveralChartTypes = function(v)
	{
		this.severalChartTypes = v;
	};

	CAscChartProp.prototype.get_SeveralChartStyles = function()
	{
		return this.severalChartStyles;
	};
	CAscChartProp.prototype.put_SeveralChartStyles = function(v)
	{
		this.severalChartStyles = v;
	};

	CAscChartProp.prototype.get_VerticalTextAlign = function()
	{
		return this.verticalTextAlign;
	};
	CAscChartProp.prototype.put_VerticalTextAlign = function(v)
	{
		this.verticalTextAlign = v;
	};

	CAscChartProp.prototype.get_Locked = function()
	{
		return this.Locked;
	};

	CAscChartProp.prototype.get_ChartProperties = function()
	{
		return this.ChartProperties;
	};

	CAscChartProp.prototype.put_ChartProperties = function(v)
	{
		this.ChartProperties = v;
	};

	CAscChartProp.prototype.get_ShapeProperties = function()
	{
		return this.ShapeProperties;
	};

	CAscChartProp.prototype.put_ShapeProperties = function(v)
	{
		this.ShapeProperties = v;
	};

	CAscChartProp.prototype.asc_getType    = function()
	{
		return this.ChartProperties.asc_getType();
	};
	CAscChartProp.prototype.asc_getSubType = function()
	{
		return this.ChartProperties.asc_getSubType();
	};

	CAscChartProp.prototype.asc_getStyleId = function()
	{
		return this.ChartProperties.asc_getStyleId();
	};

	CAscChartProp.prototype.asc_getHeight = function()
	{
		return this.Height;
	};
	CAscChartProp.prototype.asc_getWidth  = function()
	{
		return this.Width;
	};

	CAscChartProp.prototype.asc_setType    = function(v)
	{
		this.ChartProperties.asc_setType(v);
	};
	CAscChartProp.prototype.asc_setSubType = function(v)
	{
		this.ChartProperties.asc_setSubType(v);
	};

	CAscChartProp.prototype.asc_setStyleId = function(v)
	{
		this.ChartProperties.asc_setStyleId(v);
	};

	CAscChartProp.prototype.asc_setHeight = function(v)
	{
		this.Height = v;
	};
	CAscChartProp.prototype.asc_setWidth  = function(v)
	{
		this.Width = v;
	};

	CAscChartProp.prototype.asc_setTitle = function(v)
	{
		this.title = v;
	};
	CAscChartProp.prototype.asc_setDescription  = function(v)
	{
		this.description = v;
	};

	CAscChartProp.prototype.asc_getTitle = function()
	{
		return this.title;
	};
	CAscChartProp.prototype.asc_getDescription  = function()
	{
		return this.description;
	};

	CAscChartProp.prototype.asc_getPosition  = function()
	{
		return this.Position;
	};

	CAscChartProp.prototype.asc_putPosition  = function(v)
	{
		this.Position = v;
	};

	CAscChartProp.prototype.asc_getName  = function()
	{
		return this.name;
	};
	CAscChartProp.prototype.asc_putName  = function(v)
	{
		this.name = v;
	};

	CAscChartProp.prototype.getType = function()
	{
		return this.ChartProperties && this.ChartProperties.getType();
	};
	CAscChartProp.prototype.putType = function(v)
	{
		return this.ChartProperties && this.ChartProperties.putType(v);
	};

	CAscChartProp.prototype.getStyle      = function()
	{
		return this.ChartProperties && this.ChartProperties.getStyle();
	};
	CAscChartProp.prototype.putStyle      = function(v)
	{
		return this.ChartProperties && this.ChartProperties.putStyle(v);
	};
	CAscChartProp.prototype.getLockAspect = function()
	{
		return this.lockAspect;
	};
	CAscChartProp.prototype.putLockAspect = function(v)
	{
		return this.lockAspect = v;
	};

	CAscChartProp.prototype.changeType = function(v)
	{
		return this.ChartProperties && this.ChartProperties.changeType(v);
	};

	// CSearchResult - returns result of searching
	function CSearchResult(obj)
	{
		this.Object = obj;
	}

	CSearchResult.prototype.get_Text = function()
	{
		return this.Object.text;
	};

	CSearchResult.prototype.get_Navigator = function()
	{
		return this.Object.navigator;
	};

	CSearchResult.prototype.put_Navigator = function(obj)
	{
		this.Object.navigator = obj;
	};
	CSearchResult.prototype.put_Text      = function(obj)
	{
		this.Object.text = obj;
	};

	/**
	 *
	 * @param config
	 * @constructor
	 * @extends {AscCommon.baseEditorsApi}
	 */
	function asc_docs_api(config)
	{
		AscCommon.baseEditorsApi.call(this, config, AscCommon.c_oEditorId.Presentation);

		/************ private!!! **************/
		this.WordControl = null;

		this.documentFormatSave = c_oAscFileType.PPTX;

		this.ThemeLoader   = null;
		this.standartThemesStatus = 0;
		this.tmpThemesPath = null;
		this.tmpIsFreeze   = null;
		this.tmpSlideDiv   = null;
		this.tmpTextArtDiv = null;
		this.tmpViewRulers = null;
		this.tmpZoomType   = null;
        this.tmpDocumentUnits = null;

        this.DocumentUrl     = "";
		this.bNoSendComments = false;

		this.isApplyChangesOnOpen        = false;
		this.isApplyChangesOnVersionHistory = false;

        this.IsSpellCheckCurrentWord = false;

		this.IsSupportEmptyPresentation = true;

		this.ShowParaMarks        = false;
		this.ShowSmartGuides      = true;
		this.ShowGuides           = false;
		this.ShowGridlines        = false;


		this.isAddSpaceBetweenPrg = false;
		this.isPageBreakBefore    = false;
		this.isKeepLinesTogether  = false;
		this.isPresentationEditor = true;
		this.bSelectedSlidesTheme = false;
		this.bIsShowAnimTab       = false;

		this.isMarkerFormat             = false;
		this.isShowTableEmptyLine       = false;//true;
		this.isShowTableEmptyLineAttack = false;//true;

		this.bInit_word_control = false;
		this.isDocumentModify   = false;

        this.tmpFontRenderingMode = null;

		this.isPasteFonts_Images = false;

		this.nCurPointItemsLength = -1;

		this.pasteCallback       = null;
		this.pasteImageMap       = null;
		this.EndActionLoadImages = 0;

		this.isSaveFonts_Images = false;
		this.saveImageMap       = null;

		this.ServerImagesWaitComplete = false;

		this.SelectedObjectsStack = [];

		this.CoAuthoringApi.isPowerPoint = true;

		// объекты, нужные для отправки в тулбар (шрифты, стили)
		this._gui_editor_themes   = null;
		this._gui_document_themes = null;

		this.EndShowMessage = undefined;

		this.isOnlyDemonstration = false;

		this.mouseMode = "select";


		this.presentationViewMode = Asc.c_oAscPresentationViewMode.normal;

		if (window.editor == undefined)
		{
			window.editor = this;
			window.editor;
			window['editor'] = window.editor;
			Asc['editor'] = Asc.editor = this;

			if (window["NATIVE_EDITOR_ENJINE"])
				editor = window.editor;
		}

		this.reporterWindow 		= null;
		this.reporterWindowCounter 	= 0;
		this.reporterStartObject 	= null;
		this.isReporterMode = ("reporter" == config['using']) ? true : false;
		this.disableReporterEvents = false;
		this.TableStylesPreviewGenerator = null;

		if (this.isReporterMode)
		{
			var _windowOnResize = function() {
				if (undefined != window._resizeTimeout && -1 != window._resizeTimeout)
					clearTimeout(window._resizeTimeout);
				window._resizeTimeout = setTimeout(function() {
					window.editor.Resize();
					window._resizeTimeout = -1;
				}, 50);
			};

			if (window.addEventListener)
			{
				window.addEventListener("resize", _windowOnResize, false);
			}
			else if (window.attachEvent)
			{
				window.attachEvent("onresize", _windowOnResize);
			}
			else
			{
				window["onresize"] = _windowOnResize;
			}
		}

		if (this.isReporterMode)
			this.watermarkDraw = null;

		this.demoBackgroundColor = null;
		this._init();
	}

	asc_docs_api.prototype = Object.create(AscCommon.baseEditorsApi.prototype);
	asc_docs_api.prototype.constructor = asc_docs_api;

	asc_docs_api.prototype.sendEvent = function()
	{
        this.sendInternalEvent.apply(this, arguments);
		var name = arguments[0];
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = 0; i < _callbacks[name].length; ++i)
			{
				_callbacks[name][i].apply(this || window, Array.prototype.slice.call(arguments, 1));
			}
			return true;
		}
		return false;
	};
	asc_docs_api.prototype._loadSdkImages = function () {
		const aImages = AscCommon.getIconsForLoad();
		this.ImageLoader.LoadImagesWithCallback(aImages, function() {
		}, []);
	};

	/////////////////////////////////////////////////////////////////////////
	///////////////////CoAuthoring and Chat api//////////////////////////////
	/////////////////////////////////////////////////////////////////////////
	// Init CoAuthoring
	asc_docs_api.prototype._coAuthoringSetChange = function(change, oColor)
	{
		var oChange = new AscCommon.CCollaborativeChanges();
		oChange.Set_Data(change);
		oChange.Set_Color(oColor);
		AscCommon.CollaborativeEditing.Add_Changes(oChange);
	};

	asc_docs_api.prototype._coAuthoringSetChanges = function(e, oColor)
	{
		var Count = e.length;
		for (var Index = 0; Index < Count; ++Index)
			this._coAuthoringSetChange(e[Index], oColor);
	};

	asc_docs_api.prototype._coAuthoringInitEnd = function()
	{
		var t                                        = this;
		this.CoAuthoringApi.onCursor                 = function(e)
		{
			if (true === AscCommon.CollaborativeEditing.Is_Fast())
			{
				t.WordControl.m_oLogicDocument.Update_ForeignCursor(e[e.length - 1]['cursor'], e[e.length - 1]['user'], true, e[e.length - 1]['useridoriginal']);
			}
		};
		this.CoAuthoringApi.onConnectionStateChanged = function(e)
		{
			if (true === AscCommon.CollaborativeEditing.Is_Fast() && false === e['state'])
			{
				editor.WordControl.m_oLogicDocument.Remove_ForeignCursor(e['id']);
			}
			t.sendEvent("asc_onConnectionStateChanged", e);
		};
		this.CoAuthoringApi.onLocksAcquired          = function(e)
		{
			if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksAcquired, e))
			{
				return;
			}

			if (2 != e["state"])
			{

				var block_value = e["blockValue"];
				var classes     = [];
				switch (block_value["type"])
				{
					case c_oAscLockTypeElemPresentation.Object:
					{
						classes.push(block_value["objId"]);
						//classes.push(block_value["slideId"]);
						break;
					}
					case c_oAscLockTypeElemPresentation.Slide:
					{
						classes.push(block_value["val"]);
						break;
					}
					case c_oAscLockTypeElemPresentation.Presentation:
					{
						break;
					}
				}

				for (var i = 0; i < classes.length; ++i)
				{
					var Class = g_oTableId.Get_ById(classes[i]);// g_oTableId.Get_ById( Id );
					if (null != Class)
					{
						var Lock = Class.Lock;

						var OldType = Class.Lock.Get_Type();
						if (locktype_Other2 === OldType || locktype_Other3 === OldType)
						{
							Lock.Set_Type(locktype_Other3, true);
						}
						else
						{
							Lock.Set_Type(locktype_Other, true);
						}
						if (Class instanceof AscCommonSlide.PropLocker)
						{
							var object = g_oTableId.Get_ById(Class.objectId);
							if (object instanceof AscCommonSlide.Slide && Class === object.deleteLock)
							{
								editor.WordControl.m_oLogicDocument.DrawingDocument.LockSlide(object.num);
							}
						}
						// Выставляем ID пользователя, залочившего данный элемент
						Lock.Set_UserId(e["user"]);

						if (Class instanceof AscCommonSlide.PropLocker)
						{
							var object = g_oTableId.Get_ById(Class.objectId);
							if (object instanceof AscCommonSlide.CPresentation)
							{
								let oPresentation = editor.WordControl.m_oLogicDocument;
								if (Class === oPresentation.themeLock)
								{
									editor.sendEvent("asc_onLockDocumentTheme");
								}
								else if (Class === oPresentation.schemeLock)
								{
									editor.sendEvent("asc_onLockDocumentSchema");
								}
								else if (Class === oPresentation.slideSizeLock)
								{
									editor.sendEvent("asc_onLockDocumentProps");
								}
								else if (Class === oPresentation.viewPrLock)
								{
									editor.sendEvent("asc_onLockViewProps");
								}
								else if (Class === oPresentation.hdrFtrLock)
								{
									editor.sendEvent("asc_onLockSlideHdrFtrApplyToAll");
								}
							}
						}
						if (Class instanceof AscCommon.CComment)
						{
							editor.sync_LockComment(Class.Get_Id(), e["user"]);
						}
						if(Class instanceof AscCommon.CCore)
						{
							editor.sendEvent("asc_onLockCore", true);
						}

						// TODO: Здесь для ускорения надо сделать проверку, является ли текущим элемент с
						//       заданным Id. Если нет, тогда и не надо обновлять состояние.
						editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
					}
					else
					{
						if (classes[i].indexOf("new_object") > -1 && block_value["type"] === c_oAscLockTypeElemPresentation.Object)
						{
							var slide_id    = block_value["slideId"];
							var delete_lock = g_oTableId.Get_ById(slide_id);
							if (AscCommon.isRealObject(delete_lock))
							{
								var Lock    = delete_lock.Lock;
								var OldType = Lock.Get_Type();
								if (locktype_Other2 === OldType || locktype_Other3 === OldType)
								{
									Lock.Set_Type(locktype_Other3, true);
								}
								else
								{
									Lock.Set_Type(locktype_Other, true);
								}
								editor.WordControl.m_oLogicDocument.DrawingDocument.LockSlide(g_oTableId.Get_ById(delete_lock.objectId).num);
							}
							else
							{
								AscCommon.CollaborativeEditing.Add_NeedLock(slide_id, e["user"]);
							}
						}
						else
						{
							AscCommon.CollaborativeEditing.Add_NeedLock(classes[i], e["user"]);
						}
					}
				}
			}
		};
		this.CoAuthoringApi.onLocksReleased          = function(e, bChanges)
		{
			if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksReleased, e, bChanges))
			{
				return;
			}

			var Id;
			var block_value = e["block"];
			var classes     = [];
			switch (block_value["type"])
			{
				case c_oAscLockTypeElemPresentation.Object:
				{
					classes.push(block_value["objId"]);
					//classes.push(block_value["slideId"]);
					break;
				}
				case c_oAscLockTypeElemPresentation.Slide:
				{
					classes.push(block_value["val"]);
					break;
				}
				case c_oAscLockTypeElemPresentation.Presentation:
				{
					break;
				}
			}
			for (var i = 0; i < classes.length; ++i)
			{
				Id        = classes[i];
				var Class = g_oTableId.Get_ById(Id);
				if (null != Class)
				{
					var Lock = Class.Lock;

					if ("undefined" != typeof(Lock))
					{
						var CurType = Lock.Get_Type();

						var NewType = locktype_None;

						if (CurType === locktype_Other)
						{
							if (true != bChanges)
							{
								NewType = locktype_None;
							}
							else
							{
								NewType = locktype_Other2;
								AscCommon.CollaborativeEditing.Add_Unlock(Class);
							}
						}
						else if (CurType === locktype_Mine)
						{
							// Такого быть не должно
							NewType = locktype_Mine;
						}
						else if (CurType === locktype_Other2 || CurType === locktype_Other3)
						{
							NewType = locktype_Other2;
						}

						Lock.Set_Type(NewType, true);
						if (Class instanceof AscCommonSlide.PropLocker)
						{
							var object = g_oTableId.Get_ById(Class.objectId);
							if (object instanceof AscCommonSlide.Slide && Class === object.deleteLock)
							{
								if (NewType !== locktype_Mine && NewType !== locktype_None)
								{
									editor.WordControl.m_oLogicDocument.DrawingDocument.LockSlide(object.num);
								}
								else
								{
									editor.WordControl.m_oLogicDocument.DrawingDocument.UnLockSlide(object.num);
								}
							}
							if (object instanceof AscCommonSlide.CPresentation)
							{
								if (Class === object.themeLock)
								{
									if (NewType !== locktype_Mine && NewType !== locktype_None)
									{
										editor.sendEvent("asc_onLockDocumentTheme");
									}
									else
									{
										editor.sendEvent("asc_onUnLockDocumentTheme");
									}
								}
								if (Class === object.slideSizeLock)
								{
									if (NewType !== locktype_Mine && NewType !== locktype_None)
									{
										editor.sendEvent("asc_onLockDocumentProps");
									}
									else
									{
										editor.sendEvent("asc_onUnLockDocumentProps");
									}
								}
								if (Class === object.viewPrLock)
								{
									if (NewType !== locktype_Mine && NewType !== locktype_None)
									{
										editor.sendEvent("asc_onLockViewProps");
									}
									else
									{
										editor.sendEvent("asc_onUnLockViewProps");
									}
								}
								if (Class === object.hdrFtrLock)
								{
									if (NewType !== locktype_Mine && NewType !== locktype_None)
									{
										editor.sendEvent("asc_onLockSlideHdrFtrApplyToAll");
									}
									else
									{
										editor.sendEvent("asc_onUnLockSlideHdrFtrApplyToAll");
									}
								}
							}
						}

					}
				}
				else
				{
					AscCommon.CollaborativeEditing.Remove_NeedLock(Id);
				}
			}
		};
		this.CoAuthoringApi.onSaveChanges            = function(e, userId, bFirstLoad)
		{
			// bSendEvent = false - это означает, что мы загружаем имеющиеся изменения при открытии
			var Changes = new AscCommon.CCollaborativeChanges();
			Changes.Set_Data(e);
			AscCommon.CollaborativeEditing.Add_Changes(Changes);

			// т.е. если bSendEvent не задан, то посылаем  сообщение + когда загрузился документ
			if (!bFirstLoad && t.bInit_word_control)
			{
				t.sync_CollaborativeChanges();
			}
		};
		this.CoAuthoringApi.onChangesIndex = function(changesIndex)
		{
			if (t.isLiveViewer() && changesIndex >= 0 && changesIndex < AscCommon.CollaborativeEditing.GetAllChangesCount()) {
				let count = AscCommon.CollaborativeEditing.GetAllChangesCount() - changesIndex;
				AscCommon.CollaborativeEditing.UndoGlobal(count);
			}
		};
		this.CoAuthoringApi.onRecalcLocks            = function(e)
		{
			if (e && true === AscCommon.CollaborativeEditing.Is_Fast())
			{
				AscCommon.CollaborativeEditing.UpdateForeignCursorByAdditionalInfo(JSON.parse(e));
			}
		};
	};
	
	asc_docs_api.prototype.startCollaborationEditing = function()
	{
		AscCommon.CollaborativeEditing.Start_CollaborationEditing();
		this.asc_setDrawCollaborationMarks(true);
		if (this.WordControl && this.WordControl.m_oDrawingDocument)
		{
			this.WordControl.m_oDrawingDocument.Start_CollaborationEditing();
		}
	};
	asc_docs_api.prototype.endCollaborationEditing = function()
	{
		AscCommon.CollaborativeEditing.End_CollaborationEditing();
		if (this.WordControl && this.WordControl.m_oLogicDocument &&
			false !== this.WordControl.m_oLogicDocument.DrawingDocument.IsLockObjectsEnable)
		{
			this.WordControl.m_oLogicDocument.DrawingDocument.IsLockObjectsEnable = false;
			this.WordControl.m_oLogicDocument.DrawingDocument.FirePaint();
		}
	};

	asc_docs_api.prototype.isMasterMode = function() {
		return this.presentationViewMode === Asc.c_oAscPresentationViewMode.masterSlide && !this.isSlideShow();
	};
	asc_docs_api.prototype.isNormalMode = function() {
		return this.presentationViewMode === Asc.c_oAscPresentationViewMode.normal && !this.isSlideShow();
	};
	asc_docs_api.prototype.asc_IsMasterMode = function() {
		return this.isMasterMode();
	};
	asc_docs_api.prototype.asc_AddMasterSlide = function() {
		if(!this.isMasterMode()) return;


		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument) return;
		oLogicDocument.AddNewMasterSlide();
	};
	asc_docs_api.prototype.asc_AddSlideLayout = function() {
		if(!this.isMasterMode()) return;
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument) return;
		oLogicDocument.AddNewLayout();
	};
	asc_docs_api.prototype.asc_StartAddPlaceholder = function(nType, bVertical, bStart) {
		if(!this.isMasterMode()) return;

		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument) return;
		let oCurSlide = oLogicDocument.GetCurrentSlide();
		if(oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		oLogicDocument.StartAddShape("textRect", bStart, nType, bVertical);
	};
	asc_docs_api.prototype.asc_setLayoutTitle = function(bVal) {
		if(!this.isMasterMode()) return;

		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument) return;
		let oCurSlide = oLogicDocument.GetCurrentSlide();
		if(oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		oLogicDocument.SetLayoutTitle(bVal);
	};
	asc_docs_api.prototype.asc_setLayoutFooter = function(bVal) {
		if(!this.isMasterMode()) return;

		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument) return;
		let oCurSlide = oLogicDocument.GetCurrentSlide();
		if(oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		oLogicDocument.SetLayoutFooter(bVal);
	};

    //----------------------------------------------------------------------------------------------------------------------
    // SpellCheck_CallBack
    //          Функция ответа от сервера.
    //----------------------------------------------------------------------------------------------------------------------
    asc_docs_api.prototype.SpellCheck_CallBack = function(Obj)
    {
        if (undefined != Obj && undefined != Obj["ParagraphId"])
        {
            var ParaId    = Obj["ParagraphId"];
            var Paragraph = g_oTableId.Get_ById(ParaId);
            var Type      = Obj["type"];
            if (null != Paragraph)
            {
                if ("spell" === Type)
                {
                    Paragraph.SpellChecker.SpellCheckResponse(Obj["RecalcId"], Obj["usrCorrect"]);
                    Paragraph.ReDraw();
                }
                else if ("suggest" === Type)
                {
                    Paragraph.SpellChecker.SuggestResponse(Obj["RecalcId"], Obj["ElementId"], Obj["usrSuggest"][0]);
                    this.sync_SpellCheckVariantsFound();
                }
            }
			else
			{
				this.WordControl.m_oLogicDocument.Spelling.RemoveWaitingParagraphById(ParaId);
			}
        }
    };

    asc_docs_api.prototype._spellCheckDisconnect   = function()
    {
        if (this.WordControl.m_oLogicDocument)
            this.WordControl.m_oLogicDocument.TurnOffSpellCheck();
    };

    asc_docs_api.prototype.pre_Save = function(_images)
	{
		this.isSaveFonts_Images = true;
		this.saveImageMap       = _images;
		this.WordControl.m_oDrawingDocument.CheckFontNeeds();
		this.FontLoader.LoadDocumentFonts2(this.WordControl.m_oLogicDocument.Fonts);
	};

    asc_docs_api.prototype.asc_GetRevisionsChangesStack = function()
	{
		return [];
	};
	asc_docs_api.prototype.asc_undoAllChanges       = function()
	{
		this.WordControl.m_oLogicDocument.Document_Undo({All : true});
	};
	asc_docs_api.prototype.asc_CloseFile            = function()
	{
		History.Clear();
		g_oIdCounter.Clear();
		g_oTableId.Clear();
		AscCommon.CollaborativeEditing.Clear();
		this.isApplyChangesOnOpenEnabled = true;
		this.isDocumentLoadComplete = false;
		this.ServerImagesWaitComplete = false;
		this.turnOffSpecialModes();

		var oLogicDocument = this.WordControl.m_oLogicDocument;
		oLogicDocument.StopRecalculate();
		oLogicDocument.StopSpellCheck();
		AscCommon.pptx_content_loader.ImageMapChecker = {};

		this.WordControl.m_oDrawingDocument.CloseFile();
		this.sendEvent("asc_onCloseFile");
	};
	asc_docs_api.prototype.asc_SetFastCollaborative = function(isOn)
	{
		if (AscCommon.CollaborativeEditing)
			AscCommon.CollaborativeEditing.Set_Fast(isOn);
	};

	asc_docs_api.prototype.sync_CollaborativeChanges = function()
	{
		if (true !== AscCommon.CollaborativeEditing.Is_Fast())
			this.sendEvent("asc_onCollaborativeChanges");
	};

	// Эвент о пришедщих изменениях
	asc_docs_api.prototype.syncCollaborativeChanges = function()
	{
		this.sendEvent("asc_onCollaborativeChanges");
	};


	asc_docs_api.prototype.SetCollaborativeMarksShowType = function(Type)
	{
		this.CollaborativeMarksShowType = Type;
	};

	asc_docs_api.prototype.GetCollaborativeMarksShowType = function(Type)
	{
		return this.CollaborativeMarksShowType;
	};

	asc_docs_api.prototype.Clear_CollaborativeMarks = function()
	{
		AscCommon.CollaborativeEditing.Clear_CollaborativeMarks(true);
	};

	asc_docs_api.prototype._onUpdateDocumentCanSave = function()
	{
		var CollEditing = AscCommon.CollaborativeEditing;

		// Можно модифицировать это условие на более быстрое (менять самим состояние в аргументах, а не запрашивать каждый раз)
		var isCanSave = this.isDocumentModified() || (true !== CollEditing.Is_SingleUser() && 0 !== CollEditing.getOwnLocksLength());

		if (true === CollEditing.Is_Fast() && true !== CollEditing.Is_SingleUser())
			isCanSave = false;

		if (isCanSave !== this.isDocumentCanSave)
		{
			this.isDocumentCanSave = isCanSave;
			this.sendEvent('asc_onDocumentCanSaveChanged', this.isDocumentCanSave);
		}
	};
	asc_docs_api.prototype._onUpdateDocumentCanUndoRedo = function ()
	{
        if (this.WordControl && this.WordControl.m_oLogicDocument)
            this.WordControl.m_oLogicDocument.Document_UpdateUndoRedoState();
	};

	///////////////////////////////////////////
	asc_docs_api.prototype.CheckChangedDocument = function()
	{
		if (true === History.Have_Changes())
		{
			// дублирование евента. когда будет undo-redo - тогда
			// эти евенты начнут отличаться
			this.SetDocumentModified(true);
		}
		else
		{
			this.SetDocumentModified(false);
		}

		this._onUpdateDocumentCanSave();
	};
	asc_docs_api.prototype.SetUnchangedDocument = function()
	{
		this.SetDocumentModified(false);
		this._onUpdateDocumentCanSave();
	};

	asc_docs_api.prototype.SetDocumentModified = function(bValue)
	{
		this.isDocumentModify = bValue;
		this.sendEvent("asc_onDocumentModifiedChanged");

		if (undefined !== window["AscDesktopEditor"])
		{
			window["AscDesktopEditor"]["onDocumentModifiedChanged"](bValue);
		}
	};

	asc_docs_api.prototype.isDocumentModified = function()
	{
		if (!this.canSave)
		{
			// Пока идет сохранение, мы не закрываем документ
			return true;
		}
		return this.isDocumentModify;
	};

	asc_docs_api.prototype.asc_getCurrentFocusObject = function()
    {
        if (!this.WordControl || !this.WordControl.Thumbnails)
            return 1;
        return this.WordControl.Thumbnails.FocusObjType;
    };

	asc_docs_api.prototype.sync_BeginCatchSelectedElements = function()
	{
		if (0 != this.SelectedObjectsStack.length)
			this.SelectedObjectsStack.splice(0, this.SelectedObjectsStack.length);
	};
	asc_docs_api.prototype.sync_EndCatchSelectedElements   = function()
	{
		this.sendEvent("asc_onFocusObject", this.SelectedObjectsStack);
	};
	asc_docs_api.prototype.getSelectedElements             = function(bUpdate)
	{
        if (true === bUpdate){
        	if(this.WordControl && this.WordControl.m_oLogicDocument){
                this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
			}
		}

		return this.SelectedObjectsStack;
	};

	asc_docs_api.prototype.asc_getAnimatedObjectsStack = function()
	{
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation)
		{
			return this.SelectedObjectsStack;
		}
		let oDrawingsController = oPresentation.GetCurrentController();
		if(!oDrawingsController)
		{
			return this.SelectedObjectsStack;
		}
		if(oDrawingsController.getSelectedObjects().length > 0)
		{
			return this.SelectedObjectsStack;
		}
		let oCurSlide = oPresentation.GetCurrentSlide();
		if(!oCurSlide)
		{
			return [];
		}
		let oTiming = oCurSlide.timing;
		if(!oTiming)
		{
			return this.SelectedObjectsStack;
		}
		let aEffects = oTiming.getSelectedEffects();
		if(aEffects.length === 0)
		{
			return this.SelectedObjectsStack;
		}
		let aDrawings = [];
		let oMapDrawings = {};
		for(let nEffect = 0; nEffect < aEffects.length; ++nEffect)
		{
			let oEffect = aEffects[nEffect];
			let sObjectId = oEffect.getObjectId();
			if(!sObjectId)
			{
				continue;
			}
			if(oMapDrawings[sObjectId])
			{
				continue;
			}

			let oDrawing = AscCommon.g_oTableId.Get_ById(sObjectId);
			if (!oDrawing)
			{
				continue;
			}
			oMapDrawings[sObjectId] = true;
			aDrawings.push(oDrawing);
		}
		if(aDrawings.length === 0)
		{
			return this.SelectedObjectsStack;
		}

		let aStack = [];


		for(let nSelObj = 0; nSelObj < this.SelectedObjectsStack.length; ++nSelObj)
		{
			if(this.SelectedObjectsStack[nSelObj].Type === c_oAscTypeSelectElement.Slide)
			{
				aStack.push(this.SelectedObjectsStack[nSelObj]);
				break;
			}
		}
		let oDrawingPr = oDrawingsController.getDrawingPropsFromArray(aDrawings);
		let oParaPr = oDrawingsController.getParagraphParaPr();
		let oTextPr = oDrawingsController.getParagraphTextPr();


		if (oTextPr && oParaPr)
		{
			aStack.push(new asc_CSelectedObject(c_oAscTypeSelectElement.Paragraph, new Asc.asc_CParagraphProperty(oParaPr)));
		}
		let oImgPr = oDrawingPr.imageProps;
		let oSpPr = oDrawingPr.shapeProps;
		let oChartPr = oDrawingPr.chartProps;
		let oTblPr = oDrawingPr.tableProps;
		if (oImgPr)
		{
			oImgPr.Width = oImgPr.w;
			oImgPr.Height = oImgPr.h;
			oImgPr.Position = {X: oImgPr.x, Y: oImgPr.y};
			if (AscFormat.isRealBool(oImgPr.locked) && oImgPr.locked)
			{
				oImgPr.Locked = true;
			}
			aStack.push(this.getImageSelectedObject(oImgPr));
		}
		if (oSpPr)
		{
			oSpPr.Position = new Asc.CPosition({X: oSpPr.x, Y: oSpPr.y});
			aStack.push(new asc_CSelectedObject(c_oAscTypeSelectElement.Shape, AscFormat.CreateAscShapePropFromProp(oSpPr)));
		}
		if (oDrawingPr.animProps)
		{
			aStack.push(new asc_CSelectedObject(c_oAscTypeSelectElement.Animation, oDrawingPr.animProps));
		}
		if (oChartPr && oChartPr.chartProps)
		{
			if (oSpPr)
			{
				oChartPr.x = oSpPr.x;
				oChartPr.y = oSpPr.y;
				if (oSpPr.Position)
				{
					oChartPr.Position = new Asc.CPosition(oSpPr.Position);
				}
			}
			aStack.push(this.getImageSelectedObject(oChartPr));
		}
		if (oTblPr)
		{
			aStack.push(new asc_CSelectedObject(c_oAscTypeSelectElement.Table, new Asc.CTableProp(oTblPr)));
		}
		return aStack;
	};
	asc_docs_api.prototype.sync_ChangeLastSelectedElement  = function(type, obj)
	{
		var oUnkTypeObj = null;

		switch (type)
		{
			case c_oAscTypeSelectElement.Paragraph:
				oUnkTypeObj = new Asc.asc_CParagraphProperty(obj);
				break;
			case c_oAscTypeSelectElement.Image:
				oUnkTypeObj = new Asc.asc_CImgProperty(obj);
				break;
			case c_oAscTypeSelectElement.Table:
				oUnkTypeObj = new Asc.CTableProp(obj);
				break;
			case c_oAscTypeSelectElement.Shape:
				oUnkTypeObj = obj;
				break;
		}

		var _i       = this.SelectedObjectsStack.length - 1;
		var bIsFound = false;
		while (_i >= 0)
		{
			if (this.SelectedObjectsStack[_i].Type == type)
			{

				this.SelectedObjectsStack[_i].Value = oUnkTypeObj;
				bIsFound                            = true;
				break;
			}
			_i--;
		}

		if (!bIsFound)
		{
			this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(type, oUnkTypeObj);
		}
	};

	asc_docs_api.prototype.Init          = function()
	{
		this.WordControl.Init();
	};
	asc_docs_api.prototype.asc_setLocale = function(val)
	{
		this.locale = val;
	};
	asc_docs_api.prototype.asc_getLocale = function()
	{
		return this.locale;
	};
	asc_docs_api.prototype.SetThemesPath = function(path)
	{
	    if (this.standartThemesStatus == 0)
        {
            // 0 - начальное состояние
            // 1 - просто чтобы не позволить грузить два раза
            // 2 - загрузка скрипта/конец открытия документа
            // 3 - конец открытия документа/загрузка скрипта

            this.standartThemesStatus = 1;
            var t = this;
            AscCommon.loadScript(path + "/themes.js", function() {
                t.standartThemesStatus++;
                if (t.ThemeLoader)
                    t.ThemeLoader.Themes._init();
                if (2 < t.standartThemesStatus)
                    t.WordControl.m_oLogicDocument.SendThemesThumbnails();
            }, function() {
                t.standartThemesStatus++;
                if (2 < t.standartThemesStatus)
                    t.WordControl.m_oLogicDocument.SendThemesThumbnails();
            });
        }

		if (!this.isLoadFullApi)
		{
			this.tmpThemesPath = path;
			return;
		}

		this.ThemeLoader.ThemesUrl = path;
		if (this.documentOrigin)
		{
			this.ThemeLoader.ThemesUrlAbs = AscCommon.joinUrls(this.documentOrigin + this.documentPathname, path);
		}
		else
		{
			this.ThemeLoader.ThemesUrlAbs = path;
		}
	};

	asc_docs_api.prototype.CreateCSS = function()
	{
		var _head = document.getElementsByTagName('head')[0];

		var style0       = document.createElement('style');
		style0.type      = 'text/css';
		style0.innerHTML = ".block_elem { position:absolute;padding:0;margin:0; }";
		_head.appendChild(style0);

		var style1       = document.createElement('style');
		style1.type      = 'text/css';
		style1.innerHTML = ".buttonTabs {\
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABMAAAA5CAMAAADjueCuAAAABGdBTUEAALGPC/xhBQAAAEhQTFRFAAAAWFhYZWVlSEhIY2NjV1dXQ0NDYWFhYmJiTk5OVlZWYGBgVFRUS0tLbGxsRERETExMZmZmVVVVXl5eR0dHa2trPj4+u77CpAZQrwAAAAF0Uk5TAEDm2GYAAABwSURBVDjL1dHHDoAgEEVR7NLr4P//qQm6EMaFxtje8oTF5ELIpU35Fstf3GegsPEBG+uwSYpNB1qNKreoDeNw/r6dLr/tnFpbbNZj8wKbk8W/1d6ZPjfrhdHx9c4fbA9wzMYWm3OFhbQmbC2ue6z9DCH/Exf/mU3YAAAAAElFTkSuQmCC);\
background-position: 0px 0px;\
background-repeat: no-repeat;\
}";
		_head.appendChild(style1);

		var style3       = document.createElement('style');
		style3.type      = 'text/css';
		style3.innerHTML = ".buttonPrevPage {\
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAABgBAMAAADm/++TAAAABGdBTUEAALGPC/xhBQAAABJQTFRFAAAA////UVNVu77Cenp62Nrc3x8hMQAAAAF0Uk5TAEDm2GYAAABySURBVCjPY2AgETDBGEoKUAElJcJSxANjKGAwDQWDYAKMIBhDSRXCCFJSIixF0GS4M+AMExcwcCbAcIQxBEUgDEdBQcJSBE2GO4PU6IJHASxS4NGER4p28YWIAlikwKMJjxTt4gsRBbBIgUcTHini4wsAwMmIvYZODL0AAAAASUVORK5CYII=);\
background-position: 0px 0px;\
background-repeat: no-repeat;\
}";
		_head.appendChild(style3);

		var style4       = document.createElement('style');
		style4.type      = 'text/css';
		style4.innerHTML = ".buttonNextPage {\
background-image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAABgBAMAAADm/++TAAAABGdBTUEAALGPC/xhBQAAABJQTFRFAAAA////UVNVu77Cenp62Nrc3x8hMQAAAAF0Uk5TAEDm2GYAAABySURBVCjPY2AgETDBGEoKUAElJcJSxANjKGAwDQWDYAKMIBhDSRXCCFJSIixF0GS4M+AMExcwcCbAcIQxBEUgDEdBQcJSBE2GO4PU6IJHASxS4NGER4p28YWIAlikwKMJjxTt4gsRBbBIgUcTHini4wsAwMmIvYZODL0AAAAASUVORK5CYII=);\
background-position: 0px -48px;\
background-repeat: no-repeat;\
}";
		_head.appendChild(style4);
	};

	asc_docs_api.prototype.CreateComponents = function()
	{
		this.asc_setSkin(this.skinObject);
		delete this.skinObject;

		this.CreateCSS();

		var _innerHTML = "<div id=\"id_panel_thumbnails\" class=\"block_elem\" style=\"touch-action:none;-webkit-touch-callout:none;background-color:" + AscCommon.GlobalSkin.BackgroundColorThumbnails + ";\">\
									<div id=\"id_panel_thumbnails_split\" class=\"block_elem\" style=\"pointer-events:none;background-color:" + AscCommon.GlobalSkin.BackgroundColorThumbnails + ";\"></div>\
		                            <canvas id=\"id_thumbnails_background\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;z-index:1\"></canvas>\
		                            <canvas id=\"id_thumbnails\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;z-index:2\"></canvas>\
		                            <div id=\"id_vertical_scroll_thmbnl\" style=\"left:0;top:0;width:1px;overflow:hidden;position:absolute;\">\
									</div>\
		                        </div>\
		                    <div id=\"id_main_parent\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;overflow:hidden;border-left-width: 1px;border-left-color:" + AscCommon.GlobalSkin.BorderSplitterColor + "; border-left-style: solid;\" UNSELECTABLE=\"on\">\
                            <div id=\"id_main\" class=\"block_elem\" style=\"z-index:5;-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";overflow:hidden;\" UNSELECTABLE=\"on\">\
								<div id=\"id_panel_left\" class=\"block_elem\">\
									<canvas id=\"id_buttonTabs\" class=\"block_elem\"></canvas>\
									<canvas id=\"id_vert_ruler\" class=\"block_elem\"></canvas>\
								</div>\
                                <div id=\"id_panel_top\" class=\"block_elem\">\
									<canvas id=\"id_hor_ruler\" class=\"block_elem\"></canvas>\
                                </div>\
                                <div id=\"id_main_view\" class=\"block_elem\" style=\"overflow:hidden\">\
                                    <canvas id=\"id_viewer\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";z-index:6\"></canvas>\
                                    <canvas id=\"id_viewer_overlay\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;z-index:7\"></canvas>\
                                    <div id=\"id_target_cursor\" class=\"block_elem\" width=\"1\" height=\"1\" style=\"-ms-touch-action: none;-webkit-user-select: none;width:2px;height:13px;display:none;z-index:9;\"></div>\
                                </div>\
							    <div id=\"id_panel_right\" class=\"block_elem\" style=\"margin-right:1px;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";z-index:0;\">\
							        <div id=\"id_buttonRulers\" class=\"block_elem buttonRuler\"></div>\
								    <div id=\"id_vertical_scroll\" style=\"left:0;top:0;width:14px;overflow:hidden;position:absolute;\">\
								    </div>\
								    <div id=\"id_buttonPrevPage\" class=\"block_elem buttonPrevPage\"></div>\
								    <div id=\"id_buttonNextPage\" class=\"block_elem buttonNextPage\"></div>\
                                </div>\
                                <div id=\"id_horscrollpanel\" class=\"block_elem\" style=\"margin-bottom:1px;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";\">\
                                    <div id=\"id_horizontal_scroll\" style=\"left:0;top:0;height:14px;overflow:hidden;position:absolute;width:100%;\">\
                                    </div>\
                                </div>\
                            </div>";

		if (true)
		{
			_innerHTML += "<div id=\"id_bottom_pannels_container\" class=\"block_elem\" style=\"-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;overflow:hidden;background-color:" + AscCommon.GlobalSkin.BackgroundColorNotes + ";\">\
							<div id=\"id_panel_notes\" class=\"block_elem\" style=\"-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;overflow:hidden;background-color:" + AscCommon.GlobalSkin.BackgroundColorNotes + ";\">\
                                <canvas id=\"id_notes\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;background-color:" + AscCommon.GlobalSkin.BackgroundColorNotes + ";z-index:6\"></canvas>\
                                <canvas id=\"id_notes_overlay\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;z-index:7\"></canvas>\
                                <div id=\"id_vertical_scroll_notes\" style=\"left:0;top:0;width:16px;overflow:hidden;position:absolute;\">\
                                    <div id=\"panel_right_scroll_notes\" class=\"block_elem\" style=\"left:0;top:0;width:1px;height:1px;\"></div>\
                                </div>\
                            </div>\
                            <div id=\"id_panel_animation\" class=\"block_elem\" style=\"-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;overflow:hidden;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";\">\
								<div id=\"id_anim_header\" class=\"block_elem\">\
									<canvas id=\"id_anim_header_canvas\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;background-color:" + AscCommon.GlobalSkin.BackgroundColorNotes + ";z-index:8\"></canvas>\
								</div>\
								<div id=\"id_anim_list_container\" class=\"block_elem\">\
									<canvas id=\"id_anim_list_canvas\" class=\"block_elem\"></canvas>\
									<div id=\"id_anim_list_scroll\" style=\"left:0;top:0;width:16px;overflow:hidden;position:absolute;\">\
                                    	<div id=\"id_pane_anim_list_scroll\" class=\"block_elem\" style=\"left:0;top:0;width:1px;height:1px;\"></div>\
                                	</div>\
								</div>\
								<div id=\"id_anim_timeline_container\" class=\"block_elem\">\
									<canvas id=\"id_anim_timeline_canvas\" class=\"block_elem\"></canvas>\
								</div>\
                            </div>\
						</div>\
						</div>";
		}

		if (this.HtmlElement)
			_innerHTML += this.HtmlElement.innerHTML;

		if (this.HtmlElement != null)
		{
			this.HtmlElement.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			this.HtmlElement.innerHTML = _innerHTML;
		}
	};

	asc_docs_api.prototype.updateSkin = function()
	{
		var baseElem = document.getElementById(this.HtmlElementName);
		if (baseElem)
			baseElem.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;

		var obj_id_main = document.getElementById("id_main");
		if (obj_id_main)
		{
			obj_id_main.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_panel_thumbnails").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColorThumbnails;
			document.getElementById("id_panel_thumbnails_split").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColorThumbnails;
			document.getElementById("id_main_parent").style.borderLeftColor = AscCommon.GlobalSkin.BorderSplitterColor;
			document.getElementById("id_viewer").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_panel_right").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
			document.getElementById("id_horscrollpanel").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
			document.getElementById("id_panel_notes").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColorNotes;
			document.getElementById("id_bottom_pannels_container").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColorNotes;
			document.getElementById("id_bottom_pannels_container").style.borderTopColor = AscCommon.GlobalSkin.BorderSplitterColor;
			document.getElementById("id_panel_animation").style.borderTopColor = AscCommon.GlobalSkin.BorderSplitterColor;
			document.getElementById("id_notes").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColorNotes;
		}

		if (this.isReporterMode)
		{
			if (document.getElementById("dem_id_pointer") && this.WordControl)
			{
				var _head = document.getElementsByTagName('head')[0];
				var style = document.createElement('style');
				style.type = 'text/css';
				var styleContent = "";
				styleContent += ".dem_menu {background-color: " + GlobalSkin.DemBackgroundColor + "; border: 1px solid " + GlobalSkin.DemSplitterColor + ";}";
				styleContent += "#dem_id_draw_menu li>a{color:" + GlobalSkin.DemButtonTextColor +";}";
				styleContent += "#dem_id_draw_menu li>a:hover{background-color:" + GlobalSkin.DemButtonBackgroundColorHover + ";}";
				styleContent += "#dem_id_draw_menu li>a[data-checked=\"true\"]{color:" + GlobalSkin.DemButtonTextColorActive + ";background-color:" + GlobalSkin.DemButtonBackgroundColorActive + ";}";
				styleContent += "#dem_id_draw_menu >li.submenu>a:after{border-left-color:" + GlobalSkin.DemButtonTextColor + ";}";
				styleContent += ".dem_draw_menu_divider { background-color:" + GlobalSkin.DemSplitterColor + ";}";
				if (GlobalSkin.Name === "theme-classic-light") {
					styleContent += ("a[data-checked] > .menu-item-icon { background-position-x: -20px; }");
				}
				styleContent += this.WordControl.getStylesReporter();
				style.innerHTML = styleContent;
				_head.appendChild(style);

				["id_reporter_dem", "id_reporter_dem_controller", "id_main_parent", "id_reporter_dem_parent", "id_main"].forEach(function (id) {
					document.getElementById(id).style.backgroundColor = GlobalSkin.BackgroundColorThumbnails;
				});
			}
		}

		if (this.WordControl && this.WordControl.m_oBody)
			this.WordControl.OnResize(true);

	};

	asc_docs_api.prototype.initDefaultShortcuts = function()
	{
		// [[ActionType, KeyCode, Ctrl, Shift, Alt]]
		var aShortcuts =
		[
			[Asc.c_oAscPresentationShortcutType.EditSelectAll, 65, true, false, false],
			[Asc.c_oAscPresentationShortcutType.EditUndo, 90, true, false, false],
			[Asc.c_oAscPresentationShortcutType.EditRedo, 89, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Cut, 88, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Copy, 67, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Paste, 86, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Duplicate, 68, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Print, 80, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Save, 83, true, false, false],
			[Asc.c_oAscPresentationShortcutType.ShowContextMenu, 93, false, false, false],
			[Asc.c_oAscPresentationShortcutType.ShowContextMenu, 121, false, true, false],
			[Asc.c_oAscPresentationShortcutType.ShowContextMenu, 57351, false, false, false],
			[Asc.c_oAscPresentationShortcutType.ShowParaMarks, 56, true, true, false],
			[Asc.c_oAscPresentationShortcutType.Bold, 66, true, false, false],
			[Asc.c_oAscPresentationShortcutType.CopyFormat, 67, true, false, true],
			[Asc.c_oAscPresentationShortcutType.CenterAlign, 69, true, false, false],
			[Asc.c_oAscPresentationShortcutType.EuroSign, 69, true, false, true],
			[Asc.c_oAscPresentationShortcutType.Group, 71, true, false, false],
			[Asc.c_oAscPresentationShortcutType.UnGroup, 71, true, true, false],
			[Asc.c_oAscPresentationShortcutType.Italic, 73, true, false, false],
			[Asc.c_oAscPresentationShortcutType.JustifyAlign, 74, true, false, false],
			[Asc.c_oAscPresentationShortcutType.AddHyperlink, 75, true, false, false],
			[Asc.c_oAscPresentationShortcutType.BulletList, 76, true, true, false],
			[Asc.c_oAscPresentationShortcutType.LeftAlign, 76, true, false, false],
			[Asc.c_oAscPresentationShortcutType.RightAlign, 82, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Underline, 85, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Strikethrough, 53, true, false, false],
			[Asc.c_oAscPresentationShortcutType.PasteFormat, 86, true, false, true],
			[Asc.c_oAscPresentationShortcutType.Superscript, 188, true, false, false],
			[Asc.c_oAscPresentationShortcutType.Subscript, 190, true, false, false],
			[Asc.c_oAscPresentationShortcutType.EnDash, 189, true, true, false],
			[Asc.c_oAscPresentationShortcutType.EnDash, 173, true, true, false],
			[Asc.c_oAscPresentationShortcutType.DecreaseFont, 219, true, false, false],
			[Asc.c_oAscPresentationShortcutType.IncreaseFont, 221, true, false, false],
			[Asc.c_oAscPresentationShortcutType.SpeechWorker, 90, true, false, true]
		];
		this.initShortcuts(aShortcuts, false)
	};

	asc_docs_api.prototype.InitEditor = function()
	{
		this.WordControl.m_oLogicDocument                    = new AscCommonSlide.CPresentation(this.WordControl.m_oDrawingDocument);
		this.WordControl.m_oDrawingDocument.m_oLogicDocument = this.WordControl.m_oLogicDocument;

		if (this.WordControl.MobileTouchManager)
			this.WordControl.MobileTouchManager.delegate.LogicDocument = this.WordControl.m_oLogicDocument;
	};

	asc_docs_api.prototype.SetInterfaceDrawImagePlaceSlide = function(div_id)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpSlideDiv = div_id;
			return;
		}
		this.WordControl.m_oDrawingDocument.InitGuiCanvasSlide(div_id);
	};

	asc_docs_api.prototype.SetInterfaceDrawImagePlaceTextArt = function(div_id)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpTextArtDiv = div_id;
			return;
		}
		this.WordControl.m_oDrawingDocument.InitGuiCanvasTextArt(div_id);
	};

	asc_docs_api.prototype.OpenDocumentFromBin = function(url, gObject)
	{
		this.InitEditor();
		this.DocumentType = 2;

		var _loader = new AscCommon.BinaryPPTYLoader();

		_loader.Api = this;
		g_oIdCounter.Set_Load(true);
		AscFonts.IsCheckSymbols = true;
		_loader.Load(gObject, this.WordControl.m_oLogicDocument);
		this.WordControl.m_oLogicDocument.Set_FastCollaborativeEditing(true);

		if (History && History.Update_FileDescription)
			History.Update_FileDescription(_loader.stream.data, _loader.stream.size);

		this.LoadedObject = 1;
		g_oIdCounter.Set_Load(false);
		AscFonts.IsCheckSymbols = false;

		this.WordControl.m_oDrawingDocument.CheckFontNeeds();
		this.FontLoader.LoadDocumentFonts(this.WordControl.m_oLogicDocument.Fonts, false);

		g_oIdCounter.Set_Load(false);

		if (this.isMobileVersion)
		{
			AscCommon.AscBrowser.isSafariMacOs   = false;
			PasteElementsId.PASTE_ELEMENT_ID     = "wrd_pastebin";
			PasteElementsId.ELEMENT_DISPAY_STYLE = "none";
		}

		if (AscCommon.AscBrowser.isSafariMacOs)
			setInterval(AscCommon.SafariIntervalFocus, 10);
	};


	asc_docs_api.prototype.OpenDocumentFromZipNoInit = function(data)
	{

		if (!data) {
			return false;
		}
		var openXml = AscCommon.openXml;
		var StaxParser = AscCommon.StaxParser;
		var xmlParserContext = new AscCommon.XmlParserContext();
		xmlParserContext.DrawingDocument = this.WordControl.m_oDrawingDocument;

		let jsZlib = new AscCommon.ZLib();
		if (!jsZlib.open(data)) {
			return false;
		}

		var reader;
		xmlParserContext.zip = jsZlib;
		var doc = new openXml.OpenXmlPackage(jsZlib, null);

		let oTableStylesPart = doc.getPartByUri("/ppt/tableStyles.xml");
		if(oTableStylesPart) {
			let oContentTableStyles = oTableStylesPart.getDocumentContent();
			if(oContentTableStyles) {
				let oTableStylesReader = new StaxParser(oContentTableStyles, oTableStylesPart, xmlParserContext);
				this.WordControl.m_oLogicDocument.readTableStylesFromXml(oTableStylesReader);
			}
		}
		let oCommentAuthorsPart = doc.getPartByUri("/ppt/commentAuthors.xml");
		if(oCommentAuthorsPart) {

			let oContentCommentAuthors = oCommentAuthorsPart.getDocumentContent();
			if(oContentCommentAuthors) {
				let oCommentAuthorsReader = new StaxParser(oContentCommentAuthors, oCommentAuthorsPart, xmlParserContext);
				this.WordControl.m_oLogicDocument.readCommentAuthors(oCommentAuthorsReader);
			}
		}


		var documentPart = doc.getPartByRelationshipType(openXml.Types.presentation.relationType);
		var contentDocument = documentPart.getDocumentContent();
		reader = new StaxParser(contentDocument, documentPart, xmlParserContext);
		this.WordControl.m_oLogicDocument.fromXml(reader, true);
		let oAppPart = doc.getPartByRelationshipType(openXml.Types.extendedFileProperties.relationType);
		let oContentApp = oAppPart.getDocumentContent();
		let oAppReader = new StaxParser(oContentApp, oAppPart, xmlParserContext);
		this.WordControl.m_oLogicDocument.App = new AscCommon.CApp();
		this.WordControl.m_oLogicDocument.App.fromXml(oAppReader, true);

		let oPresentationPrPart = doc.getPartByRelationshipType(openXml.Types.presentationProperties.relationType);
		if(oPresentationPrPart) {
			let oContentPresentationPr = oPresentationPrPart.getDocumentContent();
			if(oContentPresentationPr) {
				let oPresentationReader = new StaxParser(oContentPresentationPr, oPresentationPrPart, xmlParserContext);
				let oPresPr = new AscFormat.CPresentationProperties(this.WordControl.m_oLogicDocument);
				oPresPr.fromXml(oPresentationReader, true);
			}
		}
		let oCorePart = doc.getPartByRelationshipType(openXml.Types.coreFileProperties.relationType);
		if(oCorePart) {
			let oContentCore = oCorePart.getDocumentContent();
			if(oContentCore) {
				let oCoreReader = new StaxParser(oContentCore, oCorePart, xmlParserContext);
				this.WordControl.m_oLogicDocument.Core = new AscCommon.CCore();
				this.WordControl.m_oLogicDocument.Core.fromXml(oCoreReader, true);
			}
		}
		let oCustomPrPart = doc.getPartByRelationshipType(openXml.Types.customFileProperties.relationType);
		if(oCustomPrPart) {
			let oContentCustomPr = oCustomPrPart.getDocumentContent();
			if(oContentCustomPr) {
				let oCustomPrReader = new StaxParser(oContentCustomPr, oCustomPrPart, xmlParserContext);
				this.WordControl.m_oLogicDocument.CustomProperties.fromXml(oCustomPrReader, true);
			}
		}
		var context = reader.context;
		this.WordControl.m_oLogicDocument.ImageMap = context.loadDataLinks();
		context.GenerateSmartArts();
		jsZlib.close();
		return true;
	};

	asc_docs_api.prototype.OpenDocumentFromZip = function(data)
	{

		this.InitEditor();
		this.DocumentType = 2;

		g_oIdCounter.Set_Load(true);
		this.OpenDocumentFromZipNoInit(data);
		g_oIdCounter.Set_Load(false);
		this.WordControl.m_oLogicDocument.Set_FastCollaborativeEditing(true);

		this.LoadedObject = 1;
		AscFonts.IsCheckSymbols = false;

		this.WordControl.m_oDrawingDocument.CheckFontNeeds();
		this.FontLoader.LoadDocumentFonts(this.WordControl.m_oLogicDocument.Fonts, false);


		if (this.isMobileVersion)
		{
			AscCommon.AscBrowser.isSafariMacOs   = false;
			PasteElementsId.PASTE_ELEMENT_ID     = "wrd_pastebin";
			PasteElementsId.ELEMENT_DISPAY_STYLE = "none";
		}
		if (AscCommon.AscBrowser.isSafariMacOs)
			setInterval(AscCommon.SafariIntervalFocus, 10);

		return true;
	};

	// Callbacks
	/* все имена callback'оф начинаются с On. Пока сделаны:
	 OnBold,
	 OnItalic,
	 OnUnderline,
	 OnTextPrBaseline(возвращается расположение строки - supstring, superstring, baseline),
	 OnPrAlign(выравнивание по ширине, правому краю, левому краю, по центру),
	 OnListType( возвращается AscCommon.asc_CListType )

	 фейк-функции ожидающие TODO:
	 Print,Undo,Redo,Copy,Cut,Paste,Share,Save,Download & callbacks
	 OnFontName, OnFontSize, OnLineSpacing

	 OnFocusObject( возвращается массив asc_CSelectedObject )
	 OnInitEditorStyles( возвращается CStylesPainter )
	 OnSearchFound( возвращается CSearchResult );
	 OnParaSpacingLine( возвращается AscCommon.asc_CParagraphSpacing )
	 OnLineSpacing( не используется? )
	 OnTextColor( возвращается AscCommon.CColor )
	 OnTextHighLight( возвращается AscCommon.CColor )
	 OnInitEditorFonts( возвращается массив объектов СFont )
	 OnFontFamily( возвращается asc_CTextFontFamily )
	 */
	var _callbacks = {};

	asc_docs_api.prototype.asc_registerCallback = function(name, callback)
	{
		if (!_callbacks.hasOwnProperty(name))
			_callbacks[name] = [];
		_callbacks[name].push(callback);
	};

	asc_docs_api.prototype.asc_unregisterCallback = function(name, callback)
	{
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = _callbacks[name].length - 1; i >= 0; --i)
			{
				if (_callbacks[name][i] == callback)
					_callbacks[name].splice(i, 1);
			}
		}
	};

	asc_docs_api.prototype.asc_checkNeedCallback = function(name)
	{
		if (_callbacks.hasOwnProperty(name))
		{
			return true;
		}
		return false;
	};

	// get functions
	asc_docs_api.prototype.get_TextProps = function()
	{
		var Doc    = this.WordControl.m_oLogicDocument;
		var ParaPr = Doc.GetCalculatedParaPr();
		var TextPr = Doc.GetCalculatedTextPr();

		// return { ParaPr: ParaPr, TextPr : TextPr };
		return new Asc.CParagraphAndTextProp(ParaPr, TextPr);	// uncomment if this method will be used externally. 20/03/2012 uncommented for testers
	};

	// -------
	// тут методы, замены евентов
	asc_docs_api.prototype.get_PropertyEditorThemes = function()
	{
		var ret = [this._gui_editor_themes, this._gui_document_themes];
		return ret;
	};

	// -------

	asc_docs_api.prototype.UpdateTextPr = function(TextPr)
	{
		if ("undefined" != typeof(TextPr))
		{
			if (TextPr.Color !== undefined)
			{
				this.WordControl.m_oDrawingDocument.TargetCursorColor.R = TextPr.Color.r;
				this.WordControl.m_oDrawingDocument.TargetCursorColor.G = TextPr.Color.g;
				this.WordControl.m_oDrawingDocument.TargetCursorColor.B = TextPr.Color.b;
			}
			if (TextPr.Bold === undefined)
				TextPr.Bold = false;
			if (TextPr.Italic === undefined)
				TextPr.Italic = false;
			if (TextPr.Underline === undefined)
				TextPr.Underline = false;
			if (TextPr.Strikeout === undefined)
				TextPr.Strikeout = false;
			if (TextPr.FontFamily === undefined)
				TextPr.FontFamily = {Index : 0, Name : ""};
			if (TextPr.FontSize === undefined)
				TextPr.FontSize = "";

			this.sync_BoldCallBack(TextPr.Bold);
			this.sync_ItalicCallBack(TextPr.Italic);
			this.sync_UnderlineCallBack(TextPr.Underline);
			this.sync_StrikeoutCallBack(TextPr.Strikeout);
			this.sync_TextPrFontSizeCallBack(TextPr.FontSize);
			this.sync_TextPrFontFamilyCallBack(TextPr.FontFamily);

			if (TextPr.VertAlign !== undefined)
				this.sync_VerticalAlign(TextPr.VertAlign);
			if (TextPr.Spacing !== undefined)
				this.sync_TextSpacing(TextPr.Spacing);
			if (TextPr.DStrikeout !== undefined)
				this.sync_TextDStrikeout(TextPr.DStrikeout);
			if (TextPr.Caps !== undefined)
				this.sync_TextCaps(TextPr.Caps);
			if (TextPr.SmallCaps !== undefined)
				this.sync_TextSmallCaps(TextPr.SmallCaps);
			if (TextPr.Position !== undefined)
				this.sync_TextPosition(TextPr.Position);
			if (TextPr.Lang !== undefined)
				this.sync_TextLangCallBack(TextPr.Lang);

			if (TextPr.Unifill !== undefined)
			{
				this.sync_TextColor2(TextPr.Unifill);
			}

			if (AscCommon.isRealObject(TextPr.HighlightColor))
			{
				var oRGB = TextPr.HighlightColor.RGBA;
				this.sendEvent("asc_onTextHighLight", new AscCommon.CColor(oRGB.R, oRGB.G, oRGB.B));
			}
			else
			{
				this.sendEvent("asc_onTextHighLight", AscCommonWord.highlight_None);
			}
		}
	};

	asc_docs_api.prototype.sync_TextSpacing      = function(Spacing)
	{
		this.sendEvent("asc_onTextSpacing", Spacing);
	};
	asc_docs_api.prototype.sync_TextDStrikeout   = function(Value)
	{
		this.sendEvent("asc_onTextDStrikeout", Value);
	};
	asc_docs_api.prototype.sync_TextCaps         = function(Value)
	{
		this.sendEvent("asc_onTextCaps", Value);
	};
	asc_docs_api.prototype.sync_TextSmallCaps    = function(Value)
	{
		this.sendEvent("asc_onTextSmallCaps", Value);
	};
	asc_docs_api.prototype.sync_TextPosition     = function(Value)
	{
		this.sendEvent("asc_onTextPosition", Value);
	};
	asc_docs_api.prototype.sync_TextLangCallBack = function(Lang)
	{
		this.sendEvent("asc_onTextLanguage", Lang.Val);
	};

	asc_docs_api.prototype.sync_VerticalTextAlign = function(align)
	{
		this.sendEvent("asc_onVerticalTextAlign", align);
	};
	asc_docs_api.prototype.sync_Vert              = function(vert)
	{
		this.sendEvent("asc_onVert", vert);
	};

	asc_docs_api.prototype.UpdateParagraphProp = function(ParaPr)
	{

		ParaPr.StyleName  = "";
		var TextPr        = editor.WordControl.m_oLogicDocument.GetCalculatedTextPr();
		var oDrawingProps = editor.WordControl.m_oLogicDocument.Get_GraphicObjectsProps();
		if (oDrawingProps.shapeProps && oDrawingProps.shapeProps.locked
			|| oDrawingProps.chartProps && oDrawingProps.chartProps.locked
			|| oDrawingProps.tableProps && oDrawingProps.tableProps.Locked)
		{
			ParaPr.Locked = true;
		}
		ParaPr.Subscript   = ( TextPr.VertAlign === AscCommon.vertalign_SubScript ? true : false );
		ParaPr.Superscript = ( TextPr.VertAlign === AscCommon.vertalign_SuperScript ? true : false );
		ParaPr.Strikeout   = TextPr.Strikeout;
		ParaPr.DStrikeout  = TextPr.DStrikeout;
		ParaPr.AllCaps     = TextPr.Caps;
		ParaPr.SmallCaps   = TextPr.SmallCaps;
		ParaPr.TextSpacing = TextPr.Spacing;
		ParaPr.Position    = TextPr.Position;
		ParaPr.ListType = AscFormat.fGetListTypeFromBullet(ParaPr.Bullet);
		this.sync_ParaSpacingLine(ParaPr.Spacing);
		this.Update_ParaInd(ParaPr.Ind);
		this.sync_PrAlignCallBack(ParaPr.Jc);
		this.sync_ParaStyleName(ParaPr.StyleName);
		this.sync_ListType(ParaPr.ListType);
		this.sync_PrPropCallback(ParaPr);
	};
	/*----------------------------------------------------------------*/
	/*functions for working with clipboard, document*/
	asc_docs_api.prototype._printDesktop = function (options)
	{
		var opt = {};
        if (options && options.advancedOptions)
		{
			if (Asc.c_oAscPrintType.Selection === options.advancedOptions.asc_getPrintType())
				opt["printOptions"] = {"selection": 1};
			opt["nativeOptions"] = options.advancedOptions.asc_getNativeOptions();
		}

		window["AscDesktopEditor"]["Print"](JSON.stringify(opt));
		return true;
	};
	asc_docs_api.prototype.Undo           = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return;
		
		this.WordControl.m_oLogicDocument.Document_Undo();
	};
	asc_docs_api.prototype.Redo           = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return;

		this.WordControl.m_oLogicDocument.Document_Redo();
	};
	asc_docs_api.prototype.Copy           = function()
	{
		return AscCommon.g_clipboardBase.Button_Copy();
	};
	asc_docs_api.prototype.Update_ParaTab = function(Default_Tab, ParaTabs)
	{
		this.WordControl.m_oDrawingDocument.Update_ParaTab(Default_Tab, ParaTabs);
	};
	asc_docs_api.prototype.Cut            = function()
	{
		return AscCommon.g_clipboardBase.Button_Cut();
	};
	asc_docs_api.prototype.Paste          = function()
	{
		if (!this.WordControl.m_oLogicDocument)
			return false;

		if (AscCommon.g_clipboardBase.IsWorking())
			return false;

		return AscCommon.g_clipboardBase.Button_Paste();
	};

	asc_docs_api.prototype.asc_ShowSpecialPasteButton = function(props)
	{
		if (window["NATIVE_EDITOR_ENJINE"])
		{
			return;
		}

		var presentation = this.WordControl.m_oLogicDocument;
		var drawingDocument = presentation.DrawingDocument;
		var notesFocus = presentation.IsFocusOnNotes();

		var htmlElement = this.WordControl.m_oEditor.HtmlElement;
		var fixPos = props.fixPosition;
		var curCoord = props.asc_getCellCoord();
		var startShapePos;

		if(fixPos && fixPos.slideId)
		{
			let oThumbnails = this.WordControl.Thumbnails;
			if(oThumbnails.m_bIsVisible)
			{
				let oPos = this.WordControl.Thumbnails.getSpecialPasteButtonCoords(fixPos.slideId);
				curCoord = new AscCommon.asc_CRect( oPos.X, oPos.Y, 1, 1 );
			}
			else
			{
				curCoord = new AscCommon.asc_CRect( -1, -1, 0, 0 );
			}
		}
		else
		{
			if(fixPos && fixPos.h && fixPos.w)
			{
				startShapePos = drawingDocument.ConvertCoordsToCursorWR(fixPos.x - fixPos.w, fixPos.y - fixPos.h, fixPos.pageNum);
			}

	

			var thumbnailsLeft = this.WordControl.m_oMainParent.AbsolutePosition.L* AscCommon.g_dKoef_mm_to_pix;
			if(!notesFocus && curCoord._x > htmlElement.width + thumbnailsLeft - AscCommon.specialPasteElemWidth)
			{
				if(startShapePos && startShapePos.X < htmlElement.width + thumbnailsLeft - AscCommon.specialPasteElemWidth)
				{
					curCoord._x = htmlElement.width - AscCommon.specialPasteElemWidth + thumbnailsLeft;
				}
				else
				{
					curCoord = new AscCommon.asc_CRect( -1, -1, 0, 0 );
				}
			}
		}

		if(curCoord)
		{
			props.asc_setCellCoord(curCoord);
		}

		this.sendEvent("asc_onShowSpecialPasteOptions", props);
	};

	asc_docs_api.prototype.asc_HideSpecialPasteButton = function()
	{
		if (window["NATIVE_EDITOR_ENJINE"])
		{
			return;
		}
		this.sendEvent("asc_onHideSpecialPasteOptions");
	};

	asc_docs_api.prototype.asc_UpdateSpecialPasteButton = function()
	{
		if (window["NATIVE_EDITOR_ENJINE"])
		{
			return;
		}
		var props = AscCommon.g_specialPasteHelper.buttonInfo;
		var presentation = editor.WordControl.m_oLogicDocument;
		var drawingDocument = presentation.DrawingDocument;
		var _coord, curCoord;

		var fixPos = props.fixPosition;
		var notesFocus = presentation.IsFocusOnNotes();
		if(props.shapeId)//при переходе между шейпами, скрываем значок спец.вставки
		{
			var targetDocContent = presentation ? presentation.Get_TargetDocContent() : null;
			if(targetDocContent && targetDocContent.Id === props.shapeId)
			{
				if(fixPos)
				{
					_coord = drawingDocument.ConvertCoordsToCursorWR(fixPos.x, fixPos.y, fixPos.pageNum);
					curCoord = new AscCommon.asc_CRect( _coord.X, _coord.Y, 0, 0 );
				}
			}
			else
			{
				curCoord = new AscCommon.asc_CRect( -1, -1, 0, 0 );
			}
		}
		else
		{
			if(true === notesFocus)
			{
				curCoord = new AscCommon.asc_CRect( -1, -1, 0, 0 );
			}
			else if(fixPos && fixPos.pageNum === presentation.CurPage)
			{
				_coord = drawingDocument.ConvertCoordsToCursorWR(fixPos.x, fixPos.y, fixPos.pageNum);
				curCoord = new AscCommon.asc_CRect( _coord.X, _coord.Y, 0, 0 );
			}
			else
			{
				curCoord = new AscCommon.asc_CRect( -1, -1, 0, 0 );
			}
		}

		if(curCoord)
		{
			props.asc_setCellCoord(curCoord);
		}

		this.asc_ShowSpecialPasteButton(props);
	};

	asc_docs_api.prototype.Share          = function()
	{

	};

	asc_docs_api.prototype.asc_CheckCopy = function(_clipboard /* CClipboardData */, _formats)
	{
		if (!this.WordControl.m_oLogicDocument)
		{
			var _text_object = (AscCommon.c_oAscClipboardDataFormat.Text & _formats) ? {Text : ""} : null;
			var _html_data   = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Copy(_text_object);

			//TEXT
			if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
			{
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, _text_object.Text);
			}
			//HTML
			if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
			{
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, _html_data);
			}
			return;
		}

		var sBase64 = null, _data;

		//TEXT
		if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
		{
			_data = this.WordControl.m_oLogicDocument.GetSelectedText(false);
			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, _data)
		}
		//HTML
		if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
		{
			var oCopyProcessor = new AscCommon.CopyProcessor(this);
			sBase64            = oCopyProcessor.Start();
			_data              = oCopyProcessor.getInnerHtml();

			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, _data)
		}
		//INTERNAL
		if (AscCommon.c_oAscClipboardDataFormat.Internal & _formats)
		{
			if (sBase64 === null)
			{
				var oCopyProcessor = new AscCommon.CopyProcessor(this);
				sBase64            = oCopyProcessor.Start();
			}

			_data = sBase64;
			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Internal, _data)
		}
	};
	asc_docs_api.prototype.getDemoManager = function() {
		return this.WordControl && this.WordControl.DemonstrationManager || null;
	};
	asc_docs_api.prototype.isSlideShow = function() {
		let oManager = this.getDemoManager();
		if(oManager && oManager.Mode) {
			return true;
		}
		return false;
	};

	asc_docs_api.prototype.asc_PasteData = function(_format, data1, data2, text_data, null_param, callback)
	{
	    if (!this.canEdit())
    	    return;

		//slide show
		if(this.isSlideShow()) {
			return;
		}

		if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content, null, false)) {
			this.WordControl.m_oLogicDocument.Create_NewHistoryPoint(AscDFH.historydescription_Document_PasteHotKey);

			window['AscCommon'].g_specialPasteHelper.Paste_Process_Start(arguments[5]);
			AscCommon.Editor_Paste_Exec(this, _format, data1, data2, text_data, undefined, callback);
		}
	};

	asc_docs_api.prototype.asc_SpecialPaste = function(props)
	{
		return AscCommon.g_specialPasteHelper.Special_Paste(props);
	};

	asc_docs_api.prototype.asc_SpecialPasteData = function(props)
	{
		if (AscCommon.CollaborativeEditing.Get_GlobalLock())
			return;

		var _logicDoc = this.WordControl.m_oLogicDocument;
		if (!_logicDoc)
			return;

		//TODO пересмотреть проверку лока и добавление новой точки(AscDFH.historydescription_Document_PasteHotKey)
		if (false === _logicDoc.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, false))
		{
			window['AscCommon'].g_specialPasteHelper.Paste_Process_Start();
			window['AscCommon'].g_specialPasteHelper.Special_Paste_Start();

			//undo previous action

            this.WordControl.m_oLogicDocument.TurnOffInterfaceEvents = true;
			this.WordControl.m_oLogicDocument.Document_Undo();
            this.WordControl.m_oLogicDocument.TurnOffInterfaceEvents = false;
			//if (!useCurrentPoint) {
			_logicDoc.Create_NewHistoryPoint(AscDFH.historydescription_Document_PasteHotKey);
			//}

			AscCommon.Editor_Paste_Exec(this, null, null, null, null, props);
		}
	};

	asc_docs_api.prototype.asc_IsFocus = function(bIsNaturalFocus)
	{
		var _ret = false;
		if (this.WordControl.IsFocus)
			_ret = true;
		if (_ret && bIsNaturalFocus && this.WordControl.TextBoxInputFocus)
			_ret = false;
		return _ret;
	};

	asc_docs_api.prototype.asc_SelectionCut = function()
	{
	    if (!this.canEdit())
            return;
		var _logicDoc = this.WordControl.m_oLogicDocument;
		if (!_logicDoc)
			return;
		_logicDoc.Remove(1, true, true);
	};

	asc_docs_api.prototype._onSaveCallbackInner = function()
	{
		var t = this;
		if (c_oAscCollaborativeMarksShowType.LastChanges === this.CollaborativeMarksShowType)
		{
			AscCommon.CollaborativeEditing.Clear_CollaborativeMarks();
		}

		// Принимаем чужие изменения
		AscCommon.CollaborativeEditing.Apply_Changes();

		this.CoAuthoringApi.onUnSaveLock = function()
		{
			t.CoAuthoringApi.onUnSaveLock = null;
			if (t.isForceSaveOnUserSave && t.IsUserSave) {
				t.forceSaveButtonContinue = t.forceSave();
			}
			if (t.forceSaveForm) {
				t.forceSaveForm();
			}
			// Выставляем, что документ не модифицирован
			t.CheckChangedDocument();
			t.canSave    = true;
			t.IsUserSave = false;
			if (!t.forceSaveButtonContinue) {
				t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.Save);
			}

			// Обновляем состояние возможности сохранения документа
			t._onUpdateDocumentCanSave();

			if (undefined !== window["AscDesktopEditor"])
			{
				window["AscDesktopEditor"]["OnSave"]();
			}

			if (t.canUnlockDocument) {
				t._unlockDocument();
			}
		};
		var CursorInfo                   = null;
		if (true === AscCommon.CollaborativeEditing.Is_Fast())
		{
			CursorInfo = History.Get_DocumentPositionBinary();
		}

		// Пересылаем свои изменения
		if (this.forceSaveUndoRequest)
		{
			AscCommon.CollaborativeEditing.Set_GlobalLock(false);
			AscCommon.CollaborativeEditing.Undo();
			this.forceSaveUndoRequest = false;
		}
		else
		{
			AscCommon.CollaborativeEditing.Send_Changes(this.IsUserSave, {
				"UserId"      : this.CoAuthoringApi.getUserConnectionId(),
				"UserShortId" : this.DocInfo.get_UserId(),
				"CursorInfo"  : CursorInfo
			}, undefined, true);
		}
	};
	asc_docs_api.prototype._autoSaveInner = function () {
		if (this.isSlideShow()) {
			return;
		}

		var _curTime = new Date();
		if (null === this.lastSaveTime) {
			this.lastSaveTime = _curTime;
		}

		if (AscCommon.CollaborativeEditing.Is_Fast() && !AscCommon.CollaborativeEditing.Is_SingleUser()) {
			this.WordControl.m_oLogicDocument.Continue_FastCollaborativeEditing();
		} else if (this.isLiveViewer()) {
			if (AscCommon.CollaborativeEditing.Have_OtherChanges()) {
				AscCommon.CollaborativeEditing.Apply_Changes();
			}
		} else {
			var _bIsWaitScheme = false;
			if (this.WordControl.m_oDrawingDocument &&
				(!this.WordControl.m_oDrawingDocument.TransitionSlide || !this.WordControl.m_oDrawingDocument.TransitionSlide.IsPlaying()) && History.Points &&
				History.Index >= 0 && History.Index < History.Points.length) {
				if ((_curTime - History.Points[History.Index].Time) < this.intervalWaitAutoSave) {
					_bIsWaitScheme = true;
				}
			}

			if (!_bIsWaitScheme) {
				var _interval = (AscCommon.CollaborativeEditing.m_nUseType <= 0) ? this.autoSaveGapSlow :
					this.autoSaveGapFast;

				if ((_curTime - this.lastSaveTime) > _interval) {
					if (History.Have_Changes(true) == true) {
						this.asc_Save(true);
					}
					this.lastSaveTime = _curTime;
				}
			}
		}
	};
	asc_docs_api.prototype._saveCheck = function() {
		return !this.isLongAction() && !(this.isSlideShow());
	};
	asc_docs_api.prototype._haveOtherChanges = function () {
		return AscCommon.CollaborativeEditing.Have_OtherChanges();
	};
	asc_docs_api.prototype.asc_DownloadAs = function(options)
	{
		if (this.isLongAction()) {
			return;
		}
		this.downloadAs(c_oAscAsyncAction.DownloadAs, options);
	};
	asc_docs_api.prototype.Resize                       = function()
	{
		if (false === this.bInit_word_control)
			return;
		this.WordControl.OnResize(false);
	};
	asc_docs_api.prototype.AddURL                       = function(url)
	{

	};
	asc_docs_api.prototype.Help                         = function()
	{

	};
	/*
	 idOption идентификатор дополнительного параметра, c_oAscAdvancedOptionsID.TXT.
	 option - какие свойства применить, пока массив. для TXT объект asc_CTextOptions(codepage)
	 exp:	asc_setAdvancedOptions(c_oAscAdvancedOptionsID.TXT, new Asc.asc_CTextOptions(1200) );
	 */
	asc_docs_api.prototype.asc_setAdvancedOptions       = function(idOption, option)
	{
        if (AscCommon.EncryptionWorker.asc_setAdvancedOptions(this, idOption, option))
            return;

		switch (idOption)
		{
			case c_oAscAdvancedOptionsID.DRM:
				this.currentPassword = option.asc_getPassword();
				var v = {
					"id": this.documentId,
					"userid": this.documentUserId,
					"format": this.documentFormat,
					"c": "reopen",
					"title": this.documentTitle,
					"password": option.asc_getPassword(),
					"nobase64": true
				};

				sendCommand(this, null, v);
				break;
		}
	};
	asc_docs_api.prototype.startGetDocInfo              = function()
	{
		/*
		 Возвращаем объект следующего вида:
		 {
		 PageCount: 12,
		 WordsCount: 2321,
		 ParagraphCount: 45,
		 SymbolsCount: 232345,
		 SymbolsWSCount: 34356
		 }
		 */
		this.sync_GetDocInfoStartCallback();

		this.WordControl.m_oLogicDocument.Statistics_Start();
	};
	asc_docs_api.prototype.stopGetDocInfo               = function()
	{
		this.sync_GetDocInfoStopCallback();
		this.WordControl.m_oLogicDocument.Statistics_Stop();
	};
	asc_docs_api.prototype.sync_CanUndoCallback         = function(canUndo)
	{
		if (!this.canUndoRedoByRestrictions())
			canUndo = false;
		
		this.sendEvent("asc_onCanUndo", canUndo);
	};
	asc_docs_api.prototype.sync_CanRedoCallback         = function(canRedo)
	{
		if (!this.canUndoRedoByRestrictions()
			|| (true === AscCommon.CollaborativeEditing.Is_Fast() && true !== AscCommon.CollaborativeEditing.Is_SingleUser()))
			canRedo = false;

		this.sendEvent("asc_onCanRedo", canRedo);
	};


	/*callbacks*/
	/*asc_docs_api.prototype.sync_CursorLockCallBack = function(isLock){
	 this.sendEvent("asc_onCursorLock",isLock);
	 }*/
	asc_docs_api.prototype.sync_UndoCallBack       = function()
	{
		this.sendEvent("asc_onUndo");
	};
	asc_docs_api.prototype.sync_RedoCallBack       = function()
	{
		this.sendEvent("asc_onRedo");
	};
	asc_docs_api.prototype.sync_CopyCallBack       = function()
	{
		this.sendEvent("asc_onCopy");
	};
	asc_docs_api.prototype.sync_CutCallBack        = function()
	{
		this.sendEvent("asc_onCut");
	};
	asc_docs_api.prototype.sync_PasteCallBack      = function()
	{
		this.sendEvent("asc_onPaste");
	};
	asc_docs_api.prototype.sync_ShareCallBack      = function()
	{
		this.sendEvent("asc_onShare");
	};
	asc_docs_api.prototype.sync_SaveCallBack       = function()
	{
		this.sendEvent("asc_onSave");
	};
	asc_docs_api.prototype.sync_DownloadAsCallBack = function()
	{
		this.sendEvent("asc_onDownload");
	};

	asc_docs_api.prototype.sync_AddURLCallback  = function()
	{
		this.sendEvent("asc_onAddURL");
	};
	asc_docs_api.prototype.sync_ErrorCallback   = function(errorID, errorLevel)
	{
		this.sendEvent("asc_onError", errorID, errorLevel);
	};
	asc_docs_api.prototype.sync_HelpCallback    = function(url)
	{
		this.sendEvent("asc_onHelp", url);
	};
	asc_docs_api.prototype.sync_UpdateZoom      = function(zoom)
	{
		this.sendEvent("asc_onZoom", zoom);
	};
	asc_docs_api.prototype.ClearPropObjCallback = function(prop)
	{//колбэк предшествующий приходу свойств объекта, prop а всякий случай

		this.sendEvent("asc_onClearPropObj", prop);
	};

	// mobile version methods:
	asc_docs_api.prototype.asc_GetDefaultTableStyles = function()
	{
		var logicDoc = this.WordControl.m_oLogicDocument;
		if (!logicDoc)
			return;
		if (logicDoc.CurPage >= logicDoc.GetSlidesCount())
			return;
		if (logicDoc.GetSlidesCount() === 0)
		{
			logicDoc.addNextSlide();
		}
		this.WordControl.m_oDrawingDocument.CheckTableStylesDefault();
	};

	asc_docs_api.prototype.CollectHeaders                  = function()
	{
		this.sync_ReturnHeadersCallback([]);
	};
	asc_docs_api.prototype.GetActiveHeader                 = function()
	{

	};
	asc_docs_api.prototype.sync_ChangeActiveHeaderCallback = function(position, header)
	{
		this.sendEvent("asc_onChangeActiveHeader", position, new Asc.CHeader(header));
	};
	asc_docs_api.prototype.sync_ReturnHeadersCallback      = function(headers)
	{
		var _headers = [];
		for (var i = 0; i < headers.length; i++)
		{
			_headers[i] = new Asc.CHeader(headers[i]);
		}

		this.sendEvent("asc_onReturnHeaders", _headers);
	};
	/*----------------------------------------------------------------*/
	/*functions for working with search*/
	/*
	 структура поиска, предварительно, выглядит так
	 {
	 text: "...<b>слово поиска</b>...",
	 pageNumber: 0, //содержит номер страницы, где находится искомая последовательность
	 X: 0,//координаты по OX начала последовательности на данной страницы
	 Y: 0//координаты по OY начала последовательности на данной страницы
	 }
	 */
	asc_docs_api.prototype.startSearchText = function(what)
	{// "what" means word(s) what we search
		this._searchCur = 0;
		this.sync_SearchStartCallback();

		if (null != this.WordControl.m_oLogicDocument)
		{

		}
		else
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.StartSearch(what);
	};

	asc_docs_api.prototype.goToNextSearchResult = function()
	{
		this.WordControl.m_oLogicDocument.goToNextSearchResult();
	};


	asc_docs_api.prototype.gotoSearchResultText = function(navigator)
	{//переход к результату.

		this.WordControl.m_oDrawingDocument.CurrentSearchNavi = navigator;
		this.WordControl.ToSearchResult();
	};
	asc_docs_api.prototype.stopSearchText       = function()
	{
		this.sync_SearchStopCallback();

		this.WordControl.m_oLogicDocument.Search_Stop();
	};

	asc_docs_api.prototype.sync_ReplaceAllCallback = function(ReplaceCount, OverallCount)
	{
		this.sendEvent("asc_onReplaceAll", OverallCount, ReplaceCount);
	};

	asc_docs_api.prototype.asc_replaceText = function(oProps, replaceWith, isReplaceAll)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oLogicDocument.Search(oProps);

		if (true === isReplaceAll)
			this.WordControl.m_oLogicDocument.ReplaceSearchElement(replaceWith, true, -1);
		else
		{
			var CurId      = this.WordControl.m_oLogicDocument.SearchEngine.CurId;
			var bDirection = this.WordControl.m_oLogicDocument.SearchEngine.Direction;
			if (-1 != CurId)
				this.WordControl.m_oLogicDocument.ReplaceSearchElement(replaceWith, false, CurId);

			var Id = this.WordControl.m_oLogicDocument.GetSearchElementId(bDirection);

			if (null != Id)
			{
				this.WordControl.m_oLogicDocument.SelectSearchElement(Id);
				return true;
			}
			else
			{
				this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
				this.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
				this.WordControl.OnUpdateOverlay();
			}

			return false;
		}
	};

	asc_docs_api.prototype.asc_findText = function(oProps, isNext, callback)
	{
		var SearchEngine = editor.WordControl.m_oLogicDocument.Search(oProps);

		var Id = this.WordControl.m_oLogicDocument.GetSearchElementId(isNext);

		if (null != Id)
			this.WordControl.m_oLogicDocument.SelectSearchElement(Id);

		var result = SearchEngine.Count;
		if (callback)
			callback(result);
		return result;
	};
	asc_docs_api.prototype.asc_endFindText = function()
	{
		return this.WordControl.m_oLogicDocument.ClearSearch();
	};

	// returns: CSearchResult
	asc_docs_api.prototype.sync_SearchFoundCallback = function(obj)
	{
		this.sendEvent("asc_onSearchFound", new CSearchResult(obj));
	};
	asc_docs_api.prototype.sync_SearchStartCallback = function()
	{
		this.sendEvent("asc_onSearchStart");
	};
	asc_docs_api.prototype.sync_SearchStopCallback  = function()
	{
		this.sendEvent("asc_onSearchStop");
	};
	asc_docs_api.prototype.sync_SearchEndCallback   = function()
	{
		this.sendEvent("asc_onSearchEnd");
	};
	asc_docs_api.prototype.asc_StartTextAroundSearch = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument || !oLogicDocument.SearchEngine)
			return;

		oLogicDocument.SearchEngine.StartTextAround();
	};
	asc_docs_api.prototype.asc_SelectSearchElement = function(sId)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument || !oLogicDocument.SearchEngine)
			return;

		oLogicDocument.SelectSearchElement(sId);
	};
	/*----------------------------------------------------------------*/
	/*functions for working with font*/
	/*setters*/
	asc_docs_api.prototype.put_TextPrFontName         = function(name)
	{
		var loader   = AscCommon.g_font_loader;
		var fontinfo = AscFonts.g_fontApplication.GetFontInfo(name);
		var isasync  = loader.LoadFont(fontinfo);
		if (false === isasync)
		{
			if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
			{
				this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
				this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
					FontFamily : {
						Name  : name,
						Index : -1
					}
				}), false);
				this.WordControl.m_oLogicDocument.FinalizeAction()
			}
		}
	};
	asc_docs_api.prototype.put_TextPrFontSize         = function(size)
	{
		if(!AscFormat.isRealNumber(size) || size < 1) {
			return;
		}
		if (this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({FontSize : Math.min(size, 300)}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
			// для мобильной версии это важно
			if (this.isMobileVersion)
				this.UpdateInterfaceState();
		}
	};
    asc_docs_api.prototype.put_TextPrLang = function(value)
    {
        if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
        {
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTextLang);
            this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Lang : {Val : value}}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();

            this.WordControl.m_oLogicDocument.Spelling.CheckCurrentParagraph();

            //if (true === this.isMarkerFormat)
            //    this.sync_MarkerFormatCallback(false);
        }
    };

	asc_docs_api.prototype.put_TextPrBold             = function(value)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Bold : value}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction()
		}
	};
	asc_docs_api.prototype.put_TextPrItalic           = function(value)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Italic : value}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextPrUnderline        = function(value)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Underline : value}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextPrStrikeout        = function(value)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
                Strikeout  : value,
                DStrikeout : false
            }), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_PrLineSpacing          = function(Type, Value)
	{
		this.WordControl.m_oLogicDocument.SetParagraphSpacing({LineRule : Type, Line : Value});
	};
	asc_docs_api.prototype.put_LineSpacingBeforeAfter = function(type, value)//"type == 0" means "Before", "type == 1" means "After"
	{
		switch (type)
		{
			case 0:
				this.WordControl.m_oLogicDocument.SetParagraphSpacing({Before : value});
				break;
			case 1:
				this.WordControl.m_oLogicDocument.SetParagraphSpacing({After : value});
				break;
		}
	};
	asc_docs_api.prototype.FontSizeIn                 = function()
	{
		this.WordControl.m_oLogicDocument.IncreaseDecreaseFontSize(true);
	};
	asc_docs_api.prototype.FontSizeOut                = function()
	{
		this.WordControl.m_oLogicDocument.IncreaseDecreaseFontSize(false);
	};

	/*callbacks*/
	asc_docs_api.prototype.sync_BoldCallBack             = function(isBold)
	{
		this.sendEvent("asc_onBold", isBold);
	};
	asc_docs_api.prototype.sync_ItalicCallBack           = function(isItalic)
	{
		this.sendEvent("asc_onItalic", isItalic);
	};
	asc_docs_api.prototype.sync_UnderlineCallBack        = function(isUnderline)
	{
		this.sendEvent("asc_onUnderline", isUnderline);
	};
	asc_docs_api.prototype.sync_StrikeoutCallBack        = function(isStrikeout)
	{
		this.sendEvent("asc_onStrikeout", isStrikeout);
	};
	asc_docs_api.prototype.sync_TextPrFontFamilyCallBack = function(FontFamily)
	{
		this.sendEvent("asc_onFontFamily", new AscCommon.asc_CTextFontFamily(FontFamily));
	};
	asc_docs_api.prototype.sync_TextPrFontSizeCallBack   = function(FontSize)
	{
		this.sendEvent("asc_onFontSize", FontSize);
	};
	asc_docs_api.prototype.sync_PrLineSpacingCallBack    = function(LineSpacing)
	{
		this.sendEvent("asc_onLineSpacing", new AscCommon.asc_CParagraphSpacing(LineSpacing));
	};

	asc_docs_api.prototype.sync_InitEditorThemes      = function(gui_editor_themes, gui_document_themes)
	{
		this._gui_editor_themes   = gui_editor_themes;
		this._gui_document_themes = gui_document_themes;
		if (!this.isViewMode) {
			this.sendEvent("asc_onInitEditorStyles", [gui_editor_themes, gui_document_themes]);
		}
	};
	asc_docs_api.prototype.sync_InitEditorTableStyles = function(styles)
	{
		if (!this.isViewMode) {
			this.sendEvent("asc_onInitTableTemplates", styles);
		}
	};
	asc_docs_api.prototype.asc_getTableStylesPreviews    = function(bUseDefault, arrIds)
	{
		let oDrawingDocument = this.WordControl ? this.WordControl.m_oDrawingDocument : null;
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oDrawingDocument || !oLogicDocument)
			return [];

		this.private_CheckTableStylesPreviewGenerator();

		let arrPreviews = [];

		if (arrIds && arrIds.length)
		{
			var aAllStyles = oLogicDocument.GetAllTableStyles();
			let oTableLookOld = this.TableStylesPreviewGenerator.TableLook;
			this.TableStylesPreviewGenerator.TableLook    = this.WordControl.m_oDrawingDocument.GetTableLook(bUseDefault);
			for (let nIndex = 0, nCount = arrIds.length; nIndex < nCount; ++nIndex)
			{
				let oStyle = null;
				for(let nStyle = 0; nStyle < aAllStyles.length; ++nStyle)
				{
					if(aAllStyles[nStyle].GetId() === arrIds[nIndex])
					{
						oStyle = aAllStyles[nStyle];
						break;
					}
				}
				if(oStyle)
				{
					let oPreview = this.TableStylesPreviewGenerator.GetPreview(oStyle);
					if (oPreview)
						arrPreviews.push(oPreview);
				}
			}
			this.TableStylesPreviewGenerator.TableLook = oTableLookOld;
		}
		else
		{
			arrPreviews = this.TableStylesPreviewGenerator.GetAllPreviews(bUseDefault);
		}

		return arrPreviews;
	};
	asc_docs_api.prototype.asc_generateTableStylesPreviews = function(bUseDefault)
	{
		this.private_CheckTableStylesPreviewGenerator();
		this.TableStylesPreviewGenerator.Begin(bUseDefault);
	};
	asc_docs_api.prototype.private_CheckTableStylesPreviewGenerator = function()
	{
		if (!this.TableStylesPreviewGenerator)
			this.TableStylesPreviewGenerator = new AscCommon.CTableStylesPreviewGenerator(this.private_GetLogicDocument())
	};
	/**
	 * @returns {?CPresentation}
	 */
	asc_docs_api.prototype.private_GetLogicDocument = function()
	{
		if (!this.WordControl || !this.WordControl.m_oLogicDocument)
			return null;

		return this.WordControl.m_oLogicDocument;
	};

	/*----------------------------------------------------------------*/
	/*functions for working with paragraph*/
	/*setters*/
	// Right = 0; Left = 1; Center = 2; Justify = 3; or using enum that written above

	/* структура для параграфа
	 Ind :
	 {
	 Left      : 0,                    // Левый отступ
	 Right     : 0,                    // Правый отступ
	 FirstLine : 0                     // Первая строка
	 }
	 Spacing :
	 {
	 Line     : 1.15,                  // Расстояние между строками внутри абзаца
	 LineRule : linerule_Auto,         // Тип расстрояния между строками
	 Before   : 0,                     // Дополнительное расстояние до абзаца
	 After    : 10 * g_dKoef_pt_to_mm  // Дополнительное расстояние после абзаца
	 },
	 KeepLines : false,                    // переносить параграф на новую страницу,
	 // если на текущей он целиком не убирается
	 PageBreakBefore : false
	 */

	asc_docs_api.prototype.paraApply = function(Props)
	{
		var _presentation = editor.WordControl.m_oLogicDocument;
		var graphicObjects = _presentation.GetCurrentController();
		if (graphicObjects)
		{
			var sLoadFont = null, sLoadText = null;
			var fCallback = function()
			{
				graphicObjects.paraApplyCallback(Props);
				_presentation.Document_UpdateInterfaceState();
			};
			var oBullet = Props.asc_getBullet();
			if(oBullet)
			{
				sLoadFont = oBullet.asc_getFont();
				sLoadText = oBullet.asc_getSymbol();
			}

			if(typeof sLoadFont === "string" && sLoadFont.length > 0
				&& typeof sLoadText === "string" && sLoadText.length > 0)
			{
				var loader   = AscCommon.g_font_loader;
				var fontinfo = AscFonts.g_fontApplication.GetFontInfo(sLoadFont);
				var isasync  = loader.LoadFont(fontinfo);
				if (false === isasync)
				{
					AscFonts.FontPickerByCharacter.checkText(sLoadText, this, function () {
						graphicObjects.checkSelectedObjectsAndCallback(fCallback, [], false, AscDFH.historydescription_Presentation_ParaApply);
					});
				}
				else
				{
					this.asyncMethodCallback = function()
					{
						AscFonts.FontPickerByCharacter.checkText(sLoadText, this, function () {
							graphicObjects.checkSelectedObjectsAndCallback(fCallback, [], false, AscDFH.historydescription_Presentation_ParaApply);
						});
					}
				}
			}
			else
			{
				graphicObjects.checkSelectedObjectsAndCallback(fCallback, [], false, AscDFH.historydescription_Presentation_ParaApply);
			}
		}
	};

	asc_docs_api.prototype.put_PrAlign        = function(value)
	{
		this.WordControl.m_oLogicDocument.SetParagraphAlign(value);
	};
	// 0- baseline, 2-subscript, 1-superscript
	asc_docs_api.prototype.put_TextPrBaseline = function(value)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({VertAlign : value}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	/* 	Маркированный список Type = 0
	 нет         - SubType = -1
	 черная точка - SubType = 1
	 круг         - SubType = 2
	 квадрат      - SubType = 3
	 картинка     - SubType = -1
	 4 ромба      - SubType = 4
	 ч/б стрелка  - SubType = 5
	 галка        - SubType = 6

	 Нумерованный список Type = 1
	 нет - SubType = -1
	 1.  - SubType = 1
	 1)  - SubType = 2
	 I.  - SubType = 3
	 A.  - SubType = 4
	 a)  - SubType = 5
	 a.  - SubType = 6
	 i.  - SubType = 7

	 Многоуровневый список Type = 2
	 нет            - SubType = -1
	 1)a)i)        - SubType = 1
	 1.1.1         - SubType = 2
	 маркированный - SubType = 3
	 */
	asc_docs_api.prototype.put_ImageBulletFromFile = function () {
		this.asc_addImage({isImageBullet: true});
	}
	asc_docs_api.prototype.put_ListType = function(type, subtype, custom)
	{
		var blipUrl = custom && custom.imageId;
		if (blipUrl) {
			var that = this;
			var checkImageUrlFromServer;
			var localUrl = AscCommon.g_oDocumentUrls.getLocal(blipUrl);
			var fullUrl = AscCommon.g_oDocumentUrls.getUrl(blipUrl);
			if (fullUrl) {
				checkImageUrlFromServer = fullUrl;
			} else if (localUrl) {
				checkImageUrlFromServer = blipUrl;
			}
			if (checkImageUrlFromServer) {
				blipUrl = checkImageUrlFromServer;
				custom.imageId = blipUrl;
				var isImageNotAttendInImageLoader = !this.ImageLoader.map_image_index[blipUrl];
				if (isImageNotAttendInImageLoader) {
					var tryToSetImageBulletAgain = function () {
						that.put_ListType(type, subtype, custom);
					}
					this.ImageLoader.LoadImagesWithCallback([blipUrl], tryToSetImageBulletAgain);
					return;
				}
			} else {
				var changeBlipFillUrlToLocalAndTrySetImageBulletAgain = function (data) {
					var uploadImageUrl = data[0].url;
					custom.imageId = uploadImageUrl;
					that.put_ListType(type, subtype, custom);
				}
				AscCommon.sendImgUrls(this, [blipUrl], changeBlipFillUrlToLocalAndTrySetImageBulletAgain, false, custom.token);
				return;
			}
		}
		var oPresentation = this.WordControl.m_oLogicDocument;
		var NumberInfo =
		{
			Type     : type,
			SubType  : subtype,
			Custom     : custom
		};
		var oBullet = AscFormat.fGetPresentationBulletByNumInfo(NumberInfo);
		var sBullet = oBullet.asc_getSymbol();
		var fCallback = function ()
		{
			oPresentation.SetParagraphNumbering(oBullet);
		};
		if(typeof sBullet === "string" && sBullet.length > 0)
		{
			AscFonts.FontPickerByCharacter.checkText(sBullet, this, fCallback);
		}
		else
		{
			fCallback();
		}
	};

	asc_docs_api.prototype.put_ShowSnapLines = function(isShow)
	{
		this.ShowSmartGuides = isShow;
	};
	asc_docs_api.prototype.get_ShowSnapLines = function()
	{
		return this.ShowSmartGuides;
	};

	asc_docs_api.prototype.asc_setShowSmartGuides = function(isShow)
	{
		this.ShowSmartGuides = isShow;
	};
	asc_docs_api.prototype.asc_getShowSmartGuides = function()
	{
		return this.ShowSmartGuides;
	};

	asc_docs_api.prototype.asc_setShowGuides = function(isShow)
	{
		if(this.ShowGuides === isShow)
		{
			return;
		}
		let oPresentation = this.WordControl.m_oLogicDocument;
		if(isShow)
		{
			oPresentation.checkEmptyGuides();
		}
		this.ShowGuides = isShow;
		oPresentation.RedrawCurSlide();
	};
	asc_docs_api.prototype.asc_getShowGuides = function()
	{
		return this.ShowGuides;
	};
	asc_docs_api.prototype.asc_getGuidesCount = function()
	{
		let oPresentation = this.private_GetLogicDocument();
		if(oPresentation)
		{
			return oPresentation.getGuidesCount();
		}
		return 0;
	};

	asc_docs_api.prototype.asc_setShowGridlines = function(isShow)
	{
		if(this.ShowGridlines === isShow)
		{
			return;
		}
		this.ShowGridlines = isShow;
		let oPresentation = this.WordControl.m_oLogicDocument;
		oPresentation.RedrawCurSlide();
	};
	asc_docs_api.prototype.asc_getShowGridlines = function()
	{
		return this.ShowGridlines;
	};
	asc_docs_api.prototype.asc_setGridSpacing = function(nSpacing)
	{
		if(!AscFormat.isRealNumber(nSpacing)) {
			return;
		}
		this.WordControl.m_oLogicDocument.setGridSpacing(nSpacing);
	};
	asc_docs_api.prototype.asc_getGridSpacing = function()
	{
		return this.WordControl.m_oLogicDocument.getGridSpacing();
	};
	asc_docs_api.prototype.asc_setSnapToGrid = function(bVal)
	{
		this.WordControl.m_oLogicDocument.setSnapToGrid(bVal);
	};
	asc_docs_api.prototype.asc_getSnapToGrid = function()
	{
		return this.WordControl.m_oLogicDocument.isSnapToGrid();
	};
	asc_docs_api.prototype.asc_addHorizontalGuide = function()
	{
		return this.WordControl.m_oLogicDocument.addHorizontalGuide();
	};
	asc_docs_api.prototype.asc_addVerticalGuide = function()
	{
		return this.WordControl.m_oLogicDocument.addVerticalGuide();
	};
	asc_docs_api.prototype.asc_canClearGuides = function()
	{
		return this.WordControl.m_oLogicDocument.canClearGuides();
	};
	asc_docs_api.prototype.asc_clearGuides = function()
	{
		return this.WordControl.m_oLogicDocument.clearGuides();
	};
	asc_docs_api.prototype.asc_deleteGuide = function(sId)
	{
		return this.WordControl.m_oLogicDocument.deleteGuide(sId);
	};
	asc_docs_api.prototype.setSlideObjectName = function(sName, nType)
	{
		let oPresentation = this.getLogicDocument();
		if(!oPresentation) return;
		let aSelectedSlides = oPresentation.GetSelectedSlideObjects();
		if(aSelectedSlides.length === 1) {
			let oSlideLikeObject = aSelectedSlides[0];
			if(oSlideLikeObject.getObjectType() === nType) {
				//if (!oPresentation.IsSelectionLocked(AscCommon.changestype_Timing)) {
				AscCommon.History.Create_NewPoint(0);
				oSlideLikeObject.setName(sName);
				oPresentation.Recalculate();
				//}
			}
		}
	};
	asc_docs_api.prototype.asc_SetMasterName = function(sName)
	{
		this.setSlideObjectName(sName, AscDFH.historyitem_type_SlideMaster);
	};
	asc_docs_api.prototype.asc_SetLayoutName = function(sName)
	{
		this.setSlideObjectName(sName, AscDFH.historyitem_type_SlideLayout);
	};

	asc_docs_api.prototype.put_ShowParaMarks      = function(isShow)
	{
		this.ShowParaMarks = isShow;
		this.WordControl.OnRePaintAttack();
		return this.ShowParaMarks;
	};
	asc_docs_api.prototype.get_ShowParaMarks      = function()
	{
		return this.ShowParaMarks;
	};
	asc_docs_api.prototype.put_ShowTableEmptyLine = function(isShow)
	{
		this.isShowTableEmptyLine = isShow;
		this.WordControl.OnRePaintAttack();

		return this.isShowTableEmptyLine;
	};
	asc_docs_api.prototype.get_ShowTableEmptyLine = function()
	{
		return this.isShowTableEmptyLine;
	};

	asc_docs_api.prototype.ShapeApply = function(prop)
	{
		// нужно определить, картинка это или нет
		var image_url = "";
		var sToken = undefined;
		prop.Width    = prop.w;
		prop.Height   = prop.h;

		var bShapeTexture = true;
		if (prop.fill != null)
		{
			if (prop.fill.fill != null && prop.fill.type == c_oAscFill.FILL_TYPE_BLIP)
			{
				image_url = prop.fill.fill.asc_getUrl();
				sToken = prop.fill.fill.token;
				var _tx_id = prop.fill.fill.asc_getTextureId();
				if (null != _tx_id && 0 <= _tx_id && _tx_id < AscCommon.g_oUserTexturePresets.length)
				{
					image_url = AscCommon.g_oUserTexturePresets[_tx_id];
				}
			}
		}
		var oFill;
		if (prop.textArtProperties)
		{
			oFill = prop.textArtProperties.asc_getFill();
			if (oFill && oFill.fill != null && oFill.type == c_oAscFill.FILL_TYPE_BLIP)
			{
				image_url = oFill.fill.asc_getUrl();
				sToken = oFill.fill.token;
				var _tx_id = oFill.fill.asc_getTextureId();
				if (null != _tx_id && 0 <= _tx_id && _tx_id < AscCommon.g_oUserTexturePresets.length)
				{
					image_url = AscCommon.g_oUserTexturePresets[_tx_id];
				}
				bShapeTexture = false;
			}
		}
		if (!AscCommon.isNullOrEmptyString(image_url))
		{
			var sImageUrl = null;
			if (!g_oDocumentUrls.getImageLocal(image_url))
			{
				sImageUrl = image_url;
			}
			var oApi           = this;
			var fApplyCallback = function()
			{
				var _image   = oApi.ImageLoader.LoadImage(image_url, 1);
				var srcLocal = g_oDocumentUrls.getImageLocal(image_url);
				if (srcLocal)
				{
					image_url = srcLocal;
				}
				if (bShapeTexture)
				{
					prop.fill.fill.asc_putUrl(image_url); // erase documentUrl
				}
				else
				{
					oFill.fill.asc_putUrl(image_url);
				}
				if (null != _image || window["NATIVE_EDITOR_ENJINE"])
				{
					oApi.WordControl.m_oLogicDocument.ShapeApply(prop);
					if (bShapeTexture)
					{
						oApi.WordControl.m_oDrawingDocument.DrawImageTextureFillShape(image_url);
					}
					else
					{
						oApi.WordControl.m_oDrawingDocument.DrawImageTextureFillTextArt(image_url);
					}
				}
				else
				{
					oApi.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
					var oProp                 = prop;
					oApi.asyncImageEndLoaded2 = function(_image)
					{
						oApi.WordControl.m_oLogicDocument.ShapeApply(oProp);
						oApi.WordControl.m_oDrawingDocument.DrawImageTextureFillShape(image_url);
						oApi.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
						oApi.asyncImageEndLoaded2 = null;
					}
				}
			};
			if (!sImageUrl)
			{
				fApplyCallback();
			}
			else
			{

				if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]())
				{
					image_url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageUrl);
					image_url = g_oDocumentUrls.getImageUrl(image_url);
					fApplyCallback();
					return;
				}
                AscCommon.sendImgUrls(this, [sImageUrl], function(data) {

                    if (data && data[0] && data[0].url !== "error")
                    {
                        image_url = data[0].url;
                        fApplyCallback();
                    }

                }, undefined, sToken);
			}
		}
		else
		{
			if (!this.noCreatePoint || this.exucuteHistory)
			{
				if (!this.noCreatePoint && !this.exucuteHistory && this.exucuteHistoryEnd)
				{
					if (-1 !== this.nCurPointItemsLength)
					{
						History.UndoLastPoint();
						var slide = this.WordControl.m_oLogicDocument.GetCurrentSlide();
						slide.graphicObjects.applyDrawingProps(prop);
						this.WordControl.m_oLogicDocument.Recalculate();
						this.WordControl.m_oDrawingDocument.OnRecalculateSlide(this.WordControl.m_oLogicDocument.CurPage);
						this.WordControl.m_oDrawingDocument.OnEndRecalculate();
					}
					else
					{
						this.WordControl.m_oLogicDocument.ShapeApply(prop);
					}
					this.exucuteHistoryEnd    = false;
					this.nCurPointItemsLength = -1;
				}
				else
				{
					this.WordControl.m_oLogicDocument.ShapeApply(prop);
				}
				if (this.exucuteHistory)
				{
					var oPoint = History.Points[History.Index];
					if (oPoint)
					{
						this.nCurPointItemsLength = oPoint.Items.length;
					}
					else
					{
						this.nCurPointItemsLength = -1;
					}
					this.exucuteHistory = false;
				}
				if(this.exucuteHistoryEnd)
				{
					this.exucuteHistoryEnd = false;
				}
			}
			else
			{
				if (this.WordControl.m_oLogicDocument.GetCurrentSlide())
				{
					if (-1 !== this.nCurPointItemsLength)
					{
						History.UndoLastPoint();
						var slide = this.WordControl.m_oLogicDocument.GetCurrentSlide();
						slide.graphicObjects.applyDrawingProps(prop);
						this.WordControl.m_oLogicDocument.Recalculate();
						this.WordControl.m_oDrawingDocument.OnRecalculateSlide(this.WordControl.m_oLogicDocument.CurPage);
						this.WordControl.m_oDrawingDocument.OnEndRecalculate();
					}
					else
					{
						this.WordControl.m_oLogicDocument.ShapeApply(prop);
						var oPoint = History.Points[History.Index];
						if (oPoint)
						{
							this.nCurPointItemsLength = oPoint.Items.length;
						}
						else
						{
							this.nCurPointItemsLength = -1;
						}
					}
				}
			}
            this.exucuteHistoryEnd = false;
		}
	};

	asc_docs_api.prototype.setStartPointHistory = function()
	{
		this.noCreatePoint  = true;
		this.exucuteHistory = true;
		this.incrementCounterLongAction();
		this.WordControl.m_oLogicDocument.TurnOff_InterfaceEvents();
	};
	asc_docs_api.prototype.setEndPointHistory = function()
	{
		this.noCreatePoint     = false;
		this.exucuteHistoryEnd = true;
		this.decrementCounterLongAction();
		this.WordControl.m_oLogicDocument.TurnOn_InterfaceEvents();
	};
	asc_docs_api.prototype.SetSlideProps = function(prop)
	{
		if (null == prop)
			return;

		const oLogicDocument = this.getLogicDocument();
		if (!oLogicDocument) {
			return;
		}

		const bApplyBackgroundToAll = prop.get_ApplyBackgroundToAll();
		if (bApplyBackgroundToAll) {
			oLogicDocument.applySlideBackgroundToAll();
			return;
		}

		const arr_ind    = oLogicDocument.GetSelectedSlides();
		const bResetBackground = prop.get_ResetBackground();
		if (bResetBackground) {
			oLogicDocument.resetSlideBackground(arr_ind);
			return;
		}

		const bNewShowMasterShapes = prop.get_ShowMasterSp();
		if (bNewShowMasterShapes !== null) {
			oLogicDocument.setShowMasterSp(bNewShowMasterShapes, arr_ind);
		}

		const _back_fill = prop.get_background();
		if (_back_fill) {
			if (_back_fill.asc_getType() == c_oAscFill.FILL_TYPE_NOFILL)
			{
				var bg       = new AscFormat.CBg();
				bg.bgPr      = new AscFormat.CBgPr();
				bg.bgPr.Fill = AscFormat.CorrectUniFill(_back_fill, null, 0);

				oLogicDocument.changeBackground(bg, arr_ind);
				return;
			}

			var _old_fill = oLogicDocument.GetCurrentSlide().backgroundFill;
			if (AscCommon.isRealObject(_old_fill))
				_old_fill = _old_fill.createDuplicate();
			var bg        = new AscFormat.CBg();
			bg.bgPr       = new AscFormat.CBgPr();
			bg.bgPr.Fill  = AscFormat.CorrectUniFill(_back_fill, _old_fill, 0);
			var image_url = "", sToken = undefined;
			if (_back_fill.asc_getType() == c_oAscFill.FILL_TYPE_BLIP && _back_fill.fill && typeof _back_fill.fill.url === "string" && _back_fill.fill.url.length > 0)
			{
				image_url = _back_fill.fill.url;
				sToken = _back_fill.fill.token;
			}
			if (image_url != "")
			{
				var sImageUrl = null;
				if (!g_oDocumentUrls.getImageLocal(image_url))
				{
					sImageUrl = image_url;
				}
				var oApi           = this;
				var fApplyCallback = function()
				{

					var _image   = oApi.ImageLoader.LoadImage(image_url, 1);
					var srcLocal = g_oDocumentUrls.getImageLocal(image_url);
					if (srcLocal)
					{
						image_url                       = srcLocal;
						bg.bgPr.Fill.fill.RasterImageId = image_url; // erase documentUrl
					}

					if (null != _image || window["NATIVE_EDITOR_ENJINE"])
					{
						if (bg.bgPr.Fill != null && bg.bgPr.Fill.fill != null && bg.bgPr.Fill.fill.type == c_oAscFill.FILL_TYPE_BLIP)
						{
							oApi.WordControl.m_oDrawingDocument.DrawImageTextureFillSlide(bg.bgPr.Fill.fill.RasterImageId);
						}

						oLogicDocument.changeBackground(bg, arr_ind);
					}
					else
					{
						oApi.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);

						var oProp                 = prop;
						oApi.asyncImageEndLoaded2 = function(_image)
						{
							if (bg.bgPr.Fill != null && bg.bgPr.Fill.fill != null && bg.bgPr.Fill.fill.type == c_oAscFill.FILL_TYPE_BLIP)
							{
								oApi.WordControl.m_oDrawingDocument.DrawImageTextureFillSlide(bg.bgPr.Fill.fill.RasterImageId);
							}

							oLogicDocument.changeBackground(bg, arr_ind);
							oApi.asyncImageEndLoaded2 = null;

							oApi.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
						}
					}
				};
				if (!sImageUrl)
				{
					fApplyCallback();
				}
				else
				{
					if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]())
					{
						image_url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageUrl);
						image_url = g_oDocumentUrls.getImageUrl(image_url);
						fApplyCallback();
						return;
					}
					AscCommon.sendImgUrls(this, [sImageUrl], function(data) {

						if (data && data[0] && data[0].url !== "error")
						{
							image_url = data[0].url;
							fApplyCallback();
						}

					}, undefined, sToken);
				}
			}
			else
			{
				if (bg.bgPr.Fill != null && bg.bgPr.Fill.fill != null && bg.bgPr.Fill.fill.type == c_oAscFill.FILL_TYPE_BLIP)
				{
					this.WordControl.m_oDrawingDocument.DrawImageTextureFillSlide(bg.bgPr.Fill.fill.RasterImageId);
				}

				if (!this.noCreatePoint || this.exucuteHistory)
				{
					if (!this.noCreatePoint && !this.exucuteHistory && this.exucuteHistoryEnd)
					{
						oLogicDocument.changeBackground(bg, arr_ind, true);
						this.exucuteHistoryEnd = false;
					}
					else
					{
						oLogicDocument.changeBackground(bg, arr_ind);
					}
					if (this.exucuteHistory)
					{
						this.exucuteHistory = false;
					}
					if(this.exucuteHistoryEnd)
					{
						this.exucuteHistoryEnd = false;
					}
				}
				else
				{
					if (oLogicDocument.GetCurrentSlide())
					{
						AscFormat.ExecuteNoHistory(function()
						{

							oLogicDocument.changeBackground(bg, arr_ind, true);
							for (var i = 0; i < arr_ind.length; ++i)
							{
								oLogicDocument.GetSlide(arr_ind[i]).recalculateBackground()
							}
							for (i = 0; i < arr_ind.length; ++i)
							{
								oLogicDocument.DrawingDocument.OnRecalculateSlide(arr_ind[i]);
							}
							oLogicDocument.DrawingDocument.OnEndRecalculate(true, false);
						}, this, []);
					}
				}
			}
		}

		var _transition = prop.get_transition();
		if (_transition)
		{
			this.ApplySlideTransition(_transition);
		}
	};

	asc_docs_api.prototype.put_LineCap  = function(_cap)
	{
		this.WordControl.m_oLogicDocument.putLineCap(_cap);
	};
	asc_docs_api.prototype.put_LineJoin = function(_join)
	{
		this.WordControl.m_oLogicDocument.putLineJoin(_join);
	};

	asc_docs_api.prototype.put_LineBeginStyle = function(_style)
	{
		this.WordControl.m_oLogicDocument.putLineBeginStyle(_style);
	};
	asc_docs_api.prototype.put_LineBeginSize  = function(_size)
	{
		this.WordControl.m_oLogicDocument.putLineBeginSize(_size);
	};

	asc_docs_api.prototype.put_LineEndStyle = function(_style)
	{
		this.WordControl.m_oLogicDocument.putLineEndStyle(_style);
	};
	asc_docs_api.prototype.put_LineEndSize  = function(_size)
	{
		this.WordControl.m_oLogicDocument.putLineEndSize(_size);
	};

	asc_docs_api.prototype.put_TextColor2 = function(r, g, b)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
				Color : {
					r : r,
					g : g,
					b : b
				}
			}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.put_TextColor  = function(color)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			var _unifill        = new AscFormat.CUniFill();
			_unifill.fill       = new AscFormat.CSolidFill();
			_unifill.fill.color = AscFormat.CorrectUniColor(color, _unifill.fill.color, 0);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Unifill : _unifill}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.put_PrIndent          = function(value, levelValue)
	{
		this.WordControl.m_oLogicDocument.SetParagraphIndent({Left : value, ChangeLevel : levelValue});
	};
	asc_docs_api.prototype.IncreaseIndent        = function()
	{
		this.WordControl.m_oLogicDocument.IncreaseDecreaseIndent(true);
	};
	asc_docs_api.prototype.DecreaseIndent        = function()
	{
		this.WordControl.m_oLogicDocument.IncreaseDecreaseIndent(false);
	};
	asc_docs_api.prototype.put_PrIndentRight     = function(value)
	{
		this.WordControl.m_oLogicDocument.SetParagraphIndent({Right : value});
	};
	asc_docs_api.prototype.put_PrFirstLineIndent = function(value)
	{
		this.WordControl.m_oLogicDocument.SetParagraphIndent({FirstLine : value});
	};
	asc_docs_api.prototype.getFocusObject        = function()
	{//возвратит тип элемента - параграф c_oAscTypeSelectElement.Paragraph, изображение c_oAscTypeSelectElement.Image, таблица c_oAscTypeSelectElement.Table, колонтитул c_oAscTypeSelectElement.Header.

	};

	/*callbacks*/
	asc_docs_api.prototype.sync_VerticalAlign           = function(typeBaseline)
	{
		this.sendEvent("asc_onVerticalAlign", typeBaseline);
	};
	asc_docs_api.prototype.sync_PrAlignCallBack         = function(value)
	{
		this.sendEvent("asc_onPrAlign", value);
	};
	asc_docs_api.prototype.sync_ListType                = function(NumPr)
	{
		this.sendEvent("asc_onListType", new AscCommon.asc_CListType(NumPr));
	};
	asc_docs_api.prototype.sync_TextColor               = function(Color)
	{
		this.sendEvent("asc_onTextColor", new AscCommon.CColor(Color.r, Color.g, Color.b));
	};
	asc_docs_api.prototype.sync_TextColor2              = function(unifill)
	{
		var _color;
		if (unifill.fill == null)
			return;
		var color;
		if (unifill.fill.type == c_oAscFill.FILL_TYPE_SOLID)
		{
			_color    = unifill.getRGBAColor();
			color = AscCommon.CreateAscColor(unifill.fill.color);
			color.asc_putR(_color.R);
			color.asc_putG(_color.G);
			color.asc_putB(_color.B);
			this.sendEvent("asc_onTextColor", color);
		}
		else if (unifill.fill.type == c_oAscFill.FILL_TYPE_GRAD)
		{
			_color    = unifill.getRGBAColor();
			if(unifill.fill.colors[0] && unifill.fill.colors[0].color)
			{
				color = AscCommon.CreateAscColor(unifill.fill.colors[0].color);
			}
			else
			{
				color = new Asc.asc_CColor();
			}
			color.asc_putR(_color.R);
			color.asc_putG(_color.G);
			color.asc_putB(_color.B);
			this.sendEvent("asc_onTextColor", color);
		}
		else
		{
			_color    = unifill.getRGBAColor();
			color = new Asc.asc_CColor();
			color.asc_putR(_color.R);
			color.asc_putG(_color.G);
			color.asc_putB(_color.B);
			this.sendEvent("asc_onTextColor", color);
		}
	};
	asc_docs_api.prototype.sync_TextHighLight           = function(HighLight)
	{
		this.sendEvent("asc_onTextHighLight", new AscCommon.CColor(HighLight.r, HighLight.g, HighLight.b));
	};
	asc_docs_api.prototype.sync_ParaStyleName           = function(Name)
	{
		this.sendEvent("asc_onParaStyleName", Name);
	};
	asc_docs_api.prototype.sync_ParaSpacingLine         = function(SpacingLine)
	{
		this.sendEvent("asc_onParaSpacingLine", new AscCommon.asc_CParagraphSpacing(SpacingLine));
	};
	asc_docs_api.prototype.sync_PageBreakCallback       = function(isBreak)
	{
		this.sendEvent("asc_onPageBreak", isBreak);
	};
	asc_docs_api.prototype.sync_KeepLinesCallback       = function(isKeepLines)
	{
		this.sendEvent("asc_onKeepLines", isKeepLines);
	};
	asc_docs_api.prototype.sync_ShowParaMarksCallback   = function()
	{
		this.sendEvent("asc_onShowParaMarks");
	};
	asc_docs_api.prototype.sync_SpaceBetweenPrgCallback = function()
	{
		this.sendEvent("asc_onSpaceBetweenPrg");
	};
	asc_docs_api.prototype.sync_PrPropCallback          = function(prProp)
	{
		var _len = this.SelectedObjectsStack.length;
		if (_len > 0)
		{
			if (this.SelectedObjectsStack[_len - 1].Type == c_oAscTypeSelectElement.Paragraph)
			{
				this.SelectedObjectsStack[_len - 1].Value = new Asc.asc_CParagraphProperty(prProp);
				return;
			}
		}

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Paragraph, new Asc.asc_CParagraphProperty(prProp));
	};

    asc_docs_api.prototype.SetDrawImagePlaceParagraph = function(element_id, props)
	{
		this.WordControl.m_oDrawingDocument.InitGuiCanvasTextProps(element_id);
		this.WordControl.m_oDrawingDocument.DrawGuiCanvasTextProps(props);
	};

	/*----------------------------------------------------------------*/


	asc_docs_api.prototype.Update_ParaInd                = function(Ind)
	{
		var FirstLine = 0;
		var Left      = 0;
		var Right     = 0;
		if ("undefined" != typeof(Ind))
		{
			if ("undefined" != typeof(Ind.FirstLine))
			{
				FirstLine = Ind.FirstLine;
			}
			if ("undefined" != typeof(Ind.Left))
			{
				Left = Ind.Left;
			}
			if ("undefined" != typeof(Ind.Right))
			{
				Right = Ind.Right;
			}
		}

		this.Internal_Update_Ind_Left(Left);
		this.Internal_Update_Ind_FirstLine(FirstLine, Left);
		this.Internal_Update_Ind_Right(Right);
	};
	asc_docs_api.prototype.Internal_Update_Ind_FirstLine = function(FirstLine, Left)
	{
		if (this.WordControl.m_oHorRuler.m_dIndentLeftFirst != (FirstLine + Left))
		{
			this.WordControl.m_oHorRuler.m_dIndentLeftFirst = (FirstLine + Left);
			this.WordControl.UpdateHorRuler();
		}
	};
	asc_docs_api.prototype.Internal_Update_Ind_Left      = function(Left)
	{
		if (this.WordControl.m_oHorRuler.m_dIndentLeft != Left)
		{
			this.WordControl.m_oHorRuler.m_dIndentLeft = Left;
			this.WordControl.UpdateHorRuler();
		}
	};
	asc_docs_api.prototype.Internal_Update_Ind_Right     = function(Right)
	{
		if (this.WordControl.m_oHorRuler.m_dIndentRight != Right)
		{
			this.WordControl.m_oHorRuler.m_dIndentRight = Right;
			this.WordControl.UpdateHorRuler();
		}
	};

	/*----------------------------------------------------------------*/
	/*functions for working with table*/
	asc_docs_api.prototype.put_Table               = function(col, row, placeholder, sStyleId)
	{
		if(placeholder)
		{
			this.WordControl.m_bIsMouseLock = false;
		}
		this.WordControl.m_oLogicDocument.Add_FlowTable(col, row, placeholder, sStyleId);
	};
	asc_docs_api.prototype.addRowAbove             = function(count)
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_AddRowAbove);
			this.WordControl.m_oLogicDocument.AddTableRow(true);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.addRowBelow             = function(count)
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_AddRowBelow);
			this.WordControl.m_oLogicDocument.AddTableRow(false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.addColumnLeft           = function(count)
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_AddColLeft);
			this.WordControl.m_oLogicDocument.AddTableColumn(true);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.addColumnRight          = function(count)
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_AddColRight);
			this.WordControl.m_oLogicDocument.AddTableColumn(false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.remRow                  = function()
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_RemoveRow);
			this.WordControl.m_oLogicDocument.RemoveTableRow();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.remColumn               = function()
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_RemoveCol);
			this.WordControl.m_oLogicDocument.RemoveTableColumn();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.remTable                = function()
	{
		var doc = this.WordControl.m_oLogicDocument;
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_RemoveTable);
			this.WordControl.m_oLogicDocument.RemoveTable();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_DistributeTableCells = function(isHorizontally)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;


		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			oLogicDocument.StartAction(AscDFH.historydescription_Document_DistributeTableCells);
			oLogicDocument.DistributeTableCells(isHorizontally)
			oLogicDocument.FinalizeAction();
		}

		return true;
	};
	asc_docs_api.prototype.selectRow               = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Row);
	};
	asc_docs_api.prototype.selectColumn            = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Column);
	};
	asc_docs_api.prototype.selectCell              = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Cell);
	};
	asc_docs_api.prototype.selectTable             = function()
	{
		this.WordControl.m_oLogicDocument.SelectTable(c_oAscTableSelectionType.Table);
	};
	asc_docs_api.prototype.setColumnWidth          = function(width)
	{

	};
	asc_docs_api.prototype.setRowHeight            = function(height)
	{

	};
	asc_docs_api.prototype.set_TblDistanceFromText = function(left, top, right, bottom)
	{

	};
	asc_docs_api.prototype.CheckBeforeMergeCells   = function()
	{
		return this.WordControl.m_oLogicDocument.CanMergeTableCells();
	};
	asc_docs_api.prototype.CheckBeforeSplitCells   = function()
	{
		return this.WordControl.m_oLogicDocument.CanSplitTableCells();
	};
	asc_docs_api.prototype.MergeCells              = function()
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_MergeCells);
			this.WordControl.m_oLogicDocument.MergeTableCells();
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.SplitCell               = function(Cols, Rows)
	{
		var doc = this.WordControl.m_oLogicDocument;
		var oController = doc.GetCurrentController();
		if(!oController){
			return;
		}
		var aAdditionalObjects = oController.getConnectorsForCheck2();
		if (doc.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_SplitCells);
			this.WordControl.m_oLogicDocument.SplitTableCells(Cols, Rows);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};
	asc_docs_api.prototype.widthTable              = function(width)
	{

	};
	asc_docs_api.prototype.put_CellsMargin         = function(left, top, right, bottom)
	{

	};
	asc_docs_api.prototype.set_TblWrap             = function(type)
	{

	};
	asc_docs_api.prototype.set_TblIndentLeft       = function(spacing)
	{

	};
	asc_docs_api.prototype.set_Borders             = function(typeBorders, size, Color)
	{//если size == 0 то границы нет.

	};
	asc_docs_api.prototype.set_TableBackground     = function(Color)
	{

	};
	asc_docs_api.prototype.set_AlignCell           = function(align)
	{// c_oAscAlignType.RIGHT, c_oAscAlignType.LEFT, c_oAscAlignType.CENTER
		switch (align)
		{
			case c_oAscAlignType.LEFT :
				break;
			case c_oAscAlignType.CENTER :
				break;
			case c_oAscAlignType.RIGHT :
				break;
		}
	};
	asc_docs_api.prototype.set_TblAlign            = function(align)
	{// c_oAscAlignType.RIGHT, c_oAscAlignType.LEFT, c_oAscAlignType.CENTER
		switch (align)
		{
			case c_oAscAlignType.LEFT :
				break;
			case c_oAscAlignType.CENTER :
				break;
			case c_oAscAlignType.RIGHT :
				break;
		}
	};
	asc_docs_api.prototype.set_SpacingBetweenCells = function(isOn, spacing)
	{// c_oAscAlignType.RIGHT, c_oAscAlignType.LEFT, c_oAscAlignType.CENTER
		if (isOn)
		{

		}
	};


	/*
	 {
	 TableWidth   : null - галочка убрана, либо заданное значение в мм
	 TableSpacing : null - галочка убрана, либо заданное значение в мм

	 TableDefaultMargins :  // маргины для всей таблицы(значение по умолчанию)
	 {
	 Left   : 1.9,
	 Right  : 1.9,
	 Top    : 0,
	 Bottom : 0
	 }

	 CellMargins :
	 {
	 Left   : 1.9, (null - неопределенное значение)
	 Right  : 1.9, (null - неопределенное значение)
	 Top    : 0,   (null - неопределенное значение)
	 Bottom : 0,   (null - неопределенное значение)
	 Flag   : 0 - У всех выделенных ячеек значение берется из TableDefaultMargins
	 1 - У выделенных ячеек есть ячейки с дефолтовыми значениями, и есть со своими собственными
	 2 - У всех ячеек свои собственные значения
	 }

	 TableAlignment : 0, 1, 2 (слева, по центру, справа)
	 TableIndent : значение в мм,
	 TableWrappingStyle : 0, 1 (inline, flow)
	 TablePaddings:
	 {
	 Left   : 3.2,
	 Right  : 3.2,
	 Top    : 0,
	 Bottom : 0
	 }

	 TableBorders : // границы таблицы
	 {
	 Bottom :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 Left :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 Right :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 Top :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 InsideH :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 InsideV :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 }
	 }

	 CellBorders : // границы выделенных ячеек
	 {
	 ForSelectedCells : true,

	 Bottom :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 Left :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 Right :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 Top :
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 InsideH : // данного элемента может не быть, если у выделенных ячеек
	 // нет горизонтальных внутренних границ
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 },

	 InsideV : // данного элемента может не быть, если у выделенных ячеек
	 // нет вертикальных внутренних границ
	 {
	 Color : { r : 0, g : 0, b : 0 },
	 Value : border_Single,
	 Size  : 0.5 * g_dKoef_pt_to_mm
	 Space :
	 }
	 }

	 TableBackground :
	 {
	 Value : тип заливки(прозрачная или нет),
	 Color : { r : 0, g : 0, b : 0 }
	 }
	 CellsBackground : null если заливка не определена для выделенных ячеек
	 {
	 Value : тип заливки(прозрачная или нет),
	 Color : { r : 0, g : 0, b : 0 }
	 }

	 Position:
	 {
	 X:0,
	 Y:0
	 }
	 }
	 */
	asc_docs_api.prototype.tblApply = function(oPr)
	{
		const oPresentation = this.private_GetLogicDocument()
		if(!oPresentation)
			return;
		const oController = oPresentation.GetCurrentController();
		if(!oController){
			return;
		}
		const aAdditionalObjects = oController.getConnectorsForCheck2();
		if (!oPresentation.Document_Is_SelectionLocked(changestype_Drawing_Props, undefined, undefined, aAdditionalObjects))
		{
			oPresentation.StartAction(AscDFH.historydescription_Presentation_TblApply);

			let oBorders = oPr.CellBorders;
			if (oPr.CellBorders)
			{
				function fCheckBorder(oBorder)
				{
					if(!oBorder || !oBorder.Color)
						return;
					oBorder.Unifill =  AscFormat.CreateUnifillFromAscColor(oBorder.Color, 0);
				}
				fCheckBorder(oBorders.Left);
				fCheckBorder(oBorders.Top);
				fCheckBorder(oBorders.Right);
				fCheckBorder(oBorders.Bottom);
				fCheckBorder(oBorders.InsideH);
				fCheckBorder(oBorders.InsideV);
			}
			let oBackground = oPr.CellsBackground;
			if (oBackground && oBackground.Color)
			{
				if(oBackground.Value === Asc.c_oAscShd.Nil)
				{
					oBackground.Value = Asc.c_oAscShd.Clear;
					oBackground.Unifill = AscFormat.CreateNoFillUniFill();
				}
				else
				{
					oBackground.Unifill = AscFormat.CreateUnifillFromAscColor(oBackground.Color, 0);
				}
			}
			this.WordControl.m_oLogicDocument.SetTableProps(oPr);
			oPresentation.FinalizeAction();
		}
	};
	/*callbacks*/
	asc_docs_api.prototype.sync_AddTableCallback            = function()
	{
		this.sendEvent("asc_onAddTable");
	};
	asc_docs_api.prototype.getDrawingObjects = function () {
		const oController = this.getGraphicController();
		if (oController) {
			return oController.drawingObjects;
		}
	};
	asc_docs_api.prototype.getDrawingDocument = function () {
		return this.WordControl.m_oDrawingDocument;
	};
	asc_docs_api.prototype.getLogicDocument = asc_docs_api.prototype.private_GetLogicDocument;
	asc_docs_api.prototype._createSmartArt = function (oSmartArt, oPlaceholder)
	{
		const oLogicDocument = this.getLogicDocument();
		const oController = this.getGraphicController();
		const oDrawingObjects = this.getDrawingObjects();
		if (!oDrawingObjects || !oController || !oLogicDocument)
		{
			return;
		}
		oSmartArt.setDrawingObjects(oDrawingObjects);
		oSmartArt.setParent(oDrawingObjects);
		oSmartArt.setRecalculateInfo();
		let oPh;
		let bLocked = false;
		if (oPlaceholder)
		{
			this.WordControl.m_bIsMouseLock = false;
			oPh = AscCommon.g_oTableId.Get_ById(oPlaceholder.id);
			if (oPh)
			{
				if (oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, undefined, undefined, [oPh]))
				{
					bLocked = true;
				}
				else
				{
					const nWidth = oPh.extX;
					const nHeight = oPh.extY;
					oSmartArt.fitForSizes(nHeight, nWidth);
					const nX = oPh.x + oPh.extX / 2 - oSmartArt.spPr.xfrm.extX / 2;
					const nY = oPh.y + oPh.extY / 2 - oSmartArt.spPr.xfrm.extY / 2;
					oSmartArt.spPr.xfrm.setOffX(nX);
					oSmartArt.spPr.xfrm.setOffY(nY);
				}
			}
			else
			{
				return;
			}
		}
		var oSlide = oLogicDocument.GetCurrentSlide();
		if (oPh && !bLocked)
		{
			oSlide.replaceSp(oPh, oSmartArt);
		}
		else
		{
			oSlide.addToSpTreeToPos(oSlide.cSld.spTree.length, oSmartArt);
		}
		oSmartArt.checkDrawingBaseCoords();
		oSmartArt.generateDrawingPart();
		oController.checkChartTextSelection();
		oController.resetSelection();
		oSmartArt.select(oController, 0);
		oController.startRecalculate();
		oDrawingObjects.sendGraphicObjectProps();
	};
	asc_docs_api.prototype.sync_AlignCellCallback           = function(align)
	{
		this.sendEvent("asc_onAlignCell", align);
	};
	asc_docs_api.prototype.sync_TblPropCallback             = function(tblProp)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Table, new Asc.CTableProp(tblProp));
	};
	asc_docs_api.prototype.sync_TblWrapStyleChangedCallback = function(style)
	{
		this.sendEvent("asc_onTblWrapStyleChanged", style);
	};
	asc_docs_api.prototype.sync_TblAlignChangedCallback     = function(style)
	{
		this.sendEvent("asc_onTblAlignChanged", style);
	};

	/*----------------------------------------------------------------*/
	/*functions for working with images*/
	asc_docs_api.prototype.ChangeImageFromFile      = function()
	{
		this.asc_addImage({isImageChangeUrl: true});
	};
	asc_docs_api.prototype.ChangeShapeImageFromFile = function(type)
	{
		this.asc_addImage({isShapeImageChangeUrl: true, textureType: type});
	};
	asc_docs_api.prototype.ChangeSlideImageFromFile = function(type)
	{
		this.asc_addImage({isSlideImageChangeUrl: true, textureType: type});
	};
	asc_docs_api.prototype.ChangeArtImageFromFile   = function(type)
	{
		this.asc_addImage({isTextArtChangeUrl: true, textureType: type});
	};

	asc_docs_api.prototype.AddImage      = function()
	{
		this.asc_addImage();
	};

	asc_docs_api.prototype.asc_AddToLayout      = function()
	{
		this.WordControl.m_oLogicDocument.AddToLayout();
	};
	asc_docs_api.prototype.asc_AddAnimation      = function(nPresetClass, nPresetId, nPresetSubtype, oColor, bReplace, bPreview)
	{

		let oColor_, bReplace_, bPreview_;
		if(arguments.length === 6) {
			oColor_ =  oColor;
			bReplace_ = bReplace;
			bPreview_ = bPreview;
		}
		else {
			oColor_ =  null;
			bReplace_ = oColor;
			bPreview_ = bReplace;
		}
		var bStartPreview = true;
		if(bPreview_ === false) {
			bStartPreview = false;
		}
		this.WordControl.m_oLogicDocument.AddAnimation(nPresetClass, nPresetId, nPresetSubtype, oColor_, bReplace_, bStartPreview);
	};
	asc_docs_api.prototype.asc_getCurSlideObjectsNames = function()
	{
		return this.WordControl.m_oLogicDocument.GetCurSlideObjectsNames();
	};

	asc_docs_api.prototype.asc_IsStartedAnimationPreview = function()
	{
		return this.WordControl.m_oLogicDocument.IsStartedPreview()
	};
	asc_docs_api.prototype.asc_StartAnimationPreview = function (bAll, bIncludeFollowing) {
		this.asc_StopAnimationPreview();
		const bStartAll = (bAll === false) ? false : true; // if "bAll" is "undefined", then run preview of all effects
		if (this.WordControl.m_oLogicDocument.StartAnimationPreview(bStartAll, bIncludeFollowing)) {
			this.sendEvent("asc_onAnimPreviewStarted");
		}
		//this.sendEvent("asc_onStartDemonstration");//todo
	};

	asc_docs_api.prototype.asc_canStartAnimationPreview = function()
	{
		return this.WordControl.m_oLogicDocument.CanStartAnimationPreview();
	};
	asc_docs_api.prototype.asc_StopAnimationPreview = function()
	{
		this.WordControl.m_oLogicDocument.StopAnimationPreview();
	};
	asc_docs_api.prototype.asc_SetAnimationProperties = function(oPr)
	{
		this.WordControl.m_oLogicDocument.SetAnimationProperties(oPr)
	};
	asc_docs_api.prototype.asc_canMoveAnimationEarlier = function(nPositions)
	{
		return this.WordControl.m_oLogicDocument.CanMoveAnimation(true, nPositions);
	};
	asc_docs_api.prototype.asc_canMoveAnimationLater = function(nPositions) 
	{
		return this.WordControl.m_oLogicDocument.CanMoveAnimation(false, nPositions);
	};
	asc_docs_api.prototype.asc_onShowAnimTab = function(bShow) 
	{
		this.bIsShowAnimTab = bShow;
		if(this.WordControl && this.WordControl.m_oLogicDocument) 
		{
			var slide = this.WordControl.m_oLogicDocument.GetSlide(this.WordControl.m_oLogicDocument.CurPage);
			if(slide) 
			{
				this.WordControl.m_oDrawingDocument.OnRecalculateSlide(this.WordControl.m_oLogicDocument.CurPage);
				if(!bShow)
				{
					if(slide.timing)
					{
						slide.timing.checkSelectedAnimMotionShapes();
					}
				}
			}
		}
	};
	asc_docs_api.prototype.isDrawAnimLabels = function() 
	{
		return this.bIsShowAnimTab || this.getIsAnimPaneShow();
	};
	asc_docs_api.prototype.asc_moveAnimationEarlier = function(nPositions)
	{
		return this.WordControl.m_oLogicDocument.MoveAnimation(true, nPositions);
	};
	asc_docs_api.prototype.asc_moveAnimationLater = function(nPositions)
	{
		return this.WordControl.m_oLogicDocument.MoveAnimation(false, nPositions);
	};
	asc_docs_api.prototype.getImageDataFromSelection = function()
	{
		return this.WordControl.m_oLogicDocument.getImageDataFromSelection();
	};
	asc_docs_api.prototype.putImageToSelection = function(sImageSrc, nWidth, nHeight, replaceMode)
	{
		return this.WordControl.m_oLogicDocument.putImageToSelection(sImageSrc, nWidth, nHeight, replaceMode);
	};

	asc_docs_api.prototype.getPluginContextMenuInfo = function ()
	{
		const oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
		{
			return new AscCommon.CPluginCtxMenuInfo();
		}
		let oCurrentController = oLogicDocument.GetCurrentController();
		if(!oCurrentController)
		{
			return new AscCommon.CPluginCtxMenuInfo();
		}
		return oCurrentController.getPluginSelectionInfo();
	};
	asc_docs_api.prototype.StartAddShape = function(prst, is_apply)
	{
		this.stopInkDrawer();
		this.cancelEyedropper();
		this.WordControl.m_oLogicDocument.StartAddShape(prst, is_apply);

		if (is_apply)
		{
			this.WordControl.m_oDrawingDocument.LockCursorType("crosshair");
		}
	};
	asc_docs_api.prototype.asc_canEditGeometry = function ()
	{
		var oPresentation = this.WordControl && this.WordControl.m_oLogicDocument;
		if(!oPresentation)
		{
			return false;
		}
		return oPresentation.CanEditGeometry();
	};

	asc_docs_api.prototype.asc_editPointsGeometry = function()
	{
		var oPresentation = this.WordControl && this.WordControl.m_oLogicDocument;
		if(!oPresentation)
		{
			return;
		}
		return oPresentation.StartEditGeometry();
	};

	asc_docs_api.prototype.asc_addSlideNumber = function()
	{
		this.WordControl.m_oLogicDocument.addSlideNumber();
	};
	asc_docs_api.prototype.asc_addDateTime = function(oPr)
	{

		var sCheck = oPr.get_DateTimeExamples()[oPr.get_DateTime()], sTextForCheck = "";
		if(typeof sCheck === "string" && sCheck.length > 0)
		{
			sTextForCheck += sCheck;
		}
		else
		{
			sCheck = oPr.get_CustomDateTime();
			if(typeof sCheck === "string" && sCheck.length > 0)
			{
				sTextForCheck += sCheck;
			}
		}

		if(sTextForCheck.length > 0)
		{
			AscFonts.FontPickerByCharacter.checkText(sTextForCheck, this, function() {
				this.WordControl.m_oLogicDocument.addDateTime(oPr);
			});
		}
		else
		{
			this.WordControl.m_oLogicDocument.addDateTime(oPr);
		}
	};
	asc_docs_api.prototype.asc_setDefaultDateTimeFormat = function(aFormat)
	{
		window['AscCommonWord'] = window['AscCommonWord'] || {};
		for(var key in aFormat)
		{
			window['AscCommonWord'].oDefaultDateTimeFormat[key] = aFormat[key];
		}
	};

	asc_docs_api.prototype.asc_getHeaderFooterProperties = function()
	{
		if(this.WordControl && this.WordControl.m_oLogicDocument)
		{
			return this.WordControl.m_oLogicDocument.getHFProperties();
		}
		return null;
	};

	asc_docs_api.prototype.asc_setHeaderFooterProperties = function(oProps, bAll)
	{

		if(oProps && this.WordControl && this.WordControl.m_oLogicDocument)
		{
			var sTextForCheck = "";
			var sCheck;
			var oSlide = oProps.get_Slide();
			var oThis = this;
			var sDateTime;
			if(oSlide)
			{
				var oDateTime = oSlide.get_DateTime();
				if(oDateTime)
				{

					sCheck = oDateTime.get_DateTimeExamples()[oDateTime.get_DateTime()];
					if(typeof sCheck === "string" && sCheck.length > 0)
					{
						sTextForCheck += sCheck;
					}
					else
					{
						sCheck = oDateTime.get_CustomDateTime();
						if(typeof sCheck === "string" && sCheck.length > 0)
						{
							sTextForCheck += sCheck;
						}
					}
				}
				sCheck = oSlide.get_Footer();
				if(typeof sCheck === "string" && sCheck.length > 0)
				{
					sTextForCheck += sCheck;
				}
				sCheck = oSlide.get_Header();
				if(typeof sCheck === "string" && sCheck.length > 0)
				{
					sTextForCheck += sCheck;
				}
			}
			//TODO: check notes
			if(sTextForCheck.length > 0)
			{
				AscFonts.FontPickerByCharacter.checkText(sTextForCheck, this, function() {
					oThis.WordControl.m_oLogicDocument.setHFProperties(oProps, bAll);
				});
			}
			else
			{

				oThis.WordControl.m_oLogicDocument.setHFProperties(oProps, bAll);
			}
		}
	};


	asc_docs_api.prototype.asc_addOleObjectAction = function(sLocalUrl, Data, sApplicationId, fWidth, fHeight, nWidthPix, nHeightPix, bSelect, arrImagesForAddToHistory)
	{
		var _image = this.ImageLoader.LoadImage(AscCommon.getFullImageSrc2(sLocalUrl), 1);
		if (null != _image)//картинка уже должна быть загружена
		{
            this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_PasteHotKey);
			this.WordControl.m_oLogicDocument.AddOleObject(fWidth, fHeight, nWidthPix, nHeightPix, sLocalUrl, Data, sApplicationId, bSelect, arrImagesForAddToHistory);
            this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_editOleObjectAction = function(oOleObject, sImageUrl, sData, fWidthMM, fHeightMM, nPixWidth, nPixHeight, arrImagesForAddToHistory)
	{
		if (oOleObject)
		{
            this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_PasteHotKey);
			this.WordControl.m_oLogicDocument.EditOleObject(oOleObject, sData, sImageUrl, fWidthMM, fHeightMM, nPixWidth, nPixHeight, arrImagesForAddToHistory);
			this.WordControl.m_oLogicDocument.Recalculate();
            this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_doubleClickOnTableOleObject    = function(obj)
	{
		this.isOleEditor = true;	// Для совместного редактирования
		this.asc_onOpenChartFrame();

		if(!window['IS_NATIVE_EDITOR']) {
			this.WordControl.onMouseUpMainSimple();
		}
		if(this.asc_checkNeedCallback("asc_doubleClickOnTableOleObject"))
		{
			this.sendEvent("asc_doubleClickOnTableOleObject", obj);
		}
		else
		{
			this.sendEvent("asc_doubleClickOnChart", obj); // TODO: change event type
		}
	};


    asc_docs_api.prototype.asc_startEditCurrentOleObject = function(){
    	if(this.WordControl.m_oLogicDocument.GetCurrentSlide())
            this.WordControl.m_oLogicDocument.GetCurrentSlide().graphicObjects.startEditCurrentOleObject();
    };

    // signatures
    asc_docs_api.prototype.asc_addSignatureLine = function (oPr, Width, Height, sImgUrl) {
        if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false){
            this.WordControl.m_oLogicDocument.AddSignatureLine(oPr, Width, Height, sImgUrl);
        }
    };

    asc_docs_api.prototype.asc_getAllSignatures = function(){
        return this.WordControl.m_oLogicDocument.GetAllSignatures();
    };


    asc_docs_api.prototype.asc_CallSignatureDblClickEvent = function(sGuid){
        return this.WordControl.m_oLogicDocument.CallSignatureDblClickEvent(sGuid);
    };
    //-------------------------------------------------------



	asc_docs_api.prototype.asc_canEditCrop = function()
	{

		return this.WordControl.m_oLogicDocument.canStartImageCrop();
	};

	asc_docs_api.prototype.asc_startEditCrop = function()
	{
		return this.WordControl.m_oLogicDocument.startImageCrop();
	};

	asc_docs_api.prototype.asc_endEditCrop = function()
	{
		return this.WordControl.m_oLogicDocument.endImageCrop();
	};

	asc_docs_api.prototype.asc_cropFit = function()
	{
		return this.WordControl.m_oLogicDocument.cropFit();
	};

	asc_docs_api.prototype.asc_cropFill = function()
	{
		return this.WordControl.m_oLogicDocument.cropFill();
	};


	asc_docs_api.prototype.AddTextArt = function(nStyle)
	{
		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.AddTextArt(nStyle);
		}
	};


	asc_docs_api.prototype.canGroup = function()
	{
		return this.WordControl.m_oLogicDocument.canGroup();
	};

	asc_docs_api.prototype.canUnGroup = function()
	{
		return this.WordControl.m_oLogicDocument.canUnGroup();
	};

	asc_docs_api.prototype._addImageUrl = function(urls, obj)
	{
		if (obj && obj.sendUrlsToFrameEditor && this.isOpenedChartFrame) {
			this.addImageUrlsFromGeneralToFrameEditor(urls);
			return;
		}
		if(obj && (obj.isImageChangeUrl || obj.isShapeImageChangeUrl || obj.isSlideImageChangeUrl || obj.isTextArtChangeUrl || obj.isImageBullet || obj.fAfterUploadOleObjectImage)){
            this.AddImageUrlAction(urls[0], undefined, obj);
		}
		else{
			if(this.ImageLoader){
				var oApi = this;
                this.ImageLoader.LoadImagesWithCallback(urls, function(){
                	var aImages = [];
                	for(var i = 0; i < urls.length; ++i){
                        var _image = oApi.ImageLoader.LoadImage(urls[i], 1);
                      	if(_image){
                            aImages.push(_image);
						}
					}
                    oApi.WordControl.m_oLogicDocument.addImages(aImages, obj);
                }, []);
			}
		}
	};

	asc_docs_api.prototype.AddImageUrlActionCallback = function(_image, obj)
	{
		var _w = AscCommon.Page_Width - (AscCommon.X_Left_Margin + AscCommon.X_Right_Margin);
		var _h = AscCommon.Page_Height - (AscCommon.Y_Top_Margin + AscCommon.Y_Bottom_Margin);
		if (_image.Image != null)
		{
			var __w = Math.max((_image.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
			var __h = Math.max((_image.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
			_w      = Math.max(5, Math.min(_w, __w));
			_h      = Math.max(5, Math.min((_w * __h / __w)));
		}

		var src = _image.src;
		if (obj && obj.isShapeImageChangeUrl)
		{
			var AscShapeProp       = new Asc.asc_CShapeProperty();
			AscShapeProp.fill      = new asc_CShapeFill();
			AscShapeProp.fill.type = c_oAscFill.FILL_TYPE_BLIP;
			AscShapeProp.fill.fill = new asc_CFillBlip();
			AscShapeProp.fill.fill.asc_putUrl(src);
			if(obj.textureType !== null && obj.textureType !== undefined){
                AscShapeProp.fill.fill.asc_putType(obj.textureType);
			}
			this.ShapeApply(AscShapeProp);
		}
		else if (obj && obj.fAfterUploadOleObjectImage)
		{
			obj.fAfterUploadOleObjectImage(src);
		}
		else if (obj && obj.isImageBullet)
		{
			var fillBlip = new Asc.asc_CFillBlip();
			fillBlip.asc_putUrl(src, null);
			this.put_ListType(undefined, undefined, fillBlip);
		}
		else if (obj && obj.isSlideImageChangeUrl)
		{
			var AscSlideProp             = new CAscSlideProps();
			AscSlideProp.Background      = new asc_CShapeFill();
			AscSlideProp.Background.type = c_oAscFill.FILL_TYPE_BLIP;
			AscSlideProp.Background.fill = new asc_CFillBlip();
			AscSlideProp.Background.fill.asc_putUrl(src);
			if(obj.textureType !== null && obj.textureType !== undefined){
                AscSlideProp.Background.fill.asc_putType(obj.textureType);
			}
			this.SetSlideProps(AscSlideProp);
		}
		else if (obj && obj.isImageChangeUrl)
		{
			var AscImageProp      = new Asc.asc_CImgProperty();
			AscImageProp.ImageUrl = src;
			this.ImgApply(AscImageProp);
		}
		else if (obj && obj.isTextArtChangeUrl)
		{
			var AscShapeProp = new Asc.asc_CShapeProperty();
			var oFill        = new asc_CShapeFill();
			oFill.type       = c_oAscFill.FILL_TYPE_BLIP;
			oFill.fill       = new asc_CFillBlip();
			oFill.fill.asc_putUrl(src);
            if(obj.textureType !== null && obj.textureType !== undefined){
				oFill.fill.asc_putType(obj.textureType);
            }
			AscShapeProp.textArtProperties = new Asc.asc_TextArtProperties();
			AscShapeProp.textArtProperties.asc_putFill(oFill);
			this.ShapeApply(AscShapeProp);
		}
		else
		{
			var srcLocal = g_oDocumentUrls.getImageLocal(src);
			if (srcLocal)
			{
				src = srcLocal;
			}

			this.WordControl.m_oLogicDocument.addImages([_image], obj);
		}
	};

	asc_docs_api.prototype.AddImageUrlAction = function(url, imgProp, obj)
	{
		var _image = this.ImageLoader.LoadImage(url, 1);
		if (null != _image)
		{
			this.AddImageUrlActionCallback(_image, obj);
		}
		else
		{
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
			this.asyncImageEndLoaded2 = function(_image)
			{
				this.AddImageUrlActionCallback(_image, obj);
				this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);

				this.asyncImageEndLoaded2 = null;
			}
		}
	};
	/* В качестве параметра  передается объект класса Asc.asc_CImgProperty, он же приходит на OnImgProp
	 Asc.asc_CImgProperty заменяет пережнюю структуру:
	 если параметр не имеет значения то передвать следует null, напримере inline-картинок: в качестве left,top,bottom,right,X,Y,ImageUrl необходимо передавать null.
	 {
	 Width: 0,
	 Height: 0,
	 WrappingStyle: 0,
	 Paddings: { Left : 0, Top : 0, Bottom: 0, Right: 0 },
	 Position : {X : 0, Y : 0},
	 ImageUrl : ""
	 }
	 */
	asc_docs_api.prototype.ImgApply = function(obj)
	{
		var ImagePr        = {};
		ImagePr.lockAspect = obj.lockAspect;
		ImagePr.Width      = null === obj.Width || undefined === obj.Width ? null : parseFloat(obj.Width);
		ImagePr.Height     = null === obj.Height || undefined === obj.Height ? null : parseFloat(obj.Height);

		ImagePr.title       = obj.title;
		ImagePr.bSetOriginalSize = obj.bSetOriginalSize;
		ImagePr.description = obj.description;
		ImagePr.name = obj.name;
		ImagePr.rot = obj.rot;
		ImagePr.rotAdd = obj.rotAdd;
		ImagePr.flipH = obj.flipH;
		ImagePr.flipV = obj.flipV;
		ImagePr.flipHInvert = obj.flipHInvert;
		ImagePr.flipVInvert = obj.flipVInvert;
		ImagePr.resetCrop = obj.resetCrop;
		ImagePr.transparent = obj.transparent;

		if (undefined != obj.Position)
		{
			ImagePr.Position =
			{
				X : null === obj.Position.X || undefined === obj.Position.X ? null : parseFloat(obj.Position.X),
				Y : null === obj.Position.Y || undefined === obj.Position.Y ? null : parseFloat(obj.Position.Y)
			};
		}
		else
		{
			ImagePr.Position = {X : null, Y : null};
		}

		ImagePr.ImageUrl = obj.ImageUrl;

		if (window["NATIVE_EDITOR_ENJINE"])
		{
		  this.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
		  return;
		}
		if (!AscCommon.isNullOrEmptyString(ImagePr.ImageUrl))
		{
			var sImageUrl = null, sToken = undefined;
			if (!g_oDocumentUrls.getImageLocal(ImagePr.ImageUrl))
			{
				sImageUrl = ImagePr.ImageUrl;
				sToken = obj.Token;
			}

			var oApi           = this;
			var fApplyCallback = function()
			{
				var _img     = oApi.ImageLoader.LoadImage(ImagePr.ImageUrl, 1);
				var srcLocal = g_oDocumentUrls.getImageLocal(ImagePr.ImageUrl);
				if (srcLocal)
				{
					ImagePr.ImageUrl = srcLocal;
				}
				if (null != _img)
				{
					oApi.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
				}
				else
				{
					oApi.asyncImageEndLoaded2 = function(_image)
					{
						oApi.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
						oApi.asyncImageEndLoaded2 = null;
					}
				}
			};
			if (!sImageUrl)
			{
				fApplyCallback();
			}
			else
			{
				if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]())
                {
                    this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
                    var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageUrl);
                    _url     = g_oDocumentUrls.getImageUrl(_url);
                    ImagePr.ImageUrl = _url;
                    fApplyCallback();
                    this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
                    return;
                }

                AscCommon.sendImgUrls(this, [sImageUrl], function(data) {

                    if (data && data[0] && data[0].url !== "error")
                    {
                        ImagePr.ImageUrl = data[0].url;
                        fApplyCallback();
                    }

                }, undefined, sToken);
			}
		}
		else
		{
			ImagePr.ImageUrl = null;
			this.WordControl.m_oLogicDocument.SetImageProps(ImagePr);
		}
	};

	asc_docs_api.prototype.ChartApply              = function(obj)
	{
		if (obj.ChartProperties && obj.ChartProperties.type === Asc.c_oAscChartTypeSettings.stock && this.WordControl.m_oLogicDocument.GetCurrentSlide())
		{
			if (!AscFormat.CheckStockChart(this.WordControl.m_oLogicDocument.GetCurrentSlide().graphicObjects, this))
			{
				return;
			}
		}
		this.WordControl.m_oLogicDocument.ChartApply(obj);
	};
	asc_docs_api.prototype.set_Size                = function(width, height)
	{

	};
	asc_docs_api.prototype.set_ConstProportions    = function(isOn)
	{
		if (isOn)
		{

		}
		else
		{

		}
	};
	asc_docs_api.prototype.set_WrapStyle           = function(type)
	{

	};
	asc_docs_api.prototype.deleteImage             = function()
	{

	};
	asc_docs_api.prototype.set_ImgDistanceFromText = function(left, top, right, bottom)
	{

	};
	asc_docs_api.prototype.set_PositionOnPage      = function(X, Y)
	{//расположение от начала страницы

	};
	asc_docs_api.prototype.get_OriginalSizeImage   = function()
	{
		for(var i = 0; i < this.SelectedObjectsStack.length; ++i){
			if(this.SelectedObjectsStack[i].Type == c_oAscTypeSelectElement.Image && this.SelectedObjectsStack[i].Value && this.SelectedObjectsStack[i].Value.ImageUrl){
                return this.SelectedObjectsStack[i].Value.asc_getOriginSize(this);
			}
		}
		return null;
	};

	asc_docs_api.prototype.asc_FitImagesToSlide   = function()
	{
		if(!this.WordControl || !this.WordControl.m_oLogicDocument)
		{
			return;
		}
		this.WordControl.m_oLogicDocument.FitImagesToSlide();
	};

	/*callbacks*/
	asc_docs_api.prototype.sync_AddImageCallback = function()
	{
		this.sendEvent("asc_onAddImage");
	};

	asc_docs_api.prototype.getImageSelectedObject  = function(imgProp)
	{
		var type = imgProp.chartProps ? c_oAscTypeSelectElement.Chart : c_oAscTypeSelectElement.Image;
		var objects;
		if (type === c_oAscTypeSelectElement.Chart)
		{
			objects = new CAscChartProp(imgProp);
		}
		else
		{
			objects = new Asc.asc_CImgProperty(imgProp);
		}
		return new asc_CSelectedObject(type, objects);
	};
	asc_docs_api.prototype.sync_ImgPropCallback  = function(imgProp)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = this.getImageSelectedObject(imgProp);
	};

	asc_docs_api.prototype.sync_MathPropCallback = function(MathProp)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Math, MathProp);
	};

	asc_docs_api.prototype.SetDrawingFreeze = function(bIsFreeze)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpIsFreeze = bIsFreeze;
			return;
		}

		this.WordControl.DrawingFreeze = bIsFreeze;

		var _elem1 = document.getElementById("id_main");
		if (_elem1)
		{
			var _elem2 = document.getElementById("id_panel_thumbnails");
			var _elem3 = document.getElementById("id_panel_notes");
			if (bIsFreeze)
			{
				_elem1.style.display = "none";
				_elem2.style.display = "none";
				_elem3.style.display = "none";
			}
			else
			{
				_elem1.style.display = "block";
				_elem2.style.display = "block";
				_elem3.style.display = "block";
			}
		}

		if (!bIsFreeze)
			this.WordControl.OnScroll();
	};


	asc_docs_api.prototype.AddShapeOnCurrentPage = function(sPreset){
		if(!this.WordControl.m_oLogicDocument){
			return;
		}
        this.WordControl.m_oLogicDocument.AddShapeOnCurrentPage(sPreset);
	}
	asc_docs_api.prototype.can_CopyCut = function(){
		if(!this.WordControl.m_oLogicDocument){
			return false;
		}
        return this.WordControl.m_oLogicDocument.Can_CopyCut();
	}

	/*----------------------------------------------------------------*/
	/*functions for working with zoom & navigation*/
	asc_docs_api.prototype.zoomIn         = function()
	{
		this.WordControl.zoom_In();
	};
	asc_docs_api.prototype.zoomOut        = function()
	{
		this.WordControl.zoom_Out();
	};
	asc_docs_api.prototype.zoomFitToPage  = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.FitToPage;
			return;
		}
		this.WordControl.zoom_FitToPage();
	};
	asc_docs_api.prototype.zoomFitToWidth = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.FitToWidth;
			return;
		}
		this.WordControl.zoom_FitToWidth();
	};
	asc_docs_api.prototype.zoomCustomMode = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.CustomMode;
			return;
		}
		this.WordControl.m_nZoomType = 0;
		this.WordControl.zoom_Fire();
	};
	asc_docs_api.prototype.zoom100        = function()
	{
		this.WordControl.m_nZoomValue = 100;
		this.WordControl.zoom_Fire();
	};
	asc_docs_api.prototype.zoom           = function(percent)
	{
		this.WordControl.m_nZoomValue = percent;
		this.WordControl.zoom_Fire(0);
	};
	asc_docs_api.prototype.goToPage       = function(number)
	{
		this.WordControl.GoToPage(number);
	};
	asc_docs_api.prototype.getCountPages  = function()
	{
		return this.getCountSlides();
	};
	asc_docs_api.prototype.getCountSlides  = function()
	{
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation) return 0;
		return oPresentation.GetSlidesCount();
	};
	asc_docs_api.prototype.getCurrentPage = function()
	{
		return this.WordControl.m_oDrawingDocument.SlideCurrent;
	};
	/*callbacks*/
	asc_docs_api.prototype.sync_zoomChangeCallback  = function(percent, type)
	{	//c_oAscZoomType.Current, c_oAscZoomType.FitWidth, c_oAscZoomType.FitPage
		this.sendEvent("asc_onZoomChange", percent, type);
	};
	asc_docs_api.prototype.sync_countPagesCallback  = function(count)
	{
		this.sendEvent("asc_onCountPages", count);
	};
	asc_docs_api.prototype.sync_currentPageCallback = function(number)
	{
		this.sendEvent("asc_onCurrentPage", number);
	};

	asc_docs_api.prototype.sync_SendThemeColors = function(colors, standart_colors)
	{
		this.sendEvent("asc_onSendThemeColors", colors, standart_colors);
	};

	asc_docs_api.prototype.getCurrentTheme = function()
	{
		if (null == this.WordControl.m_oLogicDocument)
			return null;
		return this.WordControl.m_oLogicDocument.GetCurrentTheme();
	};

	asc_docs_api.prototype.getGraphicController = function() {
		var oPresentation = this.WordControl && this.WordControl.m_oLogicDocument;
		return oPresentation && oPresentation.GetCurrentController();
	};


	asc_docs_api.prototype.ChangeColorScheme = function(sSchemeName)
	{
		var scheme = AscCommon.getColorSchemeByName(sSchemeName);
		if (!scheme)
		{
			var theme = this.getCurrentTheme();
			if (null == theme)
				return;
			scheme = theme.getExtraClrScheme(sSchemeName);
		}
		if(!scheme)
		{
			return;
		}
		this.WordControl.m_oLogicDocument.changeColorScheme(scheme);
		this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
	};

	asc_docs_api.prototype.asc_ChangeColorSchemeByIdx = function(nIdx)
	{
		var scheme = this.getColorSchemeByIdx(nIdx);
		if(!scheme) {
			return;
		}
		this.WordControl.m_oLogicDocument.changeColorScheme(scheme);
		this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
	};

	/*----------------------------------------------------------------*/
	asc_docs_api.prototype.asc_enableKeyEvents = function(value, isFromInput)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpFocus = value;
			return;
		}

		if (this.WordControl && this.WordControl.IsFocus != value)
		{
			this.WordControl.IsFocus = value;
			this.sendEvent("asc_onEnableKeyEventsChanged", value);
		}

		if (isFromInput !== true && AscCommon.g_inputContext)
			AscCommon.g_inputContext.setInterfaceEnableKeyEvents(value);

		if(!value)
		{
			this.asc_hideMediaControl();
		}
	};


	//-----------------------------------------------------------------
	// Функции для работы с комментариями
	//-----------------------------------------------------------------
	function asc_CCommentData(obj)
	{
		if (obj)
		{
			this.m_sText      = (undefined != obj.m_sText     ) ? obj.m_sText : "";
			this.m_sTime      = (undefined != obj.m_sTime     ) ? obj.m_sTime : "";
			this.m_sOOTime    = (undefined != obj.m_sOOTime   ) ? obj.m_sOOTime : "";
			this.m_sUserId    = (undefined != obj.m_sUserId   ) ? obj.m_sUserId : "";
			this.m_sQuoteText = (undefined != obj.m_sQuoteText) ? obj.m_sQuoteText : null;
			this.m_bSolved    = (undefined != obj.m_bSolved   ) ? obj.m_bSolved : false;
			this.m_sUserName  = (undefined != obj.m_sUserName ) ? obj.m_sUserName : "";
			this.m_sGuid      = (undefined != obj.m_sGuid     ) ? obj.m_sGuid : AscCommon.CreateGUID();
			this.m_nTimeZoneBias= (undefined != obj.m_nTimeZoneBias) ? obj.m_nTimeZoneBias : null;
			this.bDocument    = (undefined != obj.bDocument   ) ? obj.bDocument : false;
			this.m_sUserData  = (undefined != obj.m_sUserData ) ? obj.m_sUserData : ""; // For compatibility with 'Word Comment Control'
			this.m_aReplies   = [];
			if (undefined != obj.m_aReplies)
			{
				var Count = obj.m_aReplies.length;
				for (var Index = 0; Index < Count; Index++)
				{
					var Reply = new asc_CCommentData(obj.m_aReplies[Index]);
					this.m_aReplies.push(Reply);
				}
			}
		}
		else
		{
			this.m_sText      = "";
			this.m_sTime      = "";
			this.m_sOOTime    = "";
			this.m_sUserId    = "";
			this.m_sQuoteText = null;
			this.m_bSolved    = false;
			this.m_sUserName  = "";
			this.m_sGuid      = AscCommon.CreateGUID();
			this.m_nTimeZoneBias =  null;
			this.bDocument    = false;
			this.m_sUserData  = "";
			this.m_aReplies   = [];
		}
	}

	asc_CCommentData.prototype.asc_getText         = function()
	{
		return this.m_sText;
	};
	asc_CCommentData.prototype.asc_putText         = function(v)
	{
		this.m_sText = v ? v.slice(0, Asc.c_oAscMaxCellOrCommentLength) : v;
	};
	asc_CCommentData.prototype.asc_getTime         = function()
	{
		return this.m_sTime;
	};
	asc_CCommentData.prototype.asc_putTime         = function(v)
	{
		this.m_sTime = v;
		this.m_nTimeZoneBias = new Date().getTimezoneOffset();
	};
	asc_CCommentData.prototype.asc_getOnlyOfficeTime         = function()
	{
		return this.m_sOOTime;
	};
	asc_CCommentData.prototype.asc_putOnlyOfficeTime         = function(v)
	{
		this.m_sOOTime = v;
	};
	asc_CCommentData.prototype.asc_getUserId       = function()
	{
		return this.m_sUserId;
	};
	asc_CCommentData.prototype.asc_putUserId       = function(v)
	{
		this.m_sUserId = v;
	};
	asc_CCommentData.prototype.asc_getUserName     = function()
	{
		return this.m_sUserName;
	};
	asc_CCommentData.prototype.asc_putUserName     = function(v)
	{
		this.m_sUserName = v;
	};
	asc_CCommentData.prototype.asc_getGuid     = function()
	{
		return this.m_sGuid;
	};
	asc_CCommentData.prototype.asc_putGuid     = function(v)
	{
		this.m_sGuid = v;
	};
	asc_CCommentData.prototype.asc_putTimeZoneBias     = function(v)
	{
		this.m_nTimeZoneBias = v;
	};
	asc_CCommentData.prototype.asc_getTimeZoneBias     = function()
	{
		return this.m_nTimeZoneBias;
	};
	asc_CCommentData.prototype.asc_getQuoteText    = function()
	{
		return this.m_sQuoteText;
	};
	asc_CCommentData.prototype.asc_putQuoteText    = function(v)
	{
		this.m_sQuoteText = v;
	};
	asc_CCommentData.prototype.asc_getSolved       = function()
	{
		return this.m_bSolved;
	};
	asc_CCommentData.prototype.asc_putSolved       = function(v)
	{
		this.m_bSolved = v;
	};
	asc_CCommentData.prototype.asc_getReply        = function(i)
	{
		return this.m_aReplies[i];
	};
	asc_CCommentData.prototype.asc_addReply        = function(v)
	{
		this.m_aReplies.push(v);
	};
	asc_CCommentData.prototype.asc_getRepliesCount = function(v)
	{
		return this.m_aReplies.length;
	};
	asc_CCommentData.prototype.asc_putDocumentFlag        = function(v)
	{
		this.bDocument = v;
	};
	asc_CCommentData.prototype.asc_getDocumentFlag = function()
	{
		return this.bDocument;
	};
	asc_CCommentData.prototype.asc_putUserData        = function(v)
	{
		this.m_sUserData = v;
	};
	asc_CCommentData.prototype.asc_getUserData = function()
	{
		return this.m_sUserData;
	};
	asc_CCommentData.prototype.fromCValue = function (value) {
		if (!value) {
			return value;
		}
		let pr = new AscWord.CSdtPictureFormPr();
		pr.asc_putText(value["Text"]);
		pr.asc_putTime(value["Time"]);
		pr.asc_putOnlyOfficeTime(value["OnlyOfficeTime"]);
		pr.asc_putUserId(value["UserId"]);
		pr.asc_putUserName(value["UserName"]);
		pr.asc_putGuid(value["Guid"]);
		pr.asc_putTimeZoneBias(value["TimeZoneBias"]);
		pr.asc_putQuoteText(value["QuoteText"]);
		pr.asc_putSolved(value["Solved"]);

		if (Array.isArray(value["asc_getReplies"])) {
			for (let nIdx = 0; nIdx < value["asc_getReplies"].length; ++nIdx) {
				let reply = Asc.asc_CCommentData.prototype.fromCValue(value["asc_getReplies"][nIdx]);
				if (reply) {
					this.asc_addReply(reply);
				}
			}
		}
		pr.asc_putDocumentFlag(value["DocumentFlag"]);
		pr.asc_putUserData(value["UserData"]);
		return pr;
	};
	asc_CCommentData.prototype.toCValue = function () {
		let value = {};
		value["Text"] = this.asc_getText();
		value["Time"] = this.asc_getTime();
		value["OnlyOfficeTime"] = this.asc_getOnlyOfficeTime();
		value["UserId"] = this.asc_getUserId();
		value["UserName"] = this.asc_getUserName();
		value["Guid"] = this.asc_getGuid();
		value["TimeZoneBias"] = this.asc_getTimeZoneBias();
		value["QuoteText"] = this.asc_getQuoteText();
		value["Solved"] = this.asc_getSolved();
		value["asc_getReplies"] = [];
		let Count = this.m_aReplies.length;
		for (let nIdx = 0; nIdx < Count; nIdx++) {
			let val = this.m_aReplies[nIdx].toCValue();
			if (val) {
				value["asc_getReplies"].push(val);
			}
		}
		value["DocumentFlag"] = this.asc_getDocumentFlag();
		value["UserData"] = this.asc_getUserData();
		return value;
	};

	asc_docs_api.prototype.asc_showComments = function()
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oLogicDocument.ShowComments();
	};

	asc_docs_api.prototype.asc_hideComments = function()
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oLogicDocument.HideComments();
		editor.sync_HideComment();
	};

	asc_docs_api.prototype.asc_addComment = function(AscCommentData)
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//   return;

		if (null == this.WordControl.m_oLogicDocument) {
			return;
		}

		var CommentData = new AscCommon.CCommentData();
		CommentData.Read_FromAscCommentData(AscCommentData);

		var Comment = this.WordControl.m_oLogicDocument.AddComment(CommentData, AscCommentData.asc_getDocumentFlag());
		if (Comment) {
			return Comment.Get_Id();
		}
	};

	asc_docs_api.prototype.asc_getMasterCommentId = function()
	{
		return -1;
	};

	asc_docs_api.prototype.asc_getAnchorPosition = function()
	{
		var AnchorPos = this.WordControl.m_oLogicDocument.GetSelectionAnchorPos();
		return new AscCommon.asc_CRect(AnchorPos.X0, AnchorPos.Y, AnchorPos.X1 - AnchorPos.X0, 0);
	};

	asc_docs_api.prototype.asc_removeComment = function(Id)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		var comment = g_oTableId.Get_ById(Id);
		if(!comment)
		{
			return;
		}
		var oComments = comment.Parent;
		if(!oComments)
		{
			return;
		}
		var bPresComments = (oComments === this.WordControl.m_oLogicDocument.comments);
		var oCheckData = {
			comment: comment,
			slide: bPresComments ? null : oComments.slide
		};
		if (this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_MoveComment, [oCheckData], this.WordControl.m_oLogicDocument.IsEditCommentsMode()) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_RemoveComment);
			this.WordControl.m_oLogicDocument.RemoveComment(Id, true);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};


	//Remove All comments
	asc_docs_api.prototype.asc_RemoveAllComments = function(isMine, isCurrent)
	{

		if (!this.WordControl.m_oLogicDocument)
		{
			return;
		}
		if(isCurrent)
		{
			this.WordControl.m_oLogicDocument.RemoveCurrentComment(isMine);
		}
		else
		{
			if(isMine)
			{
				this.WordControl.m_oLogicDocument.RemoveMyComments();
			}
			else
			{
				this.WordControl.m_oLogicDocument.RemoveAllComments();
			}
		}
	};

	asc_docs_api.prototype.asc_ResolveAllComments = function(isMine, isCurrent, arrIds)
	{
		var oPresentation = this.WordControl && this.WordControl.m_oLogicDocument;
		if(!oPresentation)
		{
			return;
		}
		oPresentation.ResolveAllComments(isMine, isCurrent, arrIds);
	};

	asc_docs_api.prototype.asc_GetCommentLogicPosition = function()
	{
		return -1;
	};

	asc_docs_api.prototype.asc_onDeleteComment = function(Id, oCommentData)
	{
		var AscCommentData = new asc_CCommentData(oCommentData);
		this.sendEvent("asc_onDeleteComment", Id, AscCommentData);
	};

	asc_docs_api.prototype.asc_changeComment = function(Id, AscCommentData)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		var CommentData = new AscCommon.CCommentData();
		CommentData.Read_FromAscCommentData(AscCommentData);
		this.WordControl.m_oLogicDocument.EditComment(Id, CommentData);
	};

	asc_docs_api.prototype.asc_selectComment = function(Id)
	{
		if (null == this.WordControl.m_oLogicDocument)
			return;

		this.WordControl.m_oLogicDocument.SelectComment(Id);
	};

	asc_docs_api.prototype.asc_showComment = function(Id)
	{
		this.WordControl.m_oLogicDocument.ShowComment(Id);
	};

	asc_docs_api.prototype.can_AddQuotedComment = function()
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//    return false;

		return this.WordControl.m_oLogicDocument.CanAddComment();
	};

	asc_docs_api.prototype.sync_RemoveComment = function(Id)
	{
		this.sendEvent("asc_onRemoveComment", Id);
	};

	asc_docs_api.prototype.sync_AddComment = function(Id, CommentData)
	{
		if (this.bNoSendComments === false)
		{
			var AscCommentData = new asc_CCommentData(CommentData);
			AscCommentData.asc_putQuoteText("");
			this.sendEvent("asc_onAddComment", Id, AscCommentData);
		}
	};

	asc_docs_api.prototype.sync_ShowComment = function(Id, X, Y)
	{
		/*
		 if (this.WordControl.m_oMainContent)
		 {
		 X -= ((this.WordControl.m_oMainContent.Bounds.L * g_dKoef_mm_to_pix) >> 0);
		 }
		 */
		// TODO: Переделать на нормальный массив
		this.sendEvent("asc_onShowComment", [Id], X, Y);
	};

	asc_docs_api.prototype.sync_HideComment = function()
	{
		this.sendEvent("asc_onHideComment");
	};

	asc_docs_api.prototype.sync_UpdateCommentPosition = function(Id, X, Y)
	{
		// TODO: Переделать на нормальный массив
		this.sendEvent("asc_onUpdateCommentPosition", [Id], X, Y);
	};

	asc_docs_api.prototype.sync_ChangeCommentData = function(Id, CommentData)
	{
		var AscCommentData = new asc_CCommentData(CommentData);
		this.sendEvent("asc_onChangeCommentData", Id, AscCommentData);
	};

	asc_docs_api.prototype.sync_LockComment = function(Id, UserId)
	{
		this.sendEvent("asc_onLockComment", Id, UserId);
	};

	asc_docs_api.prototype.sync_UnLockComment = function(Id)
	{
		this.sendEvent("asc_onUnLockComment", Id);
	};


	asc_docs_api.prototype._goToComment = function(data)
	{
		if (this.WordControl && this.WordControl.m_oLogicDocument && data)
		{
			var commentId = this.WordControl.m_oLogicDocument.GetCommentIdByGuid(data);
			if (commentId) {
				this.asc_selectComment(commentId);
				this.asc_showComment(commentId);
			}
		}
	};


	// работа с шрифтами
	asc_docs_api.prototype.asyncFontsDocumentStartLoaded = function(blockType)
	{
		// здесь прокинуть евент о заморозке меню
		// и нужно вывести информацию в статус бар
		if (this.isPasteFonts_Images)
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadFont);
		else
		{
			this.sync_StartAction(undefined === blockType ? c_oAscAsyncActionType.BlockInteraction : blockType, c_oAscAsyncAction.LoadDocumentFonts);

			// заполним прогресс
			var _progress         = this.OpenDocumentProgress;
			_progress.Type        = c_oAscAsyncAction.LoadDocumentFonts;
			_progress.FontsCount  = this.FontLoader.fonts_loading.length;
			_progress.CurrentFont = 0;

			var _loader_object = this.WordControl.m_oLogicDocument;
			var _count         = 0;
			if (_loader_object !== undefined && _loader_object != null)
			{
				for (var i in _loader_object.ImageMap)
				{
					if (this.DocInfo.get_OfflineApp())
					{
						var localUrl = _loader_object.ImageMap[i];
						g_oDocumentUrls.addImageUrl(localUrl, this.documentUrl + 'media/' + localUrl);
					}
					++_count;
				}
			}

			_progress.ImagesCount  = _count + AscCommon.g_oUserTexturePresets.length;
			_progress.CurrentImage = 0;
		}
	};
	asc_docs_api.prototype.GenerateStyles                = function()
	{
		return;
	};
	asc_docs_api.prototype.asyncFontsDocumentEndLoaded   = function(blockType)
	{
		// все, шрифты загружены. Теперь нужно подгрузить картинки
		if (this.isPasteFonts_Images)
			this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadFont);
		else
			this.sync_EndAction(undefined === blockType ? c_oAscAsyncActionType.BlockInteraction : blockType, c_oAscAsyncAction.LoadDocumentFonts);

        if (undefined !== this.asyncMethodCallback)
        {
            this.asyncMethodCallback();
            this.asyncMethodCallback = undefined;
            return;
        }

		this.EndActionLoadImages = 0;
		if (this.isPasteFonts_Images)
		{
			var _count = 0;
			for (var i in this.pasteImageMap)
				++_count;

			if (_count > 0)
			{
				this.EndActionLoadImages = 2;
				this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
			}

			this.ImageLoader.LoadDocumentImages(this.pasteImageMap);
			return;
		}
		else if (this.isSaveFonts_Images)
		{
			var _count = 0;
			for (var i in this.saveImageMap)
				++_count;

			if (_count > 0)
			{
				this.EndActionLoadImages = 2;
				this.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
			}

			this.ImageLoader.LoadDocumentImages(this.saveImageMap);
			return;
		}

		this.GenerateStyles();
		// открытие после загрузки документа

		var _loader_object = this.WordControl.m_oLogicDocument;
		if (null == _loader_object)
			_loader_object = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer;

		var _count = 0;
		for (var i in _loader_object.ImageMap)
			++_count;

		// add const textures
		var _st_count = AscCommon.g_oUserTexturePresets.length;
		for (var i = 0; i < _st_count; i++)
			_loader_object.ImageMap[_count + i] = AscCommon.g_oUserTexturePresets[i];

		_loader_object.ImageMap[_count + _st_count] = AscFormat.ICON_TRIGGER;
		if (_count > 0)
		{
			this.EndActionLoadImages = 1;
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadDocumentImages);
		}

		this.ImageLoader.bIsLoadDocumentFirst = true;
		this.ImageLoader.LoadDocumentImages(_loader_object.ImageMap);
	};
	asc_docs_api.prototype.asyncImagesDocumentEndLoaded  = function()
	{
		this.ImageLoader.bIsLoadDocumentFirst = false;
		var _bIsOldPaste                      = this.isPasteFonts_Images;

		if (this.EndActionLoadImages == 1)
		{
			this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadDocumentImages);
		}
		else if (this.EndActionLoadImages == 2)
		{
			if (_bIsOldPaste)
				this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadImage);
			else
				this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadImage);
		}

		this.EndActionLoadImages = 0;

		// размораживаем меню... и начинаем считать документ
		if (this.isPasteFonts_Images)
		{
			this.isPasteFonts_Images = false;
			this.pasteImageMap       = null;
			this.pasteCallback();
			this.pasteCallback            = null;
			this.decrementCounterLongAction();
		}
		else if (this.isSaveFonts_Images)
		{
			this.isSaveFonts_Images = false;
			this.saveImageMap       = null;
			this.pre_SaveCallback();
		}
		else
		{
			this.ServerImagesWaitComplete = true;
			this._openDocumentEndCallback();
		}
	};

	asc_docs_api.prototype.IsAsyncOpenDocumentImages     = function()
	{
		return !this.isReporterMode && !this.asc_IsStartDemonstrationOnOpen();
	};
	asc_docs_api.prototype.asyncImagesDocumentStartLoaded      = function(aImages)
	{
		let oDrawingDocument = this.WordControl.m_oDrawingDocument;
		if(oDrawingDocument && oDrawingDocument.OnStartImagesLoading)
		{
			oDrawingDocument.OnStartImagesLoading(aImages);
		}
	};
	asc_docs_api.prototype.asyncImageEndLoadedBackground = function(_image)
	{
		this.WordControl.m_oDrawingDocument.CheckRasterImageOnScreen(_image.src);
	};

	asc_docs_api.prototype._openDocumentEndCallback = function()
	{
		if (this.isDocumentLoadComplete || !this.ServerImagesWaitComplete || !this.ServerIdWaitComplete ||
			!this.WordControl || !this.WordControl.m_oLogicDocument)
			return;

		var bIsScroll = false;

		if (0 == this.DocumentType)
			this.WordControl.m_oLogicDocument.LoadEmptyDocument();
		else
		{
			if (this.LoadedObject)
			{
				if (this.LoadedObject === 1)
				{
					if (this.isApplyChangesOnOpenEnabled)
					{
                        if (AscCommon.EncryptionWorker)
                        {
                            AscCommon.EncryptionWorker.init();
                            if (!AscCommon.EncryptionWorker.isChangesHandled)
                            	return AscCommon.EncryptionWorker.handleChanges(AscCommon.CollaborativeEditing.m_aChanges, this, this._openDocumentEndCallback);
                        }

						this.isApplyChangesOnOpenEnabled = false;
						this.bNoSendComments             = true;
						var OtherChanges                 = AscCommon.CollaborativeEditing.m_aChanges.length > 0;
						this._applyPreOpenLocks();
						let perfStart = performance.now();
						AscCommon.CollaborativeEditing.Apply_Changes();
						let perfEnd = performance.now();
						if (OtherChanges) {
							AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onApplyChanges", perfEnd - perfStart), this);
						}
						AscCommon.CollaborativeEditing.Release_Locks();
						this.bNoSendComments      = false;
						this.isApplyChangesOnOpen = true;
						if(OtherChanges && this.isSaveFonts_Images){
							return;
						}
					}
				}
				this.WordControl.m_oLogicDocument.Recalculate({Drawings : {All : true, Map : {}}});
				AscCommon.History.private_ClearRecalcData();
				var presentation = this.WordControl.m_oLogicDocument;

				presentation.DrawingDocument.OnEndRecalculate();

				if(!window['IS_NATIVE_EDITOR']) {
					this.asc_registerCallback('asc_doubleClickOnChart', function(){
						// next tick
						setTimeout(function() {
							window.editor.WordControl.onMouseUpMainSimple();
						}, 0);
					});
				}

				if(!window["NATIVE_EDITOR_ENJINE"]){
					this.WordControl.m_oLayoutDrawer.WidthMM  = presentation.GetWidthMM();
					this.WordControl.m_oLayoutDrawer.HeightMM = presentation.GetHeightMM();
					this.WordControl.m_oMasterDrawer.WidthMM  = presentation.GetWidthMM();
					this.WordControl.m_oMasterDrawer.HeightMM = presentation.GetHeightMM();
				}
                this.standartThemesStatus++;
                if (2 < this.standartThemesStatus)
                   this.WordControl.m_oLogicDocument.SendThemesThumbnails();

				this.sendEvent("asc_onPresentationSize", presentation.GetWidthEMU(), presentation.GetHeightEMU(), presentation.GetSizeType(), presentation.getFirstSlideNumber());

				this.WordControl.GoToPage(0);
				bIsScroll = true;
			}
		}


		this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		this.WordControl.m_oLogicDocument.Document_UpdateRulersState();
		this.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
		this.LoadedObject       = null;
		this.bInit_word_control = true;
		if (!this.bNoSendComments)
		{
			var _slides      = this.WordControl.m_oLogicDocument.Slides;
			var _slidesCount = _slides.length;
			for (var i = 0; i < _slidesCount; i++)
			{
				var slideComments = _slides[i].slideComments;
				if (slideComments)
				{
					var _comments      = slideComments.comments;
					var _commentsCount = _comments.length;
					for (var j = 0; j < _commentsCount; j++)
					{
						this.sync_AddComment(_comments[j].Get_Id(), _comments[j].Data);
					}
				}
			}
		}
		var slideComments = this.WordControl.m_oLogicDocument.comments;
		if (slideComments)
		{
			var _comments      = slideComments.comments;
			var _commentsCount = _comments.length;
			for (var j = 0; j < _commentsCount; j++)
			{
				_comments[j].Data.bDocument = true;
				this.sync_AddComment(_comments[j].Get_Id(), _comments[j].Data);
			}
		}
		this.onDocumentContentReady();

		if(this.asc_IsStartDemonstrationOnOpen())
		{
			this.sendEvent("asc_onStartDemonstration");
		}
		this.isApplyChangesOnOpen = false;

		this.WordControl.InitControl();
		if (bIsScroll)
		{
			this.WordControl.OnScroll();
		}

		if (!this.isViewMode)
		{
			this.sendStandartTextures();
			this.sendMathToMenu();
			if (this.shapeElementId)
			{
				this.WordControl.m_oDrawingDocument.InitGuiCanvasShape(this.shapeElementId);
			}
		}

		if (this.isViewMode)
			this.asc_setViewMode(true);

		// Меняем тип состояния (на никакое)
		this.advancedOptionsAction = AscCommon.c_oAscAdvancedOptionsAction.None;
		this.goTo();
	};
	asc_docs_api.prototype.asc_IsStartDemonstrationOnOpen = function()
	{
		if(this.documentFormat === "ppsx")
		{
			return true;
		}
		if (this.documentTitle && this.documentTitle.endsWith(".ppsx"))
		{
			return true;
		}
		return false;
	};

	asc_docs_api.prototype.asc_AddMath = function(Type)
	{
		var loader   = AscCommon.g_font_loader;
		var fontinfo = AscFonts.g_fontApplication.GetFontInfo("Cambria Math");
		var isasync  = loader.LoadFont(fontinfo);
		if (false === isasync)
		{
			return this.asc_AddMath2(Type);
		}
		else
		{
			this.asyncMethodCallback = function()
			{
				return this.asc_AddMath2(Type);
			}
		}
	};

	asc_docs_api.prototype.asc_AddMath2 = function(Type)
	{
		let logicDocument = this.WordControl.m_oLogicDocument
		if (false === logicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			let textPr = logicDocument.GetDirectTextPr();
			logicDocument.StartAction(AscDFH.historydescription_Document_AddMath);
			var MathElement = new AscCommonWord.MathMenu(Type, textPr ? textPr.Copy() : null);
			logicDocument.AddToParagraph(MathElement, false);
			logicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_ConvertMathView = function(isToLinear, isAll)
	{
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation)
		{
			return;
		}
		oPresentation.ConvertMathView(isToLinear, isAll);
	};

	asc_docs_api.prototype.asc_AddVideoCallback = function(sImageUrlLocal, sVideoUrl, obj)
	{
		this.addMediaCallback(sImageUrlLocal, obj, sVideoUrl, "localVideo");
	};
	asc_docs_api.prototype.asc_AddAudioCallback = function(sImageUrlLocal, sAudioUrl, obj)
	{
		this.addMediaCallback(sImageUrlLocal, obj, sAudioUrl, "localAudio");
	};



	asc_docs_api.prototype.asc_AddVideoUrlCallback = function(sImageUrlLocal, sVideoUrl, obj)
	{
		this.addMediaCallback(sImageUrlLocal, obj, sVideoUrl, "linkVideo");
	};
	asc_docs_api.prototype.asc_AddAudioUrlCallback = function(sImageUrlLocal, sAudioUrl, obj)
	{
		this.addMediaCallback(sImageUrlLocal, obj, sAudioUrl, "linkAudio");
	};

	asc_docs_api.prototype.addMediaCallback = function(sImageUrlLocal, oPlaceholder, sLink, sType)
	{
		let oApi = this;
		let sImageUrl = AscCommon.g_oDocumentUrls.getImageUrl(sImageUrlLocal);
		this.ImageLoader.LoadImagesWithCallback([sImageUrl], function(){
			let _image = oApi.ImageLoader.LoadImage(sImageUrl, 1);
			if (!_image || !_image.Image)
				return;

			let oImageObject = {};
			oImageObject.src = sImageUrl;
			oImageObject.Image = {};
			oImageObject.Image.width = 50;
			oImageObject.Image.height = 50;
			if(sType === "localVideo") {
				oImageObject.videoUrl = sLink;
			}
			else if(sType === "localAudio") {
				oImageObject.audioUrl = sLink;
			}
			else if(sType === "linkVideo") {
				oImageObject.videoLink = sLink;
			}
			else if(sType === "linkAudio") {
				oImageObject.audioLink = sLink;
			}
			oApi.WordControl.m_oLogicDocument.addImages([oImageObject], oPlaceholder);
		});
	};


	asc_docs_api.prototype.asc_AddVideoUrl = function(sImageUrlLocal, sVideoUrl, obj)
	{
		this.addMediaCallback(sImageUrlLocal, obj, sVideoUrl, "linkVideo");
	};
	asc_docs_api.prototype.asc_AddAudioUrl = function(sImageUrlLocal, sAudioUrl, obj)
	{
		this.addMediaCallback(sImageUrlLocal, obj, sAudioUrl, "linkAudio");
	};

	asc_docs_api.prototype.closeDemonstration = function(bSaveAnnotations, isNoUseFullScreen)
	{
		let oAnnotations = this.getAnnotations();
		if(oAnnotations)
		{
			if(bSaveAnnotations)
			{
				oAnnotations.saveAnnotations();
			}
			else
			{
				oAnnotations.clear();
			}
		}
		this.endDemoMode(isNoUseFullScreen);
	};

	asc_docs_api.prototype.endDemoMode = function(isNoUseFullScreen)
	{

		if (this.windowReporter)
			this.windowReporter.close();

		this.WordControl.DemonstrationManager.End(isNoUseFullScreen);
	};

	//----------------------------------------------------------------------------------------------------------------------
	// Работаем с формулами
	//----------------------------------------------------------------------------------------------------------------------
	asc_docs_api.prototype.asc_SetMathProps = function(MathProps)
	{
		this.WordControl.m_oLogicDocument.Set_MathProps(MathProps);
	};

	asc_docs_api.prototype.asc_SetHFProps = function(HFProps, bAll)
	{
		if(this.WordControl && this.WordControl.m_oLogicDocument)
		{
			this.WordControl.m_oLogicDocument.setHFProperties(HFProps, bAll);
		}
	};

	asc_docs_api.prototype.asyncFontEndLoaded = function(fontinfo)
	{
		this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);

		if (undefined !== this.asyncMethodCallback)
		{
			this.asyncMethodCallback();
			this.asyncMethodCallback = undefined;
			return;
		}

		if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) === false)
		{
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ParagraphAdd);
			this.WordControl.m_oLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({
				FontFamily : {
					Name  : fontinfo.Name,
					Index : -1
				}
			}), false);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.asc_replaceLoadImageCallback = function(fCallback)
	{
		this.asyncImageEndLoaded2 = fCallback;
	};

	asc_docs_api.prototype.asyncImageEndLoaded = function(_image, placeholder)
	{
		// отжать заморозку меню
		if (this.asyncImageEndLoaded2)
			this.asyncImageEndLoaded2(_image, placeholder);
		else
		{
			this.WordControl.m_oLogicDocument.addImages([_image], placeholder);
		}
	};

	asc_docs_api.prototype.openDocument = function(file)
	{
		let perfStart = performance.now();
		if (file.changes && this.VersionHistory)
		{
			this.VersionHistory.changes = file.changes;
			this.VersionHistory.applyChanges(this);
		}
		this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(file.data);
		if (this.isOpenOOXInBrowser) {
			this.openOOXInBrowserZip = file.data;
			this.OpenDocumentFromZip(file.data);
		} else {
			this.OpenDocumentFromBin(file.url, file.data);
		}
		let perfEnd = performance.now();
		AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onOpenDocument", perfEnd - perfStart), this);
	};

	asc_docs_api.prototype.get_PresentationWidth  = function()
	{
		if (this.WordControl.m_oLogicDocument == null)
			return 0;
		return this.WordControl.m_oLogicDocument.GetWidthEMU();
	};
	asc_docs_api.prototype.get_PresentationHeight = function()
	{
		if (this.WordControl.m_oLogicDocument == null)
			return 0;
		return this.WordControl.m_oLogicDocument.GetHeightEMU();
	};

	asc_docs_api.prototype.pre_Paste = function(_fonts, _images, callback)
	{
		if (undefined !== window["native"] && undefined !== window["native"]["GetImageUrl"])
		{
			callback();
			return;
		}

		if(window['IS_NATIVE_EDITOR'])
		{
			callback();
			return;
		}
		this.pasteCallback = callback;
		this.pasteImageMap = _images;

		var _count = 0;
		for (var i in this.pasteImageMap)
			++_count;

        AscFonts.FontPickerByCharacter.extendFonts(_fonts);
		if (0 == _count && false === this.FontLoader.CheckFontsNeedLoading(_fonts))
		{
			// никаких евентов. ничего грузить не нужно. сделано для сафари под макОс.
			// там при LongActions теряется фокус и вставляются пробелы
			this.pasteCallback();
			this.pasteCallback            = null;
			return;
		}

		this.incrementCounterLongAction();
		this.isPasteFonts_Images = true;
		this.FontLoader.LoadDocumentFonts2(_fonts);
	};

	asc_docs_api.prototype.pre_SaveCallback = function()
	{
		AscCommon.CollaborativeEditing.OnEnd_Load_Objects();

		if (this.isApplyChangesOnOpen)
		{
			this.isApplyChangesOnOpen = false;
			this._openDocumentEndCallback();
		}
		if (this.isApplyChangesOnVersionHistory)
		{
			this.isApplyChangesOnVersionHistory = false;
			this._openVersionHistoryEndCallback();
		}

		this.WordControl.SlideDrawer.CheckRecalculateSlide();
	};

	asc_docs_api.prototype.ViewScrollToX             = function(x)
	{
		this.WordControl.m_oScrollHorApi.scrollToX(x);
	};
	asc_docs_api.prototype.ViewScrollToY             = function(y)
	{
		this.WordControl.m_oScrollVerApi.scrollToY(y);
	};
	asc_docs_api.prototype.GetDocWidthPx             = function()
	{
		return this.WordControl.m_dDocumentWidth;
	};
	asc_docs_api.prototype.GetDocHeightPx            = function()
	{
		return this.WordControl.m_dDocumentHeight;
	};
	asc_docs_api.prototype.ClearSearch               = function()
	{
		return this.WordControl.m_oDrawingDocument.EndSearch(true);
	};
	asc_docs_api.prototype.GetCurrentVisiblePage     = function()
	{
		return this.WordControl.m_oDrawingDocument.SlideCurrent;
	};

	asc_docs_api.prototype.asc_SetDocumentPlaceChangedEnabled = function(bEnabled)
	{
		if (this.WordControl)
			this.WordControl.m_bDocumentPlaceChangedEnabled = bEnabled;
	};

	asc_docs_api.prototype.asc_SetViewRulers       = function(bRulers)
	{
		//if (false === this.bInit_word_control || true === this.isViewMode)
		//    return;

		if (!this.isLoadFullApi)
		{
			this.tmpViewRulers = bRulers;
			return;
		}

		if (this.WordControl.m_bIsRuler != bRulers)
		{
			this.WordControl.m_bIsRuler = bRulers;
			this.WordControl.checkNeedRules();
			this.WordControl.OnResize(true);
		}
	};
	asc_docs_api.prototype.asc_SetViewRulersChange = function()
	{
		//if (false === this.bInit_word_control || true === this.isViewMode)
		//    return;

		this.WordControl.m_bIsRuler = !this.WordControl.m_bIsRuler;
		this.WordControl.checkNeedRules();
		this.WordControl.OnResize(true);
		return this.WordControl.m_bIsRuler;
	};
	asc_docs_api.prototype.asc_GetViewRulers       = function()
	{
		return this.WordControl.m_bIsRuler;
	};
	asc_docs_api.prototype.asc_SetDocumentUnits    = function(_units)
	{
		if (this.WordControl && this.WordControl.m_oHorRuler && this.WordControl.m_oVerRuler)
		{
			this.WordControl.m_oHorRuler.Units = _units;
			this.WordControl.m_oVerRuler.Units = _units;
			this.WordControl.UpdateHorRulerBack(true);
			this.WordControl.UpdateVerRulerBack(true);
		}
		else
		{
            this.tmpDocumentUnits = _units;
		}
	};

	asc_docs_api.prototype.changeSlideSize = function(width, height, type, nFirstSlideNum)
	{
		if (this.isMobileVersion && this.WordControl.MobileTouchManager)
			this.WordControl.MobileTouchManager.BeginZoomCheck();

		this.WordControl.m_oLogicDocument.changeSlideSize(width, height, type, nFirstSlideNum);

		if (this.isMobileVersion && this.WordControl.MobileTouchManager)
			this.WordControl.MobileTouchManager.EndZoomCheck();
	};
	asc_docs_api.prototype.asc_getFirstSlideNumber = function()
	{
		const oPresentation = this.private_GetLogicDocument();
		if(!oPresentation)
		{
			return 1;
		}
		return oPresentation.getFirstSlideNumber();
	};

	asc_docs_api.prototype.AddSlide       = function(layoutIndex)
	{
		this.WordControl.m_oLogicDocument.addNextSlide(layoutIndex);
	};
	asc_docs_api.prototype.DeleteSlide    = function()
	{
		var _delete_array = this.WordControl.m_oLogicDocument.GetSelectedSlides();

		if (!this.IsSupportEmptyPresentation)
		{
			if (_delete_array.length == this.WordControl.GetSlidesCount())
				_delete_array.splice(0, 1);
		}

		if (_delete_array.length != 0)
		{
			this.WordControl.m_oLogicDocument.deleteSlides(_delete_array);
		}
	};
	asc_docs_api.prototype.DublicateSlide = function()
	{
		this.WordControl.m_oLogicDocument.DublicateSlide();
	};

	asc_docs_api.prototype.asc_DeleteMaster    = function()
	{
		this.WordControl.m_oLogicDocument.deleteMaster();
	};
	asc_docs_api.prototype.asc_DeleteLayout    = function()
	{
		this.DeleteSlide();
	};
	asc_docs_api.prototype.asc_CanDeleteMaster    = function()
	{
		return this.WordControl.m_oLogicDocument.CanDeleteSelectedMasters();
	};
	asc_docs_api.prototype.asc_CanDeleteLayout    = function()
	{
		return this.WordControl.m_oLogicDocument.CanDeleteSelectedLayouts();
	};
	asc_docs_api.prototype.asc_DuplicateMaster    = function()
	{
		this.WordControl.m_oLogicDocument.DublicateMaster();
	};
	asc_docs_api.prototype.asc_DuplicateLayout    = function()
	{
		this.WordControl.m_oLogicDocument.DublicateLayout();
	};

	asc_docs_api.prototype.SelectAllSlides = function(layoutType)
	{
		var drDoc       = this.WordControl.m_oDrawingDocument;
		var slidesCount = drDoc.GetSlidesCount();

		for (var i = 0; i < slidesCount; i++)
		{
			this.WordControl.Thumbnails.m_arrPages[i].IsSelected = true;
		}
		this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		this.WordControl.Thumbnails.OnUpdateOverlay();
	};
	asc_docs_api.prototype.getThumbnailsManager = function()
	{
		if(!this.WordControl || !this.WordControl.Thumbnails)
			return null;
		return this.WordControl.Thumbnails;
	};

	asc_docs_api.prototype.asc_moveSelectedSlidesToEnd = function()
	{
		if(!this.WordControl || !this.WordControl.m_oLogicDocument)
		{
			return;
		}
		this.WordControl.m_oLogicDocument.moveSelectedSlidesToEnd();
	};
	asc_docs_api.prototype.asc_moveSlidesNextPos = function()
	{
		if(!this.WordControl || !this.WordControl.m_oLogicDocument)
		{
			return;
		}
		this.WordControl.m_oLogicDocument.moveSlidesNextPos();
	};
	asc_docs_api.prototype.asc_moveSelectedSlidesToStart = function()
	{
		if(!this.WordControl || !this.WordControl.m_oLogicDocument)
		{
			return;
		}
		this.WordControl.m_oLogicDocument.moveSelectedSlidesToStart();
	};
	asc_docs_api.prototype.asc_moveSlidesPrevPos = function()
	{
		if(!this.WordControl || !this.WordControl.m_oLogicDocument)
		{
			return;
		}
		this.WordControl.m_oLogicDocument.moveSlidesPrevPos();
	};

	asc_docs_api.prototype.asc_IsSlideSelected = function(nIdx)
	{
		if(!this.WordControl || !this.WordControl.m_oLogicDocument)
		{
			return false;
		}
		var aSelected = this.WordControl.m_oLogicDocument.GetSelectedSlides();
		for(var nSld = 0; nSld < aSelected.length; ++nSld) {
			if(aSelected[nSld] === nIdx) {
				return true;
			}
		}
		return false;
	};
	asc_docs_api.prototype.asc_IsFirstSlideSelected = function()
	{
		return this.asc_IsSlideSelected(0);
	};
	asc_docs_api.prototype.asc_IsLastSlideSelected = function()
	{
		return this.asc_IsSlideSelected(this.getCountPages() - 1);
	};

	asc_docs_api.prototype.AddShape        = function(shapetype)
	{
	};
	asc_docs_api.prototype.ChangeShapeType = function(shapetype)
	{
		this.WordControl.m_oLogicDocument.changeShapeType(shapetype);
	};
	asc_docs_api.prototype.AddText         = function()
	{
	};

    asc_docs_api.prototype["asc_IsSpellCheckCurrentWord"]  = function()
    {
        return this.IsSpellCheckCurrentWord;
    };
    asc_docs_api.prototype["asc_putSpellCheckCurrentWord"] = function(value)
    {
        this.IsSpellCheckCurrentWord = value;
    };

	asc_docs_api.prototype.groupShapes = function()
	{
		this.WordControl.m_oLogicDocument.groupShapes();
	};

	asc_docs_api.prototype.unGroupShapes = function()
	{
		this.WordControl.m_oLogicDocument.unGroupShapes();
	};

	asc_docs_api.prototype.asc_canMergeSelectedShapes = function (operation) {
		return AscFormat.canMergeSelectedShapes(operation);
	};
	asc_docs_api.prototype.asc_mergeSelectedShapesAction = function (operation) {
		const isSelectionLocked = this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props);
		const canMerge = this.asc_canMergeSelectedShapes(operation);
		if (!isSelectionLocked && canMerge) {
			this.WordControl.m_oLogicDocument.StartAction(0);
			AscFormat.mergeSelectedShapes(operation);
			this.WordControl.m_oLogicDocument.FinalizeAction();
		}
	};

	asc_docs_api.prototype.setVerticalAlign = function(align)
	{
		this.WordControl.m_oLogicDocument.setVerticalAlign(align);
	};

	asc_docs_api.prototype.setVert = function(vert)
	{
		this.WordControl.m_oLogicDocument.setVert(vert);
	};

	asc_docs_api.prototype.sync_MouseMoveStartCallback = function()
	{
		this.sendEvent("asc_onMouseMoveStart");
	};

	asc_docs_api.prototype.sync_MouseMoveEndCallback = function()
	{
		this.sendEvent("asc_onMouseMoveEnd");
	};

	asc_docs_api.prototype.sync_MouseMoveCallback = function(Data)
	{
		if (Data.Hyperlink && typeof Data.Hyperlink.Value === "string")
		{
			var Url       = Data.Hyperlink.Value;
			if (AscCommon.IsLinkPPAction(Url))
			{
				if (Url == "ppaction://hlinkshowjump?jump=firstslide")
				{
					Data.Hyperlink.Value = AscCommon.translateManager.getValue("First Slide");
				}
				else if (Url == "ppaction://hlinkshowjump?jump=lastslide")
				{
					Data.Hyperlink.Value = AscCommon.translateManager.getValue("Last Slide");
				}
				else if (Url == "ppaction://hlinkshowjump?jump=nextslide")
				{
					Data.Hyperlink.Value = AscCommon.translateManager.getValue("Next Slide");
				}
				else if (Url == "ppaction://hlinkshowjump?jump=previousslide")
				{
					Data.Hyperlink.Value = AscCommon.translateManager.getValue("Previous Slide");
				}
				else
				{
					let mask     = "ppaction://hlinksldjumpslide";
					let indSlide = Url.indexOf(mask);
					if (0 === indSlide)
					{
						let slideNum = parseInt(Url.substring(mask.length));
						Data.Hyperlink.Value = AscCommon.translateManager.getValue("Slide") + " " + (slideNum + 1);
					}
				}
			}
		}

		this.sendEvent("asc_onMouseMove", Data);
	};

	asc_docs_api.prototype.sync_ShowForeignCursorLabel = function(UserId, X, Y, Color)
	{
		if (this.WordControl.m_oLogicDocument.IsFocusOnNotes())
		{
			Y += parseInt(this.WordControl.m_oNotesContainer.HtmlElement.style.top);
		}
		this.sendEvent("asc_onShowForeignCursorLabel", UserId, X, Y, new AscCommon.CColor(Color.r, Color.g, Color.b, 255));
	};
	asc_docs_api.prototype.sync_HideForeignCursorLabel = function(UserId)
	{
		this.sendEvent("asc_onHideForeignCursorLabel", UserId);
	};

	asc_docs_api.prototype.ShowThumbnails           = function(bIsShow)
	{
		if (bIsShow)
		{
			this.WordControl.Splitter1Pos = this.WordControl.OldSplitter1Pos;
			if (this.WordControl.Splitter1Pos == 0)
				this.WordControl.Splitter1Pos = 70;
			this.WordControl.OnResizeSplitter();
		}
		else
		{
			var old                       = this.WordControl.OldSplitter1Pos;
			this.WordControl.Splitter1Pos = 0;
			this.WordControl.OnResizeSplitter();
			this.WordControl.OldSplitter1Pos = old;
		}
	};

	asc_docs_api.prototype.asc_ShowNotes = function(bIsShow)
	{
		this.WordControl.ShowNotes(bIsShow);
	};

	asc_docs_api.prototype.asc_ShowAnimPane = function(bIsShow)
	{
		this.WordControl.ShowAnimPane(bIsShow);
	};

	asc_docs_api.prototype.asc_ZoomOutTimeline = function() {
		this.WordControl.ChangeTimelineScale(true);
	}
	asc_docs_api.prototype.asc_ZoomInTimeline = function() {
		this.WordControl.ChangeTimelineScale(false);
	}

	asc_docs_api.prototype.asc_RemoveSelectedAnimEffects = function() {
		this.asc_AddAnimation(null, AscFormat.ANIM_PRESET_NONE);
	};
	asc_docs_api.prototype.asc_SetSelectedAnimEffectsStartType = function(nStartType) {
		const selectedElements = this.getSelectedElements();
		for (let nElement = 0; nElement < selectedElements.length; nElement++) {
			if (selectedElements[nElement].Value instanceof AscFormat.CPar) {
				const par = selectedElements[nElement].Value.createDuplicate();
				par.asc_putStartType(nStartType);
				this.asc_SetAnimationProperties(par);
			}
		}
	};
	asc_docs_api.prototype.asc_DeleteVerticalScroll = function()
	{
		this.WordControl.DeleteVerticalScroll();
	};

	asc_docs_api.prototype.syncOnThumbnailsShow = function()
	{
		var bIsShow = true;
		if (0 == this.WordControl.Splitter1Pos)
			bIsShow = false;

		this.sendEvent("asc_onThumbnailsShow", bIsShow);
	};

	asc_docs_api.prototype.getTextInputOffset = function()
	{
		let ret = {
			X : 0,
			Y : 0
		};
		let wordControl = this.WordControl;
		if (wordControl && wordControl.Thumbnails)
		{
			if (wordControl.m_bIsRuler &&
				wordControl.Thumbnails.FocusObjType === AscCommon.FOCUS_OBJECT_MAIN)
			{
				ret.X += ((5 * AscCommon.g_dKoef_mm_to_pix) >> 0);
				ret.Y += ((7 * AscCommon.g_dKoef_mm_to_pix) >> 0);
			}
		}
		return ret;
	};

	asc_docs_api.prototype.getIsAnimPaneShow = function()
	{
		return this.WordControl.IsAnimPaneShown();
	};
	asc_docs_api.prototype.getIsNotesShow = function()
	{
		return this.WordControl.IsNotesShown();
	};
	asc_docs_api.prototype.syncOnNotesShow = function()
	{
		this.sendEvent("asc_onNotesShow", this.getIsNotesShow());
	};


	//-----------------------------------------------------------------
	// Функции для работы с гиперссылками
	//-----------------------------------------------------------------
	asc_docs_api.prototype.can_AddHyperlink = function()
	{
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation) return false;
		var bCanAdd = oPresentation.CanAddHyperlink(true);
		if (true === bCanAdd)
		{
			let oController = oPresentation.GetCurrentController();
			if(!oController)
			{
				return null;
			}
			let oTargetContent = oController.getTargetDocContent();
			if(!oTargetContent) return null;
			return oPresentation.GetSelectedText(true);
		}

		return false;
	};

	// HyperProps - объект CHyperlinkProperty
	asc_docs_api.prototype.add_Hyperlink = function(HyperProps)
	{
		if(null !== HyperProps.Text && undefined !== HyperProps.Text)
		{
			AscFonts.FontPickerByCharacter.checkText(HyperProps.Text, this, function() {

				this.WordControl.m_oLogicDocument.AddHyperlink(HyperProps);

			});
		}
		else
		{
			this.WordControl.m_oLogicDocument.AddHyperlink(HyperProps);
		}
	};

	// HyperProps - объект CHyperlinkProperty
	asc_docs_api.prototype.change_Hyperlink = function(HyperProps)
	{
		this.WordControl.m_oLogicDocument.ModifyHyperlink(HyperProps);
	};

	asc_docs_api.prototype.remove_Hyperlink = function()
	{
		this.WordControl.m_oLogicDocument.RemoveHyperlink();
	};

	asc_docs_api.prototype.sync_HyperlinkPropCallback = function(hyperProp)
	{
		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Hyperlink, new Asc.CHyperlinkProperty(hyperProp));
	};

	asc_docs_api.prototype.sync_HyperlinkClickCallback = function(Url)
	{
		this.sendEvent("asc_onHyperlinkClick", Url);
	};

	asc_docs_api.prototype.sync_CanAddHyperlinkCallback = function(bCanAdd)
	{
		//if ( true === CollaborativeEditing.Get_GlobalLock() )
		//    this.sendEvent("asc_onCanAddHyperlink", false);
		//else
		this.sendEvent("asc_onCanAddHyperlink", bCanAdd);
	};

	asc_docs_api.prototype.sync_DialogAddHyperlink = function(bInternal)
	{
		this.sendEvent("asc_onDialogAddHyperlink", bInternal);
	};

    //-----------------------------------------------------------------
    // Функции для работы с орфографией
    //-----------------------------------------------------------------
    asc_docs_api.prototype.sync_SpellCheckCallback = function(Word, Checked, Variants, ParaId, Element)
    {
        this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.SpellCheck, new AscCommon.asc_CSpellCheckProperty(Word, Checked, Variants, ParaId, Element));
    };

    asc_docs_api.prototype.sync_SpellCheckVariantsFound = function()
    {
        this.sendEvent("asc_onSpellCheckVariantsFound");
    };

    asc_docs_api.prototype.asc_replaceMisspelledWord = function(Word, SpellCheckProperty)
    {
        if(this.WordControl.m_oLogicDocument){
            this.WordControl.m_oLogicDocument.replaceMisspelledWord(Word, SpellCheckProperty);
		}
    };

    asc_docs_api.prototype.asc_ignoreMisspelledWord = function(SpellCheckProperty, bAll)
    {
        if (false === bAll)
        {
            var ParaId = SpellCheckProperty.ParaId;

            var Paragraph = g_oTableId.Get_ById(ParaId);
            if (null != Paragraph)
            {
                Paragraph.IgnoreMisspelledWord(SpellCheckProperty.Element);
            }
        }
        else
        {
            var LogicDocument = editor.WordControl.m_oLogicDocument;
            LogicDocument.Spelling.AddToIgnore(SpellCheckProperty.Word);
            LogicDocument.DrawingDocument.ClearCachePages();
            LogicDocument.DrawingDocument.FirePaint();
			if(LogicDocument.GetCurrentSlide())
			{
				LogicDocument.DrawingDocument.Notes_OnRecalculate(LogicDocument.CurPage,LogicDocument.NotesWidth, LogicDocument.GetCurrentSlide().getNotesHeight());
			}
        }
    };

    asc_docs_api.prototype._spellCheckRestart = function(word)
    {
		var LogicDocument = this.WordControl.m_oLogicDocument;
		if (LogicDocument)
		{
			// TODO: сделать нормальный сброс слова
			var oldWordStatus = LogicDocument.Spelling.IsIgnored(word);
			if (true !== oldWordStatus)
			{
				LogicDocument.Spelling.AddToIgnore(word);
				LogicDocument.DrawingDocument.ClearCachePages();
				LogicDocument.DrawingDocument.FirePaint();
				if(LogicDocument.GetCurrentSlide())
				{
					LogicDocument.DrawingDocument.Notes_OnRecalculate(LogicDocument.CurPage,LogicDocument.NotesWidth, LogicDocument.GetCurrentSlide().getNotesHeight());
				}
				delete LogicDocument.Spelling.Words[word];
			}
		}
    };
    asc_docs_api.prototype.asc_spellCheckClearDictionary = function()
    {
        if (window["AscDesktopEditor"])
            window["AscDesktopEditor"]["SpellCheck"]("{\"type\":\"clear\"}");
    };

    asc_docs_api.prototype.asc_setDefaultLanguage = function(Lang)
    {
        if (false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_PresDefaultLang))
        {
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetDefaultLanguage);
            this.WordControl.m_oLogicDocument.SetDefaultLanguage(Lang);
			this.WordControl.m_oLogicDocument.FinalizeAction();
        }
    };

    asc_docs_api.prototype.asc_getDefaultLanguage = function()
    {
        return this.WordControl.m_oLogicDocument.GetDefaultLanguage();
    };

    asc_docs_api.prototype.asc_getKeyboardLanguage = function()
    {
        if (undefined !== window["asc_current_keyboard_layout"])
            return window["asc_current_keyboard_layout"];
        return -1;
    };

	asc_docs_api.prototype.asc_getInputLanguage = function()
	{
		let keyboardLang = this.asc_getKeyboardLanguage();
		if (-1 !== keyboardLang)
			return keyboardLang;

		let oPresentation = this.private_GetLogicDocument();
		if (!oPresentation)
			return lcid_enUS;
		return oPresentation.GetInputLanguage();
	};

    asc_docs_api.prototype.asc_setSpellCheck = function(isOn)
    {
        if (editor.WordControl.m_oLogicDocument)
        {
        	var _presentation = editor.WordControl.m_oLogicDocument;
            _presentation.Spelling.Use = isOn;
            var _drawing_document = editor.WordControl.m_oDrawingDocument;
            _drawing_document.ClearCachePages();
            _drawing_document.FirePaint();
            if(_presentation.GetCurrentSlide() && _presentation.GetCurrentSlide().notes){
                _drawing_document.Notes_OnRecalculate(_presentation.CurPage, _presentation.GetCurrentSlide().NotesWidth, _presentation.GetCurrentSlide().getNotesHeight());
            }
        }
    };

	asc_docs_api.prototype.asc_restartCheckSpelling = function()
	{
		if (!this.WordControl || !this.WordControl.m_oLogicDocument)
			return null;

		this.WordControl.m_oLogicDocument.RestartSpellCheck();
	};

	asc_docs_api.prototype.asc_setSpellCheckSettings = function(oSettings)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		let oSpellCheck = oLogicDocument.GetSpellCheckManager();
		if (oSpellCheck.SetSettings(oSettings))
			oLogicDocument.RestartSpellCheck();
	};
	asc_docs_api.prototype.asc_getSpellCheckSettings = function()
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return new AscCommon.CSpellCheckSettings();

		return oLogicDocument.GetSpellCheckManager().GetSettings();
	};

	asc_docs_api.prototype.sync_shapePropCallback = function(pr)
	{
		var obj = AscFormat.CreateAscShapePropFromProp(pr);
		if (pr.fill != null && pr.fill.fill != null && pr.fill.fill.type == c_oAscFill.FILL_TYPE_BLIP)
		{
			this.WordControl.m_oDrawingDocument.DrawImageTextureFillShape(pr.fill.fill.RasterImageId);
		}
		else
		{
			this.WordControl.m_oDrawingDocument.DrawImageTextureFillShape(null);
		}

		var oTextArtProperties = pr.textArtProperties;
		if (oTextArtProperties && oTextArtProperties.Fill && oTextArtProperties.Fill.fill && oTextArtProperties.Fill.fill.type == c_oAscFill.FILL_TYPE_BLIP)
		{
			this.WordControl.m_oDrawingDocument.DrawImageTextureFillTextArt(oTextArtProperties.Fill.fill.RasterImageId);
		}
		else
		{
			this.WordControl.m_oDrawingDocument.DrawImageTextureFillTextArt(null);
		}


		var _len = this.SelectedObjectsStack.length;
		if (_len > 0)
		{
			if (this.SelectedObjectsStack[_len - 1].Type == c_oAscTypeSelectElement.Shape)
			{
				this.SelectedObjectsStack[_len - 1].Value = obj;
				return;
			}
		}

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Shape, obj);
	};

	asc_docs_api.prototype.sync_animPropCallback = function(pr)
	{
		var _len = this.SelectedObjectsStack.length;
		if (_len > 0)
		{
			if (this.SelectedObjectsStack[_len - 1].Type == c_oAscTypeSelectElement.Animation)
			{
				this.SelectedObjectsStack[_len - 1].Value = pr;
				return;
			}
		}

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Animation, pr);
	};

	asc_docs_api.prototype.sync_slidePropCallback = function(slide)
	{
		if(!this.WordControl)
		{
			return;
		}
		if(!this.WordControl.m_oLogicDocument)
		{
			return;
		}
		let obj = new CAscSlideProps();
		let oPresentation = this.WordControl.m_oLogicDocument, i;
		let aSlides = oPresentation.GetSelectedSlideObjects();
		if(this.WordControl.Thumbnails)
		{
			var oTh = editor.WordControl.Thumbnails;
			var aSelectedArray = oTh.GetSelectedArray();
			obj.isHidden = oTh.IsSlideHidden(aSelectedArray);
		}
		else
		{
			obj.isHidden = false;
		}
		if (!slide)
			return;
		if(aSlides.length === 0)
		{
			aSlides.push(slide);
		}

		for(let nIdx = 0; nIdx < aSlides.length; ++nIdx)
		{
			let oSlide = aSlides[nIdx];
			if(oSlide.getObjectType() === AscDFH.historyitem_type_SlideMaster)
			{
				obj.IsMaster = true;
				break;
			}
		}


		var bgFill = aSlides[0].backgroundFill ? aSlides[0].backgroundFill.createDuplicate() : aSlides[0].backgroundFill;
		for(i = 1; i < aSlides.length; ++i)
		{
			bgFill = AscFormat.CompareUniFill(bgFill, aSlides[i].backgroundFill);
			if(!bgFill){
				break;
			}
		}
		if (!bgFill)
		{
			obj.Background      = new asc_CShapeFill();
			obj.Background.type = c_oAscFill.FILL_TYPE_NOFILL;

			this.WordControl.m_oDrawingDocument.DrawImageTextureFillSlide(null);
		}
		else
		{
			obj.Background = AscFormat.CreateAscFill(bgFill);

			if (bgFill != null && bgFill.fill != null && bgFill.fill.type == c_oAscFill.FILL_TYPE_BLIP)
			{
				this.WordControl.m_oDrawingDocument.DrawImageTextureFillSlide(bgFill.fill.RasterImageId);
			}
			else
			{
				this.WordControl.m_oDrawingDocument.DrawImageTextureFillSlide(null);
			}
		}
		var transition = aSlides[0].transition ? aSlides[0].transition.createDuplicate() : aSlides[0].transition;
		for(i = 1; i < aSlides.length; ++i)
		{
			transition = AscCommonSlide.CompareTransitions(transition, aSlides[i].transition);
			if(!transition){
				break;
			}
		}

        if(transition){
            obj.Transition = transition.createDuplicate();
        }
        else{
            obj.Transition = new Asc.CAscSlideTransition();
        }
        obj.Transition.ShowLoop = this.WordControl.m_oLogicDocument.isLoopShowMode();

		let isCanResetBackground = false;
		let showMasterSp = true;
		for (let i = 0; i < aSlides.length; i += 1) {
			const oSlide = aSlides[i];
			isCanResetBackground = isCanResetBackground || !!oSlide.cSld.Bg;
			showMasterSp         = showMasterSp && oSlide.showMasterSp;

			obj.lockDelete     = obj.lockDelete || !(oSlide.deleteLock.Lock.Type === locktype_Mine || oSlide.deleteLock.Lock.Type === locktype_None);
			obj.lockLayout     = obj.lockLayout || !(oSlide.layoutLock.Lock.Type === locktype_Mine || oSlide.layoutLock.Lock.Type === locktype_None);
			obj.lockTiming     = obj.lockTiming || !(oSlide.timingLock.Lock.Type === locktype_Mine || oSlide.timingLock.Lock.Type === locktype_None);
			obj.lockTransition = obj.lockTransition || !(oSlide.transitionLock.Lock.Type === locktype_Mine || oSlide.transitionLock.Lock.Type === locktype_None);
			obj.lockBackground = obj.lockBackground || !(oSlide.backgroundLock.Lock.Type === locktype_Mine || oSlide.backgroundLock.Lock.Type === locktype_None);
			obj.lockRemove     = obj.lockRemove || obj.lockDelete ||
				obj.lockLayout ||
				obj.lockTiming ||
				obj.lockTransition ||
				obj.lockBackground || oSlide.isLockedObject();
		}
		obj.showMasterSp = showMasterSp;
		obj.lockResetBackground = obj.lockBackground || !isCanResetBackground;
		obj.lockApplyBackgroundToAll = obj.lockBackground || oPresentation.getLockApplyBackgroundToAll();

		if(slide && slide.Layout && slide.Layout.Master){
			var aLayouts = slide.Layout.Master.sldLayoutLst;
			for(i = 0; i < aLayouts.length; ++i){
				if(slide.Layout === aLayouts[i]){
                    obj.LayoutIndex = i;
					break;
				}
			}
		}
		var _len = this.SelectedObjectsStack.length;
		if (_len > 0)
		{
			if (this.SelectedObjectsStack[_len - 1].Type == c_oAscTypeSelectElement.Slide)
			{
				this.SelectedObjectsStack[_len - 1].Value = obj;
				return;
			}
		}

		if(aSlides.length === 1) {
			let oSlideLikeObject = aSlides[0];
			if(oSlideLikeObject.isMaster()) {
				obj.put_MasterName(oSlideLikeObject.getName())
			}
			else if(oSlideLikeObject.isLayout()) {
				obj.put_LayoutName(oSlideLikeObject.getName())
			}
		}

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new asc_CSelectedObject(c_oAscTypeSelectElement.Slide, obj);
	};

	asc_docs_api.prototype.GetCurrentPixOffsetY = function()
	{
		return this.WordControl.m_dScrollY;
	};

	asc_docs_api.prototype.retrieveFormatPainterData = function()
	{
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation)
		{
			return null;
		}
		return oPresentation.GetFormattingPasteData();
	};
	asc_docs_api.prototype.SetPaintFormat = function(value)
	{
		const oPresentation = this.WordControl.m_oLogicDocument;
		if(!oPresentation)
			return;
		const oDrawingObjects = oPresentation.GetCurrentController();
		if(!oDrawingObjects)
			return;
		if (oDrawingObjects.isTrackingDrawings()) {
			oDrawingObjects.endTrackNewShape();
			this.sync_EndAddShape();
		}
		this.stopInkDrawer();
		this.cancelEyedropper();
		this.formatPainter.putState(value);
		this.WordControl.m_oLogicDocument.Document_Format_Copy();
	};

	asc_docs_api.prototype.onInkDrawerChangeState = function()
	{
		const oPresentation = this.WordControl.m_oLogicDocument;
		if(!oPresentation)
			return;
		this.cancelEyedropper();
		if (this.isFormatPainterOn())
		{
			this.sync_PaintFormatCallback(AscCommon.c_oAscFormatPainterState.kOff);
		}
		oPresentation.OnInkDrawerChangeState();
	};

	asc_docs_api.prototype.sync_PaintFormatCallback = function(value)
	{
		this.formatPainter.putState(value);
		return this.sendEvent("asc_onPaintFormatChanged", value);
	};
	asc_docs_api.prototype.SetMarkerFormat          = function(value, is_flag, r, g, b)
	{
		this.isMarkerFormat = value;
		if (this.isMarkerFormat)
		{
			this.WordControl.m_oLogicDocument.SetParagraphHighlight(is_flag, r, g, b);
		}
	};

	asc_docs_api.prototype.sync_MarkerFormatCallback = function(value)
	{
		this.isMarkerFormat = value;
		return this.sendEvent("asc_onMarkerFormatChanged", value);
	};

	asc_docs_api.prototype.ClearFormating           = function()
	{
		this.WordControl.m_oLogicDocument.ClearParagraphFormatting(false, true);
	};

	window.ID_KEYBOARD_AREA = undefined;
	window.ID_KEYBOARD_AREA;
	asc_docs_api.prototype.SetDeviceInputHelperId = function(idKeyboard)
	{
		if (window.ID_KEYBOARD_AREA === undefined && this.WordControl.m_oMainView != null)
		{
			window.ID_KEYBOARD_AREA = document.getElementById(idKeyboard);

			window.ID_KEYBOARD_AREA.onkeypress = function(e)
			{
				if (false === editor.WordControl.IsFocus)
				{
					editor.WordControl.IsFocus = true;
					var ret                    = editor.WordControl.onKeyPress(e);
					editor.WordControl.IsFocus = false;
					return ret;
				}
			};
			window.ID_KEYBOARD_AREA.onkeydown  = function(e)
			{
				if (false === editor.WordControl.IsFocus)
				{
					editor.WordControl.IsFocus = true;
					var ret                    = editor.WordControl.onKeyDown(e);
					editor.WordControl.IsFocus = false;
					return ret;
				}
			};
		}
		window.ID_KEYBOARD_AREA.focus();
	};
	asc_docs_api.prototype.asc_setViewMode        = function(isViewMode)
	{
		this.isViewMode = !!isViewMode;
		if (!this.isLoadFullApi)
		{
			return;
		}

		this.WordControl.setNotesEnable((this.isMobileVersion && !this.isReporterMode) ? false : true);
		if (isViewMode)
		{
			this.ShowParaMarks          = false;
			this.WordControl.m_bIsRuler = false;
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.HideRulers();

            AscCommon.CollaborativeEditing.Set_GlobalLock(true);
			if (null != this.WordControl.m_oLogicDocument)
			{
				this.WordControl.m_oLogicDocument.viewMode = true;
			}
			this.turnOffSpecialModes();
		}
		else
		{
			this.WordControl.checkNeedRules();
			this.WordControl.m_oDrawingDocument.ClearCachePages();
			this.WordControl.OnResize(true);

			if (null != this.WordControl.m_oLogicDocument)
			{
				this.WordControl.m_oLogicDocument.viewMode = false;
			}
		}

		this.checkInputMode();
	};

	asc_docs_api.prototype.sync_HyperlinkClickCallback = function(Url)
	{
		if (AscCommon.IsLinkPPAction(Url))
		{
			if (Url == "ppaction://hlinkshowjump?jump=firstslide")
			{
				this.WordControl.GoToPage(0);
			}
			else if (Url == "ppaction://hlinkshowjump?jump=lastslide")
			{
				this.WordControl.GoToPage(this.WordControl.GetSlidesCount() - 1);
			}
			else if (Url == "ppaction://hlinkshowjump?jump=nextslide")
			{
				this.WordControl.onNextPage();
			}
			else if (Url == "ppaction://hlinkshowjump?jump=previousslide")
			{
				this.WordControl.onPrevPage();
			}
			else
			{
				var mask     = "ppaction://hlinksldjumpslide";
				var indSlide = Url.indexOf(mask);
				if (0 == indSlide)
				{
					var slideNum = parseInt(Url.substring(mask.length));
					if (slideNum >= 0 && slideNum < this.WordControl.GetSlidesCount())
						this.WordControl.GoToPage(slideNum);
				}
			}
			return;
		}

		this.sendEvent("asc_onHyperlinkClick", Url);
	};

	asc_docs_api.prototype.asc_GoToInternalHyperlink = function(url)
	{
		for(var i = 0; i < this.SelectedObjectsStack.length; ++i){
			if(this.SelectedObjectsStack[i].Type === c_oAscTypeSelectElement.Hyperlink){
				var oHyperProp = this.SelectedObjectsStack[i].Value;
				if(typeof oHyperProp.Value === "string" && oHyperProp.Value.indexOf("ppaction://hlink") === 0){
					this.sync_HyperlinkClickCallback(oHyperProp.Value);
				}
				return;
			}
		}
	};

	asc_docs_api.prototype.UpdateInterfaceState = function()
	{
		if (this.WordControl.m_oLogicDocument != null)
		{
			this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
			this.WordControl.CheckLayouts(true);
		}
	};

	asc_docs_api.prototype.OnMouseUp = function(x, y)
	{
		var _e = AscCommon.CreateMouseUpEventObject(x, y);
		AscCommon.Window_OnMouseUp(_e);

		//this.WordControl.onMouseUpExternal(x, y);
	};

	asc_docs_api.prototype.asyncImageEndLoaded2 = null;

	asc_docs_api.prototype.ChangeTheme = function(indexTheme, bSelectedSlides)
	{
		if (true === AscCommon.CollaborativeEditing.Get_GlobalLock())
			return;

		if (!this.isViewMode && this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Theme) === false)
		{
			AscCommon.CollaborativeEditing.Set_GlobalLock(true);

			this.inkDrawer.startSilentMode();
			this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Presentation_ChangeTheme);
			this.inkDrawer.endSilentMode();
			this.bSelectedSlidesTheme = (bSelectedSlides === true);
			this.ThemeLoader.StartLoadTheme(indexTheme);
		}
	};

	asc_docs_api.prototype.StartLoadTheme = function()
	{
	};
	asc_docs_api.prototype.EndLoadTheme   = function(theme_load_info)
	{
		AscCommon.CollaborativeEditing.Set_GlobalLock(false);

		// применение темы
		var _array = this.WordControl.m_oLogicDocument.GetSelectedSlides();
		this.WordControl.m_oLogicDocument.changeTheme(theme_load_info, (_array.length <= 1 && !this.bSelectedSlidesTheme) ? null : _array);
		this.WordControl.ThemeGenerateThumbnails(theme_load_info.Master);
		// меняем шаблоны в меню
		this.WordControl.CheckLayouts();
		this.WordControl.m_oLogicDocument.FinalizeAction(true);

		this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.LoadTheme);
	};

	asc_docs_api.prototype.ChangeLayout = function(layout_index)
	{
		var _array = this.WordControl.m_oLogicDocument.GetSelectedSlides();
		this.WordControl.m_oLogicDocument.changeLayout(_array, layout_index);
	};
	asc_docs_api.prototype.ResetSlide = function()
	{
		let _array = this.WordControl.m_oLogicDocument.GetSelectedSlides();
		this.WordControl.m_oLogicDocument.changeLayout(_array, undefined);
	};

	asc_docs_api.prototype.put_ShapesAlign = function(type, alignType)
	{
		if(!AscFormat.isRealNumber(alignType))
		{
			alignType = Asc.c_oAscObjectsAlignType.Slide;
		}
		switch (type)
		{
			case c_oAscAlignShapeType.ALIGN_LEFT:
			{
				this.shapes_alignLeft(alignType);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_RIGHT:
			{
				this.shapes_alignRight(alignType);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_TOP:
			{
				this.shapes_alignTop(alignType);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_BOTTOM:
			{
				this.shapes_alignBottom(alignType);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_CENTER:
			{
				this.shapes_alignCenter(alignType);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_MIDDLE:
			{
				this.shapes_alignMiddle(alignType);
				break;
			}
			default:
				break;
		}
	};
	asc_docs_api.prototype.DistributeHorizontally = function(alignType)
	{
		if(!AscFormat.isRealNumber(alignType))
		{
			alignType = Asc.c_oAscObjectsAlignType.Slide;
		}
		this.WordControl.m_oLogicDocument.distributeHor(alignType);
	};
	asc_docs_api.prototype.DistributeVertically   = function(alignType)
	{
		if(!AscFormat.isRealNumber(alignType))
		{
			alignType = Asc.c_oAscObjectsAlignType.Slide;
		}
		this.WordControl.m_oLogicDocument.distributeVer(alignType);
	};
	asc_docs_api.prototype.shapes_alignLeft       = function(alignType)
	{
		this.WordControl.m_oLogicDocument.alignLeft(alignType);
	};

	asc_docs_api.prototype.shapes_alignRight = function(alignType)
	{
		this.WordControl.m_oLogicDocument.alignRight(alignType);
	};

	asc_docs_api.prototype.shapes_alignTop = function(alignType)
	{
		this.WordControl.m_oLogicDocument.alignTop(alignType);

	};

	asc_docs_api.prototype.shapes_alignBottom = function(alignType)
	{
		this.WordControl.m_oLogicDocument.alignBottom(alignType);

	};

	asc_docs_api.prototype.shapes_alignCenter = function(alignType)
	{
		this.WordControl.m_oLogicDocument.alignCenter(alignType);
	};

	asc_docs_api.prototype.shapes_alignMiddle = function(alignType)
	{
		this.WordControl.m_oLogicDocument.alignMiddle(alignType);
	};

	asc_docs_api.prototype.shapes_bringToFront = function()
	{
		this.WordControl.m_oLogicDocument.bringToFront();
	};

	asc_docs_api.prototype.shapes_bringForward = function()
	{
		this.WordControl.m_oLogicDocument.bringForward();
	};

	asc_docs_api.prototype.shapes_bringToBack = function()
	{
		this.WordControl.m_oLogicDocument.sendToBack();
	};

	asc_docs_api.prototype.shapes_bringBackward = function()
	{
		this.WordControl.m_oLogicDocument.bringBackward();
	};

	asc_docs_api.prototype.asc_setLoopShow = function(isLoop)
	{
		this.WordControl.m_oLogicDocument.setShowLoop(isLoop);
	};

	asc_docs_api.prototype.sync_endDemonstration          = function()
	{
		this.sendEvent("asc_onEndDemonstration");
	};
	asc_docs_api.prototype.sync_DemonstrationSlideChanged = function(slideNum)
	{
		this.sendEvent("asc_onDemonstrationSlideChanged", slideNum);
	};
	asc_docs_api.prototype.getAnnotations = function ()
	{
		if(!this.isSlideShow()) return null;
		return this.WordControl.DemonstrationManager.SlideAnnotations;
	};
	asc_docs_api.prototype.isDrawSlideshowAnnotations = function ()
	{
		if(this.isSlideShow() && this.isInkDrawerOn())
			return true;
		return false;
	};
	asc_docs_api.prototype.StartDemonstration = function(div_id, slidestart_num, reporterStartObject)
	{
		if (window.g_asc_plugins)
			window.g_asc_plugins.stopWorked();

		const bIsReporter = (reporterStartObject && !this.isReporterMode);

		if (!bIsReporter)
		{
			this.turnOffSpecialModes();
		}

		if (bIsReporter)
			this.DemonstrationReporterStart(reporterStartObject);

		if (bIsReporter && (this.reporterWindow || window["AscDesktopEditor"]))
			this.WordControl.DemonstrationManager.StartWaitReporter(div_id, slidestart_num, true);
		else
			this.WordControl.DemonstrationManager.Start(div_id, slidestart_num, true);

        if (undefined !== this.EndShowMessage)
        {
            this.WordControl.DemonstrationManager.EndShowMessage = this.EndShowMessage;
            this.EndShowMessage = undefined;
        }
	};

	asc_docs_api.prototype.EndDemonstration = function(isNoUseFullScreen)
	{
		let oAnnotations = this.getAnnotations();
		if(oAnnotations && !oAnnotations.isEmpty())
		{
			this.sendEvent("asc_onEndDemoWithAnnotations", function (bKeep) {
				Asc.editor.closeDemonstration(bKeep, isNoUseFullScreen);
			});
		}
		else
		{
			this.closeDemonstration(false, isNoUseFullScreen);
		}
	};

	asc_docs_api.prototype.DemonstrationReporterStart = function(startObject)
	{
		this.reporterStartObject = startObject;
		this.reporterStartObject["translate"] = AscCommon.translateManager.mapTranslate;
		this.reporterStartObject["skin"] = AscCommon.GlobalSkin;
		this.reporterStartObject["canEditMain"] = this.canEdit();

		if (window["AscDesktopEditor"])
		{
			window["AscDesktopEditor"]["startReporter"](window.location.href);
			this.reporterWindow = {};
			return;
		}

		var dualScreenLeft = (window.screenLeft != undefined) ? window.screenLeft : screen.left;
		var dualScreenTop = (window.screenTop != undefined) ? window.screenTop : screen.top;

		var width = window.innerWidth ? window.innerWidth : document.documentElement.clientWidth ? document.documentElement.clientWidth : screen.width;
		var height = window.innerHeight ? window.innerHeight : document.documentElement.clientHeight ? document.documentElement.clientHeight : screen.height;

		var w = 800;
		var h = 600;
		var left = ((width / 2) - (w / 2)) + dualScreenLeft;
		var top = ((height / 2) - (h / 2)) + dualScreenTop;

		var _windowPos = "width=" + w + ",height=" + h + ",left=" + left + ",top=" + top;
		var _url = "index.reporter.html";
		if (this.locale)
			_url += ("?lang=" + this.locale);
		this.reporterWindow = window.open(_url, "_blank", "resizable=yes,status=0,toolbar=0,location=0,menubar=0,directories=0,scrollbars=0," + _windowPos);

		if (!this.reporterWindow)
			return;

		this.reporterWindowCounter = 0;
		if (!AscCommon.AscBrowser.isSafariMacOs)
		{
			this.reporterWindow.onbeforeunload = function ()
			{
				window.editor.EndDemonstration();
			};
		}

		this.reporterWindow.onunload = function () {
			window.editor.reporterWindowCounter++;
			if (1 < window.editor.reporterWindowCounter)
			{
				window.editor.EndDemonstration();
			}
		};

		if ( this.reporterWindow.attachEvent )
			this.reporterWindow.attachEvent('onmessage', this.DemonstrationReporterMessages);
		else
			this.reporterWindow.addEventListener('message', this.DemonstrationReporterMessages, false);
	};

	asc_docs_api.prototype.DemonstrationReporterEnd = function()
	{
		if (window["AscDesktopEditor"])
		{
			window["AscDesktopEditor"]["endReporter"]();
			this.reporterWindow = null;
			return;
		}

		try
		{
			this.reporterWindowCounter = 0;
			if (!this.reporterWindow)
				return;

			if (this.reporterWindow.attachEvent)
				this.reporterWindow.detachEvent('onmessage', this.DemonstrationReporterMessages);
			else
				this.reporterWindow.removeEventListener('message', this.DemonstrationReporterMessages, false);

			this.reporterWindow.close();
			this.reporterWindow = null;
			this.reporterStartObject = null;
		}
		catch (err)
		{
			this.reporterWindow = null;
			this.reporterStartObject = null;
		}
	};

	asc_docs_api.prototype.DemonstrationReporterMessages = function(e)
	{
		var _this = window.editor;
		if ( e.data == 'i:am:ready' )
		{
			var _msg_ = {
				type: 'file:open',
				data: _this.reporterStartObject
			};

			if (AscCommon.EncryptionWorker.isPasswordCryptoPresent)
			{
                _msg_.data["cryptoCurrentPassword"] = this.currentPassword;
                _msg_.data["cryptoCurrentDocumentHash"] = this.currentDocumentHash;
                _msg_.data["cryptoCurrentDocumentInfo"] = this.currentDocumentInfo;
            }

			this.reporterStartObject = null;
			_this.sendToReporter(JSON.stringify(_msg_));

			return;
		}

		try
		{
			var _obj = JSON.parse(e.data);

			if (undefined == _obj["reporter_command"])
				return;

			switch (_obj["reporter_command"])
			{
				case "end":
				{
					_this.EndDemonstration();
					break;
				}
				case "next":
				{
					_this.WordControl.DemonstrationManager.OnNextSlide();
					break;
				}
				case "prev":
				{
					_this.WordControl.DemonstrationManager.OnPrevSlide();
					break;
				}
				case "go_to_slide":
				{
					_this.WordControl.DemonstrationManager.GoToSlide(_obj["slide"]);
					break;
				}
				case "start_show":
				{
					_this.WordControl.DemonstrationManager.EndWaitReporter();
					break;
				}
				case "pointer_move":
				{
					_this.WordControl.DemonstrationManager.PointerMove(_obj["x"], _obj["y"], _obj["w"], _obj["h"]);
					break;
				}
				case "pointer_remove":
				{
					_this.WordControl.DemonstrationManager.PointerRemove();
					break;
				}
				case "pause":
				{
					_this.WordControl.DemonstrationManager.Pause();
					_this.sendEvent("asc_onDemonstrationStatus", "pause");
					break;
				}
				case "play":
				{
					_this.WordControl.DemonstrationManager.Play();
					_this.sendEvent("asc_onDemonstrationStatus", "play");
					break;
				}
				case "resize":
				{
					_this.WordControl.DemonstrationManager.Resize(true);
					break;
				}
				case "on_mouse_down":
				{
					_this.WordControl.DemonstrationManager.CheckMouseDown(_obj["x"], _obj["y"], _obj["page"]);
					break;
				}
				case "annotation":
				{
					let oAnnotations = _this.getAnnotations();
					if(oAnnotations)
					{
						oAnnotations.handleMessage(_obj["value"]);
					}
					break;
				}
				default:
					break;
			}
		}
		catch (err)
		{
		}
	};

	asc_docs_api.prototype.preloadReporter = function(data)
	{
		if (data["translate"])
			this.translateManager = AscCommon.translateManager.init(data["translate"]);
		if (data["skin"]) {
			let oSkin = data["skin"];
			oSkin.SupportNotes = true;
			this.asc_setSkin(data["skin"]);
		}

		this.reporterTranslates = [
			data["translations"]["reset"],
			data["translations"]["slideOf"],
			data["translations"]["endSlideshow"],
			data["translations"]["finalMessage"],
			data["translations"]["pen"],
			data["translations"]["highlighter"],
			data["translations"]["inkColor"],
			data["translations"]["eraser"],
			data["translations"]["eraseScreen"]
		];

        if (data["cryptoCurrentPassword"])
        {
            this.currentPassword = data["cryptoCurrentPassword"];
            this.currentDocumentHash = data["cryptoCurrentDocumentHash"];
            this.currentDocumentInfo = data["cryptoCurrentDocumentInfo"];

            if (this.pluginsManager)
                this.pluginsManager.checkCryptoReporter();
            else
                this.isCheckCryptoReporter = true;
        }

        this.asc_registerCallback('asc_onHyperlinkClick', function(url){
            if (url) {
                window.open(url);
            }
        });

		if (!this.WordControl)
			return;

		this.WordControl.reporterTranslates = this.reporterTranslates;
        this.WordControl.DemonstrationManager.EndShowMessage = this.reporterTranslates[3];

		var _button1 = document.getElementById("dem_id_reset");
		var _button2 = document.getElementById("dem_id_end");

		let canEditMain = data["canEditMain"];
		if(!canEditMain)
		{
			let drawButton = document.getElementById("dem_id_draw_menu_trigger");
			drawButton.style.display = "none";
		}
		else
		{
			var _miPen = document.querySelector("#dem_id_draw_menu a[data-tool=\"pen\"]");
			var _miHighlighter = document.querySelector("#dem_id_draw_menu a[data-tool=\"highlighter\"]");
			var _miInkColor = document.querySelector("#dem_id_draw_color_menu_trigger > a");
			var _miEraser = document.querySelector("#dem_id_draw_menu a[data-tool=\"eraser\"]");
			var _miEraseAll = document.querySelector("#dem_id_draw_menu a[data-tool=\"erase-all\"]");
		}


		if (_button1)
			_button1.innerHTML = this.reporterTranslates[0];
		if (_button2)
		{
			_button2.innerHTML = this.reporterTranslates[2];
			this.WordControl.OnResizeReporter();
		}

		if (_miPen) _miPen.childNodes[1].textContent = this.reporterTranslates[4];
		if (_miHighlighter) _miHighlighter.childNodes[1].textContent = this.reporterTranslates[5];
		if (_miInkColor) _miInkColor.textContent = this.reporterTranslates[6];
		if (_miEraser) _miEraser.childNodes[1].textContent = this.reporterTranslates[7];
		if (_miEraseAll) _miEraseAll.childNodes[1].textContent = this.reporterTranslates[8];

		this.WordControl.UpdateBottomControlsParams();
	};

	asc_docs_api.prototype.sendToReporter = function(value)
	{
		if (this.disableReporterEvents)
			return;

		if (window["AscDesktopEditor"])
		{
			window["AscDesktopEditor"]["sendToReporter"](value);
			return;
		}

		if (this.reporterWindow)
			this.reporterWindow.postMessage(value, "*");
	};

	asc_docs_api.prototype.sendFromReporter = function(value)
	{
		if (this.disableReporterEvents)
			return;

		if (window["AscDesktopEditor"])
		{
			window["AscDesktopEditor"]["sendFromReporter"](value);
			return;
		}

		window.postMessage(value, "*");
	};

	asc_docs_api.prototype.DemonstrationToReporterMessages = function(e)
	{
		var _this = window.editor;

		try
		{
			var _obj = JSON.parse(e.data);

			if (window["AscDesktopEditor"] && (_obj["type"] == "file:open"))
			{
				window.postMessage(e.data, "*");
				return;
			}

			if (undefined == _obj["main_command"])
				return;

			if (undefined !== _obj["keyCode"])
			{
				_this.WordControl.DemonstrationManager.onKeyDownCode(_obj["keyCode"]);
			}
			else if (undefined !== _obj["mouseUp"])
			{
				_this.WordControl.DemonstrationManager.onMouseUp({}, true, true);
			}
			else if (undefined !== _obj["mouseWhell"])
			{
				_this.WordControl.DemonstrationManager.onMouseWheelDelta(_obj["mouseWhell"]);
			}
			else if (undefined !== _obj["resize"])
			{
				_this.WordControl.DemonstrationManager.Resize(true);
			}
			else if (true === _obj["next"])
			{
				_this.WordControl.DemonstrationManager.OnNextSlide(true);
			}
			else if (true === _obj["prev"])
			{
				_this.WordControl.DemonstrationManager.OnPrevSlide(true);
			}
			else if (undefined !== _obj["go_to_slide"])
			{
				_this.WordControl.DemonstrationManager.GoToSlide(_obj["go_to_slide"], true);
			}
			else if (true === _obj["play"])
			{
				var _isNowPlaying = _this.WordControl.DemonstrationManager.IsPlayMode;
				_this.WordControl.DemonstrationManager.Play(true);
				var _elem = document.getElementById("dem_id_play_span");
				if (_elem && !_isNowPlaying)
				{
					_elem.classList.remove("btn-play");
					_elem.classList.add("btn-pause");

					_this.WordControl.reporterTimerLastStart = new Date().getTime();

					_this.WordControl.reporterTimer = setInterval(_this.WordControl.reporterTimerFunc, 1000);
				}
			}
			else if (true === _obj["pause"])
			{
				var _isNowPlaying = _this.WordControl.DemonstrationManager.IsPlayMode;
				_this.WordControl.DemonstrationManager.Pause();
				var _elem = document.getElementById("dem_id_play_span");
				if (_elem && _isNowPlaying)
				{
					_elem.classList.remove("btn-pause");
					_elem.classList.add("btn-play");

					if (-1 != _this.WordControl.reporterTimer)
					{
						clearInterval(_this.WordControl.reporterTimer);
						_this.WordControl.reporterTimer = -1;
					}

					_this.WordControl.reporterTimerAdd = _this.WordControl.reporterTimerFunc(true);
				}
			}
			else if(true === _obj["on_mouse_down"])
			{
				_this.WordControl.DemonstrationManager.CheckMouseDown(_obj["x"], _obj["y"], _obj["page"]);
			}
			else if(undefined !== _obj["annotation"])
			{
				let oAnnotations = _this.getAnnotations();
				if(oAnnotations)
				{
					oAnnotations.handleMessage(_obj["annotation"])
				}
			}
		}
		catch (err)
		{
		}
	};

	asc_docs_api.prototype.DemonstrationPlay = function()
	{
		if (undefined !== this.EndShowMessage)
		{
			this.WordControl.DemonstrationManager.EndShowMessage = this.EndShowMessage;
			this.EndShowMessage = undefined;
		}
		this.WordControl.DemonstrationManager.Play(true);

		if (this.reporterWindow)
			this.sendToReporter("{ \"main_command\" : true, \"play\" : true }");
	};

	asc_docs_api.prototype.DemonstrationPause = function()
	{
		this.WordControl.DemonstrationManager.Pause();
		if (this.reporterWindow)
			this.sendToReporter("{ \"main_command\" : true, \"pause\" : true }");
	};

	asc_docs_api.prototype.DemonstrationEndShowMessage = function(message)
	{
		if (!this.WordControl)
			this.EndShowMessage = message;
		else
			this.WordControl.DemonstrationManager.EndShowMessage = message;
	};

	asc_docs_api.prototype.DemonstrationNextSlide = function()
	{
		this.WordControl.DemonstrationManager.OnNextSlide();
		if (this.reporterWindow)
			this.sendToReporter("{ \"main_command\" : true, \"next\" : true }");
	};

	asc_docs_api.prototype.DemonstrationPrevSlide = function()
	{
		this.WordControl.DemonstrationManager.OnPrevSlide();
		if (this.reporterWindow)
			this.sendToReporter("{ \"main_command\" : true, \"prev\" : true }");
	};

	asc_docs_api.prototype.DemonstrationGoToSlide = function(slideNum)
	{
		this.WordControl.DemonstrationManager.GoToSlide(slideNum);

		if (this.isReporterMode)
			this.sendFromReporter("{ \"reporter_command\" : \"go_to_slide\", \"slide\" : " + slideNum + " }");

		if (this.reporterWindow)
			this.sendToReporter("{ \"main_command\" : true, \"go_to_slide\" : " + slideNum + " }");
	};

	asc_docs_api.prototype.SetDemonstrationModeOnly = function()
	{
		this.isOnlyDemonstration = true;
	};

	asc_docs_api.prototype.ApplySlideTransition      = function(oTransition)
	{
		if(!oTransition)
			return;
		if(this.isMasterMode()) return;

		const oPresentation = this.private_GetLogicDocument();
		if(!oPresentation)
			return;
		if(oPresentation.IsEmpty())
			return;

		if (!oPresentation.IsSelectionLocked(AscCommon.changestype_SlideTransition))
		{
			oPresentation.StartAction(AscDFH.historydescription_Presentation_ApplyTransition);
            const aSelectedSlides = oPresentation.GetSelectedSlides();
            for(let nSld = 0; nSld < aSelectedSlides.length; ++nSld)
            {
                let oSlide = oPresentation.GetSlide(aSelectedSlides[nSld]);
				if(oSlide)
				{
					oSlide.applyTransition(oTransition);
				}
			}
			const bNewIsShowLoop = oTransition.get_ShowLoop();
			if(AscFormat.isRealBool(bNewIsShowLoop) && bNewIsShowLoop !== oPresentation.isLoopShowMode())
			{
				oPresentation.setShowLoop(bNewIsShowLoop);
			}
			oPresentation.UpdateInterface();
			oPresentation.FinalizeAction();
		}
	};
	asc_docs_api.prototype.SlideTransitionApplyToAll = function()
	{
		const oPresentation = this.private_GetLogicDocument();

		if(!oPresentation)
			return;
		if(oPresentation.IsEmpty())
			return;

		if(this.isMasterMode())
			return;
		const oCurSlide = oPresentation.GetCurrentSlide();
		if(!oCurSlide)
			return;

		let oTrToApply = null;
		if(oCurSlide.transition)
		{
			oTrToApply = oCurSlide.transition.createDuplicate();
		}
		if(!oTrToApply)
			return;


		if (!oPresentation.IsSelectionLocked(AscCommon.changestype_SlideTransition, {All : true}))
		{
			oPresentation.StartAction(AscDFH.historydescription_Presentation_ApplyTransitionToAll);
			let aSlides = oPresentation.Slides;
			for(let nSld = 0; nSld < aSlides.length; ++nSld)
			{
				let oSlide = aSlides[nSld];
				if(oSlide !== oCurSlide)
				{
					oSlide.applyTransition(oTrToApply);
				}
			}
			oPresentation.UpdateInterface();
			oPresentation.FinalizeAction();
		}
	};
	asc_docs_api.prototype.SlideTransitionPlay   = function(endCallback)
	{
		var _count = this.WordControl.GetSlidesCount();
		var _cur   = this.WordControl.m_oDrawingDocument.SlideCurrent;
		if (_cur < 0 || _cur >= _count)
			return;
		var _transition = this.WordControl.m_oLogicDocument.Slides[_cur].transition;

		var _tr      = this.WordControl.m_oDrawingDocument.TransitionSlide;
		let oTypeAndOption = _transition.getTypeAndOption();
		_tr.Type     = oTypeAndOption.Type;
		_tr.Param    = oTypeAndOption.Option;
		_tr.Duration = _transition.TransitionDuration;

		_tr.Start(true, endCallback);
		return _tr;
	};

	asc_docs_api.prototype.asc_HideSlides   = function(isHide)
	{
		this.WordControl.m_oLogicDocument.hideSlides(isHide);
	};

	asc_docs_api.prototype.sync_EndAddShape = function()
	{
		editor.sendEvent("asc_onEndAddShape");
		if (this.WordControl.m_oDrawingDocument.m_sLockedCursorType == "crosshair")
		{
			this.WordControl.m_oDrawingDocument.UnlockCursorType();
		}
		if(this.WordControl.m_oLogicDocument.TurnOffInterfaceEvents) {
			this.WordControl.m_oLogicDocument.TurnOn_InterfaceEvents(false);
		}
	};

	// Вставка диаграмм
	asc_docs_api.prototype.asc_getChartObject = function(type)
	{
        if (!AscFormat.isRealNumber(type))
        {
            this.asc_onOpenChartFrame();
            this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props);
        }
		return this.WordControl.m_oLogicDocument.GetChartObject(type);
	};

	asc_docs_api.prototype.asc_addChartDrawingObject = function(chartBinary, Placeholder)
	{
		/**/

		// Приводим бинарик к объекту типа CChartAsGroup и добавляем объект
		if (AscFormat.isObject(chartBinary))
		{
			//if ( false === this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(changestype_Drawing_Props) )
			{
				AscFonts.IsCheckSymbols = true;
				this.WordControl.m_oLogicDocument.addChart(chartBinary, true, Placeholder);
				AscFonts.IsCheckSymbols = false;
			}
		}
	};

	asc_docs_api.prototype.asc_editChartDrawingObject = function(chartBinary)
	{
		/**/

		// Находим выделенную диаграмму и накатываем бинарник
		if (AscCommon.isRealObject(chartBinary))
		{
			this.WordControl.m_oLogicDocument.EditChart(chartBinary);
		}
	};

	asc_docs_api.prototype.asc_onCloseChartFrame               = function()
	{
		AscCommon.baseEditorsApi.prototype.asc_onCloseChartFrame.call(this);
		this.WordControl.m_bIsMouseLock = false;
	};

	asc_docs_api.prototype.sync_closeChartEditor = function()
	{
		this.sendEvent("asc_onCloseChartEditor");
	};
	asc_docs_api.prototype.sync_closeOleEditor = function()
	{
		this.sendEvent("asc_onCloseOleEditor");
	};
	asc_docs_api.prototype.asc_setDrawCollaborationMarks = function()
	{
	};

	asc_docs_api.prototype.asc_getSelectedDrawingObjectsCount = function()
	{
		if(!this.WordControl)
		{
			return 0;
		}
		if(!this.WordControl.m_oLogicDocument)
		{
			return 0;
		}
		return this.WordControl.m_oLogicDocument.getSelectedDrawingObjectsCount();
	};
	asc_docs_api.prototype.asc_ChangeTextCase = function(nType)
	{
		var oLogicDocument = this.WordControl && this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;
		oLogicDocument.ChangeTextCase(nType);
	};

	//-----------------------------------------------------------------
	// События контекстного меню
	//-----------------------------------------------------------------

	function CContextMenuData(oData)
	{
		if (AscCommon.isRealObject(oData))
		{
			this.Type          = oData.Type;
			this.X_abs         = oData.X_abs;
			this.Y_abs         = oData.Y_abs;
			this.IsSlideSelect = oData.IsSlideSelect;
			this.IsSlideHidden = oData.IsSlideHidden;
			this.Guide         = oData.Guide;
		}
		else
		{
			this.Type          = Asc.c_oAscContextMenuTypes.Main;
			this.X_abs         = 0;
			this.Y_abs         = 0;
			this.IsSlideSelect = true;
            this.IsSlideHidden = false;
            this.Guide         = null;
		}
	}

	CContextMenuData.prototype.get_Type          = function()
	{
		return this.Type;
	};
	CContextMenuData.prototype.get_X             = function()
	{
		return this.X_abs;
	};
	CContextMenuData.prototype.get_Y             = function()
	{
		return this.Y_abs;
	};
	CContextMenuData.prototype.get_IsSlideSelect = function()
	{
		return this.IsSlideSelect;
	};

	CContextMenuData.prototype.get_IsSlideHidden = function()
	{
		return this.IsSlideHidden;
	};
	CContextMenuData.prototype.get_Guide = function()
	{
		return this.Guide;
	};
	CContextMenuData.prototype.get_EffectStartType = function() {
		return this.EffectStartType;
	};
	CContextMenuData.prototype.get_ButtonWidth = function() {
		return this.ButtonWidth;
	}
	CContextMenuData.prototype.get_ButtonHeight = function() {
		return this.ButtonHeight;
	}

	asc_docs_api.prototype.sync_ContextMenuCallback = function(Data)
	{
		this.sendEvent("asc_onContextMenu", Data);
	};

	asc_docs_api.prototype._onNeedParams  = function(data, opt_isPassword)
	{
		if (opt_isPassword) {
			if (this.asc_checkNeedCallback("asc_onAdvancedOptions")) {
				this.sendEvent("asc_onAdvancedOptions", c_oAscAdvancedOptionsID.DRM, data ? data.message : undefined);
			} else {
				this.sendEvent("asc_onError", c_oAscError.ID.ConvertationPassword, c_oAscError.Level.Critical);
			}
		}
	};

	asc_docs_api.prototype._onEndLoadSdk  = function()
	{
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk.call(this);

		History           = AscCommon.History;
		PasteElementsId   = AscCommon.PasteElementsId;
		global_mouseEvent = AscCommon.global_mouseEvent;

		this.WordControl      = new AscCommonSlide.CEditorPage(this);
		this.WordControl.Name = this.HtmlElementName;

		this.ThemeLoader     = new AscCommonSlide.CThemeLoader();
		this.ThemeLoader.Api = this;

		//выставляем тип copypaste
		PasteElementsId.g_bIsDocumentCopyPaste = false;

		this.CreateComponents();
		this.WordControl.Init();

		if (AscCommon.g_oTextMeasurer.SetParams)
		{
			AscCommon.g_oTextMeasurer.SetParams({ mode : "slide" });
		}

        if (this.tmpFontRenderingMode)
        {
            this.SetFontRenderingMode(this.tmpFontRenderingMode);
        }
		if (this.tmpThemesPath)
		{
			this.SetThemesPath(this.tmpThemesPath);
		}
		if (null !== this.tmpIsFreeze)
		{
			this.SetDrawingFreeze(this.tmpIsFreeze);
		}
		if (this.tmpSlideDiv)
		{
			this.SetInterfaceDrawImagePlaceSlide(this.tmpSlideDiv);
		}
		if (this.tmpTextArtDiv)
		{
			this.SetInterfaceDrawImagePlaceTextArt(this.tmpTextArtDiv);
		}
		if (null !== this.tmpViewRulers)
		{
			this.asc_SetViewRulers(this.tmpViewRulers);
		}
		if (null !== this.tmpZoomType)
		{
			switch (this.tmpZoomType)
			{
				case AscCommon.c_oZoomType.FitToPage:
					this.zoomFitToPage();
					break;
				case AscCommon.c_oZoomType.FitToWidth:
					this.zoomFitToWidth();
					break;
				case AscCommon.c_oZoomType.CustomMode:
					this.zoomCustomMode();
					break;
			}
		}
        if (null != this.tmpDocumentUnits)
        {
            this.asc_SetDocumentUnits(this.tmpDocumentUnits);
            this.tmpDocumentUnits = null;
        }

		this.asc_setViewMode(this.isViewMode);

		if (this.isReporterMode)
		{
			var _onbeforeunload = function ()
			{
				window.editor.EndDemonstration();
			};
			if (window.attachEvent)
				window.attachEvent('onbeforeunload', _onbeforeunload);
			else
				window.addEventListener('beforeunload', _onbeforeunload, false);
		}

        if (this.openFileCryptBinary)
        {
            this.openFileCryptCallback(this.openFileCryptBinary);
        }
	};
	asc_docs_api.prototype.initCollaborativeEditing = function()
	{
		if (AscCommon.CollaborativeEditing)
			return;
		
		AscCommon.CollaborativeEditing = new AscCommon.SlideCollaborativeEditing();
	};
	
	asc_docs_api.prototype._downloadAs = function(actionType, options, oAdditionalData, dataContainer, downloadType)
	{
		var t = this;
		var fileType = options.fileType;

		if (this.isCloudSaveAsLocalToDrawingFormat(actionType, fileType))
		{
			this.localSaveToDrawingFormat(this.WordControl.m_oDrawingDocument.ToRendererPart(false, options.isPdfPrint), fileType);
			return true;
		}

		if (c_oAscFileType.PDF === fileType || c_oAscFileType.PDFA === fileType)
		{
			var isSelection = false;
			if (options.advancedOptions && options.advancedOptions && (Asc.c_oAscPrintType.Selection === options.advancedOptions.asc_getPrintType()))
				isSelection = true;

			var dd             = this.WordControl.m_oDrawingDocument;
			dataContainer.data = dd.ToRendererPart(oAdditionalData["nobase64"], isSelection);
		}
		else if (this.insertDocumentUrlsData) {
			var last = this.insertDocumentUrlsData.documents.shift();
			oAdditionalData['url'] = last.url;
			oAdditionalData['format'] = last.format;
			if (last.token) {
				oAdditionalData['tokenDownload'] = last.token;
				//remove to reduce message size
				oAdditionalData['tokenSession'] = undefined;
			}
			oAdditionalData['outputurls']= true;
			// ToDo select txt params
			oAdditionalData["codepage"] = AscCommon.c_oAscCodePageUtf8;
			dataContainer.data = last.data;
		}
		else if(this.isOpenOOXInBrowser && this["asc_isSupportFeature"]("ooxml"))
		{
			var title = this.documentTitle;
			this.saveLogicDocumentToZip(undefined, undefined,
				function(data) {
					if (data) {
						if (c_oAscFileType.PPTX === fileType && !window.isCloudCryptoDownloadAs) {
							AscCommon.DownloadFileFromBytes(data, title, AscCommon.openXml.GetMimeType("pptx"));
						} else {
							dataContainer.data = data;
							if (window.isCloudCryptoDownloadAs)
							{
								window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType);
								return true;
							}
							t._downloadAsUsingServer(actionType, options, oAdditionalData, dataContainer, downloadType);
							return;
						}
					} else {
						t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
					}
					if (actionType)
					{
						t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, actionType);
					}
				});
			return true;
		}
		else
		{
			dataContainer.data = this.WordControl.SaveDocument(oAdditionalData["nobase64"]);
		}

        if (window.isCloudCryptoDownloadAs)
        {
        	window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType);
			return true;
        }
	};

    asc_docs_api.prototype.SetFontRenderingMode         = function(mode)
    {
        if (!this.isLoadFullApi)
        {
            this.tmpFontRenderingMode = mode;
            return;
        }

        if (c_oAscFontRenderingModeType.noHinting === mode)
            AscCommon.g_fontManager.SetHintsProps(false, false);
        else if (c_oAscFontRenderingModeType.hinting === mode)
            AscCommon.g_fontManager.SetHintsProps(true, false);
        else if (c_oAscFontRenderingModeType.hintingAndSubpixeling === mode)
            AscCommon.g_fontManager.SetHintsProps(true, true);

        if (AscCommon.g_fontManager2 !== undefined && AscCommon.g_fontManager2 !== null)
            AscCommon.g_fontManager2.ClearFontsRasterCache();

        this.WordControl.m_oDrawingDocument.ClearCachePages();

        if (this.bInit_word_control)
            this.WordControl.OnScroll();
    };

	asc_docs_api.prototype.processSavedFile             = function(url, downloadType, filetype)
	{
		if (this.insertDocumentUrlsData && this.insertDocumentUrlsData.convertCallback)
		{
			this.insertDocumentUrlsData.convertCallback(this, url);
		}
		else
		{
			AscCommon.baseEditorsApi.prototype.processSavedFile.call(this, url, downloadType, filetype);
		}
	};

	asc_docs_api.prototype.asc_Recalculate = function(bIsUpdateInterface)
	{
		if (!this.WordControl.m_oLogicDocument)
			return;
		this.WordControl.m_oLogicDocument.Recalculate({Drawings : {All : true, Map : {}}});
		this.WordControl.m_oLogicDocument.DrawingDocument.OnEndRecalculate();
	};

	asc_docs_api.prototype.asc_canPaste = function()
	{
		if (!this.WordControl ||
			!this.WordControl.m_oLogicDocument ||
			this.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props))
			return false;

		this.WordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_PasteHotKey);
		return true;
	};
	asc_docs_api.prototype.canRunBuilderScript = function()
	{
		let logicDocument = this.getLogicDocument();
		if (!logicDocument)
			return false;

		logicDocument.StartAction(AscDFH.historydescription_BuilderScript);
		return true;
	};
	asc_docs_api.prototype._onEndBuilderScript = function(callback)
	{
		this.asc_Recalculate();
		let logicDocument = this.getLogicDocument();
		logicDocument.FinalizeAction(true, true);
		
		if (callback)
			callback(true)
		
		return true;
	};

	asc_docs_api.prototype.asc_SetAutomaticBulletedLists = function(isAuto)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutomaticBulletedLists(isAuto);
	};
	asc_docs_api.prototype.asc_SetAutomaticNumberedLists = function(isAuto)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutomaticNumberedLists(isAuto);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectSmartQuotes = function(isSmartQuotes)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectSmartQuotes(isSmartQuotes);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectHyphensWithDash = function(isReplace)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectHyphensWithDash(isReplace);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfSentences = function(isCorrect)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectFirstLetterOfSentences(isCorrect);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectHyperlinks = function(isCorrect)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectHyperlinks(isCorrect);
	};
	asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfCells = function(isCorrect)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectFirstLetterOfCells(isCorrect);
	};

	asc_docs_api.prototype.asc_SetAutoCorrectDoubleSpaceWithPeriod = function(isCorrect)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		return oLogicDocument.SetAutoCorrectDoubleSpaceWithPeriod(isCorrect);
	};

	asc_docs_api.prototype.asc_SetFirstLetterAutoCorrectExceptions = function(exceptions, lang)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		let exceptionsManager = oLogicDocument.GetAutoCorrectSettings().GetFirstLetterExceptionManager();
		return exceptionsManager.SetExceptions(exceptions, lang);
	};
	asc_docs_api.prototype.asc_GetFirstLetterAutoCorrectExceptions = function(lang)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return [];

		let exceptionsManager = oLogicDocument.GetAutoCorrectSettings().GetFirstLetterExceptionManager();
		return exceptionsManager.GetExceptions(lang);
	};
	asc_docs_api.prototype.asc_GetAutoCorrectSettings = function()
	{
		let logicDocument = this.WordControl.m_oLogicDocument;
		if (!logicDocument)
			return [];

		return logicDocument.GetAutoCorrectSettings();
	};

	asc_docs_api.prototype.asc_GetSelectedText = function(bClearText, select_Pr)
	{
		bClearText = typeof(bClearText) === "boolean" ? bClearText : false;
		var oPresentation = this.WordControl && this.WordControl.m_oLogicDocument;
		if(oPresentation)
		{
			return oPresentation.GetSelectedText(bClearText, select_Pr);
		}
		return "";
	};

	// input
	asc_docs_api.prototype.asc_enterText = function(codePoints)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;

		return logicDocument.EnterText(codePoints);
	};

	asc_docs_api.prototype.asc_correctEnterText = function(oldValue, newValue)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;

		return logicDocument.CorrectEnterText(oldValue, newValue);
	};

	asc_docs_api.prototype.Begin_CompositeInput = function()
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.Begin_CompositeInput();
		return null;
	};
	asc_docs_api.prototype.Add_CompositeText = function(nCharCode)
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.Add_CompositeText(nCharCode);
		return null;
	};
	asc_docs_api.prototype.Remove_CompositeText = function(nCount)
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.Remove_CompositeText(nCount);
		return null;
	};
	asc_docs_api.prototype.Replace_CompositeText = function(arrCharCodes)
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.Replace_CompositeText(arrCharCodes);
		return null;
	};
	asc_docs_api.prototype.Set_CursorPosInCompositeText = function(nPos)
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.Set_CursorPosInCompositeText(nPos);
		return null;
	};
	asc_docs_api.prototype.Get_CursorPosInCompositeText = function()
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.Get_CursorPosInCompositeText();
		return 0;
	};
	asc_docs_api.prototype.End_CompositeInput = function()
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.End_CompositeInput();
		return null;
	};
	asc_docs_api.prototype.Get_MaxCursorPosInCompositeText = function()
	{
		if (this.WordControl.m_oLogicDocument)
			return this.WordControl.m_oLogicDocument.Get_MaxCursorPosInCompositeText();
		return 0;
	};
	asc_docs_api.prototype.Input_UpdatePos = function()
	{
		if (this.WordControl.m_oLogicDocument)
			this.WordControl.m_oDrawingDocument.MoveTargetInInputContext();
	};

	asc_docs_api.prototype.onKeyDown = function(e)
	{
		return this.WordControl.onKeyDown(e);
	};
	asc_docs_api.prototype.onKeyPress = function(e)
	{
		return this.WordControl.onKeyPress(e);
	};
	asc_docs_api.prototype.onKeyUp = function(e)
	{
		return this.WordControl.onKeyUp(e);
	};

    asc_docs_api.prototype.getAddedTextOnKeyDown = function(e)
    {
        var oLogicDocument = this.WordControl.m_oLogicDocument;
        if (!oLogicDocument)
            return [];

        return oLogicDocument.GetAddedTextOnKeyDown(e);
    };
	asc_docs_api.prototype.sync_OnConvertEquationToMath = function(oEquation)
	{
		this.sendEvent("asc_onConvertEquationToMath", oEquation);
	};
	asc_docs_api.prototype.asc_ConvertEquationToMath = function(oEquation, isAll)
	{
		var oLogicDocument = this.WordControl.m_oLogicDocument;

		// TODO: Вообще здесь нужно запрашивать шрифты, которые использовались в старой формуле,
		//      но пока это только 1 шрифт "Cambria Math".
		var loader   = AscCommon.g_font_loader;
		var fontinfo = AscFonts.g_fontApplication.GetFontInfo("Cambria Math");
		var isasync  = loader.LoadFont(fontinfo, function()
		{
			oLogicDocument.ConvertEquationToMath(oEquation, isAll);
		}, this);

		if (false === isasync)
		{
			oLogicDocument.ConvertEquationToMath(oEquation, isAll);
		}
	};

    //test
	window["asc_docs_api"]                                 = asc_docs_api;
	window["asc_docs_api"].prototype["asc_nativeOpenFile"] = function(base64File, version)
	{
		this.SpellCheckUrl = '';

		this.User = new AscCommon.asc_CUser();
		this.User.setId("TM");
		this.User.setUserName("native");

		this.WordControl.m_bIsRuler = false;
		this.WordControl.Init();

		this.InitEditor();

		g_oIdCounter.Set_Load(true);

		this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(base64File);
		if (this.isOpenOOXInBrowser) {
			//slice because array contains garbage after end of function
			this.openOOXInBrowserZip = base64File.slice();
			this.OpenDocumentFromZipNoInit(base64File);
		} else {
			var _loader = new AscCommon.BinaryPPTYLoader();
			_loader.Api = this;

			_loader.Load(base64File, this.WordControl.m_oLogicDocument);
		}

		this.LoadedObject = 1;
		g_oIdCounter.Set_Load(false);
	};

	window["asc_docs_api"].prototype["asc_nativeCalculateFile"] = function()
	{
		this.bNoSendComments = false;
		this.ShowParaMarks   = false;

		var presentation = this.WordControl.m_oLogicDocument;
		if(presentation){
			presentation.Recalculate({Drawings : {All : true, Map : {}}});
			presentation.DrawingDocument.OnEndRecalculate();
		}

	};

	window["asc_docs_api"].prototype["asc_nativeApplyChanges"] = function(changes)
	{
		var _len = changes.length;
		for (var i = 0; i < _len; i++)
		{
			var Changes = new AscCommon.CCollaborativeChanges();
			Changes.Set_Data(changes[i]);
			AscCommon.CollaborativeEditing.Add_Changes(Changes);
		}
		AscCommon.CollaborativeEditing.Apply_OtherChanges();
	};

	window["asc_docs_api"].prototype["asc_nativeApplyChanges2"] = function(data, isFull)
	{
		// Чтобы заново созданные параграфы не отображались залоченными
		g_oIdCounter.Set_Load(true);

		var stream = new AscCommon.FT_Stream2(data, data.length);
		stream.obj = null;
		var Loader = {Reader : stream, Reader2 : null};
		var _color = new AscCommonWord.CDocumentColor(191, 255, 199);

		// Применяем изменения, пока они есть
		var _count = Loader.Reader.GetLong();

		var _pos = 4;
		for (var i = 0; i < _count; i++)
		{
			if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["CheckNextChange"])
			{
				if (!window["native"]["CheckNextChange"]())
					break;
			}

            var nChangeLen = stream.GetLong();
            _pos += 4;
            stream.size = _pos + nChangeLen;

            var ClassId = stream.GetString2();
            var Class   = AscCommon.g_oTableId.Get_ById(ClassId);

            var nReaderPos  = stream.GetCurPos();
            var nChangeType = stream.GetLong();

            if (Class)
            {
                var fChangesClass = AscDFH.changesFactory[nChangeType];
                if (fChangesClass)
                {
                    var oChange = new fChangesClass(Class);
                    oChange.ReadFromBinary(stream);

                    if (true === AscCommon.CollaborativeEditing.private_AddOverallChange(oChange, false))
                        oChange.Load(_color);
                }
                else
                {
                    AscCommon.CollaborativeEditing.private_AddOverallChange(data, false);

                    stream.Seek(nReaderPos);
                    stream.Seek2(nReaderPos);

                    Class.Load_Changes(stream, null, _color);
                }
            }

            _pos += nChangeLen;
            stream.Seek2(_pos);
            stream.size = data.length;
		}

		if (isFull)
		{
			AscCommon.CollaborativeEditing.m_aChanges = [];
			
			// Делаем проверки корректности новых изменений
			AscCommon.CollaborativeEditing.Check_MergeData();

			AscCommon.CollaborativeEditing.OnEnd_ReadForeignChanges();

			if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["AddImageInChanges"])
			{
				var _new_images     = AscCommon.CollaborativeEditing.m_aNewImages;
				var _new_images_len = _new_images.length;

				for (var nImage = 0; nImage < _new_images_len; nImage++)
					window["native"]["AddImageInChanges"](_new_images[nImage]);
			}
		}

		g_oIdCounter.Set_Load(false);
	};

	window["asc_docs_api"].prototype["asc_nativeGetFile"] = function()
	{
		var writer = new AscCommon.CBinaryFileWriter();
		this.WordControl.m_oLogicDocument.CalculateComments();
		return writer.WriteDocument(this.WordControl.m_oLogicDocument);
	};

    window["asc_docs_api"].prototype.asc_nativeGetFile3 = function()
    {
        var writer = new AscCommon.CBinaryFileWriter();
        this.WordControl.m_oLogicDocument.CalculateComments();
        return { data: writer.WriteDocument3(this.WordControl.m_oLogicDocument, true), header: ("PPTY;v10;" + writer.pos + ";") };
    };

	window["asc_docs_api"].prototype["asc_nativeGetFileData"] = function()
	{
		if (this.isOpenOOXInBrowser && this["asc_isSupportFeature"]("ooxml")) {
			let res;
			this.saveLogicDocumentToZip(undefined, undefined, function(data) {
				res = data;
			});
			if (res) {
				window["native"]["Save_End"](";v10;", res.length);
				return res;
			}
			return new Uint8Array(0);
		} else {
			var writer = new AscCommon.CBinaryFileWriter();
			this.WordControl.m_oLogicDocument.CalculateComments();
			writer.WriteDocument3(this.WordControl.m_oLogicDocument);

			var _header = "PPTY;v10;" + writer.pos + ";";
			window["native"]["Save_End"](_header, writer.pos);

			return writer.data;
		}
	};

	window["asc_docs_api"].prototype["asc_nativeCalculate"] = function()
	{
	};

	window["asc_docs_api"].prototype["asc_nativePrint"] = function(_printer, _page, _options)
	{
		if (undefined === _printer && _page === undefined)
		{
			if (undefined !== window["AscDesktopEditor"])
			{
				var isSelection = (_options && _options["printOptions"] && _options["printOptions"]["selection"]) ? true : false;
				var _drawing_document = this.WordControl.m_oDrawingDocument;
				var pagescount        = _drawing_document.GetSlidesCount();
                if (isSelection)
                    pagescount = this.WordControl.Thumbnails.GetSelectedArray().length;

				window["AscDesktopEditor"]["Print_Start"](this.DocumentUrl, pagescount, this.ThemeLoader.ThemesUrl, this.getCurrentPage());

				var oDocRenderer                         = new AscCommon.CDocumentRenderer();
                oDocRenderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
				oDocRenderer.VectorMemoryForPrint        = new AscCommon.CMemory();
				var bOldShowMarks                        = this.ShowParaMarks;
				this.ShowParaMarks                       = false;
				oDocRenderer.IsNoDrawingEmptyPlaceholder = true;

                pagescount = _drawing_document.GetSlidesCount();
				for (var i = 0; i < pagescount; i++)
				{
					if (isSelection && !this.WordControl.Thumbnails.isSelectedPage(i))
						continue;

					oDocRenderer.Memory.Seek(0);
					oDocRenderer.VectorMemoryForPrint.ClearNoAttack();

					oDocRenderer.BeginPage(_drawing_document.m_oLogicDocument.GetWidthMM(), _drawing_document.m_oLogicDocument.GetHeightMM());
					this.WordControl.m_oLogicDocument.DrawPage(i, oDocRenderer);
					oDocRenderer.EndPage();

					window["AscDesktopEditor"]["Print_Page"](oDocRenderer.Memory.GetBase64Memory(), _drawing_document.m_oLogicDocument.GetWidthMM(), _drawing_document.m_oLogicDocument.GetHeightMM());
				}

				if (0 == pagescount)
				{
					oDocRenderer.BeginPage(_drawing_document.m_oLogicDocument.GetWidthMM(), _drawing_document.m_oLogicDocument.GetHeightMM());
					oDocRenderer.EndPage();

					window["AscDesktopEditor"]["Print_Page"](oDocRenderer.Memory.GetBase64Memory());
				}

				this.ShowParaMarks = bOldShowMarks;

				window["AscDesktopEditor"]["Print_End"]();
			}
			return;
		}

		let _logic_doc = this.WordControl.m_oLogicDocument;
		let pageW = _logic_doc.GetWidthMM();
		let pageH = _logic_doc.GetHeightMM();

		_printer.BeginPage(pageW, pageH);
		_logic_doc.DrawPage(_page, _printer);
		if (this.watermarkDraw)
			this.watermarkDraw.DrawOnRenderer(_printer, pageW, pageH);
		_printer.EndPage();
	};

	window["asc_docs_api"].prototype["asc_nativePrintPagesCount"] = function()
	{
		return this.WordControl.GetSlidesCount();
	};

	window["asc_docs_api"].prototype["asc_nativeGetPDF"] = function(options)
	{
		if (options && options["watermark"])
		{
			this.watermarkDraw = new AscCommon.CWatermarkOnDraw(options["watermark"], this);
			this.watermarkDraw.generateNative();
		}
		else
		{
			this.watermarkDraw = null;
		}

		var pagescount = this["asc_nativePrintPagesCount"]();
        if (options && options["printOptions"] && options["printOptions"]["onlyFirstPage"])
            pagescount = 1;

		var _renderer                         = new AscCommon.CDocumentRenderer();
        _renderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
		_renderer.VectorMemoryForPrint        = new AscCommon.CMemory();
		_renderer.DocInfo(this.asc_getCoreProps());
		var _bOldShowMarks                    = this.ShowParaMarks;
		this.ShowParaMarks                    = false;
		_renderer.IsNoDrawingEmptyPlaceholder = true;

		let nativeOptions = options ? options["nativeOptions"] : undefined;
		let pages = nativeOptions ? AscCommon.getNativePrintRanges(nativeOptions["pages"], nativeOptions["currentPage"], pagescount) : undefined;

		for (var i = 0; i < pagescount; i++)
		{
			if (pages !== undefined && !pages[i])
				continue;
			this["asc_nativePrint"](_renderer, i, options);
		}

		this.ShowParaMarks = _bOldShowMarks;

		window["native"]["Save_End"]("", _renderer.Memory.GetCurPosition());

		return _renderer.Memory.data;
	};

    window["asc_docs_api"].prototype["asc_nativeGetThemeThumbnail"] = function(params)
    {
        if (!this.WordControl.m_oLogicDocument ||
            !this.WordControl.m_oLogicDocument.slideMasters ||
            !this.WordControl.m_oLogicDocument.slideMasters[0] ||
			!this.WordControl.m_oLogicDocument.slideMasters[0].Theme)
		{
			return null;
		}

		var _pres = this.WordControl.m_oLogicDocument;
        var _master = this.WordControl.m_oLogicDocument.slideMasters[0];

    	var _renderer                         = new AscCommon.CDocumentRenderer();
        _renderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
        _renderer.VectorMemoryForPrint        = new AscCommon.CMemory();
        var _bOldShowMarks                    = this.ShowParaMarks;
        this.ShowParaMarks                    = false;
        _renderer.IsNoDrawingEmptyPlaceholder = true;

        var pxW = AscCommon.GlobalSkin.THEMES_THUMBNAIL_WIDTH; if (params && params.length && params[0]) pxW = params[0];
        var pxH = AscCommon.GlobalSkin.THEMES_THUMBNAIL_HEIGHT; if (params && params.length && params[1]) pxH = params[1];
        var mmW = pxW * AscCommon.g_dKoef_pix_to_mm;
        var mmH = pxH * AscCommon.g_dKoef_pix_to_mm;

        _renderer.BeginPage(mmW, mmH);
        //var oldEngine = window["NATIVE_EDITOR_ENJINE"];
        //window["NATIVE_EDITOR_ENJINE"] = undefined;
        this.WordControl.m_oMasterDrawer.WidthMM = mmW;
        this.WordControl.m_oMasterDrawer.HeightMM = mmH;
        this.WordControl.m_oMasterDrawer.WidthPx = pxW;
        this.WordControl.m_oMasterDrawer.HeightPx = pxH;
        this.WordControl.m_oMasterDrawer.Draw2(_renderer, _master, undefined, undefined, params);
        //window["NATIVE_EDITOR_ENJINE"] = oldEngine;
        _renderer.EndPage();

        this.ShowParaMarks = _bOldShowMarks;

        var objectRet = {};
        objectRet["name"] = _master.Theme.name;
        objectRet["data"] = _renderer.Memory.data;
        objectRet["dataLen"] = _renderer.Memory.GetCurPosition();
        return objectRet;
    };

	asc_docs_api.prototype.asc_OnHideContextMenu = function()
	{
		if (this.WordControl.MobileTouchManager)
		{
			this.WordControl.checkBodyOffset();
			this.WordControl.MobileTouchManager.showKeyboard();
		}
	};
	asc_docs_api.prototype.asc_OnShowContextMenu = function()
	{
		if (this.WordControl.MobileTouchManager)
		{
			this.WordControl.checkBodyOffset();
		}
	};

	asc_docs_api.prototype.getDefaultFontFamily = function () {
		//TODO переделать и отдавать дефолтовый шрифт
		var defaultFont = "Arial";
		return defaultFont;
	};

	asc_docs_api.prototype.getDefaultFontSize = function () {
		//TODO переделать и отдавать дефолтовый шрифт
		var defaultSize = 11;
		return defaultSize;
	};

	asc_docs_api.prototype.asc_getAppProps = function()
	{
		return this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.App || null;
	};

	asc_docs_api.prototype.getInternalCoreProps = function()
	{
		return this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.Core;
	};

	asc_docs_api.prototype.asc_setCoreProps = function(oProps)
	{
		var oCore = this.getInternalCoreProps();
		if(!oCore)
		{
			return;
		}
		var oLogicDocument = this.WordControl.m_oLogicDocument;
		if(false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_CorePr, null))
		{
			oLogicDocument.StartAction(AscDFH.historydescription_SetCoreproperties);
			oCore.setProps(oProps);
			oLogicDocument.FinalizeAction(true);
			this.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		}
	};
	asc_docs_api.prototype.asc_Remove = function()
	{
		var oPresentation = (this.WordControl && this.WordControl.m_oLogicDocument);
		if(oPresentation)
		{
			var oController = oPresentation.GetCurrentController();
			if(oController)
			{
				oController.resetChartElementsSelection();
				AscCommon.baseEditorsApi.prototype.asc_Remove.call(this);
			}
		}
	};

	asc_docs_api.prototype.turnOffSpecialModes = function() {
		let bResult = false;
		const oPresentation = this.WordControl.m_oLogicDocument;
		if(!oPresentation)
			return;
		const oDrawingObjects = oPresentation.GetCurrentController();
		if(!oDrawingObjects)
			return;
		if (oDrawingObjects.isTrackingDrawings()) {
			oDrawingObjects.endTrackNewShape();
			this.sync_EndAddShape();
		}
		if(this.isEyedropperStarted())
		{
			this.cancelEyedropper();
			bResult = true;
		}
		if (this.isFormatPainterOn()) {
			this.formatPainter.putState(AscCommon.c_oAscFormatPainterState.kOff);
			bResult = true;
		}
		if(this.isInkDrawerOn()) {
			this.stopInkDrawer();
			bResult = true;
		}
		return bResult;
	};

	asc_docs_api.prototype.onUpdateRestrictions = function()
	{
		if (!this.WordControl || !this.WordControl.m_oLogicDocument || !this.WordControl.m_oDrawingDocument)
			return;

		var logicDocument = this.WordControl.m_oLogicDocument;
		var slide = (logicDocument.CurPage >= 0) ? logicDocument.Slides[logicDocument.CurPage] : null;
		if (slide)
		{
			if (this.WordControl.m_oDrawingDocument.placeholders)
				this.WordControl.m_oDrawingDocument.placeholders.update(slide.getPlaceholdersControls());
		}
		this.turnOffSpecialModes();
		logicDocument.Document_UpdateInterfaceState();
	};

	asc_docs_api.prototype.isShowShapeAdjustments = function()
	{
		return this.canEdit();
	};
	asc_docs_api.prototype.isShowTableAdjustments = function()
	{
		return this.canEdit();
	};
	asc_docs_api.prototype.isShowEquationTrack = function()
	{
		return this.canEdit();
	};

	asc_docs_api.prototype.asc_EditSelectAll = function()
	{
		let oPresentation = (this.WordControl && this.WordControl.m_oLogicDocument);
		if (!oPresentation)
			return;

		oPresentation.SelectAll();
	};

	// print-preview
	asc_docs_api.prototype.asc_initPrintPreview = function(containerId, options)
	{
		if (this.printPreview)
			return;
		this.printPreview = new AscCommon.CPrintPreview(this, containerId);
	};
	asc_docs_api.prototype.asc_drawPrintPreview = function(index, paperSize)
	{
		if (this.printPreview)
		{
			this.printPreview.page = index;
			this.printPreview.update(paperSize);
		}
	};
	asc_docs_api.prototype.asc_closePrintPreview = function()
	{
		if (this.printPreview)
		{
			this.printPreview.close();
			delete this.printPreview;
		}
	};
	asc_docs_api.prototype.getEyedropperImgData = function()
	{
		const oViewerCanvas = document.getElementById("id_viewer");
		const oOverlayCanvas = document.getElementById("id_viewer_overlay");
		if(!oViewerCanvas || !oOverlayCanvas)
		{
			return null;
		}
		let oCanvas = document.createElement("canvas");
		oCanvas.width = oViewerCanvas.width;
		oCanvas.height = oViewerCanvas.height;
		const oCtx = oCanvas.getContext("2d");
		oCtx.drawImage(oViewerCanvas, 0, 0);
		oCtx.drawImage(oOverlayCanvas, 0, 0);
		return oCtx.getImageData(0, 0, oCanvas.width, oCanvas.height);
	};
	asc_docs_api.prototype.getSelectionState = function()
	{
		const oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
		{
			return null;
		}
		return oLogicDocument.GetSelectionState();
	};
	asc_docs_api.prototype.getSpeechDescription = function(prevState, action)
	{
		const oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
		{
			return null;
		}
		return oLogicDocument.GetSpeechDescription(prevState, action);
	};


	asc_docs_api.prototype.asc_changePresentationViewMode = function(mode) {
		if(this.presentationViewMode === mode) return;
		this.presentationViewMode = mode;
		this.updateViewMode();
		this.sendEvent("asc_onChangeViewMode", mode);
	};
	asc_docs_api.prototype.updateViewMode = function() {
		if(!this.WordControl) {
			return;
		}
		window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide(true);
		this.WordControl.UpdateViewMode();
	};


	
	asc_docs_api.prototype.getCustomProperties = function() {
		let oPresentation = this.private_GetLogicDocument();
		if(!oPresentation)
			return null;
		return oPresentation.CustomProperties;
	};

	asc_docs_api.prototype.addCustomProperty = function(name, type, value)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return;
		oLogicDocument.AddCustomProperty(name, type, value);
	};

	asc_docs_api.prototype.modifyCustomProperty = function(idx, name, type, value)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return;
		oLogicDocument.ModifyCustomProperty(idx, name, type, value);
	};

	asc_docs_api.prototype.removeCustomProperty = function(idx)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return;
		oLogicDocument.RemoveCustomProperty(idx);
	};

	asc_docs_api.prototype.asc_setViewerTargetType = function(mode)
	{
		this.mouseMode = mode;
		if (this.WordControl)
			this.WordControl.setMouseMode(mode);
	};


	asc_docs_api.prototype.asc_setDemoBackgroundColor = function(sColor)
	{
		this.demoBackgroundColor = sColor;
		if(this.isSlideShow())
		{
			this.WordControl.DemonstrationManager.CheckBackgroundColor();
		}
	};

	asc_docs_api.prototype.asc_EraseAllInksOnSlide = function()
	{
		let oLogicDocument = this.getLogicDocument();
		if(!oLogicDocument) return;

		oLogicDocument.EraseAllInksOnSlide();
	};
	asc_docs_api.prototype.removeAllInks = function () {
		let oLogicDocument = this.getLogicDocument();
		if(!oLogicDocument) return;

		oLogicDocument.RemoveAllInks();
	};
	asc_docs_api.prototype.haveInks = function() {
		let oLogicDocument = this.getLogicDocument();
		if(!oLogicDocument) return;

		return !!oLogicDocument.GetAllInks().length;
	};

	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']                                                 = window['Asc'] || {};
	window['AscCommonSlide']                                      = window['AscCommonSlide'] || {};
	window['Asc']['asc_docs_api'] = window['Asc'].asc_docs_api    = asc_docs_api;
	asc_docs_api.prototype['asc_GetFontThumbnailsPath']           = asc_docs_api.prototype.asc_GetFontThumbnailsPath;
	asc_docs_api.prototype['pre_Save']                            = asc_docs_api.prototype.pre_Save;
	asc_docs_api.prototype['sync_CollaborativeChanges']           = asc_docs_api.prototype.sync_CollaborativeChanges;
	asc_docs_api.prototype['asc_coAuthoringDisconnect']           = asc_docs_api.prototype.asc_coAuthoringDisconnect;
	asc_docs_api.prototype['asc_coAuthoringChatSendMessage']      = asc_docs_api.prototype.asc_coAuthoringChatSendMessage;
	asc_docs_api.prototype['asc_coAuthoringChatGetMessages']      = asc_docs_api.prototype.asc_coAuthoringChatGetMessages;
	asc_docs_api.prototype['asc_coAuthoringGetUsers']             = asc_docs_api.prototype.asc_coAuthoringGetUsers;
	asc_docs_api.prototype['syncCollaborativeChanges']            = asc_docs_api.prototype.syncCollaborativeChanges;
	asc_docs_api.prototype['SetCollaborativeMarksShowType']       = asc_docs_api.prototype.SetCollaborativeMarksShowType;
	asc_docs_api.prototype['GetCollaborativeMarksShowType']       = asc_docs_api.prototype.GetCollaborativeMarksShowType;
	asc_docs_api.prototype['Clear_CollaborativeMarks']            = asc_docs_api.prototype.Clear_CollaborativeMarks;
	asc_docs_api.prototype['_onUpdateDocumentCanSave']            = asc_docs_api.prototype._onUpdateDocumentCanSave;
	asc_docs_api.prototype['SetUnchangedDocument']                = asc_docs_api.prototype.SetUnchangedDocument;
	asc_docs_api.prototype['SetDocumentModified']                 = asc_docs_api.prototype.SetDocumentModified;
	asc_docs_api.prototype['isDocumentModified']                  = asc_docs_api.prototype.isDocumentModified;
	asc_docs_api.prototype['asc_isDocumentCanSave']               = asc_docs_api.prototype.asc_isDocumentCanSave;
	asc_docs_api.prototype['asc_getCanUndo']                      = asc_docs_api.prototype.asc_getCanUndo;
	asc_docs_api.prototype['asc_getCanRedo']                      = asc_docs_api.prototype.asc_getCanRedo;
	asc_docs_api.prototype['sync_BeginCatchSelectedElements']     = asc_docs_api.prototype.sync_BeginCatchSelectedElements;
	asc_docs_api.prototype['sync_EndCatchSelectedElements']       = asc_docs_api.prototype.sync_EndCatchSelectedElements;
	asc_docs_api.prototype['getSelectedElements']                 = asc_docs_api.prototype.getSelectedElements;
	asc_docs_api.prototype['asc_getAnimatedObjectsStack']         = asc_docs_api.prototype.asc_getAnimatedObjectsStack;
	asc_docs_api.prototype['sync_ChangeLastSelectedElement']      = asc_docs_api.prototype.sync_ChangeLastSelectedElement;
	asc_docs_api.prototype['asc_getEditorPermissions']            = asc_docs_api.prototype.asc_getEditorPermissions;
	asc_docs_api.prototype['asc_setDocInfo']                      = asc_docs_api.prototype.asc_setDocInfo;
	asc_docs_api.prototype['asc_changeDocInfo']                   = asc_docs_api.prototype.asc_changeDocInfo;
	asc_docs_api.prototype['asc_setLocale']                       = asc_docs_api.prototype.asc_setLocale;
	asc_docs_api.prototype['asc_getLocale']                       = asc_docs_api.prototype.asc_getLocale;
	asc_docs_api.prototype['asc_LoadDocument']                    = asc_docs_api.prototype.asc_LoadDocument;
	asc_docs_api.prototype['SetThemesPath']                       = asc_docs_api.prototype.SetThemesPath;
	asc_docs_api.prototype['InitEditor']                          = asc_docs_api.prototype.InitEditor;
	asc_docs_api.prototype['SetInterfaceDrawImagePlaceSlide']     = asc_docs_api.prototype.SetInterfaceDrawImagePlaceSlide;
	asc_docs_api.prototype['SetInterfaceDrawImagePlaceTextArt']   = asc_docs_api.prototype.SetInterfaceDrawImagePlaceTextArt;
	asc_docs_api.prototype['OpenDocumentFromBin']                 = asc_docs_api.prototype.OpenDocumentFromBin;
	asc_docs_api.prototype['asc_getDocumentName']                 = asc_docs_api.prototype.asc_getDocumentName;
	asc_docs_api.prototype['asc_getAppProps']                     = asc_docs_api.prototype.asc_getAppProps;
	asc_docs_api.prototype['asc_getCoreProps']                    = asc_docs_api.prototype.asc_getCoreProps;
	asc_docs_api.prototype['asc_setCoreProps']                    = asc_docs_api.prototype.asc_setCoreProps;
	asc_docs_api.prototype['asc_registerCallback']                = asc_docs_api.prototype.asc_registerCallback;
	asc_docs_api.prototype['asc_unregisterCallback']              = asc_docs_api.prototype.asc_unregisterCallback;
	asc_docs_api.prototype['asc_checkNeedCallback']               = asc_docs_api.prototype.asc_checkNeedCallback;
	asc_docs_api.prototype['get_TextProps']                       = asc_docs_api.prototype.get_TextProps;
	asc_docs_api.prototype['asc_getPropertyEditorShapes']         = asc_docs_api.prototype.asc_getPropertyEditorShapes;
	asc_docs_api.prototype['asc_getPropertyEditorTextArts']       = asc_docs_api.prototype.asc_getPropertyEditorTextArts;
	asc_docs_api.prototype['get_PropertyEditorThemes']            = asc_docs_api.prototype.get_PropertyEditorThemes;
	asc_docs_api.prototype['get_ContentCount']                    = asc_docs_api.prototype.get_ContentCount;
	asc_docs_api.prototype['UpdateTextPr']                        = asc_docs_api.prototype.UpdateTextPr;
	asc_docs_api.prototype['sync_TextSpacing']                    = asc_docs_api.prototype.sync_TextSpacing;
	asc_docs_api.prototype['sync_TextDStrikeout']                 = asc_docs_api.prototype.sync_TextDStrikeout;
	asc_docs_api.prototype['sync_TextCaps']                       = asc_docs_api.prototype.sync_TextCaps;
	asc_docs_api.prototype['sync_TextSmallCaps']                  = asc_docs_api.prototype.sync_TextSmallCaps;
	asc_docs_api.prototype['sync_TextPosition']                   = asc_docs_api.prototype.sync_TextPosition;
	asc_docs_api.prototype['sync_TextLangCallBack']               = asc_docs_api.prototype.sync_TextLangCallBack;
	asc_docs_api.prototype['sync_VerticalTextAlign']              = asc_docs_api.prototype.sync_VerticalTextAlign;
	asc_docs_api.prototype['sync_Vert']                           = asc_docs_api.prototype.sync_Vert;
	asc_docs_api.prototype['UpdateParagraphProp']                 = asc_docs_api.prototype.UpdateParagraphProp;
	asc_docs_api.prototype['Undo']                                = asc_docs_api.prototype.Undo;
	asc_docs_api.prototype['Redo']                                = asc_docs_api.prototype.Redo;
	asc_docs_api.prototype['Copy']                                = asc_docs_api.prototype.Copy;
	asc_docs_api.prototype['Update_ParaTab']                      = asc_docs_api.prototype.Update_ParaTab;
	asc_docs_api.prototype['Cut']                                 = asc_docs_api.prototype.Cut;
	asc_docs_api.prototype['Paste']                               = asc_docs_api.prototype.Paste;
	asc_docs_api.prototype['Share']                               = asc_docs_api.prototype.Share;
	asc_docs_api.prototype['asc_Save']                            = asc_docs_api.prototype.asc_Save;
	asc_docs_api.prototype['forceSave']                           = asc_docs_api.prototype.forceSave;
	asc_docs_api.prototype['asc_setIsForceSaveOnUserSave']        = asc_docs_api.prototype.asc_setIsForceSaveOnUserSave;
	asc_docs_api.prototype['asc_DownloadAs']                      = asc_docs_api.prototype.asc_DownloadAs;
	asc_docs_api.prototype['Resize']                              = asc_docs_api.prototype.Resize;
	asc_docs_api.prototype['AddURL']                              = asc_docs_api.prototype.AddURL;
	asc_docs_api.prototype['Help']                                = asc_docs_api.prototype.Help;
	asc_docs_api.prototype['asc_setAdvancedOptions']              = asc_docs_api.prototype.asc_setAdvancedOptions;
    asc_docs_api.prototype['SetFontRenderingMode']                = asc_docs_api.prototype.SetFontRenderingMode;
	asc_docs_api.prototype['stopGetDocInfo']                      = asc_docs_api.prototype.stopGetDocInfo;
	asc_docs_api.prototype['sync_CanUndoCallback']                = asc_docs_api.prototype.sync_CanUndoCallback;
	asc_docs_api.prototype['sync_CanRedoCallback']                = asc_docs_api.prototype.sync_CanRedoCallback;
	asc_docs_api.prototype['sync_CursorLockCallBack']             = asc_docs_api.prototype.sync_CursorLockCallBack;
	asc_docs_api.prototype['sync_UndoCallBack']                   = asc_docs_api.prototype.sync_UndoCallBack;
	asc_docs_api.prototype['sync_RedoCallBack']                   = asc_docs_api.prototype.sync_RedoCallBack;
	asc_docs_api.prototype['sync_CopyCallBack']                   = asc_docs_api.prototype.sync_CopyCallBack;
	asc_docs_api.prototype['sync_CutCallBack']                    = asc_docs_api.prototype.sync_CutCallBack;
	asc_docs_api.prototype['sync_PasteCallBack']                  = asc_docs_api.prototype.sync_PasteCallBack;
	asc_docs_api.prototype['sync_ShareCallBack']                  = asc_docs_api.prototype.sync_ShareCallBack;
	asc_docs_api.prototype['sync_SaveCallBack']                   = asc_docs_api.prototype.sync_SaveCallBack;
	asc_docs_api.prototype['sync_DownloadAsCallBack']             = asc_docs_api.prototype.sync_DownloadAsCallBack;
	asc_docs_api.prototype['sync_StartAction']                    = asc_docs_api.prototype.sync_StartAction;
	asc_docs_api.prototype['sync_EndAction']                      = asc_docs_api.prototype.sync_EndAction;
	asc_docs_api.prototype['sync_AddURLCallback']                 = asc_docs_api.prototype.sync_AddURLCallback;
	asc_docs_api.prototype['sync_ErrorCallback']                  = asc_docs_api.prototype.sync_ErrorCallback;
	asc_docs_api.prototype['sync_HelpCallback']                   = asc_docs_api.prototype.sync_HelpCallback;
	asc_docs_api.prototype['sync_UpdateZoom']                     = asc_docs_api.prototype.sync_UpdateZoom;
	asc_docs_api.prototype['ClearPropObjCallback']                = asc_docs_api.prototype.ClearPropObjCallback;
	asc_docs_api.prototype['CollectHeaders']                      = asc_docs_api.prototype.CollectHeaders;
	asc_docs_api.prototype['GetActiveHeader']                     = asc_docs_api.prototype.GetActiveHeader;
	asc_docs_api.prototype['sync_ChangeActiveHeaderCallback']     = asc_docs_api.prototype.sync_ChangeActiveHeaderCallback;
	asc_docs_api.prototype['sync_ReturnHeadersCallback']          = asc_docs_api.prototype.sync_ReturnHeadersCallback;
	asc_docs_api.prototype['startSearchText']                     = asc_docs_api.prototype.startSearchText;
	asc_docs_api.prototype['goToNextSearchResult']                = asc_docs_api.prototype.goToNextSearchResult;
	asc_docs_api.prototype['gotoSearchResultText']                = asc_docs_api.prototype.gotoSearchResultText;
	asc_docs_api.prototype['stopSearchText']                      = asc_docs_api.prototype.stopSearchText;
	asc_docs_api.prototype['asc_findText']                        = asc_docs_api.prototype.asc_findText;
	asc_docs_api.prototype['asc_endFindText']                     = asc_docs_api.prototype.asc_endFindText;
	asc_docs_api.prototype['asc_replaceText']                     = asc_docs_api.prototype.asc_replaceText;
	asc_docs_api.prototype['sync_SearchFoundCallback']            = asc_docs_api.prototype.sync_SearchFoundCallback;
	asc_docs_api.prototype['sync_SearchStartCallback']            = asc_docs_api.prototype.sync_SearchStartCallback;
	asc_docs_api.prototype['sync_SearchStopCallback']             = asc_docs_api.prototype.sync_SearchStopCallback;
	asc_docs_api.prototype['sync_SearchEndCallback']              = asc_docs_api.prototype.sync_SearchEndCallback;
	asc_docs_api.prototype['asc_StartTextAroundSearch']           = asc_docs_api.prototype.asc_StartTextAroundSearch;
	asc_docs_api.prototype['asc_SelectSearchElement']             = asc_docs_api.prototype.asc_SelectSearchElement;
	asc_docs_api.prototype['put_TextPrFontName']                  = asc_docs_api.prototype.put_TextPrFontName;
	asc_docs_api.prototype['put_TextPrFontSize']                  = asc_docs_api.prototype.put_TextPrFontSize;
	asc_docs_api.prototype['put_TextPrLang']                      = asc_docs_api.prototype.put_TextPrLang;
	asc_docs_api.prototype['put_TextPrBold']                      = asc_docs_api.prototype.put_TextPrBold;
	asc_docs_api.prototype['put_TextPrItalic']                    = asc_docs_api.prototype.put_TextPrItalic;
	asc_docs_api.prototype['put_TextPrUnderline']                 = asc_docs_api.prototype.put_TextPrUnderline;
	asc_docs_api.prototype['put_TextPrStrikeout']                 = asc_docs_api.prototype.put_TextPrStrikeout;
	asc_docs_api.prototype['put_PrLineSpacing']                   = asc_docs_api.prototype.put_PrLineSpacing;
	asc_docs_api.prototype['put_LineSpacingBeforeAfter']          = asc_docs_api.prototype.put_LineSpacingBeforeAfter;
	asc_docs_api.prototype['FontSizeIn']                          = asc_docs_api.prototype.FontSizeIn;
	asc_docs_api.prototype['FontSizeOut']                         = asc_docs_api.prototype.FontSizeOut;
	asc_docs_api.prototype['sync_BoldCallBack']                   = asc_docs_api.prototype.sync_BoldCallBack;
	asc_docs_api.prototype['sync_ItalicCallBack']                 = asc_docs_api.prototype.sync_ItalicCallBack;
	asc_docs_api.prototype['sync_UnderlineCallBack']              = asc_docs_api.prototype.sync_UnderlineCallBack;
	asc_docs_api.prototype['sync_StrikeoutCallBack']              = asc_docs_api.prototype.sync_StrikeoutCallBack;
	asc_docs_api.prototype['sync_TextPrFontFamilyCallBack']       = asc_docs_api.prototype.sync_TextPrFontFamilyCallBack;
	asc_docs_api.prototype['sync_TextPrFontSizeCallBack']         = asc_docs_api.prototype.sync_TextPrFontSizeCallBack;
	asc_docs_api.prototype['sync_PrLineSpacingCallBack']          = asc_docs_api.prototype.sync_PrLineSpacingCallBack;
	asc_docs_api.prototype['sync_InitEditorThemes']               = asc_docs_api.prototype.sync_InitEditorThemes;
	asc_docs_api.prototype['paraApply']                           = asc_docs_api.prototype.paraApply;
	asc_docs_api.prototype['put_PrAlign']                         = asc_docs_api.prototype.put_PrAlign;
	asc_docs_api.prototype['put_TextPrBaseline']                  = asc_docs_api.prototype.put_TextPrBaseline;
	asc_docs_api.prototype['put_ListType']                        = asc_docs_api.prototype.put_ListType;
	asc_docs_api.prototype['put_ShowSnapLines']                   = asc_docs_api.prototype.put_ShowSnapLines;
	asc_docs_api.prototype['get_ShowSnapLines']                   = asc_docs_api.prototype.get_ShowSnapLines;
	asc_docs_api.prototype['asc_setShowSmartGuides']              = asc_docs_api.prototype.asc_setShowSmartGuides;
	asc_docs_api.prototype['asc_getShowSmartGuides']              = asc_docs_api.prototype.asc_getShowSmartGuides;
	asc_docs_api.prototype['asc_setShowGuides']                   = asc_docs_api.prototype.asc_setShowGuides;
	asc_docs_api.prototype['asc_getShowGuides']                   = asc_docs_api.prototype.asc_getShowGuides;
	asc_docs_api.prototype['asc_getGuidesCount']                  = asc_docs_api.prototype.asc_getGuidesCount;
	asc_docs_api.prototype['asc_setShowGridlines']                = asc_docs_api.prototype.asc_setShowGridlines;
	asc_docs_api.prototype['asc_getShowGridlines']                = asc_docs_api.prototype.asc_getShowGridlines;
	asc_docs_api.prototype['asc_setGridSpacing']                  = asc_docs_api.prototype.asc_setGridSpacing;
	asc_docs_api.prototype['asc_getGridSpacing']                  = asc_docs_api.prototype.asc_getGridSpacing;
	asc_docs_api.prototype['asc_setSnapToGrid']                   = asc_docs_api.prototype.asc_setSnapToGrid;
	asc_docs_api.prototype['asc_getSnapToGrid']                   = asc_docs_api.prototype.asc_getSnapToGrid;
	asc_docs_api.prototype['asc_addHorizontalGuide']              = asc_docs_api.prototype.asc_addHorizontalGuide;
	asc_docs_api.prototype['asc_addVerticalGuide']                = asc_docs_api.prototype.asc_addVerticalGuide;
	asc_docs_api.prototype['asc_canClearGuides']                  = asc_docs_api.prototype.asc_canClearGuides;
	asc_docs_api.prototype['asc_clearGuides']                     = asc_docs_api.prototype.asc_clearGuides;
	asc_docs_api.prototype['asc_deleteGuide']                     = asc_docs_api.prototype.asc_deleteGuide;
	asc_docs_api.prototype['asc_SetMasterName']                   = asc_docs_api.prototype.asc_SetMasterName;
	asc_docs_api.prototype['asc_SetLayoutName']                   = asc_docs_api.prototype.asc_SetLayoutName;
	asc_docs_api.prototype['put_ShowParaMarks']                   = asc_docs_api.prototype.put_ShowParaMarks;
	asc_docs_api.prototype['get_ShowParaMarks']                   = asc_docs_api.prototype.get_ShowParaMarks;
	asc_docs_api.prototype['put_ShowTableEmptyLine']              = asc_docs_api.prototype.put_ShowTableEmptyLine;
	asc_docs_api.prototype['get_ShowTableEmptyLine']              = asc_docs_api.prototype.get_ShowTableEmptyLine;
	asc_docs_api.prototype['ShapeApply']                          = asc_docs_api.prototype.ShapeApply;
	asc_docs_api.prototype['setStartPointHistory']                = asc_docs_api.prototype.setStartPointHistory;
	asc_docs_api.prototype['setEndPointHistory']                  = asc_docs_api.prototype.setEndPointHistory;
	asc_docs_api.prototype['SetSlideProps']                       = asc_docs_api.prototype.SetSlideProps;
	asc_docs_api.prototype['put_LineCap']                         = asc_docs_api.prototype.put_LineCap;
	asc_docs_api.prototype['put_LineJoin']                        = asc_docs_api.prototype.put_LineJoin;
	asc_docs_api.prototype['put_LineBeginStyle']                  = asc_docs_api.prototype.put_LineBeginStyle;
	asc_docs_api.prototype['put_LineBeginSize']                   = asc_docs_api.prototype.put_LineBeginSize;
	asc_docs_api.prototype['put_LineEndStyle']                    = asc_docs_api.prototype.put_LineEndStyle;
	asc_docs_api.prototype['put_LineEndSize']                     = asc_docs_api.prototype.put_LineEndSize;
	asc_docs_api.prototype['put_TextColor2']                      = asc_docs_api.prototype.put_TextColor2;
	asc_docs_api.prototype['put_TextColor']                       = asc_docs_api.prototype.put_TextColor;
	asc_docs_api.prototype['put_PrIndent']                        = asc_docs_api.prototype.put_PrIndent;
	asc_docs_api.prototype['IncreaseIndent']                      = asc_docs_api.prototype.IncreaseIndent;
	asc_docs_api.prototype['DecreaseIndent']                      = asc_docs_api.prototype.DecreaseIndent;
	asc_docs_api.prototype['put_PrIndentRight']                   = asc_docs_api.prototype.put_PrIndentRight;
	asc_docs_api.prototype['put_PrFirstLineIndent']               = asc_docs_api.prototype.put_PrFirstLineIndent;
	asc_docs_api.prototype['getFocusObject']                      = asc_docs_api.prototype.getFocusObject;
	asc_docs_api.prototype['sync_VerticalAlign']                  = asc_docs_api.prototype.sync_VerticalAlign;
	asc_docs_api.prototype['sync_PrAlignCallBack']                = asc_docs_api.prototype.sync_PrAlignCallBack;
	asc_docs_api.prototype['sync_ListType']                       = asc_docs_api.prototype.sync_ListType;
	asc_docs_api.prototype['sync_TextColor']                      = asc_docs_api.prototype.sync_TextColor;
	asc_docs_api.prototype['sync_TextColor2']                     = asc_docs_api.prototype.sync_TextColor2;
	asc_docs_api.prototype['sync_TextHighLight']                  = asc_docs_api.prototype.sync_TextHighLight;
	asc_docs_api.prototype['sync_ParaStyleName']                  = asc_docs_api.prototype.sync_ParaStyleName;
	asc_docs_api.prototype['sync_ParaSpacingLine']                = asc_docs_api.prototype.sync_ParaSpacingLine;
	asc_docs_api.prototype['sync_PageBreakCallback']              = asc_docs_api.prototype.sync_PageBreakCallback;
	asc_docs_api.prototype['sync_KeepLinesCallback']              = asc_docs_api.prototype.sync_KeepLinesCallback;
	asc_docs_api.prototype['sync_ShowParaMarksCallback']          = asc_docs_api.prototype.sync_ShowParaMarksCallback;
	asc_docs_api.prototype['sync_SpaceBetweenPrgCallback']        = asc_docs_api.prototype.sync_SpaceBetweenPrgCallback;
	asc_docs_api.prototype['sync_PrPropCallback']                 = asc_docs_api.prototype.sync_PrPropCallback;
	asc_docs_api.prototype['SetDrawImagePlaceParagraph']          = asc_docs_api.prototype.SetDrawImagePlaceParagraph;
	asc_docs_api.prototype['put_AddPageBreak']                    = asc_docs_api.prototype.put_AddPageBreak;
	asc_docs_api.prototype['Update_ParaInd']                      = asc_docs_api.prototype.Update_ParaInd;
	asc_docs_api.prototype['Internal_Update_Ind_FirstLine']       = asc_docs_api.prototype.Internal_Update_Ind_FirstLine;
	asc_docs_api.prototype['Internal_Update_Ind_Left']            = asc_docs_api.prototype.Internal_Update_Ind_Left;
	asc_docs_api.prototype['Internal_Update_Ind_Right']           = asc_docs_api.prototype.Internal_Update_Ind_Right;
	asc_docs_api.prototype['put_Table']                           = asc_docs_api.prototype.put_Table;
	asc_docs_api.prototype['addRowAbove']                         = asc_docs_api.prototype.addRowAbove;
	asc_docs_api.prototype['addRowBelow']                         = asc_docs_api.prototype.addRowBelow;
	asc_docs_api.prototype['addColumnLeft']                       = asc_docs_api.prototype.addColumnLeft;
	asc_docs_api.prototype['addColumnRight']                      = asc_docs_api.prototype.addColumnRight;
	asc_docs_api.prototype['remRow']                              = asc_docs_api.prototype.remRow;
	asc_docs_api.prototype['remColumn']                           = asc_docs_api.prototype.remColumn;
	asc_docs_api.prototype['remTable']                            = asc_docs_api.prototype.remTable;
	asc_docs_api.prototype['asc_DistributeTableCells']            = asc_docs_api.prototype.asc_DistributeTableCells;
	asc_docs_api.prototype['selectRow']                           = asc_docs_api.prototype.selectRow;
	asc_docs_api.prototype['selectColumn']                        = asc_docs_api.prototype.selectColumn;
	asc_docs_api.prototype['selectCell']                          = asc_docs_api.prototype.selectCell;
	asc_docs_api.prototype['selectTable']                         = asc_docs_api.prototype.selectTable;
	asc_docs_api.prototype['setColumnWidth']                      = asc_docs_api.prototype.setColumnWidth;
	asc_docs_api.prototype['setRowHeight']                        = asc_docs_api.prototype.setRowHeight;
	asc_docs_api.prototype['set_TblDistanceFromText']             = asc_docs_api.prototype.set_TblDistanceFromText;
	asc_docs_api.prototype['CheckBeforeMergeCells']               = asc_docs_api.prototype.CheckBeforeMergeCells;
	asc_docs_api.prototype['CheckBeforeSplitCells']               = asc_docs_api.prototype.CheckBeforeSplitCells;
	asc_docs_api.prototype['MergeCells']                          = asc_docs_api.prototype.MergeCells;
	asc_docs_api.prototype['SplitCell']                           = asc_docs_api.prototype.SplitCell;
	asc_docs_api.prototype['widthTable']                          = asc_docs_api.prototype.widthTable;
	asc_docs_api.prototype['put_CellsMargin']                     = asc_docs_api.prototype.put_CellsMargin;
	asc_docs_api.prototype['set_TblWrap']                         = asc_docs_api.prototype.set_TblWrap;
	asc_docs_api.prototype['set_TblIndentLeft']                   = asc_docs_api.prototype.set_TblIndentLeft;
	asc_docs_api.prototype['set_Borders']                         = asc_docs_api.prototype.set_Borders;
	asc_docs_api.prototype['set_TableBackground']                 = asc_docs_api.prototype.set_TableBackground;
	asc_docs_api.prototype['set_AlignCell']                       = asc_docs_api.prototype.set_AlignCell;
	asc_docs_api.prototype['set_TblAlign']                        = asc_docs_api.prototype.set_TblAlign;
	asc_docs_api.prototype['set_SpacingBetweenCells']             = asc_docs_api.prototype.set_SpacingBetweenCells;
	asc_docs_api.prototype['tblApply']                            = asc_docs_api.prototype.tblApply;
	asc_docs_api.prototype['sync_AddTableCallback']               = asc_docs_api.prototype.sync_AddTableCallback;
	asc_docs_api.prototype['sync_AlignCellCallback']              = asc_docs_api.prototype.sync_AlignCellCallback;
	asc_docs_api.prototype['sync_TblPropCallback']                = asc_docs_api.prototype.sync_TblPropCallback;
	asc_docs_api.prototype['sync_TblWrapStyleChangedCallback']    = asc_docs_api.prototype.sync_TblWrapStyleChangedCallback;
	asc_docs_api.prototype['sync_TblAlignChangedCallback']        = asc_docs_api.prototype.sync_TblAlignChangedCallback;
	asc_docs_api.prototype['ChangeImageFromFile']                 = asc_docs_api.prototype.ChangeImageFromFile;
	asc_docs_api.prototype['ChangeShapeImageFromFile']            = asc_docs_api.prototype.ChangeShapeImageFromFile;
	asc_docs_api.prototype['ChangeSlideImageFromFile']            = asc_docs_api.prototype.ChangeSlideImageFromFile;
	asc_docs_api.prototype['ChangeArtImageFromFile']              = asc_docs_api.prototype.ChangeArtImageFromFile;
	asc_docs_api.prototype['AddImage']                            = asc_docs_api.prototype.AddImage;
	asc_docs_api.prototype['asc_addImage']                        = asc_docs_api.prototype.asc_addImage;
	asc_docs_api.prototype['asc_AddToLayout']                     = asc_docs_api.prototype.asc_AddToLayout;
	asc_docs_api.prototype['asc_AddAnimation']                    = asc_docs_api.prototype.asc_AddAnimation;
	asc_docs_api.prototype['asc_StartAnimationPreview']           = asc_docs_api.prototype.asc_StartAnimationPreview;
	asc_docs_api.prototype['asc_canStartAnimationPreview']        = asc_docs_api.prototype.asc_canStartAnimationPreview;
	asc_docs_api.prototype['asc_getCurSlideObjectsNames']         = asc_docs_api.prototype.asc_getCurSlideObjectsNames;
	asc_docs_api.prototype['asc_StopAnimationPreview']            = asc_docs_api.prototype.asc_StopAnimationPreview;
	asc_docs_api.prototype['asc_SetAnimationProperties']          = asc_docs_api.prototype.asc_SetAnimationProperties;
	asc_docs_api.prototype['asc_canMoveAnimationEarlier']         = asc_docs_api.prototype.asc_canMoveAnimationEarlier;
	asc_docs_api.prototype['asc_canMoveAnimationLater']           = asc_docs_api.prototype.asc_canMoveAnimationLater;
	asc_docs_api.prototype['asc_moveAnimationEarlier']            = asc_docs_api.prototype.asc_moveAnimationEarlier;
	asc_docs_api.prototype['asc_moveAnimationLater']              = asc_docs_api.prototype.asc_moveAnimationLater;
	asc_docs_api.prototype['asc_onShowAnimTab']                   = asc_docs_api.prototype.asc_onShowAnimTab;


	asc_docs_api.prototype['StartAddShape']                       = asc_docs_api.prototype.StartAddShape;
	asc_docs_api.prototype['asc_canEditGeometry']                 = asc_docs_api.prototype.asc_canEditGeometry;
	asc_docs_api.prototype['asc_editPointsGeometry']              = asc_docs_api.prototype.asc_editPointsGeometry;
	asc_docs_api.prototype['AddTextArt']                          = asc_docs_api.prototype.AddTextArt;
	asc_docs_api.prototype['asc_canEditCrop']                     = asc_docs_api.prototype.asc_canEditCrop;
	asc_docs_api.prototype['asc_startEditCrop']                   = asc_docs_api.prototype.asc_startEditCrop;
	asc_docs_api.prototype['asc_endEditCrop']                     = asc_docs_api.prototype.asc_endEditCrop;
	asc_docs_api.prototype['asc_cropFit']                         = asc_docs_api.prototype.asc_cropFit;
	asc_docs_api.prototype['asc_cropFill']                        = asc_docs_api.prototype.asc_cropFill;
	asc_docs_api.prototype['canGroup']                            = asc_docs_api.prototype.canGroup;
	asc_docs_api.prototype['canUnGroup']                          = asc_docs_api.prototype.canUnGroup;
	asc_docs_api.prototype['AddImageUrl']                         = asc_docs_api.prototype.AddImageUrl;
	asc_docs_api.prototype['AddImageUrlActionCallback']           = asc_docs_api.prototype.AddImageUrlActionCallback;
	asc_docs_api.prototype['AddImageUrlAction']                   = asc_docs_api.prototype.AddImageUrlAction;
	asc_docs_api.prototype['ImgApply']                            = asc_docs_api.prototype.ImgApply;
	asc_docs_api.prototype['ChartApply']                          = asc_docs_api.prototype.ChartApply;
	asc_docs_api.prototype['set_Size']                            = asc_docs_api.prototype.set_Size;
	asc_docs_api.prototype['set_ConstProportions']                = asc_docs_api.prototype.set_ConstProportions;
	asc_docs_api.prototype['set_WrapStyle']                       = asc_docs_api.prototype.set_WrapStyle;
	asc_docs_api.prototype['deleteImage']                         = asc_docs_api.prototype.deleteImage;
	asc_docs_api.prototype['set_ImgDistanceFromText']             = asc_docs_api.prototype.set_ImgDistanceFromText;
	asc_docs_api.prototype['set_PositionOnPage']                  = asc_docs_api.prototype.set_PositionOnPage;
	asc_docs_api.prototype['get_OriginalSizeImage']               = asc_docs_api.prototype.get_OriginalSizeImage;
	asc_docs_api.prototype['asc_FitImagesToSlide']                = asc_docs_api.prototype.asc_FitImagesToSlide;
	asc_docs_api.prototype['asc_onCloseChartFrame']               = asc_docs_api.prototype.asc_onCloseChartFrame;
	asc_docs_api.prototype['sync_AddImageCallback']               = asc_docs_api.prototype.sync_AddImageCallback;
	asc_docs_api.prototype['sync_ImgPropCallback']                = asc_docs_api.prototype.sync_ImgPropCallback;
	asc_docs_api.prototype['SetDrawingFreeze']                    = asc_docs_api.prototype.SetDrawingFreeze;
	asc_docs_api.prototype['zoomIn']                              = asc_docs_api.prototype.zoomIn;
	asc_docs_api.prototype['zoomOut']                             = asc_docs_api.prototype.zoomOut;
	asc_docs_api.prototype['zoomFitToPage']                       = asc_docs_api.prototype.zoomFitToPage;
	asc_docs_api.prototype['zoomFitToWidth']                      = asc_docs_api.prototype.zoomFitToWidth;
	asc_docs_api.prototype['zoomCustomMode']                      = asc_docs_api.prototype.zoomCustomMode;
	asc_docs_api.prototype['zoom100']                             = asc_docs_api.prototype.zoom100;
	asc_docs_api.prototype['zoom']                                = asc_docs_api.prototype.zoom;
	asc_docs_api.prototype['goToPage']                            = asc_docs_api.prototype.goToPage;
	asc_docs_api.prototype['getCountPages']                       = asc_docs_api.prototype.getCountPages;
	asc_docs_api.prototype['getCurrentPage']                      = asc_docs_api.prototype.getCurrentPage;
	asc_docs_api.prototype['sync_countPagesCallback']             = asc_docs_api.prototype.sync_countPagesCallback;
	asc_docs_api.prototype['sync_currentPageCallback']            = asc_docs_api.prototype.sync_currentPageCallback;
	asc_docs_api.prototype['sync_SendThemeColors']                = asc_docs_api.prototype.sync_SendThemeColors;
	asc_docs_api.prototype['ChangeColorScheme']                   = asc_docs_api.prototype.ChangeColorScheme;
	asc_docs_api.prototype['asc_ChangeColorSchemeByIdx']          = asc_docs_api.prototype.asc_ChangeColorSchemeByIdx;
	asc_docs_api.prototype['asc_enableKeyEvents']                 = asc_docs_api.prototype.asc_enableKeyEvents;
	asc_docs_api.prototype['asc_showComments']                    = asc_docs_api.prototype.asc_showComments;
	asc_docs_api.prototype['asc_hideComments']                    = asc_docs_api.prototype.asc_hideComments;
	asc_docs_api.prototype['asc_addComment']                      = asc_docs_api.prototype.asc_addComment;
	asc_docs_api.prototype['asc_getMasterCommentId']              = asc_docs_api.prototype.asc_getMasterCommentId;
	asc_docs_api.prototype['asc_getAnchorPosition']               = asc_docs_api.prototype.asc_getAnchorPosition;
	asc_docs_api.prototype['asc_removeComment']                   = asc_docs_api.prototype.asc_removeComment;
	asc_docs_api.prototype['asc_RemoveAllComments']               = asc_docs_api.prototype.asc_RemoveAllComments;
	asc_docs_api.prototype['asc_ResolveAllComments']              = asc_docs_api.prototype.asc_ResolveAllComments;
	asc_docs_api.prototype['asc_GetCommentLogicPosition']         = asc_docs_api.prototype.asc_GetCommentLogicPosition;
	asc_docs_api.prototype['asc_changeComment']                   = asc_docs_api.prototype.asc_changeComment;
	asc_docs_api.prototype['asc_selectComment']                   = asc_docs_api.prototype.asc_selectComment;
	asc_docs_api.prototype['asc_showComment']                     = asc_docs_api.prototype.asc_showComment;
	asc_docs_api.prototype['can_AddQuotedComment']                = asc_docs_api.prototype.can_AddQuotedComment;
	asc_docs_api.prototype['sync_RemoveComment']                  = asc_docs_api.prototype.sync_RemoveComment;
	asc_docs_api.prototype['sync_AddComment']                     = asc_docs_api.prototype.sync_AddComment;
	asc_docs_api.prototype['sync_ShowComment']                    = asc_docs_api.prototype.sync_ShowComment;
	asc_docs_api.prototype['sync_HideComment']                    = asc_docs_api.prototype.sync_HideComment;
	asc_docs_api.prototype['sync_UpdateCommentPosition']          = asc_docs_api.prototype.sync_UpdateCommentPosition;
	asc_docs_api.prototype['sync_ChangeCommentData']              = asc_docs_api.prototype.sync_ChangeCommentData;
	asc_docs_api.prototype['sync_LockComment']                    = asc_docs_api.prototype.sync_LockComment;
	asc_docs_api.prototype['sync_UnLockComment']                  = asc_docs_api.prototype.sync_UnLockComment;
	asc_docs_api.prototype['GenerateStyles']                      = asc_docs_api.prototype.GenerateStyles;
	asc_docs_api.prototype['asyncFontsDocumentEndLoaded']         = asc_docs_api.prototype.asyncFontsDocumentEndLoaded;
	asc_docs_api.prototype['asyncFontEndLoaded']                  = asc_docs_api.prototype.asyncFontEndLoaded;
	asc_docs_api.prototype['asyncImageEndLoaded']                 = asc_docs_api.prototype.asyncImageEndLoaded;
	asc_docs_api.prototype['get_PresentationWidth']               = asc_docs_api.prototype.get_PresentationWidth;
	asc_docs_api.prototype['get_PresentationHeight']              = asc_docs_api.prototype.get_PresentationHeight;
	asc_docs_api.prototype['pre_Paste']                           = asc_docs_api.prototype.pre_Paste;
	asc_docs_api.prototype['ViewScrollToX']                       = asc_docs_api.prototype.ViewScrollToX;
	asc_docs_api.prototype['ViewScrollToY']                       = asc_docs_api.prototype.ViewScrollToY;
	asc_docs_api.prototype['GetDocWidthPx']                       = asc_docs_api.prototype.GetDocWidthPx;
	asc_docs_api.prototype['GetDocHeightPx']                      = asc_docs_api.prototype.GetDocHeightPx;
	asc_docs_api.prototype['ClearSearch']                         = asc_docs_api.prototype.ClearSearch;
	asc_docs_api.prototype['GetCurrentVisiblePage']               = asc_docs_api.prototype.GetCurrentVisiblePage;
	asc_docs_api.prototype['asc_setAutoSaveGap']                  = asc_docs_api.prototype.asc_setAutoSaveGap;
	asc_docs_api.prototype['asc_SetDocumentPlaceChangedEnabled']  = asc_docs_api.prototype.asc_SetDocumentPlaceChangedEnabled;
	asc_docs_api.prototype['asc_SetViewRulers']                   = asc_docs_api.prototype.asc_SetViewRulers;
	asc_docs_api.prototype['asc_SetViewRulersChange']             = asc_docs_api.prototype.asc_SetViewRulersChange;
	asc_docs_api.prototype['asc_GetViewRulers']                   = asc_docs_api.prototype.asc_GetViewRulers;
	asc_docs_api.prototype['asc_SetDocumentUnits']                = asc_docs_api.prototype.asc_SetDocumentUnits;
	asc_docs_api.prototype['changeSlideSize']                     = asc_docs_api.prototype.changeSlideSize;
	asc_docs_api.prototype['asc_getFirstSlideNumber']             = asc_docs_api.prototype.asc_getFirstSlideNumber;
	asc_docs_api.prototype['AddSlide']                            = asc_docs_api.prototype.AddSlide;
	asc_docs_api.prototype['DeleteSlide']                         = asc_docs_api.prototype.DeleteSlide;
	asc_docs_api.prototype['DublicateSlide']                      = asc_docs_api.prototype.DublicateSlide;
	asc_docs_api.prototype['SelectAllSlides']                     = asc_docs_api.prototype.SelectAllSlides;
	asc_docs_api.prototype['asc_moveSelectedSlidesToEnd']         = asc_docs_api.prototype.asc_moveSelectedSlidesToEnd;
	asc_docs_api.prototype['asc_moveSlidesNextPos']               = asc_docs_api.prototype.asc_moveSlidesNextPos;
	asc_docs_api.prototype['asc_moveSelectedSlidesToStart']       = asc_docs_api.prototype.asc_moveSelectedSlidesToStart;
	asc_docs_api.prototype['asc_moveSlidesPrevPos']               = asc_docs_api.prototype.asc_moveSlidesPrevPos;
	asc_docs_api.prototype['asc_IsSlideSelected']                 = asc_docs_api.prototype.asc_IsSlideSelected;
	asc_docs_api.prototype['asc_IsFirstSlideSelected']            = asc_docs_api.prototype.asc_IsFirstSlideSelected;
	asc_docs_api.prototype['asc_IsLastSlideSelected']             = asc_docs_api.prototype.asc_IsLastSlideSelected;
	asc_docs_api.prototype['AddShape']                            = asc_docs_api.prototype.AddShape;
	asc_docs_api.prototype['ChangeShapeType']                     = asc_docs_api.prototype.ChangeShapeType;
	asc_docs_api.prototype['AddText']                             = asc_docs_api.prototype.AddText;
	asc_docs_api.prototype['groupShapes']                         = asc_docs_api.prototype.groupShapes;
	asc_docs_api.prototype['unGroupShapes']                       = asc_docs_api.prototype.unGroupShapes;
	asc_docs_api.prototype['asc_canMergeSelectedShapes']          = asc_docs_api.prototype.asc_canMergeSelectedShapes;
	asc_docs_api.prototype['setVerticalAlign']                    = asc_docs_api.prototype.setVerticalAlign;
	asc_docs_api.prototype['setVert']                             = asc_docs_api.prototype.setVert;
	asc_docs_api.prototype['sync_MouseMoveStartCallback']         = asc_docs_api.prototype.sync_MouseMoveStartCallback;
	asc_docs_api.prototype['sync_MouseMoveEndCallback']           = asc_docs_api.prototype.sync_MouseMoveEndCallback;
	asc_docs_api.prototype['sync_MouseMoveCallback']              = asc_docs_api.prototype.sync_MouseMoveCallback;
	asc_docs_api.prototype['ShowThumbnails']                      = asc_docs_api.prototype.ShowThumbnails;
	asc_docs_api.prototype['asc_ShowNotes']                       = asc_docs_api.prototype.asc_ShowNotes;
	asc_docs_api.prototype['asc_ShowAnimPane']                    = asc_docs_api.prototype.asc_ShowAnimPane;
	asc_docs_api.prototype['asc_ZoomOutTimeline']                 = asc_docs_api.prototype.asc_ZoomOutTimeline;
	asc_docs_api.prototype['asc_ZoomInTimeline']                  = asc_docs_api.prototype.asc_ZoomInTimeline;
	asc_docs_api.prototype['asc_SetSelectedAnimEffectsStartType'] = asc_docs_api.prototype.asc_SetSelectedAnimEffectsStartType;
	asc_docs_api.prototype['asc_RemoveSelectedAnimEffects']       = asc_docs_api.prototype.asc_RemoveSelectedAnimEffects;
	asc_docs_api.prototype['asc_DeleteVerticalScroll']            = asc_docs_api.prototype.asc_DeleteVerticalScroll;
	asc_docs_api.prototype['syncOnThumbnailsShow']                = asc_docs_api.prototype.syncOnThumbnailsShow;
	asc_docs_api.prototype['can_AddHyperlink']                    = asc_docs_api.prototype.can_AddHyperlink;
	asc_docs_api.prototype['add_Hyperlink']                       = asc_docs_api.prototype.add_Hyperlink;
	asc_docs_api.prototype['change_Hyperlink']                    = asc_docs_api.prototype.change_Hyperlink;
	asc_docs_api.prototype['remove_Hyperlink']                    = asc_docs_api.prototype.remove_Hyperlink;
	asc_docs_api.prototype['sync_HyperlinkPropCallback']          = asc_docs_api.prototype.sync_HyperlinkPropCallback;
	asc_docs_api.prototype['sync_HyperlinkClickCallback']         = asc_docs_api.prototype.sync_HyperlinkClickCallback;
	asc_docs_api.prototype['asc_GoToInternalHyperlink']           = asc_docs_api.prototype.asc_GoToInternalHyperlink;
	asc_docs_api.prototype['sync_CanAddHyperlinkCallback']        = asc_docs_api.prototype.sync_CanAddHyperlinkCallback;
	asc_docs_api.prototype['sync_DialogAddHyperlink']             = asc_docs_api.prototype.sync_DialogAddHyperlink;
    asc_docs_api.prototype['sync_SpellCheckCallback']             = asc_docs_api.prototype.sync_SpellCheckCallback;
    asc_docs_api.prototype['sync_SpellCheckVariantsFound']        = asc_docs_api.prototype.sync_SpellCheckVariantsFound;
    asc_docs_api.prototype['asc_replaceMisspelledWord']           = asc_docs_api.prototype.asc_replaceMisspelledWord;
    asc_docs_api.prototype['asc_ignoreMisspelledWord']            = asc_docs_api.prototype.asc_ignoreMisspelledWord;
    asc_docs_api.prototype['asc_spellCheckAddToDictionary']       = asc_docs_api.prototype.asc_spellCheckAddToDictionary;
    asc_docs_api.prototype['asc_spellCheckClearDictionary']       = asc_docs_api.prototype.asc_spellCheckClearDictionary;
    asc_docs_api.prototype['asc_setDefaultLanguage']              = asc_docs_api.prototype.asc_setDefaultLanguage;
    asc_docs_api.prototype['asc_getDefaultLanguage']              = asc_docs_api.prototype.asc_getDefaultLanguage;
    asc_docs_api.prototype['asc_getKeyboardLanguage']             = asc_docs_api.prototype.asc_getKeyboardLanguage;
    asc_docs_api.prototype['asc_getInputLanguage']                = asc_docs_api.prototype.asc_getInputLanguage;
    asc_docs_api.prototype['asc_setSpellCheck']                   = asc_docs_api.prototype.asc_setSpellCheck;
	asc_docs_api.prototype['asc_setSpellCheckSettings']           = asc_docs_api.prototype.asc_setSpellCheckSettings;
	asc_docs_api.prototype['asc_getSpellCheckSettings']           = asc_docs_api.prototype.asc_getSpellCheckSettings;
	asc_docs_api.prototype['sync_shapePropCallback']              = asc_docs_api.prototype.sync_shapePropCallback;
	asc_docs_api.prototype['sync_slidePropCallback']              = asc_docs_api.prototype.sync_slidePropCallback;
	asc_docs_api.prototype['GetCurrentPixOffsetY']                = asc_docs_api.prototype.GetCurrentPixOffsetY;
	asc_docs_api.prototype['SetPaintFormat']                      = asc_docs_api.prototype.SetPaintFormat;
	asc_docs_api.prototype['sync_PaintFormatCallback']            = asc_docs_api.prototype.sync_PaintFormatCallback;
	asc_docs_api.prototype['SetMarkerFormat']                     = asc_docs_api.prototype.SetMarkerFormat;
	asc_docs_api.prototype['sync_MarkerFormatCallback']           = asc_docs_api.prototype.sync_MarkerFormatCallback;
	asc_docs_api.prototype['ClearFormating']                      = asc_docs_api.prototype.ClearFormating;
	asc_docs_api.prototype['SetDeviceInputHelperId']              = asc_docs_api.prototype.SetDeviceInputHelperId;
	asc_docs_api.prototype['asc_setViewMode']                     = asc_docs_api.prototype.asc_setViewMode;
	asc_docs_api.prototype['sync_HyperlinkClickCallback']         = asc_docs_api.prototype.sync_HyperlinkClickCallback;
	asc_docs_api.prototype['UpdateInterfaceState']                = asc_docs_api.prototype.UpdateInterfaceState;
	asc_docs_api.prototype['OnMouseUp']                           = asc_docs_api.prototype.OnMouseUp;
	asc_docs_api.prototype['asyncImageEndLoaded2']                = asc_docs_api.prototype.asyncImageEndLoaded2;
	asc_docs_api.prototype['ChangeTheme']                         = asc_docs_api.prototype.ChangeTheme;
	asc_docs_api.prototype['StartLoadTheme']                      = asc_docs_api.prototype.StartLoadTheme;
	asc_docs_api.prototype['EndLoadTheme']                        = asc_docs_api.prototype.EndLoadTheme;
	asc_docs_api.prototype['ResetSlide']                          = asc_docs_api.prototype.ResetSlide;
	asc_docs_api.prototype['ChangeLayout']                        = asc_docs_api.prototype.ChangeLayout;
	asc_docs_api.prototype['put_ShapesAlign']                     = asc_docs_api.prototype.put_ShapesAlign;
	asc_docs_api.prototype['DistributeHorizontally']              = asc_docs_api.prototype.DistributeHorizontally;
	asc_docs_api.prototype['DistributeVertically']                = asc_docs_api.prototype.DistributeVertically;
	asc_docs_api.prototype['shapes_alignLeft']                    = asc_docs_api.prototype.shapes_alignLeft;
	asc_docs_api.prototype['shapes_alignRight']                   = asc_docs_api.prototype.shapes_alignRight;
	asc_docs_api.prototype['shapes_alignTop']                     = asc_docs_api.prototype.shapes_alignTop;
	asc_docs_api.prototype['shapes_alignBottom']                  = asc_docs_api.prototype.shapes_alignBottom;
	asc_docs_api.prototype['shapes_alignCenter']                  = asc_docs_api.prototype.shapes_alignCenter;
	asc_docs_api.prototype['shapes_alignMiddle']                  = asc_docs_api.prototype.shapes_alignMiddle;
	asc_docs_api.prototype['shapes_bringToFront']                 = asc_docs_api.prototype.shapes_bringToFront;
	asc_docs_api.prototype['shapes_bringForward']                 = asc_docs_api.prototype.shapes_bringForward;
	asc_docs_api.prototype['shapes_bringToBack']                  = asc_docs_api.prototype.shapes_bringToBack;
	asc_docs_api.prototype['shapes_bringBackward']                = asc_docs_api.prototype.shapes_bringBackward;
	asc_docs_api.prototype['sync_endDemonstration']               = asc_docs_api.prototype.sync_endDemonstration;
	asc_docs_api.prototype['sync_DemonstrationSlideChanged']      = asc_docs_api.prototype.sync_DemonstrationSlideChanged;
	asc_docs_api.prototype['StartDemonstration']                  = asc_docs_api.prototype.StartDemonstration;
	asc_docs_api.prototype['EndDemonstration']                    = asc_docs_api.prototype.EndDemonstration;
	asc_docs_api.prototype['DemonstrationPlay']                   = asc_docs_api.prototype.DemonstrationPlay;
	asc_docs_api.prototype['DemonstrationPause']                  = asc_docs_api.prototype.DemonstrationPause;
	asc_docs_api.prototype['DemonstrationEndShowMessage']         = asc_docs_api.prototype.DemonstrationEndShowMessage;
	asc_docs_api.prototype['DemonstrationNextSlide']              = asc_docs_api.prototype.DemonstrationNextSlide;
	asc_docs_api.prototype['DemonstrationPrevSlide']              = asc_docs_api.prototype.DemonstrationPrevSlide;
	asc_docs_api.prototype['DemonstrationGoToSlide']              = asc_docs_api.prototype.DemonstrationGoToSlide;
	asc_docs_api.prototype['sendFromReporter']              	  = asc_docs_api.prototype.sendFromReporter;
	asc_docs_api.prototype['SetDemonstrationModeOnly']            = asc_docs_api.prototype.SetDemonstrationModeOnly;
	asc_docs_api.prototype['ApplySlideTransition']                = asc_docs_api.prototype.ApplySlideTransition;
	asc_docs_api.prototype['SlideTransitionApplyToAll']           = asc_docs_api.prototype.SlideTransitionApplyToAll;
	asc_docs_api.prototype['SlideTransitionPlay']                 = asc_docs_api.prototype.SlideTransitionPlay;
	asc_docs_api.prototype['asc_HideSlides']                      = asc_docs_api.prototype.asc_HideSlides;
	asc_docs_api.prototype['SetTextBoxInputMode']                 = asc_docs_api.prototype.SetTextBoxInputMode;
	asc_docs_api.prototype['GetTextBoxInputMode']                 = asc_docs_api.prototype.GetTextBoxInputMode;
	asc_docs_api.prototype['sync_EndAddShape']                    = asc_docs_api.prototype.sync_EndAddShape;
	asc_docs_api.prototype['asc_getChartObject']                  = asc_docs_api.prototype.asc_getChartObject;
	asc_docs_api.prototype['asc_addChartDrawingObject']           = asc_docs_api.prototype.asc_addChartDrawingObject;
	asc_docs_api.prototype['asc_editChartDrawingObject']          = asc_docs_api.prototype.asc_editChartDrawingObject;
	asc_docs_api.prototype['asc_getChartPreviews']                = asc_docs_api.prototype.asc_getChartPreviews;
	asc_docs_api.prototype['asc_getTextArtPreviews']              = asc_docs_api.prototype.asc_getTextArtPreviews;
	asc_docs_api.prototype['sync_closeChartEditor']               = asc_docs_api.prototype.sync_closeChartEditor;
	asc_docs_api.prototype['sync_closeOleEditor']                 = asc_docs_api.prototype.sync_closeOleEditor;
	asc_docs_api.prototype['asc_getSelectedDrawingObjectsCount']  = asc_docs_api.prototype.asc_getSelectedDrawingObjectsCount;
	asc_docs_api.prototype['asc_ChangeTextCase']                  = asc_docs_api.prototype.asc_ChangeTextCase;
	asc_docs_api.prototype['asc_stopSaving']                      = asc_docs_api.prototype.asc_stopSaving;
	asc_docs_api.prototype['asc_continueSaving']                  = asc_docs_api.prototype.asc_continueSaving;
	asc_docs_api.prototype['asc_undoAllChanges']                  = asc_docs_api.prototype.asc_undoAllChanges;
	asc_docs_api.prototype['sync_ContextMenuCallback']            = asc_docs_api.prototype.sync_ContextMenuCallback;
	asc_docs_api.prototype['asc_SetFastCollaborative']            = asc_docs_api.prototype.asc_SetFastCollaborative;
	asc_docs_api.prototype['asc_isOffline']                       = asc_docs_api.prototype.asc_isOffline;
	asc_docs_api.prototype['asc_getUrlType']                      = asc_docs_api.prototype.asc_getUrlType;
	asc_docs_api.prototype['asc_prepareUrl']                      = asc_docs_api.prototype.asc_prepareUrl;
	asc_docs_api.prototype['asc_getSessionToken']                 = asc_docs_api.prototype.asc_getSessionToken;
	asc_docs_api.prototype["asc_setInterfaceDrawImagePlaceShape"] = asc_docs_api.prototype.asc_setInterfaceDrawImagePlaceShape;
	asc_docs_api.prototype["asc_nativeInitBuilder"]               = asc_docs_api.prototype.asc_nativeInitBuilder;
	asc_docs_api.prototype["asc_SetSilentMode"]                   = asc_docs_api.prototype.asc_SetSilentMode;
	asc_docs_api.prototype["asc_pluginsRegister"]                 = asc_docs_api.prototype.asc_pluginsRegister;
	asc_docs_api.prototype["asc_pluginRun"]                       = asc_docs_api.prototype.asc_pluginRun;
	asc_docs_api.prototype["asc_pluginStop"]                      = asc_docs_api.prototype.asc_pluginStop;
	asc_docs_api.prototype["asc_pluginResize"]                    = asc_docs_api.prototype.asc_pluginResize;
	asc_docs_api.prototype["asc_pluginButtonClick"]               = asc_docs_api.prototype.asc_pluginButtonClick;
	asc_docs_api.prototype["asc_pluginEnableMouseEvents"]         = asc_docs_api.prototype.asc_pluginEnableMouseEvents;

	asc_docs_api.prototype['asc_SetAutomaticBulletedLists']             = asc_docs_api.prototype.asc_SetAutomaticBulletedLists;
	asc_docs_api.prototype['asc_SetAutomaticNumberedLists']             = asc_docs_api.prototype.asc_SetAutomaticNumberedLists;
	asc_docs_api.prototype['asc_SetAutoCorrectSmartQuotes']             = asc_docs_api.prototype.asc_SetAutoCorrectSmartQuotes;
	asc_docs_api.prototype['asc_SetAutoCorrectHyphensWithDash']         = asc_docs_api.prototype.asc_SetAutoCorrectHyphensWithDash;
	asc_docs_api.prototype['asc_SetAutoCorrectFirstLetterOfSentences']  = asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfSentences;
	asc_docs_api.prototype['asc_SetAutoCorrectHyperlinks']              = asc_docs_api.prototype.asc_SetAutoCorrectHyperlinks;
	asc_docs_api.prototype['asc_SetAutoCorrectFirstLetterOfCells']      = asc_docs_api.prototype.asc_SetAutoCorrectFirstLetterOfCells;
	asc_docs_api.prototype['asc_SetAutoCorrectDoubleSpaceWithPeriod']   = asc_docs_api.prototype.asc_SetAutoCorrectDoubleSpaceWithPeriod;

	asc_docs_api.prototype['asc_SetFirstLetterAutoCorrectExceptions']   = asc_docs_api.prototype.asc_SetFirstLetterAutoCorrectExceptions;
	asc_docs_api.prototype['asc_GetFirstLetterAutoCorrectExceptions']   = asc_docs_api.prototype.asc_GetFirstLetterAutoCorrectExceptions;
	asc_docs_api.prototype['asc_GetAutoCorrectSettings']                = asc_docs_api.prototype.asc_GetAutoCorrectSettings;

	asc_docs_api.prototype["asc_GetSelectedText"]                 = asc_docs_api.prototype.asc_GetSelectedText;

	asc_docs_api.prototype["asc_addSlideNumber"]                  = asc_docs_api.prototype.asc_addSlideNumber;
	asc_docs_api.prototype["asc_addDateTime"]                     = asc_docs_api.prototype.asc_addDateTime;
	asc_docs_api.prototype["asc_setDefaultDateTimeFormat"]        = asc_docs_api.prototype.asc_setDefaultDateTimeFormat;
	asc_docs_api.prototype["asc_getHeaderFooterProperties"]       = asc_docs_api.prototype.asc_getHeaderFooterProperties;
	asc_docs_api.prototype["asc_setHeaderFooterProperties"]       = asc_docs_api.prototype.asc_setHeaderFooterProperties;

	asc_docs_api.prototype["asc_startEditCurrentOleObject"]       = asc_docs_api.prototype.asc_startEditCurrentOleObject;
	asc_docs_api.prototype["asc_doubleClickOnTableOleObject"]     = asc_docs_api.prototype.asc_doubleClickOnTableOleObject;
	asc_docs_api.prototype["asc_editOleObjectAction"]             = asc_docs_api.prototype.asc_editOleObjectAction;
	asc_docs_api.prototype["asc_addOleObjectAction"]              = asc_docs_api.prototype.asc_addOleObjectAction;
	asc_docs_api.prototype["asc_InputClearKeyboardElement"]       = asc_docs_api.prototype.asc_InputClearKeyboardElement;

	asc_docs_api.prototype["asc_getCurrentFocusObject"]           = asc_docs_api.prototype.asc_getCurrentFocusObject;
	asc_docs_api.prototype["asc_AddMath"]           			  = asc_docs_api.prototype.asc_AddMath;
	asc_docs_api.prototype["asc_ConvertMathView"]           	  = asc_docs_api.prototype.asc_ConvertMathView;
	asc_docs_api.prototype["asc_SetMathProps"]           		  = asc_docs_api.prototype.asc_SetMathProps;

    asc_docs_api.prototype['sendEvent']								= asc_docs_api.prototype.sendEvent;
	
	asc_docs_api.prototype['SetDrawImagePreviewBulletForMenu']		= asc_docs_api.prototype.SetDrawImagePreviewBulletForMenu;
	asc_docs_api.prototype['asc_EditSelectAll']		                = asc_docs_api.prototype.asc_EditSelectAll;

	// mobile
	asc_docs_api.prototype["asc_GetDefaultTableStyles"]           	= asc_docs_api.prototype.asc_GetDefaultTableStyles;
	asc_docs_api.prototype["asc_Remove"] 							= asc_docs_api.prototype.asc_Remove;
	asc_docs_api.prototype["AddShapeOnCurrentPage"] 				= asc_docs_api.prototype.AddShapeOnCurrentPage;
	asc_docs_api.prototype["can_CopyCut"] 							= asc_docs_api.prototype.can_CopyCut;

	asc_docs_api.prototype["asc_OnHideContextMenu"] 				= asc_docs_api.prototype.asc_OnHideContextMenu;
	asc_docs_api.prototype["asc_OnShowContextMenu"] 				= asc_docs_api.prototype.asc_OnShowContextMenu;

	asc_docs_api.prototype["DemonstrationReporterMessages"] 		= asc_docs_api.prototype.DemonstrationReporterMessages;
	asc_docs_api.prototype["DemonstrationToReporterMessages"] 		= asc_docs_api.prototype.DemonstrationToReporterMessages;
	asc_docs_api.prototype["preloadReporter"]						= asc_docs_api.prototype.preloadReporter;

	asc_docs_api.prototype["asc_SpecialPaste"]						= asc_docs_api.prototype.asc_SpecialPaste;

	// signatures
	asc_docs_api.prototype["asc_addSignatureLine"] 					= asc_docs_api.prototype.asc_addSignatureLine;
	asc_docs_api.prototype["asc_CallSignatureDblClickEvent"] 		= asc_docs_api.prototype.asc_CallSignatureDblClickEvent;
	asc_docs_api.prototype["asc_getRequestSignatures"] 				= asc_docs_api.prototype.asc_getRequestSignatures;
	asc_docs_api.prototype["asc_AddSignatureLine2"]             	= asc_docs_api.prototype.asc_AddSignatureLine2;
	asc_docs_api.prototype["asc_Sign"]             					= asc_docs_api.prototype.asc_Sign;
	asc_docs_api.prototype["asc_RequestSign"]             			= asc_docs_api.prototype.asc_RequestSign;
	asc_docs_api.prototype["asc_ViewCertificate"] 					= asc_docs_api.prototype.asc_ViewCertificate;
	asc_docs_api.prototype["asc_SelectCertificate"] 				= asc_docs_api.prototype.asc_SelectCertificate;
	asc_docs_api.prototype["asc_GetDefaultCertificate"] 			= asc_docs_api.prototype.asc_GetDefaultCertificate;
	asc_docs_api.prototype["asc_getSignatures"] 					= asc_docs_api.prototype.asc_getSignatures;
	asc_docs_api.prototype["asc_isSignaturesSupport"] 				= asc_docs_api.prototype.asc_isSignaturesSupport;
    asc_docs_api.prototype["asc_isProtectionSupport"] 				= asc_docs_api.prototype.asc_isProtectionSupport;
	asc_docs_api.prototype["asc_isAnonymousSupport"] 				= asc_docs_api.prototype.asc_isAnonymousSupport;
	asc_docs_api.prototype["asc_RemoveSignature"] 					= asc_docs_api.prototype.asc_RemoveSignature;
	asc_docs_api.prototype["asc_RemoveAllSignatures"] 				= asc_docs_api.prototype.asc_RemoveAllSignatures;
	asc_docs_api.prototype["asc_gotoSignature"] 					= asc_docs_api.prototype.asc_gotoSignature;
	asc_docs_api.prototype["asc_getSignatureSetup"] 				= asc_docs_api.prototype.asc_getSignatureSetup;

	asc_docs_api.prototype["sync_OnConvertEquationToMath"] 		    = asc_docs_api.prototype.sync_OnConvertEquationToMath;
	asc_docs_api.prototype["asc_ConvertEquationToMath"] 		    = asc_docs_api.prototype.asc_ConvertEquationToMath;
	asc_docs_api.prototype["asc_getTableStylesPreviews"] 		    = asc_docs_api.prototype.asc_getTableStylesPreviews;
	asc_docs_api.prototype["asc_generateTableStylesPreviews"] 		= asc_docs_api.prototype.asc_generateTableStylesPreviews;

	// print-preview
	asc_docs_api.prototype["asc_initPrintPreview"] 	= asc_docs_api.prototype.asc_initPrintPreview;
	asc_docs_api.prototype["asc_drawPrintPreview"] 	= asc_docs_api.prototype.asc_drawPrintPreview;
	asc_docs_api.prototype["asc_closePrintPreview"] = asc_docs_api.prototype.asc_closePrintPreview;


	asc_docs_api.prototype["asc_changePresentationViewMode"] = asc_docs_api.prototype.asc_changePresentationViewMode;
	asc_docs_api.prototype["asc_AddMasterSlide"] = asc_docs_api.prototype.asc_AddMasterSlide;
	asc_docs_api.prototype["asc_AddSlideLayout"] = asc_docs_api.prototype.asc_AddSlideLayout;
	asc_docs_api.prototype["asc_StartAddPlaceholder"] = asc_docs_api.prototype.asc_StartAddPlaceholder;
	asc_docs_api.prototype["asc_setLayoutTitle"] = asc_docs_api.prototype.asc_setLayoutTitle;
	asc_docs_api.prototype["asc_setLayoutFooter"] = asc_docs_api.prototype.asc_setLayoutFooter;
	asc_docs_api.prototype["asc_DeleteMaster"] = asc_docs_api.prototype.asc_DeleteMaster;
	asc_docs_api.prototype["asc_DeleteLayout"] = asc_docs_api.prototype.asc_DeleteLayout;
	asc_docs_api.prototype["asc_DuplicateMaster"] = asc_docs_api.prototype.asc_DuplicateMaster;
	asc_docs_api.prototype["asc_DuplicateLayout"] = asc_docs_api.prototype.asc_DuplicateLayout;
	asc_docs_api.prototype["asc_CanDeleteMaster"] = asc_docs_api.prototype.asc_CanDeleteMaster;
	asc_docs_api.prototype["asc_CanDeleteLayout"] = asc_docs_api.prototype.asc_CanDeleteLayout;
	asc_docs_api.prototype["asc_IsMasterMode"] = asc_docs_api.prototype.asc_IsMasterMode;

	asc_docs_api.prototype["asc_setViewerTargetType"] = asc_docs_api.prototype.asc_setViewerTargetType;

	asc_docs_api.prototype["asc_setDemoBackgroundColor"] = asc_docs_api.prototype.asc_setDemoBackgroundColor;
	asc_docs_api.prototype["asc_EraseAllInksOnSlide"] = asc_docs_api.prototype.asc_EraseAllInksOnSlide;


	window['Asc']['asc_CCommentData'] = window['Asc'].asc_CCommentData = asc_CCommentData;
	asc_CCommentData.prototype['asc_getText']         = asc_CCommentData.prototype.asc_getText;
	asc_CCommentData.prototype['asc_putText']         = asc_CCommentData.prototype.asc_putText;
	asc_CCommentData.prototype['asc_getTime']         = asc_CCommentData.prototype.asc_getTime;
	asc_CCommentData.prototype['asc_putTime']         = asc_CCommentData.prototype.asc_putTime;
	asc_CCommentData.prototype['asc_getOnlyOfficeTime']         = asc_CCommentData.prototype.asc_getOnlyOfficeTime;
	asc_CCommentData.prototype['asc_putOnlyOfficeTime']         = asc_CCommentData.prototype.asc_putOnlyOfficeTime;
	asc_CCommentData.prototype['asc_getUserId']       = asc_CCommentData.prototype.asc_getUserId;
	asc_CCommentData.prototype['asc_putUserId']       = asc_CCommentData.prototype.asc_putUserId;
	asc_CCommentData.prototype['asc_getUserName']     = asc_CCommentData.prototype.asc_getUserName;
	asc_CCommentData.prototype['asc_putUserName']     = asc_CCommentData.prototype.asc_putUserName;
	asc_CCommentData.prototype['asc_getGuid']         = asc_CCommentData.prototype.asc_getGuid;
	asc_CCommentData.prototype['asc_putGuid']         = asc_CCommentData.prototype.asc_putGuid;
	asc_CCommentData.prototype['asc_getTimeZoneBias'] = asc_CCommentData.prototype.asc_getTimeZoneBias;
	asc_CCommentData.prototype['asc_putTimeZoneBias'] = asc_CCommentData.prototype.asc_putTimeZoneBias;
	asc_CCommentData.prototype['asc_getQuoteText']    = asc_CCommentData.prototype.asc_getQuoteText;
	asc_CCommentData.prototype['asc_putQuoteText']    = asc_CCommentData.prototype.asc_putQuoteText;
	asc_CCommentData.prototype['asc_getSolved']       = asc_CCommentData.prototype.asc_getSolved;
	asc_CCommentData.prototype['asc_putSolved']       = asc_CCommentData.prototype.asc_putSolved;
	asc_CCommentData.prototype['asc_getReply']        = asc_CCommentData.prototype.asc_getReply;
	asc_CCommentData.prototype['asc_addReply']        = asc_CCommentData.prototype.asc_addReply;
	asc_CCommentData.prototype['asc_getRepliesCount'] = asc_CCommentData.prototype.asc_getRepliesCount;
	asc_CCommentData.prototype["asc_putDocumentFlag"] = asc_CCommentData.prototype.asc_putDocumentFlag;
	asc_CCommentData.prototype["asc_getDocumentFlag"] = asc_CCommentData.prototype.asc_getDocumentFlag;
	asc_CCommentData.prototype["asc_putUserData"]     = asc_CCommentData.prototype.asc_putUserData;
	asc_CCommentData.prototype["asc_getUserData"]     = asc_CCommentData.prototype.asc_getUserData;

	window['AscCommonSlide'].CContextMenuData         = CContextMenuData;
	CContextMenuData.prototype['get_Type']            = CContextMenuData.prototype.get_Type;
	CContextMenuData.prototype['get_X']               = CContextMenuData.prototype.get_X;
	CContextMenuData.prototype['get_Y']               = CContextMenuData.prototype.get_Y;
	CContextMenuData.prototype['get_IsSlideSelect']   = CContextMenuData.prototype.get_IsSlideSelect;
	CContextMenuData.prototype['get_IsSlideHidden']   = CContextMenuData.prototype.get_IsSlideHidden;
	CContextMenuData.prototype['get_Guide']           = CContextMenuData.prototype.get_Guide;
	CContextMenuData.prototype['get_EffectStartType'] = CContextMenuData.prototype.get_EffectStartType;
	CContextMenuData.prototype['get_ButtonWidth']     = CContextMenuData.prototype.get_ButtonWidth;
	CContextMenuData.prototype['get_ButtonHeight']    = CContextMenuData.prototype.get_ButtonHeight;

	window['Asc']['CAscSlideProps'] = window['Asc'].CAscSlideProps = CAscSlideProps;
	CAscSlideProps.prototype['get_background']        = CAscSlideProps.prototype.get_background;
	CAscSlideProps.prototype['put_background']        = CAscSlideProps.prototype.put_background;
	CAscSlideProps.prototype['get_LayoutName']        = CAscSlideProps.prototype.get_LayoutName;
	CAscSlideProps.prototype['put_LayoutName']        = CAscSlideProps.prototype.put_LayoutName;
	CAscSlideProps.prototype['get_MasterName']        = CAscSlideProps.prototype.get_MasterName;
	CAscSlideProps.prototype['put_MasterName']        = CAscSlideProps.prototype.put_MasterName;
	CAscSlideProps.prototype['get_LayoutIndex']        = CAscSlideProps.prototype.get_LayoutIndex;
	CAscSlideProps.prototype['put_LayoutIndex']        = CAscSlideProps.prototype.put_LayoutIndex;
	CAscSlideProps.prototype['get_transition']         = CAscSlideProps.prototype.get_transition;
	CAscSlideProps.prototype['put_transition']         = CAscSlideProps.prototype.put_transition;
	CAscSlideProps.prototype['get_LockDelete']        = CAscSlideProps.prototype.get_LockDelete;
	CAscSlideProps.prototype['put_LockDelete']        = CAscSlideProps.prototype.put_LockDelete;
	CAscSlideProps.prototype['get_LockLayout']        = CAscSlideProps.prototype.get_LockLayout;
	CAscSlideProps.prototype['put_LockLayout']        = CAscSlideProps.prototype.put_LockLayout;
	CAscSlideProps.prototype['get_LockTiming']        = CAscSlideProps.prototype.get_LockTiming;
	CAscSlideProps.prototype['put_LockTiming']        = CAscSlideProps.prototype.put_LockTiming;
	CAscSlideProps.prototype['get_LockBackground']    = CAscSlideProps.prototype.get_LockBackground;
	CAscSlideProps.prototype['put_LockBackground']    = CAscSlideProps.prototype.put_LockBackground;
	CAscSlideProps.prototype['get_LockTransition']    = CAscSlideProps.prototype.get_LockTransition;
	CAscSlideProps.prototype['put_LockTransition']    = CAscSlideProps.prototype.put_LockTransition;
	CAscSlideProps.prototype['get_LockRemove']        = CAscSlideProps.prototype.get_LockRemove;
	CAscSlideProps.prototype['put_LockRemove']        = CAscSlideProps.prototype.put_LockRemove;
	CAscSlideProps.prototype['get_IsHidden']          = CAscSlideProps.prototype.get_IsHidden;
	CAscSlideProps.prototype['get_LockApplyBackgroundToAll'] = CAscSlideProps.prototype.get_LockApplyBackgroundToAll;
	CAscSlideProps.prototype['put_LockApplyBackgroundToAll'] = CAscSlideProps.prototype.put_LockApplyBackgroundToAll;
	CAscSlideProps.prototype['get_LockResetBackground']      = CAscSlideProps.prototype.get_LockResetBackground;
	CAscSlideProps.prototype['put_LockResetBackground']      = CAscSlideProps.prototype.put_LockResetBackground;
	CAscSlideProps.prototype['get_ResetBackground']          = CAscSlideProps.prototype.get_ResetBackground;
	CAscSlideProps.prototype['put_ResetBackground']          = CAscSlideProps.prototype.put_ResetBackground;
	CAscSlideProps.prototype['get_ApplyBackgroundToAll']     = CAscSlideProps.prototype.get_ApplyBackgroundToAll;
	CAscSlideProps.prototype['put_ApplyBackgroundToAll']     = CAscSlideProps.prototype.put_ApplyBackgroundToAll;
	CAscSlideProps.prototype['get_ShowMasterSp']             = CAscSlideProps.prototype.get_ShowMasterSp;
	CAscSlideProps.prototype['put_ShowMasterSp']             = CAscSlideProps.prototype.put_ShowMasterSp;
	CAscSlideProps.prototype['get_IsMasterSelected']         = CAscSlideProps.prototype.get_IsMasterSelected;
	window['Asc']['CAscChartProp'] = window['Asc'].CAscChartProp = CAscChartProp;
	CAscChartProp.prototype['get_ChangeLevel']        = CAscChartProp.prototype.get_ChangeLevel;
	CAscChartProp.prototype['put_ChangeLevel']        = CAscChartProp.prototype.put_ChangeLevel;
	CAscChartProp.prototype['get_CanBeFlow']          = CAscChartProp.prototype.get_CanBeFlow;
	CAscChartProp.prototype['get_Width']              = CAscChartProp.prototype.get_Width;
	CAscChartProp.prototype['put_Width']              = CAscChartProp.prototype.put_Width;
	CAscChartProp.prototype['get_Height']             = CAscChartProp.prototype.get_Height;
	CAscChartProp.prototype['put_Height']             = CAscChartProp.prototype.put_Height;
	CAscChartProp.prototype['get_WrappingStyle']      = CAscChartProp.prototype.get_WrappingStyle;
	CAscChartProp.prototype['put_WrappingStyle']      = CAscChartProp.prototype.put_WrappingStyle;
	CAscChartProp.prototype['get_Paddings']           = CAscChartProp.prototype.get_Paddings;
	CAscChartProp.prototype['put_Paddings']           = CAscChartProp.prototype.put_Paddings;
	CAscChartProp.prototype['get_AllowOverlap']       = CAscChartProp.prototype.get_AllowOverlap;
	CAscChartProp.prototype['put_AllowOverlap']       = CAscChartProp.prototype.put_AllowOverlap;
	CAscChartProp.prototype['get_Position']           = CAscChartProp.prototype.get_Position;
	CAscChartProp.prototype['put_Position']           = CAscChartProp.prototype.put_Position;
	CAscChartProp.prototype['get_PositionH']          = CAscChartProp.prototype.get_PositionH;
	CAscChartProp.prototype['put_PositionH']          = CAscChartProp.prototype.put_PositionH;
	CAscChartProp.prototype['get_PositionV']          = CAscChartProp.prototype.get_PositionV;
	CAscChartProp.prototype['put_PositionV']          = CAscChartProp.prototype.put_PositionV;
	CAscChartProp.prototype['get_Value_X']            = CAscChartProp.prototype.get_Value_X;
	CAscChartProp.prototype['get_Value_Y']            = CAscChartProp.prototype.get_Value_Y;
	CAscChartProp.prototype['get_ImageUrl']           = CAscChartProp.prototype.get_ImageUrl;
	CAscChartProp.prototype['put_ImageUrl']           = CAscChartProp.prototype.put_ImageUrl;
	CAscChartProp.prototype['get_Group']              = CAscChartProp.prototype.get_Group;
	CAscChartProp.prototype['put_Group']              = CAscChartProp.prototype.put_Group;
	CAscChartProp.prototype['asc_getFromGroup']       = CAscChartProp.prototype.asc_getFromGroup;
	CAscChartProp.prototype['asc_putFromGroup']       = CAscChartProp.prototype.asc_putFromGroup;
	CAscChartProp.prototype['get_isChartProps']       = CAscChartProp.prototype.get_isChartProps;
	CAscChartProp.prototype['put_isChartPross']       = CAscChartProp.prototype.put_isChartPross;
	CAscChartProp.prototype['get_SeveralCharts']      = CAscChartProp.prototype.get_SeveralCharts;
	CAscChartProp.prototype['put_SeveralCharts']      = CAscChartProp.prototype.put_SeveralCharts;
	CAscChartProp.prototype['get_SeveralChartTypes']  = CAscChartProp.prototype.get_SeveralChartTypes;
	CAscChartProp.prototype['put_SeveralChartTypes']  = CAscChartProp.prototype.put_SeveralChartTypes;
	CAscChartProp.prototype['get_SeveralChartStyles'] = CAscChartProp.prototype.get_SeveralChartStyles;
	CAscChartProp.prototype['put_SeveralChartStyles'] = CAscChartProp.prototype.put_SeveralChartStyles;
	CAscChartProp.prototype['get_VerticalTextAlign']  = CAscChartProp.prototype.get_VerticalTextAlign;
	CAscChartProp.prototype['put_VerticalTextAlign']  = CAscChartProp.prototype.put_VerticalTextAlign;
	CAscChartProp.prototype['get_Locked']             = CAscChartProp.prototype.get_Locked;
	CAscChartProp.prototype['get_ChartProperties']    = CAscChartProp.prototype.get_ChartProperties;
	CAscChartProp.prototype['put_ChartProperties']    = CAscChartProp.prototype.put_ChartProperties;
	CAscChartProp.prototype['get_ShapeProperties']    = CAscChartProp.prototype.get_ShapeProperties;
	CAscChartProp.prototype['put_ShapeProperties']    = CAscChartProp.prototype.put_ShapeProperties;
	CAscChartProp.prototype['asc_getType']            = CAscChartProp.prototype.asc_getType;
	CAscChartProp.prototype['asc_getSubType']         = CAscChartProp.prototype.asc_getSubType;
	CAscChartProp.prototype['asc_getStyleId']         = CAscChartProp.prototype.asc_getStyleId;
	CAscChartProp.prototype['asc_getHeight']          = CAscChartProp.prototype.asc_getHeight;
	CAscChartProp.prototype['asc_getWidth']           = CAscChartProp.prototype.asc_getWidth;
	CAscChartProp.prototype['asc_setType']            = CAscChartProp.prototype.asc_setType;
	CAscChartProp.prototype['asc_setSubType']         = CAscChartProp.prototype.asc_setSubType;
	CAscChartProp.prototype['asc_setStyleId']         = CAscChartProp.prototype.asc_setStyleId;
	CAscChartProp.prototype['asc_setHeight']          = CAscChartProp.prototype.asc_setHeight;
	CAscChartProp.prototype['asc_setWidth']           = CAscChartProp.prototype.asc_setWidth;
	CAscChartProp.prototype['asc_putTitle']           = CAscChartProp.prototype['put_Title']           = CAscChartProp.prototype['asc_setTitle']           = CAscChartProp.prototype.asc_setTitle;
	CAscChartProp.prototype['asc_putDescription']     = CAscChartProp.prototype['put_Description']     = CAscChartProp.prototype['asc_setDescription']     = CAscChartProp.prototype.asc_setDescription;
	CAscChartProp.prototype['asc_getTitle']           = CAscChartProp.prototype.asc_getTitle;
	CAscChartProp.prototype['asc_getDescription']     = CAscChartProp.prototype.asc_getDescription;
	CAscChartProp.prototype['asc_getPosition']     = CAscChartProp.prototype.asc_getPosition;
	CAscChartProp.prototype['asc_putPosition']     = CAscChartProp.prototype.asc_putPosition;
	CAscChartProp.prototype['asc_getName']     = CAscChartProp.prototype.asc_getName;
	CAscChartProp.prototype['asc_putName']     = CAscChartProp.prototype.asc_putName;
	CAscChartProp.prototype['getType']                = CAscChartProp.prototype.getType;
	CAscChartProp.prototype['putType']                = CAscChartProp.prototype.putType;
	CAscChartProp.prototype['getStyle']               = CAscChartProp.prototype.getStyle;
	CAscChartProp.prototype['putStyle']               = CAscChartProp.prototype.putStyle;
	CAscChartProp.prototype['putLockAspect']          = CAscChartProp.prototype['asc_putLockAspect'] = CAscChartProp.prototype.putLockAspect;
	CAscChartProp.prototype['getLockAspect'] = CAscChartProp.prototype['asc_getLockAspect'] = CAscChartProp.prototype.getLockAspect;
	CAscChartProp.prototype['changeType']        = CAscChartProp.prototype.changeType;
	CSearchResult.prototype['get_Text']          = CSearchResult.prototype.get_Text;
	CSearchResult.prototype['get_Navigator']     = CSearchResult.prototype.get_Navigator;
	CSearchResult.prototype['put_Navigator']     = CSearchResult.prototype.put_Navigator;
	CSearchResult.prototype['put_Text']          = CSearchResult.prototype.put_Text;
	
	AscCommon['SlideEditorApi'] = AscCommon.SlideEditorApi = asc_docs_api;
})(window, window.document);
