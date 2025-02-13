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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
	function(window, undefined)
{
	var c_oAscError = {
		Level : {
			Critical   : -1,
			NoCritical : 0
		},
		ID    : {
			ServerSaveComplete   : 3,
			ConvertationProgress : 2,
			DownloadProgress     : 1,
			No                   : 0,
			Unknown              : -1,
			ConvertationTimeout  : -2,

			DownloadError        : -4,
			UnexpectedGuid       : -5,
			Database             : -6,
			FileRequest          : -7,
			FileVKey             : -8,
			UplImageSize         : -9,
			UplImageExt          : -10,
			UplImageFileCount    : -11,
			NoSupportClipdoard   : -12,
			UplImageUrl          : -13,
			DirectUrl            : -14,


			MaxDataPointsError    : -16,
			StockChartError       : -17,
			CoAuthoringDisconnect : -18,
			ConvertationPassword  : -19,
			VKeyEncrypt           : -20,
			KeyExpire             : -21,
			UserCountExceed       : -22,
			AccessDeny            : -23,
			LoadingScriptError    : -24,
			EditingError          :	-25,
			LoadingFontError      : -26,
			LoadingBinError       : -27,

			SplitCellMaxRows     : -30,
			SplitCellMaxCols     : -31,
			SplitCellRowsDivider : -32,

			MobileUnexpectedCharCount : -35,

			// Mail Merge
			MailMergeLoadFile : -40,
			MailMergeSaveFile : -41,

			// Data Validate
			DataValidate : -45,
			MoreOneTypeDataValidate: -46,
			ContainsCellsWithoutDataValidate: -47,

			// for AutoFilter
			AutoFilterDataRangeError         : -50,
			AutoFilterChangeFormatTableError : -51,
			AutoFilterChangeError            : -52,
			AutoFilterMoveToHiddenRangeError : -53,
			LockedAllError                   : -54,
			LockedWorksheetRename            : -55,
			FTChangeTableRangeError          : -56,
			FTRangeIncludedOtherTables       : -57,
			ChangeFilteredRangeError         : -58,

			CanNotPasteImage: -63,
			PasteMaxRangeError   : -64,
			PastInMergeAreaError : -65,
			CopyMultiselectAreaError : -66,
			PasteSlicerError: 67,
			MoveSlicerError: 68,
			PasteMultiSelectError : -69,

			NoValues         : -70,
			NoSingleRowCol   : -71,
			InvalidReference : -72,
			ErrorInFormula   : -73,
			CannotMoveRange  : -74,
			DataRangeError   : -75,

			MustSingleCell          : -76,
			MustContainFormula      : -77,
			MustFormulaResultNumber : -78,
			MustContainValue        : -79,

			MaxDataSeriesError : -80,
			CannotFillRange    : -81,

			ConvertationOpenError      : -82,
            ConvertationSaveError      : -83,
			ConvertationOpenLimitError : -84,
			ConvertationOpenFormat     : -85,

			CannotSaveWatermark     : -86,

			UserDrop : -100,
			Warning  : -101,
			UpdateVersion : -102,

			PrintMaxPagesCount : -110,

			SessionAbsolute: -120,
			SessionIdle: -121,
			SessionToken: -122,

			LockedCellGoalSeek          : -200,

			/* для формул */
			FrmlMaxReference            : -297,
			FrmlMaxLength               : -298,
			FrmlMaxTextLength           : -299,
			FrmlWrongCountParentheses   : -300,
			FrmlWrongOperator           : -301,
			FrmlWrongMaxArgument        : -302,
			FrmlWrongCountArgument      : -303,
			FrmlWrongFunctionName       : -304,
			FrmlAnotherParsingError     : -305,
			FrmlWrongArgumentRange      : -306,
			FrmlOperandExpected         : -307,
			FrmlParenthesesCorrectCount : -308,
			FrmlWrongReferences         : -309,

			InvalidReferenceOrName : -310,
			LockCreateDefName      : -311,

			LockedCellPivot				: -312,
			PivotLabledColumns			: -313,
			PivotOverlap				: -314,
			PivotGroup					: -315,
			PivotWithoutUnderlyingData	: -316,
			PivotFieldNameExists        : -317,
			FormulaInPivotFieldName     : -318,
			PasteInPivot                : -319,

			CircularReference			: -320,
			PivotItemNameNotFound       : -321,
			CalculatedItemInPageField   : -322,
			NotUniqueFieldWithCalculated: -323,
			WrongDataFieldSubtotalForCalculatedItems: -324,
			PivotFieldCustomSubtotalsWithCalculatedItems: -325,

			ForceSaveButton: -331,
			ForceSaveTimeout: -332,
			Submit: -333,

			OpenWarning : 500,

            DataEncrypted : -600,

			CannotChangeFormulaArray: -450,
			MultiCellsInTablesFormulaArray: -451,

			MailToClientMissing	: -452,

			NoDataToParse : -601,

			CannotCompareInCoEditing : 651,

			CannotUngroupError : -700,

			UplDocumentSize         : -751,
			UplDocumentExt          : -752,
			UplDocumentFileCount    : -753,

			CustomSortMoreOneSelectedError: -800,
			CustomSortNotOriginalSelectError: -801,

			// Data Validate
			DataValidateNotNumeric: -830,
			DataValidateNegativeTextLength: -831,
			DataValidateMustEnterValue: -832,
			DataValidateMinGreaterMax: 833,
			DataValidateInvalid: 834,
			NamedRangeNotFound: 835,
			FormulaEvaluateError: 836,
			DataValidateInvalidList: 837,


			RemoveDuplicates : -850,

			LargeRangeWarning: -900,

			LockedEditView: -950,

			Password : -1000,

			ComplexFieldEmptyTOC : -1101,
			ComplexFieldNoTOC    : -1102,

			TextFormWrongFormat : -1201,

			SecondaryAxis: 1001,
			ComboSeriesError: 1002,

			//conditional formatting
			NotValidPercentile : 1003,
			CannotAddConditionalFormatting: 1004,
			NotValidPercentage: 1005,
			NotSingleReferenceCannotUsed: 1006,
			CannotUseRelativeReference: 1007,
			ValueMustBeGreaterThen: 1008,
			IconDataRangesOverlap: 1009,
			ErrorTop10Between: 1010,

			SingleColumnOrRowError: 1020,
			LocationOrDataRangeError: 1021,

			ChangeOnProtectedSheet: 1030,
			PasswordIsNotCorrect: 1031,
			DeleteColumnContainsLockedCell: 1032,
			DeleteRowContainsLockedCell: 1033,
			CannotUseCommandProtectedSheet: 1034,

			FillAllRowsWarning: 1040,

			ProtectedRangeByOtherUser: 1050,

			TraceDependentsNoFormulas: 1060,
			TracePrecedentsNoValidReference: 1061,

			MustIntegerOrDecimalNumber: 1070,

			DocumentAndChangeMismatch: 1080,
			EditProtectedRange       : 1081,


			// pdf forms
			PDFFormsLocked: -1200
		}
	};

	var prot;
	window['Asc'] = window['Asc'] || {};
	window['Asc']['c_oAscError'] = window['Asc'].c_oAscError = c_oAscError;
	prot                                     = c_oAscError;
	prot['Level']                            = prot.Level;
	prot['ID']                               = prot.ID;
	prot                                     = c_oAscError.Level;
	prot['Critical']                         = prot.Critical;
	prot['NoCritical']                       = prot.NoCritical;
	prot                                     = c_oAscError.ID;
	prot['ServerSaveComplete']               = prot.ServerSaveComplete;
	prot['ConvertationProgress']             = prot.ConvertationProgress;
	prot['DownloadProgress']                 = prot.DownloadProgress;
	prot['No']                               = prot.No;
	prot['Unknown']                          = prot.Unknown;
	prot['ConvertationTimeout']              = prot.ConvertationTimeout;
	prot['DownloadError']                    = prot.DownloadError;
	prot['UnexpectedGuid']                   = prot.UnexpectedGuid;
	prot['Database']                         = prot.Database;
	prot['FileRequest']                      = prot.FileRequest;
	prot['FileVKey']                         = prot.FileVKey;
	prot['UplImageSize']                     = prot.UplImageSize;
	prot['UplImageExt']                      = prot.UplImageExt;
	prot['UplImageFileCount']                = prot.UplImageFileCount;
	prot['NoSupportClipdoard']               = prot.NoSupportClipdoard;
	prot['UplImageUrl']                      = prot.UplImageUrl;
	prot['DirectUrl']                        = prot.DirectUrl;
	prot['MaxDataPointsError']               = prot.MaxDataPointsError;
	prot['StockChartError']                  = prot.StockChartError;
	prot['CoAuthoringDisconnect']            = prot.CoAuthoringDisconnect;
	prot['ConvertationPassword']             = prot.ConvertationPassword;
	prot['VKeyEncrypt']                      = prot.VKeyEncrypt;
	prot['KeyExpire']                        = prot.KeyExpire;
	prot['UserCountExceed']                  = prot.UserCountExceed;
	prot['AccessDeny']                       = prot.AccessDeny;
	prot['LoadingScriptError']               = prot.LoadingScriptError;
	prot['EditingError']                     = prot.EditingError;
	prot['LoadingFontError']                 = prot.LoadingFontError;
	prot['LoadingBinError']                  = prot.LoadingBinError;
	prot['SplitCellMaxRows']                 = prot.SplitCellMaxRows;
	prot['SplitCellMaxCols']                 = prot.SplitCellMaxCols;
	prot['SplitCellRowsDivider']             = prot.SplitCellRowsDivider;
	prot['MobileUnexpectedCharCount']        = prot.MobileUnexpectedCharCount;
	prot['MailMergeLoadFile']                = prot.MailMergeLoadFile;
	prot['MailMergeSaveFile']                = prot.MailMergeSaveFile;
	prot['DataValidate']                     = prot.DataValidate;
	prot['MoreOneTypeDataValidate']          = prot.MoreOneTypeDataValidate;
	prot['ContainsCellsWithoutDataValidate'] = prot.ContainsCellsWithoutDataValidate;
	prot['AutoFilterDataRangeError']         = prot.AutoFilterDataRangeError;
	prot['AutoFilterChangeFormatTableError'] = prot.AutoFilterChangeFormatTableError;
	prot['AutoFilterChangeError']            = prot.AutoFilterChangeError;
	prot['AutoFilterMoveToHiddenRangeError'] = prot.AutoFilterMoveToHiddenRangeError;
	prot['LockedAllError']                   = prot.LockedAllError;
	prot['LockedWorksheetRename']            = prot.LockedWorksheetRename;
	prot['FTChangeTableRangeError']          = prot.FTChangeTableRangeError;
	prot['FTRangeIncludedOtherTables']       = prot.FTRangeIncludedOtherTables;
	prot['ChangeFilteredRangeError']         = prot.ChangeFilteredRangeError;
	prot['PasteMaxRangeError']               = prot.PasteMaxRangeError;
	prot['PastInMergeAreaError']             = prot.PastInMergeAreaError;
	prot['CopyMultiselectAreaError']         = prot.CopyMultiselectAreaError;
	prot['PasteSlicerError']                 = prot.PasteSlicerError;
	prot['MoveSlicerError']                  = prot.MoveSlicerError;
	prot['PasteMultiSelectError']            = prot.PasteMultiSelectError;
	prot['CanNotPasteImage']                 = prot.CanNotPasteImage;
	prot['DataRangeError']                   = prot.DataRangeError;
	prot['MustSingleCell']                   = prot.MustSingleCell;
	prot['MustContainFormula']               = prot.MustContainFormula;
	prot['MustFormulaResultNumber']          = prot.MustFormulaResultNumber;
	prot['MustContainValue']                 = prot.MustContainValue;
	prot['LockedCellGoalSeek']               = prot.LockedCellGoalSeek;
	prot['NoValues']                         = prot.NoValues;
	prot['NoSingleRowCol']                   = prot.NoSingleRowCol;
	prot['InvalidReference']                 = prot.InvalidReference;
	prot['ErrorInFormula']                   = prot.ErrorInFormula;
	prot['CannotMoveRange']                  = prot.CannotMoveRange;
	prot['MaxDataSeriesError']               = prot.MaxDataSeriesError;
	prot['CannotFillRange']                  = prot.CannotFillRange;
	prot['ConvertationOpenError']            = prot.ConvertationOpenError;
	prot['ConvertationSaveError']            = prot.ConvertationSaveError;
	prot['ConvertationOpenLimitError']       = prot.ConvertationOpenLimitError;
	prot['ConvertationOpenFormat']       	 = prot.ConvertationOpenFormat;
	prot['CannotSaveWatermark']       	     = prot.CannotSaveWatermark;
	prot['UserDrop']                         = prot.UserDrop;
	prot['Warning']                          = prot.Warning;
	prot['UpdateVersion']                    = prot.UpdateVersion;
	prot['PrintMaxPagesCount']               = prot.PrintMaxPagesCount;
	prot['SessionAbsolute']                  = prot.SessionAbsolute;
	prot['SessionIdle']                      = prot.SessionIdle;
	prot['SessionToken']                     = prot.SessionToken;
	prot['FrmlMaxTextLength']                = prot.FrmlMaxTextLength;
	prot['FrmlMaxLength']                    = prot.FrmlMaxLength;
	prot['FrmlMaxReference']                 = prot.FrmlMaxReference;
	prot['FrmlWrongCountParentheses']        = prot.FrmlWrongCountParentheses;
	prot['FrmlWrongOperator']                = prot.FrmlWrongOperator;
	prot['FrmlWrongMaxArgument']             = prot.FrmlWrongMaxArgument;
	prot['FrmlWrongCountArgument']           = prot.FrmlWrongCountArgument;
	prot['FrmlWrongFunctionName']            = prot.FrmlWrongFunctionName;
	prot['FrmlAnotherParsingError']          = prot.FrmlAnotherParsingError;
	prot['FrmlWrongArgumentRange']           = prot.FrmlWrongArgumentRange;
	prot['FrmlOperandExpected']              = prot.FrmlOperandExpected;
	prot['FrmlParenthesesCorrectCount']      = prot.FrmlParenthesesCorrectCount;
	prot['FrmlWrongReferences']              = prot.FrmlWrongReferences;
	prot['InvalidReferenceOrName']           = prot.InvalidReferenceOrName;
	prot['LockCreateDefName']                = prot.LockCreateDefName;
	prot['LockedCellPivot']                  = prot.LockedCellPivot;
	prot['PivotFieldNameExists']             = prot.PivotFieldNameExists;
	prot['FormulaInPivotFieldName']          = prot.FormulaInPivotFieldName;
	prot['PivotItemNameNotFound']            = prot.PivotItemNameNotFound;
	prot['CalculatedItemInPageField']        = prot.CalculatedItemInPageField;
	prot['NotUniqueFieldWithCalculated']     = prot.NotUniqueFieldWithCalculated
	prot['WrongDataFieldSubtotalForCalculatedItems'] = prot.WrongDataFieldSubtotalForCalculatedItems;
	prot['PivotFieldCustomSubtotalsWithCalculatedItems'] = prot.PivotFieldCustomSubtotalsWithCalculatedItems;
	prot['PasteInPivot']                     = prot.PasteInPivot;
	prot['PivotLabledColumns']               = prot.PivotLabledColumns;
	prot['PivotOverlap']                     = prot.PivotOverlap;
	prot['PivotGroup']                       = prot.PivotGroup;
	prot['PivotWithoutUnderlyingData']       = prot.PivotWithoutUnderlyingData;
	prot['CircularReference']                = prot.CircularReference;
	prot['ForceSaveButton']                  = prot.ForceSaveButton;
	prot['ForceSaveTimeout']                 = prot.ForceSaveTimeout;
	prot['Submit']                           = prot.Submit;
	prot['CannotChangeFormulaArray']         = prot.CannotChangeFormulaArray;
	prot['MultiCellsInTablesFormulaArray']   = prot.MultiCellsInTablesFormulaArray;
	prot['MailToClientMissing']				 = prot.MailToClientMissing;
	prot['OpenWarning']                      = prot.OpenWarning;
	prot['DataEncrypted']                    = prot.DataEncrypted;
	prot['NoDataToParse']                    = prot.NoDataToParse;
	prot['CannotCompareInCoEditing']         = prot.CannotCompareInCoEditing;
	prot['CannotUngroupError']               = prot.CannotUngroupError;
	prot['UplDocumentSize']                  = prot.UplDocumentSize;
	prot['UplDocumentExt']                   = prot.UplDocumentExt;
	prot['UplDocumentFileCount']             = prot.UplDocumentFileCount;
	prot['CustomSortMoreOneSelectedError']   = prot.CustomSortMoreOneSelectedError;
	prot['CustomSortNotOriginalSelectError'] = prot.CustomSortNotOriginalSelectError;
	prot['RemoveDuplicates']                 = prot.RemoveDuplicates;
	prot['LargeRangeWarning']                = prot.LargeRangeWarning;
	prot['LockedEditView']                   = prot.LockedEditView;
	prot['Password']                         = prot.Password;
	prot['ComplexFieldEmptyTOC']             = prot.ComplexFieldEmptyTOC;
	prot['ComplexFieldNoTOC']                = prot.ComplexFieldNoTOC;
	prot['TextFormWrongFormat']              = prot.TextFormWrongFormat;
	prot['SecondaryAxis']                    = prot.SecondaryAxis;
	prot['ComboSeriesError']                 = prot.ComboSeriesError;

	prot['DataValidateNotNumeric']           = prot.DataValidateNotNumeric;
	prot['DataValidateNegativeTextLength']   = prot.DataValidateNegativeTextLength;
	prot['DataValidateMustEnterValue']       = prot.DataValidateMustEnterValue;
	prot['DataValidateMinGreaterMax']        = prot.DataValidateMinGreaterMax;
	prot['DataValidateInvalid']              = prot.DataValidateInvalid;
	prot['NamedRangeNotFound']               = prot.NamedRangeNotFound;
	prot['FormulaEvaluateError']             = prot.FormulaEvaluateError;
	prot['DataValidateInvalidList']          = prot.DataValidateInvalidList;

	prot['NotValidPercentile']               = prot.NotValidPercentile;
	prot['CannotAddConditionalFormatting']   = prot.CannotAddConditionalFormatting;
	prot['NotValidPercentage']               = prot.NotValidPercentage;
	prot['NotSingleReferenceCannotUsed']     = prot.NotSingleReferenceCannotUsed;
	prot['CannotUseRelativeReference']       = prot.CannotUseRelativeReference;
	prot['ValueMustBeGreaterThen']           = prot.ValueMustBeGreaterThen;
	prot['IconDataRangesOverlap']            = prot.IconDataRangesOverlap;
	prot['ErrorTop10Between']                = prot.ErrorTop10Between;
	prot['SingleColumnOrRowError']           = prot.SingleColumnOrRowError;
	prot['LocationOrDataRangeError']         = prot.LocationOrDataRangeError;
	prot['ChangeOnProtectedSheet']           = prot.ChangeOnProtectedSheet;
	prot['PasswordIsNotCorrect']             = prot.PasswordIsNotCorrect;
	prot['DeleteColumnContainsLockedCell']   = prot.DeleteColumnContainsLockedCell;
	prot['DeleteRowContainsLockedCell']      = prot.DeleteRowContainsLockedCell;
	prot['FillAllRowsWarning']               = prot.FillAllRowsWarning;
	prot['CannotUseCommandProtectedSheet']   = prot.CannotUseCommandProtectedSheet;
	prot['ProtectedRangeByOtherUser']        = prot.ProtectedRangeByOtherUser;
	prot['TraceDependentsNoFormulas']        = prot.TraceDependentsNoFormulas;
	prot['TracePrecedentsNoValidReference']  = prot.TracePrecedentsNoValidReference;
	prot['MustIntegerOrDecimalNumber']       = prot.MustIntegerOrDecimalNumber;
	prot['DocumentAndChangeMismatch']        = prot.DocumentAndChangeMismatch;
	prot['EditProtectedRange']               = prot.EditProtectedRange;
	prot['PDFFormsLocked']                   = prot.PDFFormsLocked;



})(window);
