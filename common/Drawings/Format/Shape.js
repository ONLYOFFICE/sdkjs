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

(
	/**
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function (window, undefined) {

// Import

		var c_oAscSizeRelFromH = AscCommon.c_oAscSizeRelFromH;
		var c_oAscSizeRelFromV = AscCommon.c_oAscSizeRelFromV;
		var CMatrix = AscCommon.CMatrix;
		var isRealObject = AscCommon.isRealObject;
		var global_mouseEvent = AscCommon.global_mouseEvent;
		var History = AscCommon.History;
		var global_MatrixTransformer = AscCommon.global_MatrixTransformer;

		var checkNormalRotate = AscFormat.checkNormalRotate;
		var HitInLine = AscFormat.HitInLine;
		var MOVE_DELTA = AscFormat.MOVE_DELTA;


		var pHText = [];

		pHText[AscFormat.phType_body] = "Slide text";             //"Текст слайда" ;                              ;
		pHText[AscFormat.phType_chart] = "Chart";         // "Диаграмма" ;                                     ;
		pHText[AscFormat.phType_clipArt] = "Clip Art";// "Текст слайда" ; //(Clip Art)                   ;
		pHText[AscFormat.phType_ctrTitle] = "Slide title";// "Заголовок слайда" ; //(Centered Title)     ;
		pHText[AscFormat.phType_dgm] = "Diagram";// "Диаграмма";// (Diagram)                        ;
		pHText[AscFormat.phType_dt] = "Date and time";// "Дата и время";// (Date and Time)         ;
		pHText[AscFormat.phType_ftr] = "Footer";// "Нижний колонтитул";// (Footer)                  ;
		pHText[AscFormat.phType_hdr] = "Header";// "Верхний колонтитул"; //(Header)                 ;
		pHText[AscFormat.phType_media] = "Media";// "Текст слайда"; //(Media)                         ;
		pHText[AscFormat.phType_obj] = "Slide text";// "Текст слайда"; //(Object)                   ;
		pHText[AscFormat.phType_pic] = "Picture";// "Вставка рисунка"; //(Picture)                  ;
		pHText[AscFormat.phType_sldImg] = "Image";// "Вставка рисунка"; //(Slide Image)                ;
		pHText[AscFormat.phType_sldNum] = "Slide number";// "Номер слайда"; //(Slide Number)           ;
		pHText[AscFormat.phType_subTitle] = "Slide subtitle";// "Подзаголовок слайда"; //(Subtitle)      ;
		pHText[AscFormat.phType_tbl] = "Table";// "Таблица"; //(Table)                              ;
		pHText[AscFormat.phType_title] = "Slide title";// "Заголовок слайда" ;  //(Title)             ;

		var c_oAscFill = Asc.c_oAscFill;

		var dTextFitDelta = 3;// mm

		function CheckObjectLine(obj) {
			return (obj instanceof CShape && obj.spPr && obj.spPr.geometry && AscFormat.CheckLinePresetForParagraphAdd(obj.spPr.geometry.preset));
		}


		function CheckWordArtTextPr(oRun) {
			if (oRun instanceof AscCommonWord.ParaRun) {
				var oTextPr = oRun.Get_CompiledPr();
				if (oTextPr.TextFill || (oTextPr.TextOutline && oTextPr.TextOutline.isVisible()) ||
					(oTextPr.Unifill && oTextPr.Unifill.fill && (oTextPr.Unifill.fill.type !== c_oAscFill.FILL_TYPE_SOLID || oTextPr.Unifill.transparent != null && oTextPr.Unifill.transparent < 254.5)))
					return true;
			}
			return false;
		}


		function hitInRect(x, y, l, t, r, b) {
			return x >= l && x <= r && y >= t && y <= b;
		}

		function hitToCropHandles(x, y, object) {
			var invert_transform = object.getInvertTransform();
			if (!invert_transform) {
				return -1;
			}
			var t_x, t_y;
			t_x = invert_transform.TransformPointX(x, y);
			t_y = invert_transform.TransformPointY(x, y);
			var fCoeff = object.convertPixToMM(1);
			var fCoeff2 = 1 / fCoeff;


			var widthCorner = (object.extX * fCoeff2 + 1) >> 1;
			var isCentralMarkerX = widthCorner > 40;
			if (widthCorner > 17)
				widthCorner = 17;
			var heightCorner = (object.extY * fCoeff2 + 1) >> 1;
			var isCentralMarkerY = heightCorner > 40;
			if (heightCorner > 17)
				heightCorner = 17;

			widthCorner *= fCoeff;
			heightCorner *= fCoeff;
			var markerWidth = 5 * fCoeff;

			if (hitInRect(t_x, t_y, 0, 0, widthCorner, markerWidth)) {
				return 0;
			}
			if (hitInRect(t_x, t_y, 0, 0, markerWidth, heightCorner)) {
				return 0;
			}

			if (isCentralMarkerX) {
				if (hitInRect(t_x, t_y, object.extX / 2 - widthCorner / 2, 0, object.extX / 2 + widthCorner / 2, markerWidth)) {
					return 1;
				}
				if (hitInRect(t_x, t_y, object.extX / 2 - widthCorner / 2, object.extY - markerWidth, object.extX / 2 + widthCorner / 2, object.extY)) {
					return 5;
				}
			}

			if (hitInRect(t_x, t_y, object.extX - widthCorner, 0, object.extX, markerWidth)) {
				return 2;
			}
			if (hitInRect(t_x, t_y, object.extX - markerWidth, 0, object.extX, heightCorner)) {
				return 2;
			}

			if (isCentralMarkerY) {
				if (hitInRect(t_x, t_y, object.extX - markerWidth, object.extY / 2 - heightCorner / 2, object.extX, object.extY / 2 + heightCorner / 2)) {
					return 3;
				}
				if (hitInRect(t_x, t_y, 0, object.extY / 2 - heightCorner / 2, markerWidth, object.extY / 2 + heightCorner / 2)) {
					return 7;
				}
			}

			if (hitInRect(t_x, t_y, object.extX - markerWidth, object.extY - heightCorner, object.extX, object.extY)) {
				return 4;
			}
			if (hitInRect(t_x, t_y, object.extX - widthCorner, object.extY - markerWidth, object.extX, object.extY)) {
				return 4;
			}

			if (hitInRect(t_x, t_y, 0, object.extY - heightCorner, markerWidth, object.extY)) {
				return 6;
			}
			if (hitInRect(t_x, t_y, 0, object.extY - markerWidth, widthCorner, object.extY)) {
				return 6;
			}

			return -1;

		}

		function hitToHandles(x, y, object) {
			var oApi = Asc.editor || editor;
			var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;

			if (isDrawHandles && object && object.isForm && object.isForm() && object.getInnerForm() && object.getInnerForm().IsFormLocked())
				isDrawHandles = false;

			if (isDrawHandles === false) {
				return -1;
			}

			if (object.cropObject) {
				return hitToCropHandles(x, y, object);
			}
			var invert_transform = object.getInvertTransform();
			if (!invert_transform) {
				return -1;
			}
			var t_x, t_y;
			t_x = invert_transform.TransformPointX(x, y);
			t_y = invert_transform.TransformPointY(x, y);
			var radius = object.convertPixToMM(AscCommon.TRACK_CIRCLE_RADIUS);

			if (typeof global_mouseEvent !== "undefined" && isRealObject(global_mouseEvent) && AscFormat.isRealNumber(global_mouseEvent.KoefPixToMM)) {
				radius *= global_mouseEvent.KoefPixToMM;
			}

			if (global_mouseEvent && global_mouseEvent.AscHitToHandlesEpsilon) {
				radius = global_mouseEvent.AscHitToHandlesEpsilon;
			}

			// чтобы не считать корни
			radius *= radius;

			// считаем ближайший маркер, так как окрестность может быть большой, и пересекаться.

			var _min_dist = 2 * radius; // главное что больше
			var _ret_value = -1;

			var check_line = CheckObjectLine(object);

			var sqr_x = t_x * t_x, sqr_y = t_y * t_y;
			var _tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist) {
				_min_dist = _tmp_dist;
				_ret_value = 0;
			}

			var hc = object.extX * 0.5;
			var dist_x = t_x - hc;
			sqr_x = dist_x * dist_x;
			_tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist && !check_line) {
				_min_dist = _tmp_dist;
				_ret_value = 1;
			}

			dist_x = t_x - object.extX;
			sqr_x = dist_x * dist_x;
			_tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist && !check_line) {
				_min_dist = _tmp_dist;
				_ret_value = 2;
			}

			var vc = object.extY * 0.5;
			var dist_y = t_y - vc;
			sqr_y = dist_y * dist_y;
			_tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist && !check_line) {
				_min_dist = _tmp_dist;
				_ret_value = 3;
			}

			dist_y = t_y - object.extY;
			sqr_y = dist_y * dist_y;
			_tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist) {
				_min_dist = _tmp_dist;
				_ret_value = 4;
			}

			dist_x = t_x - hc;
			sqr_x = dist_x * dist_x;
			_tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist && !check_line) {
				_min_dist = _tmp_dist;
				_ret_value = 5;
			}

			dist_x = t_x;
			sqr_x = dist_x * dist_x;
			_tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist && !check_line) {
				_min_dist = _tmp_dist;
				_ret_value = 6;
			}

			dist_y = t_y - vc;
			sqr_y = dist_y * dist_y;
			_tmp_dist = sqr_x + sqr_y;
			if (_tmp_dist < _min_dist && !check_line) {
				_min_dist = _tmp_dist;
				_ret_value = 7;
			}

			if (object.canRotate && object.canRotate() && !check_line) {
				var rotate_distance = object.convertPixToMM(AscCommon.TRACK_DISTANCE_ROTATE);
				dist_y = t_y + rotate_distance;
				sqr_y = dist_y * dist_y;
				dist_x = t_x - hc;
				sqr_x = dist_x * dist_x;

				_tmp_dist = sqr_x + sqr_y;
				if (_tmp_dist < _min_dist) {
					_min_dist = _tmp_dist;
					_ret_value = 8;
				}
			}

			// теперь смотрим расстояние до центра фигуры, чтобы можно было двигать маленькую
			dist_x = t_x - hc;
			dist_y = t_y - vc;
			_tmp_dist = dist_x * dist_x + dist_y * dist_y;
			if (_tmp_dist < _min_dist && !check_line) {
				_min_dist = _tmp_dist;
				_ret_value = -1;
			}

			if (_min_dist < radius)
				return _ret_value;

			return -1;
		}

		function CreateUniFillByUniColorCopy(uniColor) {
			var ret = new AscFormat.CUniFill();
			ret.setFill(new AscFormat.CSolidFill());
			ret.fill.setColor(uniColor.createDuplicate());
			return ret;
		}

		function CreateUniFillByUniColor(uniColor) {
			var ret = new AscFormat.CUniFill();
			ret.setFill(new AscFormat.CSolidFill());
			ret.fill.setColor(uniColor);
			return ret;
		}

		function CopyRunToPPTX(Run, Paragraph, bHyper) {
			var NewRun = new ParaRun(Paragraph, false);
			var RunPr = Run.Pr.Copy();
			if (RunPr.RStyle != undefined) {
				RunPr.RStyle = undefined;
			}
			RunPr.FontScale = undefined;

			if (bHyper) {
				if (!RunPr.Unifill) {
					RunPr.Unifill = AscFormat.CreateUniFillSchemeColorWidthTint(11, 0);
				}
				RunPr.Underline = true;
			}
			if (RunPr.TextFill) {
				RunPr.Unifill = RunPr.TextFill;
				RunPr.TextFill = undefined;
			}

			NewRun.Set_Pr(RunPr);

			var PosToAdd = 0;
			for (var CurPos = 0; CurPos < Run.Content.length; CurPos++) {
				var Item = Run.Content[CurPos];
				if (Item.Type !== para_End && Item.Type !== para_Drawing && Item.Type !== para_Comment
					&& Item.Type !== para_PageCount && Item.Type !== para_FootnoteRef && Item.Type !== para_FootnoteReference
					&& Item.Type !== para_PageNum && Item.Type !== para_FieldChar && Item.Type !== para_Bookmark
					&& Item.Type !== para_RevisionMove && Item.Type !== para_InstrText
					&& Item.Type !== para_EndnoteReference && Item.Type !== para_EndnoteRef) {
					NewRun.Add_ToContent(PosToAdd, Item.Copy(), false);
					++PosToAdd;
				}
			}
			return NewRun;
		}


		function ConvertParagraphContentToPPTX(aOrigContent, oNewParagraph, bIsAddMath, bRemoveHyperlink) {
			var Count = aOrigContent.length;
			for (var Index = 0; Index < Count; Index++) {
				var Item = aOrigContent[Index];
				if (Item.Type === para_Run) {
					oNewParagraph.Internal_Content_Add(oNewParagraph.Content.length, CopyRunToPPTX(Item, oNewParagraph), false);
				} else if (Item.Type === para_Hyperlink) {
					if (bRemoveHyperlink === true) {
						for (var j = 0; j < Item.Content.length; ++j) {
							if (Item.Content[j].Type === para_Run) {
								oNewParagraph.Internal_Content_Add(oNewParagraph.Content.length, CopyRunToPPTX(Item.Content[j], oNewParagraph), false);
							}
						}
					} else {
						var aChildContent = ConvertHyperlinkToPPTX(Item, oNewParagraph);
						for (var nChildIdx = 0; nChildIdx < aChildContent.length; ++nChildIdx) {
							oNewParagraph.Internal_Content_Add(oNewParagraph.Content.length, aChildContent[nChildIdx], false);
						}
					}

				} else if (Item.Type === para_InlineLevelSdt) {
					ConvertParagraphContentToPPTX(Item.Content, oNewParagraph, bIsAddMath, bRemoveHyperlink)
				} else if (true === bIsAddMath && Item.Type === para_Math) {
					oNewParagraph.Internal_Content_Add(oNewParagraph.Content.length, Item.Copy(), false);
				}
			}
		}

		function ConvertParagraphToPPTX(paragraph, drawingDocument, newParent, bIsAddMath, bRemoveHyperlink) {
			var _drawing_document = isRealObject(drawingDocument) ? drawingDocument : paragraph.DrawingDocument;
			var _new_parent = isRealObject(newParent) ? newParent : null;

			var new_paragraph = new Paragraph(_drawing_document, _new_parent, true);
			if (!(paragraph instanceof Paragraph))
				return new_paragraph;
			var oCopyPr = paragraph.Pr.Copy();

			oCopyPr.ContextualSpacing = undefined;
			oCopyPr.KeepLines = undefined;
			oCopyPr.KeepNext = undefined;
			oCopyPr.PageBreakBefore = undefined;
			oCopyPr.Shd = undefined;
			oCopyPr.Brd.First = undefined;
			oCopyPr.Brd.Last = undefined;
			oCopyPr.Brd.Between = undefined;
			oCopyPr.Brd.Bottom = undefined;
			oCopyPr.Brd.Left = undefined;
			oCopyPr.Brd.Right = undefined;
			oCopyPr.Brd.Top = undefined;
			oCopyPr.WidowControl = undefined;
			oCopyPr.Tabs = undefined;
			oCopyPr.NumPr = undefined;
			oCopyPr.PStyle = undefined;
			oCopyPr.FramePr = undefined;


			new_paragraph.Set_Pr(oCopyPr);
			var oNewEndPr = paragraph.TextPr.Value.Copy();
			if (oNewEndPr.TextFill) {
				oNewEndPr.Unifill = oNewEndPr.TextFill;
				oNewEndPr.TextFill = undefined;
			}
			new_paragraph.TextPr.Set_Value(oNewEndPr);
			new_paragraph.Internal_Content_Remove2(0, new_paragraph.Content.length);
			ConvertParagraphContentToPPTX(paragraph.Content, new_paragraph, bIsAddMath, bRemoveHyperlink);
			var EndRun = new ParaRun(new_paragraph);
			EndRun.Add_ToContent(0, new AscWord.CRunParagraphMark());
			new_paragraph.Internal_Content_Add(new_paragraph.Content.length, EndRun, false);
			return new_paragraph;
		}


		function ConvertElementsToPPTX(aResult, aElements, drawingDocument, newParent, bIsAddMath, bRemoveHyperlink) {
			var i, j, oElement;
			for (i = 0; i < aElements.length; ++i) {
				oElement = aElements[i];
				if (oElement instanceof AscCommonWord.Paragraph) {
					aResult.push(ConvertParagraphToPPTX(oElement));
				} else if (oElement instanceof AscCommonWord.CTable) {
					var paragraphs = [];
					oElement.GetAllParagraphs({All: true}, paragraphs);
					for (j = 0; j < paragraphs.length; j++) {
						aResult.push(AscFormat.ConvertParagraphToPPTX(paragraphs[j], drawingDocument,
							newParent, bIsAddMath, bRemoveHyperlink));
					}
				} else if (oElement instanceof AscCommonWord.CBlockLevelSdt) {
					ConvertElementsToPPTX(aResult, oElement.Content.Content, drawingDocument, newParent, bIsAddMath, bRemoveHyperlink)
				}
			}
		}

		function ConvertHyperlinkToPPTX(hyperlink, paragraph) {
			var hyperlink_ret = new ParaHyperlink(), i, item, pos = 0;
			hyperlink_ret.SetValue(hyperlink.Value);
			hyperlink_ret.SetToolTip(hyperlink.ToolTip);
			for (i = 0; i < hyperlink.Content.length; ++i) {
				item = hyperlink.Content[i];
				if (item.Type === para_Run) {
					hyperlink_ret.Add_ToContent(pos++, CopyRunToPPTX(item, paragraph, true));
				} else if (item.Type === para_Hyperlink) {
					var aConvertedContent = ConvertHyperlinkToPPTX(item, paragraph);
					for (var nChildIdx = 0; nChildIdx < aConvertedContent.length; ++nChildIdx) {
						hyperlink_ret.Add_ToContent(pos++, aConvertedContent[nChildIdx]);
					}
				}
			}
			if (typeof hyperlink.Value === "string" && hyperlink.Value.length > Asc.c_nMaxHyperlinkLength) {
				return hyperlink_ret.Content;
			}
			return [hyperlink_ret];
		}

		function ConvertParagraphToWord(paragraph, docContent) {
			var _docContent = isRealObject(docContent) ? docContent : paragraph.Parent;
			var oldFlag = paragraph.bFromDocument;
			paragraph.bFromDocument = true;
			var new_paragraph = paragraph.Copy(_docContent);
			CheckWordParagraphContent(new_paragraph.Content, new_paragraph.Pr.DefaultRunPr);
			var NewRPr = CheckWordRunPr(new_paragraph.TextPr.Value);
			var oCopyDefaultPr;
			if (NewRPr) {
				if (new_paragraph.Pr.DefaultRunPr) {
					oCopyDefaultPr = new_paragraph.Pr.DefaultRunPr.Copy();
					oCopyDefaultPr.Merge(NewRPr);
					NewRPr = CheckWordRunPr(oCopyDefaultPr);
					if (!NewRPr) {
						NewRPr = oCopyDefaultPr;
					}
				}
				new_paragraph.TextPr.Apply_TextPr(NewRPr);
			} else {
				if (new_paragraph.Pr.DefaultRunPr) {
					oCopyDefaultPr = new_paragraph.Pr.DefaultRunPr.Copy();
					oCopyDefaultPr.Merge(new_paragraph.TextPr.Value);
					NewRPr = CheckWordRunPr(oCopyDefaultPr);
					if (!NewRPr) {
						NewRPr = oCopyDefaultPr;
					}
					new_paragraph.TextPr.Apply_TextPr(NewRPr);
				}
			}
			paragraph.bFromDocument = oldFlag;
			return new_paragraph;
		}

		function CheckWordRunPr(Pr, bMath) {
			var NewRPr = null;
			if (Pr.Unifill && Pr.Unifill.fill) {
				switch (Pr.Unifill.fill.type) {
					case c_oAscFill.FILL_TYPE_SOLID: {
						if (Pr.Unifill.fill.color && Pr.Unifill.fill.color.color) {
							switch (Pr.Unifill.fill.color.color.type) {
								case Asc.c_oAscColor.COLOR_TYPE_SCHEME: {
									if (Pr.Unifill.fill.color.Mods && Pr.Unifill.fill.color.Mods.Mods.length !== 0) {
										if (!Pr.Unifill.fill.color.canConvertPPTXModsToWord()) {
											NewRPr = Pr.Copy();
											NewRPr.TextFill = NewRPr.Unifill;
											NewRPr.Unifill = undefined;
										} else {
											NewRPr = Pr.Copy();
											NewRPr.Unifill.convertToWordMods();
										}
									}
									break;
								}
								case Asc.c_oAscColor.COLOR_TYPE_SRGB: {

									NewRPr = Pr.Copy();
									var RGBA = Pr.Unifill.fill.color.color.RGBA;
									NewRPr.Color = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B);
									NewRPr.Unifill = undefined;
									break;
								}
								default: {
									NewRPr = Pr.Copy();
									NewRPr.TextFill = NewRPr.Unifill;
									NewRPr.Unifill = undefined;
								}
							}
						}
						break;
					}
					case c_oAscFill.FILL_TYPE_PATT:
					case c_oAscFill.FILL_TYPE_BLIP: {
						NewRPr = Pr.Copy();
						NewRPr.TextFill = AscFormat.CreateUnfilFromRGB(0, 0, 0);
						NewRPr.Unifill = undefined;
						break;
					}
					default : {
						NewRPr = Pr.Copy();
						NewRPr.TextFill = NewRPr.Unifill;
						NewRPr.Unifill = undefined;
						break;
					}
				}
			}

			if (bMath) {
				NewRPr = Pr.Copy();
				NewRPr.RFonts.SetAll("Cambria Math", -1);
			}
			return NewRPr;
		}

		function CheckWordParagraphContent(aContent, oTextPr) {
			var NewRPr, MergePr;
			for (var i = 0; i < aContent.length; ++i) {
				var oItem = aContent[i];
				switch (oItem.Type) {
					case para_Run: {
						NewRPr = CheckWordRunPr(oItem.Pr);
						if (NewRPr) {
							MergePr = NewRPr;
							if (oTextPr) {
								MergePr = oTextPr.Copy();
								MergePr.Merge(NewRPr);
								NewRPr = CheckWordRunPr(MergePr);
								if (!NewRPr) {
									NewRPr = MergePr;
								}
							}
							oItem.Set_Pr(NewRPr);
						} else {
							if (oTextPr) {
								MergePr = oTextPr.Copy();
								MergePr.Merge(oItem.Pr);
								NewRPr = CheckWordRunPr(MergePr);
								if (!NewRPr) {
									NewRPr = MergePr;
								}
								oItem.Set_Pr(NewRPr);
							}
						}
						break;
					}
					case para_Hyperlink: {
						CheckWordParagraphContent(oItem.Content);
						break;
					}
					case para_Math: {
						if (oItem.Root && oItem.Root.Content) {
							CheckWordParagraphContent(oItem.Root.Content);
						}
						break;
					}
					case para_Math_Run: {
						NewRPr = CheckWordRunPr(oItem.Pr, true);
						if (NewRPr) {
							MergePr = NewRPr;
							if (oTextPr) {
								MergePr = oTextPr.Copy();
								MergePr.Merge(NewRPr);
								NewRPr = CheckWordRunPr(MergePr);
								if (!NewRPr) {
									NewRPr = MergePr;
								}
							}
							oItem.Set_Pr(NewRPr);
						} else {
							if (oTextPr) {
								MergePr = oTextPr.Copy();
								MergePr.Merge(oItem.Pr);
								NewRPr = CheckWordRunPr(MergePr);
								if (!NewRPr) {
									NewRPr = MergePr;
								}
								oItem.Set_Pr(NewRPr);
							}
						}
						break;
					}
				}

			}
		}

		function ConvertGraphicFrameToWordTable(oGraphicFrame, oDocument) {
			oGraphicFrame.setWordFlag(false, oDocument);
			return oGraphicFrame.graphicObject.Copy(oDocument);
		}

		function ConvertTableToGraphicFrame(oTable, oPresentation) {
			var oGraphicFrame = new AscFormat.CGraphicFrame();
			var oTable2 = new CTable(oPresentation.DrawingDocument, oGraphicFrame, false, 0, [].concat(oTable.TableGrid), oTable.TableGrid.length, true);
			oTable2.Reset(0, 0, 50, 100000, 0, 0, 1);
			oTable2.SetTableLayout(tbllayout_Fixed);
			oTable2.Set_Pr(oTable.Pr.Copy());
			oTable2.Set_TableLook(oTable.TableLook.Copy());
			for (var i = 0; i < oTable.Content.length; ++i) {
				var oRow = oTable.Content[i];
				var oNewRow = new CTableRow(oTable2, oRow.Content.length, oTable2.TableGrid);
				for (var j = 0; j < oRow.Content.length; ++j) {
					var oContent = oRow.Content[j].Content;
					var oNewContent = oNewRow.Content[j].Content;
					for (var t = 0; t < oContent.Content.length; ++t) {
						if (oContent.Content[t].Get_Type() === type_Paragraph) {
							oNewContent.Internal_Content_Add(oNewContent.Content.length, AscFormat.ConvertParagraphToPPTX(oContent.Content[t], oPresentation.DrawingDocument, oNewContent));
						}
					}
				}
				var nIndex = oTable2.Content.length;
				oTable2.Content[nIndex] = oNewRow;
				History.Add(new CChangesTableAddRow(oTable2, nIndex, [oNewRow]));
				oTable2.private_UpdateTableGrid();
			}

			if (!oGraphicFrame.spPr) {
				oGraphicFrame.setSpPr(new AscFormat.CSpPr());
				oGraphicFrame.spPr.setParent(oGraphicFrame);
			}
			oGraphicFrame.spPr.setXfrm(new AscFormat.CXfrm());
			oGraphicFrame.spPr.xfrm.setExtX(50);
			oGraphicFrame.spPr.xfrm.setExtY(50);
			oGraphicFrame.spPr.xfrm.setParent(oGraphicFrame.spPr);
			var _nvGraphicFramePr = new AscFormat.UniNvPr();
			oGraphicFrame.setNvSpPr(_nvGraphicFramePr);
			if (AscCommon.isRealObject(_nvGraphicFramePr) && AscFormat.isRealNumber(_nvGraphicFramePr.locks)) {
				oGraphicFrame.setLocks(_nvGraphicFramePr.locks);
			}
			oGraphicFrame.setGraphicObject(oTable2);
			oGraphicFrame.setBDeleted(false);
			return oGraphicFrame;
		}


		function fHandleContent(aContent, oMax) {
			for (var i = 0; i < aContent.length; ++i) {
				var oContentElement = aContent[i];
				if (oContentElement.Get_Type() === type_Paragraph) {
					var paragraph_lines = aContent[i].Lines;
					for (var j = 0; j < paragraph_lines.length; ++j) {
						if (paragraph_lines[j].Ranges[0].W > oMax.max_width)
							oMax.max_width = paragraph_lines[j].Ranges[0].X + paragraph_lines[j].Ranges[0].W;
					}
				} else if (oContentElement.Get_Type() === type_Table) {
					if (oContentElement.Bounds.Right > oMax.max_width) {
						oMax.max_width = oContentElement.Bounds.Right;
					}
				} else if (oContentElement.Get_Type() === type_BlockLevelSdt) {
					if (oContentElement && oContentElement.Content) {
						fHandleContent(oContentElement.Content.Content, oMax);
					}
				}
			}
		}

		function RecalculateDocContentByMaxLine(oDocContent, dMaxWidth, bNeedRecalcAllDrawings) {
			var oMaxWidth = {max_width: 0}, i;
			oDocContent.Reset(0, 0, dMaxWidth, 20000);
			if (bNeedRecalcAllDrawings) {
				var aAllDrawings = oDocContent.GetAllDrawingObjects();
				for (i = 0; i < aAllDrawings.length; ++i) {
					aAllDrawings[i].GraphicObj.recalculate();
				}
			}
			oDocContent.Recalculate_Page(0, true);
			fHandleContent(oDocContent.Content, oMaxWidth);
			if (oMaxWidth.max_width === 0) {
				if (oDocContent.Is_Empty()) {
					if (oDocContent.Content[0] && oDocContent.Content[0].Content[0] && oDocContent.Content[0].Content[0].Content[0]) {
						return oDocContent.Content[0].Content[0].Content[0].GetWidthVisible();
					}
				}
				return 0.001;
			}
			return oMaxWidth.max_width;
		}


		function CheckExcelDrawingXfrm(xfrm) {
			var rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;

			if (checkNormalRotate(rot)) {
				if (xfrm.offX < 0) {
					xfrm.setOffX(0);
				}
				if (xfrm.offY < 0) {
					xfrm.setOffY(0);
				}
			} else {
				var dPosX = xfrm.offX + xfrm.extX / 2 - xfrm.extY / 2;
				var dPosY = xfrm.offY + xfrm.extY / 2 - xfrm.extX / 2;
				if (dPosX < 0) {
					xfrm.setOffX(xfrm.offX - dPosX);
				}
				if (dPosY < 0) {
					xfrm.setOffY(xfrm.offY - dPosY);
				}
			}
		}


		function SetXfrmFromMetrics(oDrawing, metrics) {
			AscFormat.CheckSpPrXfrm(oDrawing);
			var rot = AscFormat.isRealNumber(oDrawing.spPr.xfrm.rot) ? AscFormat.normalizeRotate(oDrawing.spPr.xfrm.rot) : 0;

			var metricExtX, metricExtY;
			if (oDrawing.getObjectType() !== AscDFH.historyitem_type_GroupShape) {
				metricExtX = metrics.extX;
				metricExtY = metrics.extY;
				if (checkNormalRotate(rot)) {
					oDrawing.spPr.xfrm.setExtX(metrics.extX);
					oDrawing.spPr.xfrm.setExtY(metrics.extY);
				} else {
					oDrawing.spPr.xfrm.setExtX(metrics.extY);
					oDrawing.spPr.xfrm.setExtY(metrics.extX);
				}
			} else {
				if (AscFormat.isRealNumber(oDrawing.spPr.xfrm.extX) && AscFormat.isRealNumber(oDrawing.spPr.xfrm.extY)) {
					metricExtX = oDrawing.spPr.xfrm.extX;
					metricExtY = oDrawing.spPr.xfrm.extY;
				} else {
					metricExtX = metrics.extX;
					metricExtY = metrics.extY;
				}
			}

			if (checkNormalRotate(rot)) {
				oDrawing.spPr.xfrm.setOffX(metrics.x);
				oDrawing.spPr.xfrm.setOffY(metrics.y);
			} else {
				oDrawing.spPr.xfrm.setOffX(metrics.x + metricExtX / 2 - metricExtY / 2);
				oDrawing.spPr.xfrm.setOffY(metrics.y + metricExtY / 2 - metricExtX / 2);
			}
		}


		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetNvSpPr] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetSpPr] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetShapeSmartArtPointInfo] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetTxXfrm] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetSmartArtPoint] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetStyle] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetTxBody] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetTextBoxContent] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetBodyPr] = AscDFH.CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetBFromSerialize] = AscDFH.CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetParent] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetGroup] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetWordShape] = AscDFH.CChangesDrawingsBool;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetModelId] = AscDFH.CChangesDrawingsString;
		AscDFH.changesFactory[AscDFH.historyitem_ShapeSetSignature] = AscDFH.CChangesDrawingsObjectNoId;


		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetNvSpPr] = function (oClass, value) {
			oClass.nvSpPr = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetSmartArtPoint] = function (oClass, value) {
			oClass.point = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetShapeSmartArtPointInfo] = function (oClass, value) {
			oClass.shapeSmartArtInfo = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetTxXfrm] = function (oClass, value) {
			oClass.txXfrm = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetModelId] = function (oClass, value) {
			oClass.modelId = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetSpPr] = function (oClass, value) {
			oClass.spPr = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetStyle] = function (oClass, value) {
			oClass.style = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetTxBody] = function (oClass, value) {
			oClass.txBody = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetTextBoxContent] = function (oClass, value) {
			oClass.textBoxContent = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetBodyPr] = function (oClass, value) {
			oClass.bodyPr = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetBFromSerialize] = function (oClass, value) {
			oClass.fromSerialize = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetParent] = function (oClass, value) {
			oClass.oldParent = oClass.parent;
			oClass.parent = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetGroup] = function (oClass, value) {
			oClass.group = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetWordShape] = function (oClass, value) {
			oClass.bWordShape = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ShapeSetSignature] = function (oClass, value) {
			var oldSignature = oClass.signatureLine;
			var newSignature = value;
			oClass.signatureLine = value;
			//interface updating
			if (!AscCommon.isFileBuild()) {
				var oApi = window["Asc"] && window["Asc"]["editor"] || editor;
				if (oApi) {
					if (oldSignature && oldSignature.id) {
						oApi.sendEvent("asc_onRemoveSignature", oldSignature.id);
					}
					if (newSignature && newSignature.id) {
						oApi.sendEvent("asc_onAddSignature", newSignature.id);
					}
				}
			}
		};

		function CSignatureLine() {
			this.id = null;
			this.signer = null;
			this.signer2 = null;
			this.email = null;
			this.showDate = null;
			this.instructions = null;
		}

		CSignatureLine.prototype.Write_ToBinary = function (writer) {
			AscFormat.writeString(writer, this.id);
			AscFormat.writeString(writer, this.signer);
			AscFormat.writeString(writer, this.signer2);
			AscFormat.writeString(writer, this.email);
			AscFormat.writeBool(writer, this.showDate);
			AscFormat.writeString(writer, this.instructions);
		};
		CSignatureLine.prototype.Read_FromBinary = function (reader) {
			this.id = AscFormat.readString(reader);
			this.signer = AscFormat.readString(reader);
			this.signer2 = AscFormat.readString(reader);
			this.email = AscFormat.readString(reader);
			this.showDate = AscFormat.readBool(reader);
			this.instructions = AscFormat.readString(reader);
		};
		CSignatureLine.prototype.copy = function () {
			var ret = new CSignatureLine();
			ret.id = AscCommon.CreateGUID();
			ret.signer = this.signer;
			ret.signer2 = this.signer2;
			ret.email = this.email;
			ret.showDate = this.showDate;
			ret.instructions = this.instructions;
			return ret;
		};
		CSignatureLine.prototype.copyWithId = function () {
			var sId = this.id;
			var oCopy = this.copy();
			oCopy.id = sId;
			return oCopy;
		};
		CSignatureLine.prototype.setProperties = function (oPr) {
			this.signer = oPr.asc_getSigner1();
			this.signer2 = oPr.asc_getSigner2();
			this.email = oPr.asc_getEmail();
			this.showDate = oPr.asc_getShowDate();
			this.instructions = oPr.asc_getInstructions();
		};

		AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ShapeSetBodyPr] = AscFormat.CBodyPr;
		AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ShapeSetSignature] = CSignatureLine;


		const TEXT_RECT_ERROR = 0.01;

		function CShape() {
			AscFormat.CGraphicObjectBase.call(this);
			this.nvSpPr = null;
			this.style = null;
			this.txBody = null;
			this.bodyPr = null;
			this.textBoxContent = null;
			this.drawingBase = null;//DrawingBase в Excell'е
			this.bWordShape = null;//если этот флаг стоит в true то автофигура имеет формат как в редакторе документов
			this.bCheckAutoFitFlag = false;
			this.signatureLine = null;
			this.txXfrm = null;
			this.modelId = null;


			this.transformText = new CMatrix();
			this.invertTransformText = null;

			this.localTransformText = new CMatrix();
			this.worksheet = null;
			this.cachedImage = null;

			this.txWarpStruct = null;
			this.txWarpStructParamarks = null;

			this.txWarpStructNoTransform = null;
			this.txWarpStructParamarksNoTransform = null;

			this.tmpFontScale = undefined;
			this.tmpLnSpcReduction = undefined;
			this.shapeSmartArtInfo = null;
		}

		AscFormat.InitClass(CShape, AscFormat.CGraphicObjectBase, AscDFH.historyitem_type_Shape);
		CShape.prototype.setCustT = function (value) {
			var pointContent = this.getSmartArtPointContent();
			if (pointContent) {
				pointContent.forEach(function (point) {
					if (point.prSet && point.prSet.custT !== value) {
						point.prSet.setCustT(value);
					}
				})
			}
		}

		CShape.prototype.setShapeSmartArtInfo = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetShapeSmartArtPointInfo, this.shapeSmartArtInfo, pr));
			this.shapeSmartArtInfo = pr;
			this.shapeSmartArtInfo.setParent(this);
		}
		CShape.prototype.isActiveBlipFillPlaceholder = function () {
			var shapePoint = this.getSmartArtShapePoint();
			if (shapePoint) {
				var isNotBlipFill = shapePoint.isBlipFillPlaceholder() && (shapePoint.spPr && !shapePoint.spPr.Fill || !shapePoint.spPr);
				return isNotBlipFill;
			}
		}
		CShape.prototype.getSmartArtInfo = function () {
			return this.shapeSmartArtInfo;
		}

		CShape.prototype.GetAllDrawingObjects = function (DrawingObjects) {
			var oContent = this.getDocContent();
			if (oContent) {
				oContent.GetAllDrawingObjects(DrawingObjects);
			}
		};
		CShape.prototype.setSignature = function (oSignature) {
			History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ShapeSetSignature, this.signatureLine, oSignature));
			this.signatureLine = oSignature;
		};
		CShape.prototype.setSignaturePr = function (oPr, sUrl) {
			if (!oPr || !this.signatureLine) {
				return;
			}
			var oCopy = this.signatureLine.copyWithId();
			oCopy.setProperties(oPr);
			this.setSignature(oCopy);
			if (sUrl) {
				if (this.spPr) {
					var oBlipFillUnifill = AscFormat.CreateBlipFillUniFillFromUrl(sUrl);
					this.spPr.setFill(oBlipFillUnifill);
				}
			}
		};

		CShape.prototype.convertToWord = function (document) {
			this.setBDeleted(true);
			this.convertFromSmartArt();
			var c = new CShape();
			c.setWordShape(true);
			c.setBDeleted(false);
			if (this.nvSpPr) {
				c.setNvSpPr(this.nvSpPr.createDuplicate());
			}
			if (this.spPr) {
				c.setSpPr(this.spPr.createDuplicate());
				if (!c.spPr.geometry) {
					c.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
				}
				c.spPr.setParent(c);
			}
			if (this.style) {
				c.setStyle(this.style.createDuplicate());
			}
			if (this.txBody) {
				if (this.txBody.bodyPr) {
					c.setBodyPr(this.txBody.bodyPr.createDuplicate());
				}
				if (this.txBody.content) {
					var new_content = new CDocumentContent(c, document.DrawingDocument, 0, 0, 0, 20000, false, false, false);
					var paragraphs = this.txBody.content.Content;

					new_content.Internal_Content_RemoveAll();
					for (var i = 0; i < paragraphs.length; ++i) {
						var cur_par = paragraphs[i];
						var new_paragraph = ConvertParagraphToWord(cur_par, new_content);
						new_content.Internal_Content_Add(i, new_paragraph, false);
					}
					c.setTextBoxContent(new_content);
				}
			}
			if (this.signatureLine) {
				c.setSignature(this.signatureLine.copy());
			}
			return c;
		};

		CShape.prototype.convertToPPTX = function (drawingDocument, worksheet, bIsAddMath) {
			let c = new CShape();
			c.setWordShape(false);
			c.setBDeleted(false);
			c.setWorksheet(worksheet);
			if (this.nvSpPr) {
				c.setNvSpPr(this.nvSpPr.createDuplicate());
			}
			if (this.spPr) {
				c.setSpPr(this.spPr.createDuplicate());
				c.spPr.setParent(c);
			}
			if (this.style) {
				c.setStyle(this.style.createDuplicate());
			}
			if (this.textBoxContent) {
				let tx_body = new AscFormat.CTextBody();
				tx_body.setParent(c);
				if (this.bodyPr) {
					tx_body.setBodyPr(this.bodyPr.createDuplicate());
				}
				let new_content = new AscFormat.CDrawingDocContent(tx_body, drawingDocument, 0, 0, 0, 0, false, false, true);
				let aContent = this.textBoxContent.Content;
				let aNewParagraphs = [];
				for (let nIdx = 0; nIdx < aContent.length; ++nIdx) {
					let oCurElement = aContent[nIdx];
					if (oCurElement instanceof AscCommonWord.Paragraph) {
						let oParagraph = ConvertParagraphToPPTX(oCurElement, drawingDocument, new_content, bIsAddMath);
						aNewParagraphs.push(oParagraph);
					}
				}
				if (aNewParagraphs.length > 0) {
					new_content.Internal_Content_RemoveAll();
					for (let nIdx = 0; nIdx < aNewParagraphs.length; ++nIdx) {
						let oParagraph = aNewParagraphs[nIdx];
						new_content.Internal_Content_Add(nIdx, oParagraph, false);
					}
				}
				tx_body.setContent(new_content);
				c.setTxBody(tx_body);
			}
			if (worksheet) {
				if (this.signatureLine) {
					c.setSignature(this.signatureLine.copy());
				}
			}
			return c;
		};
		CShape.prototype.convertFromSmartArt = function (bForce) {
			if (AscFormat.SmartArt && !bForce) {
				return this;
			}
			var txXfrm = this.txXfrm;
			if (txXfrm) {
				if (AscFormat.isRealNumber(txXfrm.rot) && this.txBody) {
					var oCopyBodyPr;
					var rot2 = txXfrm.rot;
					while (rot2 < 0) {
						rot2 += 2 * Math.PI;
					}
					var nSquare = ((2.0 * rot2 / Math.PI + 0.5) >> 0);
					while (nSquare < 0) {
						nSquare += 4;
					}
					switch (nSquare) {
						case 0: {
							oCopyBodyPr = this.txBody.bodyPr ? this.txBody.bodyPr.createDuplicate() : new AscFormat.CBodyPr();
							oCopyBodyPr.rot = (rot2 / AscFormat.cToRad + 0.5) >> 0;
							this.txBody.setBodyPr(oCopyBodyPr);
							break;
						}
						case 1: {
							oCopyBodyPr = this.txBody.bodyPr ? this.txBody.bodyPr.createDuplicate() : new AscFormat.CBodyPr();
							oCopyBodyPr.vert = AscFormat.nVertTTvert;
							this.txBody.setBodyPr(oCopyBodyPr);
							break;
						}
						case 2: {
							oCopyBodyPr = this.txBody.bodyPr ? this.txBody.bodyPr.createDuplicate() : new AscFormat.CBodyPr();
							oCopyBodyPr.rot = (rot2 / AscFormat.cToRad + 0.5) >> 0;
							this.txBody.setBodyPr(oCopyBodyPr);
							break;
						}
						case 3: {
							oCopyBodyPr = this.txBody.bodyPr ? this.txBody.bodyPr.createDuplicate() : new AscFormat.CBodyPr();
							oCopyBodyPr.vert = AscFormat.nVertTTvert270;
							this.txBody.setBodyPr(oCopyBodyPr);
							break;
						}
					}
				}
				this.setTxXfrm(null);
			}
			return this;
		};
		CShape.prototype.handleAllContents = function (fCallback) {
			var content = this.getDocContent();
			if (content) {
				fCallback(content);
			}
		};


		CShape.prototype.isTextBox = function () {
			return this.getTxBox() === true;
		};

		CShape.prototype.documentGetAllFontNames = function (AllFonts) {
			//TODO
			var content = this.getDocContent();
			if (content) {
				content.Document_Get_AllFontNames(AllFonts);
			}
		};

		CShape.prototype.documentCreateFontMap = function (map) {
			var content = this.getDocContent();
			if (content) {
				content.Document_CreateFontMap(map);
			}
		};


		CShape.prototype.setNvSpPr = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetNvSpPr, this.nvSpPr, pr));
			this.nvSpPr = pr;
		};

		CShape.prototype.setTxXfrm = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetTxXfrm, this.txXfrm, pr));
			this.txXfrm = pr;
			if (this.txXfrm) {
				this.txXfrm.setParent(this);
			}
		};

		CShape.prototype.setModelId = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ShapeSetModelId, this.modelId, pr));
			this.modelId = pr;
		};

		CShape.prototype.setSpPr = function (spPr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetSpPr, this.spPr, spPr));
			this.spPr = spPr;
			if (spPr) {
				spPr.setParent(this);
			}
		};

		CShape.prototype.setStyle = function (style) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetStyle, this.style, style));
			this.style = style;
			var content = this.getDocContent();

			this.recalcInfo.recalculateShapeStyleForParagraph = true;
			this.recalcInfo.recalculateContent2 = true;
			if (this.recalcTextStyles)
				this.recalcTextStyles();
			if (content) {
				content.Recalc_AllParagraphs_CompiledPr();
			}
		};

		CShape.prototype.setTxBody = function (txBody) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetTxBody, this.txBody, txBody));
			this.txBody = txBody;
			if (txBody && txBody.parent !== this) {
				txBody.setParent(this);
			}
		};

		CShape.prototype.setTextBoxContent = function (textBoxContent) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetTextBoxContent, this.textBoxContent, textBoxContent));
			this.textBoxContent = textBoxContent;
		};

		CShape.prototype.setBodyPr = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ShapeSetBodyPr, this.bodyPr, pr));
			this.bodyPr = pr;
			this.recalcInfo.recalculateContent = true;
			this.recalcInfo.recalculateTransformText = true;
			this.addToRecalculate();
		};

		CShape.prototype.createTextBody = function () {
			var tx_body = new AscFormat.CTextBody();
			tx_body.setParent(this);
			tx_body.setContent(new AscFormat.CDrawingDocContent(tx_body, this.getDrawingDocument(), 0, 0, 0, 20000, false, false, true));
			var oBodyPr = new AscFormat.CBodyPr();
			if (this.worksheet) {
				oBodyPr.vertOverflow = AscFormat.nVOTClip;
				oBodyPr.horzOverflow = AscFormat.nHOTClip;
			}
			tx_body.setBodyPr(oBodyPr);
			tx_body.content.Content[0].Set_DocumentIndex(0);
			tx_body.content.MoveCursorToStartPos(false);
			this.setTxBody(tx_body);
		};

		CShape.prototype.createTextBoxContent = function () {
			var body_pr = new AscFormat.CBodyPr();
			body_pr.setAnchor(1);
			this.setBodyPr(body_pr);
			this.setTextBoxContent(new CDocumentContent(this, this.getDrawingDocument(), 0, 0, 0, 20000, false, false));
			this.textBoxContent.SetParagraphAlign(AscCommon.align_Center);
			this.textBoxContent.MoveCursorToStartPos(false);
			this.textBoxContent.Content[0].Set_DocumentIndex(0);
		};

		CShape.prototype.paragraphAdd = function (paraItem, bRecalculate) {
			var content_to_add = this.getDocContent();
			if (!content_to_add) {
				if (this.canEditText()) {
					if (this.bWordShape) {
						this.createTextBoxContent();
					} else {
						this.createTextBody();
					}
					content_to_add = this.getDocContent();
				}
			}
			if (content_to_add) {
				content_to_add.AddToParagraph(paraItem, bRecalculate);
			}
		};

		CShape.prototype.applyTextFunction = function (docContentFunction, tableFunction, args) {
			var content_to_add = this.getDocContent();
			if (!content_to_add) {
				if (this.canEditText()) {

					if (this.bWordShape) {
						this.createTextBoxContent();
					} else {
						this.createTextBody();
					}
					content_to_add = this.getDocContent();
					content_to_add.MoveCursorToStartPos();
				}
			}
			if (content_to_add) {
				if (this.isForm && this.isForm() && !content_to_add.IsCursorInSpecialForm())
					return;

				docContentFunction.apply(content_to_add, args);
			}
			if (!editor || !editor.noCreatePoint || editor.exucuteHistory) {
				var fontSize = args[0] && args[0].Value && args[0].Value.FontSize;
				if (fontSize) {
					this.setCustT(true);
				}
				this.checkExtentsByDocContent();
			}
		};

		CShape.prototype.copyTextInfoFromShapeToPoint = function (paddings) {
			var txBody = this.txBody;
			var pointContent = this.getSmartArtPointContent();
			var options = {};
			if (txBody && pointContent && pointContent.length !== 0) {
				options.pointContentLength = pointContent.length;
				var pointsCopy;

				for (var i = 0; i < pointContent.length; i += 1) {
					var point = pointContent[i];

					if (point.prSet && point.prSet.custT) {
						options.custT = true;
					}
					var bodyPr = point.t && point.t.bodyPr;
					if (bodyPr) {
						if (typeof bodyPr.lIns === 'number') {
							options.lIns = true;
						}
						if (typeof bodyPr.rIns === 'number') {
							options.rIns = true;
						}
						if (typeof bodyPr.bIns === 'number') {
							options.bIns = true;
						}
						if (typeof bodyPr.tIns === 'number') {
							options.tIns = true;
						}
					}
				}
				if (paddings) {
					if (typeof paddings.Left === 'number') {
						options.lIns = true;
					}
					if (typeof paddings.Right === 'number') {
						options.rIns = true;
					}
					if (typeof paddings.Bottom === 'number') {
						options.bIns = true;
					}
					if (typeof paddings.Top === 'number') {
						options.tIns = true;
					}
				}
				pointsCopy = txBody.createDuplicateForSmartArt(options);
				pointContent.forEach(function (point, idx) {
					point.setT(pointsCopy[idx])
				});
			}
		}

		CShape.prototype.clearContent = function () {
			var content = this.getDocContent();
			if (content) {
				content.SetApplyToAll(true);
				content.Remove(-1);
				content.AddToParagraph(new AscCommonWord.ParaTextPr({Lang: {Val: undefined}}), false);
				content.SetApplyToAll(false);
			}
		};

		CShape.prototype.getDocContent = function () {
			if (this.txBody) {
				return this.txBody.content;
			} else if (this.textBoxContent) {
				return this.textBoxContent;
			}
			return null;
		};

		CShape.prototype.getCurrentDocContentInSmartArt = function () {
			var content;
			if (this.txBody) {
				if (this.isPlaceholderInSmartArt()) {
					content = this.txBody.content2;
				} else {
					content = this.txBody.content;
				}
			}
			return content;
		}

		CShape.prototype.getBodyPr = function () {
			return AscFormat.ExecuteNoHistory(function () {

				let ret;
				if (this.bWordShape) {
					ret = new AscFormat.CBodyPr();
					ret.setDefault();
					if (this.bodyPr)
						ret.merge(this.bodyPr);
				} else {
					if (this.txBody && this.txBody.bodyPr) {
						ret = this.txBody.getCompiledBodyPr();
					} else {
						ret = new AscFormat.CBodyPr();
						ret.setDefault();
					}
				}
				let dScale = this.getScaleCoefficient();
				ret.lIns *= dScale;
				ret.tIns *= dScale;
				ret.rIns *= dScale;
				ret.bIns *= dScale;
				return ret;
			}, this, []);
		};

		CShape.prototype.GetRevisionsChangeElement = function (SearchEngine) {
			var oContent = this.getDocContent();
			if (oContent) {
				oContent.GetRevisionsChangeElement(SearchEngine);
			}
		};

		CShape.prototype.Search = function (SearchEngine, Type) {
			if (this.textBoxContent) {
				var dd = this.getDrawingDocument();
				dd.StartSearchTransform(this.transformText);
				this.textBoxContent.Search(SearchEngine, Type);
				dd.EndSearchTransform();
			} else if (this.txBody && this.txBody.content) {
				//var dd = this.getDrawingDocument();
				//dd.StartSearchTransform(this.transformText);
				this.txBody.content.Search(SearchEngine, Type);
				//dd.EndSearchTransform();
			}
		};

		CShape.prototype.GetSearchElementId = function (bNext, bCurrent) {
			if (this.textBoxContent)
				return this.textBoxContent.GetSearchElementId(bNext, bCurrent);

			else if (this.txBody && this.txBody.content) {
				return this.txBody.content.GetSearchElementId(bNext, bCurrent);
			}

			return null;
		};

		CShape.prototype.FindNextFillingForm = function (isNext, isCurrent) {
			if (this.textBoxContent)
				return this.textBoxContent.FindNextFillingForm(isNext, isCurrent, isCurrent);
			else if (this.txBody && this.txBody.content)
				return this.txBody.content.FindNextFillingForm(isNext, isCurrent, isCurrent);

			return null;
		};

		CShape.prototype.documentUpdateRulersState = function () {
			var content = this.getDocContent();
			if (!content)
				return;
			var xc, yc;
			var l, t, r, b;
			var body_pr = this.getBodyPr();
			var l_ins, t_ins, r_ins, b_ins;
			if (typeof body_pr.lIns === "number")
				l_ins = body_pr.lIns;
			else
				l_ins = 2.54;

			if (typeof body_pr.tIns === "number")
				t_ins = body_pr.tIns;
			else
				t_ins = 1.27;

			if (typeof body_pr.rIns === "number")
				r_ins = body_pr.rIns;
			else
				r_ins = 2.54;

			if (typeof body_pr.bIns === "number")
				b_ins = body_pr.bIns;
			else
				b_ins = 1.27;

			var oRect;
			if (this.getTextRect) {
				oRect = this.getTextRect();
			} else {
				oRect = {
					l: 0,
					t: 0,
					r: this.extX,
					b: this.extY
				};
			}
			l = oRect.l + l_ins;
			t = oRect.t + t_ins;
			r = oRect.r - r_ins;
			b = oRect.b - b_ins;

			var x_lt, y_lt, x_rt, y_rt, x_rb, y_rb, x_lb, y_lb;
			var tr = this.transform;
			x_lt = tr.TransformPointX(l, t);
			y_lt = tr.TransformPointY(l, t);

			x_rb = tr.TransformPointX(r, b);
			y_rb = tr.TransformPointY(r, b);


			xc = (x_lt + x_rb) * 0.5;
			yc = (y_lt + y_rb) * 0.5;

			var hc = (r - l) * 0.5;
			var vc = (b - t) * 0.5;

			this.getDrawingDocument().Set_RulerState_Paragraph({L: xc - hc, T: yc - vc, R: xc + hc, B: yc + vc});
			content.Document_UpdateRulersState(AscFormat.isRealNumber(this.selectStartPage) ? this.selectStartPage : 0);
		};

		CShape.prototype.setParent = function (parent) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetParent, this.parent, parent));
			this.parent = parent;
		};

		CShape.prototype.setGroup = function (group) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ShapeSetGroup, this.group, group));
			this.group = group;
		};

		CShape.prototype.getAllImages = function (images) {
			if (this.spPr && this.spPr.Fill && this.spPr.Fill.fill instanceof AscFormat.CBlipFill && typeof this.spPr.Fill.fill.RasterImageId === "string") {
				images[AscCommon.getFullImageSrc2(this.spPr.Fill.fill.RasterImageId)] = true;
			}
			const oContent = this.getDocContent && this.getDocContent();
			if (oContent) {
				oContent.getBulletImages(images);
			}
		};

		CShape.prototype.getAllFonts = function (fonts) {
			if (this.txBody) {
				this.txBody.content.Document_Get_AllFontNames(fonts);
				if (this.txBody && this.txBody.lstStyle) {
					this.txBody.lstStyle.Document_Get_AllFontNames(fonts);
				}
				delete fonts["+mj-lt"];
				delete fonts["+mn-lt"];
				delete fonts["+mj-ea"];
				delete fonts["+mn-ea"];
				delete fonts["+mj-cs"];
				delete fonts["+mn-cs"];
			}
		};

		CShape.prototype.canFill = function () {
			if (this.spPr && this.spPr.geometry) {
				return this.spPr.geometry.canFill();
			}
			return true;
		};


		CShape.prototype.getHierarchy = function (bIsSingleBody, info) {
			//if(this.recalcInfo.recalculateShapeHierarchy)
			{
				this.compiledHierarchy = [];
				if (this.parent) {
					var hierarchy = this.compiledHierarchy;
					if (this.isPlaceholder()) {
						var ph_type = this.getPlaceholderType();
						var ph_index = this.getPlaceholderIndex();
						var b_is_single_body;
						if (AscFormat.isRealBool(bIsSingleBody)) {
							b_is_single_body = bIsSingleBody;
						} else {
							b_is_single_body = this.getIsSingleBody && this.getIsSingleBody();
						}
						switch (this.parent.kind) {
							case AscFormat.TYPE_KIND.SLIDE: {
								hierarchy.push(this.parent.Layout.getMatchingShape(ph_type, ph_index, b_is_single_body, info));
								hierarchy.push(this.parent.Layout.Master.getMatchingShape(ph_type, ph_index, true));
								break;
							}

							case AscFormat.TYPE_KIND.LAYOUT: {
								hierarchy.push(this.parent.Master.getMatchingShape(ph_type, ph_index, true));
								break;
							}

							case AscFormat.TYPE_KIND.NOTES: {
								if (this.parent.Master) {
									hierarchy.push(this.parent.Master.getMatchingShape(ph_type, ph_index, true));
								}
								break;
							}
						}
					}
					this.recalcInfo.recalculateShapeHierarchy = true;
				}
			}
			return this.compiledHierarchy;
		};


		CShape.prototype.getPaddings = function () {
			var paddings = null;
			var shape = this;
			var body_pr;
			if (shape.txBody) {
				if (shape.txBody.compiledBodyPr) {
					body_pr = shape.txBody.compiledBodyPr;
				} else {
					body_pr = shape.txBody.getCompiledBodyPr();
				}
			} else if (shape.textBoxContent) {
				body_pr = shape.bodyPr;
			}
			if (body_pr) {
				paddings = new Asc.asc_CPaddings();
				if (typeof body_pr.lIns === "number")
					paddings.Left = body_pr.lIns;
				else
					paddings.Left = 2.54;

				if (typeof body_pr.tIns === "number")
					paddings.Top = body_pr.tIns;
				else
					paddings.Top = 1.27;

				if (typeof body_pr.rIns === "number")
					paddings.Right = body_pr.rIns;
				else
					paddings.Right = 2.54;

				if (typeof body_pr.bIns === "number")
					paddings.Bottom = body_pr.bIns;
				else
					paddings.Bottom = 1.27;
			}
			return paddings;
		};

		CShape.prototype.getCompiledFill = function () {
			if (this.recalcInfo.recalculateFill) {
				this.compiledFill = null;
				if (isRealObject(this.spPr) && isRealObject(this.spPr.Fill) && isRealObject(this.spPr.Fill.fill)) {
					if (this.spPr.Fill.fill instanceof AscFormat.CGradFill && this.spPr.Fill.fill.colors.length === 0) {
						var parent_objects = this.getParentObjects();
						var theme = parent_objects.theme;
						var fmt_scheme = theme.themeElements.fmtScheme;
						var fill_style_lst = fmt_scheme.fillStyleLst;
						for (var i = fill_style_lst.length - 1; i > -1; --i) {
							if (fill_style_lst[i] && fill_style_lst[i].fill instanceof AscFormat.CGradFill) {
								this.spPr.Fill = fill_style_lst[i].createDuplicate();
								break;
							}
						}
					}
					this.compiledFill = this.spPr.Fill.createDuplicate();
					if (this.compiledFill && this.compiledFill.fill && this.compiledFill.fill.type === c_oAscFill.FILL_TYPE_GRP) {
						if (this.group) {
							var group_compiled_fill = this.group.getCompiledFill();
							if (isRealObject(group_compiled_fill) && isRealObject(group_compiled_fill.fill)) {
								this.compiledFill = group_compiled_fill.createDuplicate();
							} else {
								this.compiledFill = null;
							}
						} else {
							this.compiledFill = null;
						}
					}
				} else if (isRealObject(this.group)) {
					var group_compiled_fill = this.group.getCompiledFill();
					if (isRealObject(group_compiled_fill) && isRealObject(group_compiled_fill.fill)) {
						this.compiledFill = group_compiled_fill.createDuplicate();
					} else {
						var hierarchy = this.getHierarchy();
						for (var i = 0; i < hierarchy.length; ++i) {
							if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.Fill) && isRealObject(hierarchy[i].spPr.Fill.fill)) {
								this.compiledFill = hierarchy[i].spPr.Fill.createDuplicate();
								break;
							}
						}
					}
				} else {
					var hierarchy = this.getHierarchy();
					for (var i = 0; i < hierarchy.length; ++i) {
						if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.Fill) && isRealObject(hierarchy[i].spPr.Fill.fill)) {
							this.compiledFill = hierarchy[i].spPr.Fill.createDuplicate();
							break;
						}
					}
				}
				this.recalcInfo.recalculateFill = false;
			}
			return this.compiledFill;
		};

		CShape.prototype.getMargins = function () {
			if (this.txBody) {
				return this.txBody.getMargins()
			} else {
				return null;
			}
		};
		CShape.prototype.Document_UpdateRulersState = function (margins) {
			if (this.txBody && this.txBody.content) {
				this.txBody.content.Document_UpdateRulersState(this.parent.num, this.getMargins());
			}
		};

		CShape.prototype.getCompiledLine = function () {
			if (this.recalcInfo.recalculateLine) {
				this.compiledLine = null;
				if (isRealObject(this.spPr) && isRealObject(this.spPr.ln) && isRealObject(this.spPr.ln)) {
					this.compiledLine = this.spPr.ln.createDuplicate();
				} else if (isRealObject(this.group)) {
					var group_compiled_line = this.group.getCompiledLine();
					if (isRealObject(group_compiled_line) && isRealObject(group_compiled_line.fill)) {
						this.compiledLine = group_compiled_line.createDuplicate();
					} else {
						var hierarchy = this.getHierarchy();
						for (var i = 0; i < hierarchy.length; ++i) {
							if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.ln)) {
								this.compiledLine = hierarchy[i].spPr.ln.createDuplicate();
								break;
							}
						}
					}
				} else {
					var hierarchy = this.getHierarchy();
					for (var i = 0; i < hierarchy.length; ++i) {
						if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.ln)) {
							this.compiledLine = hierarchy[i].spPr.ln.createDuplicate();
							break;
						}
					}
				}
				this.recalcInfo.recalculateLine = false;
			}
			return this.compiledLine;
		};

		CShape.prototype.getCompiledTransparent = function () {
			if (this.recalcInfo.recalculateTransparent) {
				this.compiledTransparent = null;
				if (isRealObject(this.spPr) && isRealObject(this.spPr.Fill)) {
					if (AscFormat.isRealNumber(this.spPr.Fill.transparent)) {
						this.compiledTransparent = this.spPr.Fill.transparent;
					} else {
						if (this.spPr.Fill && this.spPr.Fill.fill && this.spPr.Fill.fill.type === c_oAscFill.FILL_TYPE_GRP) {
							if (this.group && this.group.spPr && this.group.spPr.Fill && AscFormat.isRealNumber(this.group.spPr.Fill.transparent)) {
								this.compiledTransparent = this.group.spPr.Fill.transparent;
							}
						}
					}

				}
				if (null !== this.compiledTransparent) {

					this.recalcInfo.recalculateTransparent = false;
					return this.compiledTransparent;
				}
				if (isRealObject(this.group)) {
					var group_transparent = this.group.getCompiledTransparent();
					if (AscFormat.isRealNumber(group_transparent)) {
						this.compiledTransparent = group_transparent;
					} else {
						var hierarchy = this.getHierarchy();
						for (var i = 0; i < hierarchy.length; ++i) {
							if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.Fill) && AscFormat.isRealNumber(hierarchy[i].spPr.Fill.transparent)) {
								this.compiledTransparent = hierarchy[i].spPr.Fill.transparent;
								break;
							}

						}
					}
				} else {
					var hierarchy = this.getHierarchy();
					for (var i = 0; i < hierarchy.length; ++i) {
						if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.Fill) && AscFormat.isRealNumber(hierarchy[i].spPr.Fill.transparent)) {
							this.compiledTransparent = hierarchy[i].spPr.Fill.transparent;
							break;
						}

					}
				}
				this.recalcInfo.recalculateTransparent = false;
			}
			return this.compiledTransparent;
		};

		CShape.prototype.getPlaceholderType = function () {
			return this.isPlaceholder() ? this.nvSpPr.nvPr.ph.type : null;
		};

		CShape.prototype.getPlaceholderIndex = function () {
			return this.isPlaceholder() ? this.nvSpPr.nvPr.ph.idx : null;
		};

		CShape.prototype.getPhType = function () {
			var point = this.getSmartArtShapePoint();
			if (point) {
				return AscFormat.phType_pic;
			}
			return this.isPlaceholder() ? this.nvSpPr.nvPr.ph.type : null;
		};

		CShape.prototype.getPhIndex = function () {
			return this.isPlaceholder() ? this.nvSpPr.nvPr.ph.idx : null;
		};

		CShape.prototype.setVerticalAlign = function (align) {
			var content_to_add = this.getDocContent();
			if (!content_to_add) {
				if (this.canEditText()) {
					if (this.bWordShape) {
						this.createTextBoxContent();
					} else {
						this.createTextBody();
					}
				}
			}
			var new_body_pr = this.getBodyPr();
			if (new_body_pr) {
				new_body_pr = new_body_pr.createDuplicate();
				new_body_pr.anchor = align;
				if (this.bWordShape) {
					this.setBodyPr(new_body_pr);
				} else {
					if (this.txBody) {
						this.txBody.setBodyPr(new_body_pr);
					}
				}
				if (this.isObjectInSmartArt()) {
					this.copyTextInfoFromShapeToPoint();
				}
			}
		};
		CShape.prototype.setVert = function (vert) {
			var content_to_add = this.getDocContent();
			if (!content_to_add) {
				if (this.canEditText()) {
					if (this.bWordShape) {
						this.createTextBoxContent();
					} else {
						this.createTextBody();
					}
				}
			}
			var new_body_pr = this.getBodyPr();
			if (new_body_pr) {
				new_body_pr = new_body_pr.createDuplicate();
				new_body_pr.vert = vert;
				if (this.bWordShape) {
					this.setBodyPr(new_body_pr);
				} else {
					if (this.txBody) {
						this.txBody.setBodyPr(new_body_pr);
					}
				}
			}
			this.checkExtentsByDocContent && this.checkExtentsByDocContent();
		};


		CShape.prototype.setTextFitType = function (type) {
			if (AscFormat.isRealNumber(type)) {
				var new_body_pr = this.getBodyPr();
				if (new_body_pr) {
					new_body_pr = new_body_pr.createDuplicate();
					new_body_pr.textFit = new AscFormat.CTextFit();
					new_body_pr.textFit.type = type;

					if (this.bWordShape) {
						this.setBodyPr(new_body_pr);
					} else {
						if (this.txBody) {
							this.txBody.setBodyPr(new_body_pr);
						}
					}
				}
				this.checkExtentsByDocContent(true, true);
			}
		};


		CShape.prototype.setVertOverflowType = function (type) {
			if (AscFormat.isRealNumber(type)) {
				var new_body_pr = this.getBodyPr();
				if (new_body_pr) {
					new_body_pr = new_body_pr.createDuplicate();
					new_body_pr.vertOverflow = type;

					if (this.bWordShape) {
						this.setBodyPr(new_body_pr);
					} else {
						if (this.txBody) {
							this.txBody.setBodyPr(new_body_pr);
						}
					}
				}
				this.checkExtentsByDocContent(true, true);
			}
		};


		CShape.prototype.setPaddings = function (paddings, props) {
			props = props || {};
			if (paddings) {
				var new_body_pr = this.getBodyPr();
				if (new_body_pr) {
					new_body_pr = new_body_pr.createDuplicate();
					if (AscFormat.isRealNumber(paddings.Left)) {
						new_body_pr.lIns = paddings.Left;
					}

					if (AscFormat.isRealNumber(paddings.Top)) {
						new_body_pr.tIns = paddings.Top;
					}

					if (AscFormat.isRealNumber(paddings.Right)) {
						new_body_pr.rIns = paddings.Right;
					}
					if (AscFormat.isRealNumber(paddings.Bottom)) {
						new_body_pr.bIns = paddings.Bottom;
					}

					if (this.bWordShape) {
						this.setBodyPr(new_body_pr);
					} else {
						if (this.txBody) {
							this.txBody.setBodyPr(new_body_pr);
						}
					}
					if (this.isObjectInSmartArt() && !props.bNotCopyToPoints) {
						this.copyTextInfoFromShapeToPoint(paddings);
					}
				}
			}
		};

		CShape.prototype.recalculateTransformText = function () {

			var oContent = this.getDocContent();
			if (!oContent)
				return;

			var oBodyPr = this.getBodyPr();
			this.clipRect = this.checkTransformTextMatrix(this.localTransformText, oContent, oBodyPr, false);
			if (this.isForm && this.isForm()) {
				this.clipRect = {x: 0, y: -0.2, w: this.extX, h: this.extY + 0.4};
			}
			this.transformText = this.localTransformText.CreateDublicate();
			this.invertTransformText = global_MatrixTransformer.Invert(this.transformText);

			if (this.txBody && this.txBody.content2) {
				this.transformText2 = new CMatrix();
				this.clipRect2 = this.checkTransformTextMatrix(this.transformText2, this.txBody.content2, oBodyPr, false);
				this.invertTransformText2 = global_MatrixTransformer.Invert(this.transformText2);
				this.localTransformText2 = this.transformText2.CreateDublicate();
			}
			//if (oBodyPr.prstTxWarp) {
			var bNoTextNoShape = oBodyPr.prstTxWarp && oBodyPr.prstTxWarp.preset !== "textNoShape";
			/*if (this.bWordShape) {
            this.transformTextWordArt = this.transformText;
            this.invertTransformTextWordArt = this.invertTransformText;
        }
        else*/
			{
				this.localTransformTextWordArt = new CMatrix();
				this.checkTransformTextMatrix(this.localTransformTextWordArt, oContent, oBodyPr, bNoTextNoShape, !this.bWordShape && bNoTextNoShape);
				this.transformTextWordArt = this.localTransformTextWordArt.CreateDublicate();
				this.invertTransformTextWordArt = global_MatrixTransformer.Invert(this.transformTextWordArt);
			}
			if (this.txBody && this.txBody.content2) {
				this.checkTransformTextMatrix(this.transformText2, this.txBody.content2, oBodyPr, bNoTextNoShape, !this.bWordShape && bNoTextNoShape);
				this.transformTextWordArt2 = new CMatrix();
				this.checkTransformTextMatrix(this.transformTextWordArt2, this.txBody.content2, oBodyPr, bNoTextNoShape, !this.bWordShape && bNoTextNoShape);
			}
			// }

			if (this.checkPosTransformText) {
				this.checkPosTransformText();
			}
			if (this.checkContentDrawings) {
				this.checkContentDrawings();
			}
		};

		CShape.prototype.getFullFlip = function () {
			var _transform = this.localTransform;
			var _full_rotate = this.getFullRotate();
			var _full_pos_x_lt = _transform.TransformPointX(0, 0);
			var _full_pos_y_lt = _transform.TransformPointY(0, 0);

			var _full_pos_x_rt = _transform.TransformPointX(this.extX, 0);
			var _full_pos_y_rt = _transform.TransformPointY(this.extX, 0);

			var _full_pos_x_rb = _transform.TransformPointX(this.extX, this.extY);
			var _full_pos_y_rb = _transform.TransformPointY(this.extX, this.extY);

			var _rotate_matrix = new CMatrix();
			global_MatrixTransformer.RotateRadAppend(_rotate_matrix, _full_rotate);

			var _rotated_pos_x_lt = _rotate_matrix.TransformPointX(_full_pos_x_lt, _full_pos_y_lt);

			var _rotated_pos_x_rt = _rotate_matrix.TransformPointX(_full_pos_x_rt, _full_pos_y_rt);
			var _rotated_pos_y_rt = _rotate_matrix.TransformPointY(_full_pos_x_rt, _full_pos_y_rt);

			var _rotated_pos_y_rb = _rotate_matrix.TransformPointY(_full_pos_x_rb, _full_pos_y_rb);
			return {
				flipH: _rotated_pos_x_lt > _rotated_pos_x_rt,
				flipV: _rotated_pos_y_rt > _rotated_pos_y_rb
			};
		};

		CShape.prototype.getTextRect = function () {
			let oRect;
			if (this.txXfrm && this.spPr && this.spPr.xfrm) {
				var newL = this.txXfrm.offX - this.spPr.xfrm.offX;
				var newT = this.txXfrm.offY - this.spPr.xfrm.offY;
				var newR = newL + this.txXfrm.extX;
				var newB = newT + this.txXfrm.extY;
				oRect = {};
				let dScale = this.getScaleCoefficient();
				oRect.l = newL * dScale;
				oRect.t = newT * dScale;
				oRect.r = newR * dScale;
				oRect.b = newB * dScale;
			} else {
				let _r = this.spPr && this.spPr.geometry && this.spPr.geometry.rect;
				if (_r) {
					oRect = {
						l: _r.l,
						t: _r.t,
						r: _r.r,
						b: _r.b
					};
				}
			}
			if (oRect) {
				return oRect;
			}
			return {
				l: 0,
				t: 0,
				r: this.extX,
				b: this.extY
			};
		};
		CShape.prototype.checkTransformTextMatrixSmartArt = function (oMatrix, oContent, oBodyPr, bWordArtTransform, bIgnoreInsets) {
			if (this.txXfrm && (this.isObjectInSmartArt && this.isObjectInSmartArt())) {
				var oSmartArt = this.group.group;
				const bForceSlideTransform = oSmartArt.bForceSlideTransform;
				var diffX = 0;
				var diffY = 0;
				if (oSmartArt.group) {
					if (bForceSlideTransform || (this.parent && this.parent.getObjectType() === AscDFH.historyitem_type_Slide || this.worksheet)) {
						const oMainGroupRelativePosition = oSmartArt.group.getRelativePosition();
						diffX = oMainGroupRelativePosition.x;
						diffY = oMainGroupRelativePosition.y;
					} else {
						const oMainGroup = oSmartArt.getMainGroup();
						const oMainGroupRelativePosition = oSmartArt.getRelativePosition();
						diffX = oMainGroupRelativePosition.x - oMainGroup.x;
						diffY = oMainGroupRelativePosition.y - oMainGroup.y;
					}
				}
				var oRect = this.getTextRect();
				var oRectShape = new AscFormat.CShape();
				oRectShape.setBDeleted(false);
				oRectShape.setSpPr(new AscFormat.CSpPr());
				oRectShape.spPr.setParent(oRectShape);
				oRectShape.spPr.setXfrm(new AscFormat.CXfrm());
				oRectShape.spPr.xfrm.setParent(oRectShape.spPr);

				var defaultRot = this.getDefaultRotSA();
				var deltaRot = AscFormat.normalizeRotate(this.rot - defaultRot);
				var deltaShape = new AscFormat.CShape();
				deltaShape.setBDeleted(false);
				deltaShape.setSpPr(new AscFormat.CSpPr());
				deltaShape.spPr.setParent(deltaShape);
				deltaShape.spPr.setXfrm(new AscFormat.CXfrm());
				deltaShape.spPr.xfrm.setParent(deltaShape.spPr);
				deltaShape.spPr.xfrm.setOffX(this.spPr.xfrm.offX - diffX);
				deltaShape.spPr.xfrm.setOffY(this.spPr.xfrm.offY - diffY);
				deltaShape.spPr.xfrm.setExtX(this.spPr.xfrm.extX);
				deltaShape.spPr.xfrm.setExtY(this.spPr.xfrm.extY);
				if (deltaRot) {
					deltaShape.spPr.xfrm.setRot(deltaRot);
				}
				deltaShape.setGroup(this.group);
				deltaShape.parent = this.parent;
				// deltaShape.changeFlipH(this.spPr.xfrm.flipH); TODO: repair this
				// deltaShape.changeFlipV(this.spPr.xfrm.flipV);
				deltaShape.recalculateLocalTransform(deltaShape.localTransform);

				var extX = (oRect.r - oRect.l) / 2;
				var extY = (oRect.b - oRect.t) / 2;
				var deltaTranslateX = 0, deltaTranslateY = 0;
				if (bForceSlideTransform || (deltaShape.parent && deltaShape.parent.getObjectType() === AscDFH.historyitem_type_Slide || this.worksheet)) {
					deltaTranslateX = deltaShape.group.group.x;
					deltaTranslateY = deltaShape.group.group.y;
				}
				var xc = deltaShape.localTransform.TransformPointX(oRect.l + extX, oRect.t + extY) - deltaTranslateX;
				var yc = deltaShape.localTransform.TransformPointY(oRect.l + extX, oRect.t + extY) - deltaTranslateY;

				let dScale = this.getScaleCoefficient();
				oRectShape.spPr.xfrm.setOffX((xc - extX) / dScale);
				oRectShape.spPr.xfrm.setOffY((yc - extY) / dScale);
				oRectShape.spPr.xfrm.setExtX(this.txXfrm.extX);
				oRectShape.spPr.xfrm.setExtY(this.txXfrm.extY);
				// oRectShape.changeFlipH(this.spPr.xfrm.flipH); TODO: repair this
				// oRectShape.changeFlipV(this.spPr.xfrm.flipV);
				oRectShape.spPr.xfrm.setRot(AscFormat.normalizeRotate(this.txXfrm.rot + this.spPr.xfrm.rot));
				oRectShape.setGroup(this.group);
				oRectShape.parent = this.parent;
				oRectShape.recalculateLocalTransform(oRectShape.localTransform);
				return oRectShape.checkTransformTextMatrix(oMatrix, oContent, oBodyPr, bWordArtTransform, bIgnoreInsets);
			}
		}

		CShape.prototype.getFormRelRect = function (isUsePaddings) {
			var oSpTransform = this.transform;
			var oInvTextTransform = this.invertTransformText;

			var nX = 0, nW = this.extX;
			var nY = 0, nH = this.extY;

			if (isUsePaddings) {
				let nFormHorPadding = this.getFormHorPadding();
				nX += nFormHorPadding;
				nW -= 2 * nFormHorPadding;
			}

			var aX = [nX, nW];
			var aY = [nY, nH];
			var fX0, fY0;

			if (!oSpTransform || !oInvTextTransform) {
				return {
					X: nX,
					Y: nY,
					W: nW,
					H: nH,
					Page: this.parent.PageNum
				};
			}

			var aRelX = [], aRelY = [];
			for (var nX = 0; nX < aX.length; ++nX) {
				fX0 = aX[nX];
				for (var nY = 0; nY < aY.length; ++nY) {
					fY0 = aY[nY];
					var fX = oSpTransform.TransformPointX(fX0, fY0);
					var fY = oSpTransform.TransformPointY(fX0, fY0);
					var fRelX = oInvTextTransform.TransformPointX(fX, fY);
					var fRelY = oInvTextTransform.TransformPointY(fX, fY);
					aRelX.push(fRelX);
					aRelY.push(fRelY);
				}
			}

			return {
				X: Math.min.apply(Math, aRelX),
				Y: Math.min.apply(Math, aRelY),
				W: nW,
				H: nH,
				Page: this.parent.PageNum
			};
		};

		CShape.prototype.checkTransformTextMatrix = function (oMatrix, oContent, oBodyPr, bWordArtTransform, bIgnoreInsets) {
			oMatrix.Reset();
			var _shape_transform = this.localTransform;
			var _content_height = oContent.GetSummaryHeight();
			var _l, _t, _r, _b;
			var _t_x_lt, _t_y_lt, _t_x_rt, _t_y_rt, _t_x_lb, _t_y_lb, _t_x_rb, _t_y_rb;
			var l_ins = bIgnoreInsets ? 0 : (AscFormat.isRealNumber(oBodyPr.lIns) ? oBodyPr.lIns : 2.54);
			var t_ins = bIgnoreInsets ? 0 : (AscFormat.isRealNumber(oBodyPr.tIns) ? oBodyPr.tIns : 1.27);
			var r_ins = bIgnoreInsets ? 0 : (AscFormat.isRealNumber(oBodyPr.rIns) ? oBodyPr.rIns : 2.54);
			var b_ins = bIgnoreInsets ? 0 : (AscFormat.isRealNumber(oBodyPr.bIns) ? oBodyPr.bIns : 1.27);

			var oRect = this.getTextRect();
			if (this.txXfrm) {
				return this.checkTransformTextMatrixSmartArt(oMatrix, oContent, oBodyPr, bWordArtTransform, bIgnoreInsets);
			}

			if (this.bWordShape) {
				var oPen = this.pen;
				if (oPen) {
					var penW = (oPen.w == null) ? 12700 : parseInt(oPen.w);
					penW /= 36000.0;
					switch (oPen.algn) {
						case 1: {
							break;
						}
						default: {
							penW /= 2.0;
							break;
						}
					}

					l_ins += penW;
					r_ins += penW;
					t_ins += penW;
					b_ins += penW;
				}
			}

			let oForm = this.isForm && this.isForm() ? this.getInnerForm() : null;
			if (oForm) {
				let nFormHorPadding = this.getFormHorPadding();
				l_ins = nFormHorPadding;
				r_ins = nFormHorPadding;
				t_ins = 0;
				b_ins = 0;
			}

			_l = oRect.l + l_ins;
			_t = oRect.t + t_ins;
			_r = oRect.r - r_ins;
			_b = oRect.b - b_ins;

			if (_l >= _r) {
				var _c = (_l + _r) * 0.5;
				_l = _c - 0.01;
				_r = _c + 0.01;
			}

			if (_t >= _b) {
				_c = (_t + _b) * 0.5;
				_t = _c - 0.01;
				_b = _c + 0.01;
			}

			var XC = oContent.XLimit / 2.0;
			var YC = _content_height / 2.0;

			var _rot_angle = AscFormat.normalizeRotate((AscFormat.isRealNumber(oBodyPr.rot) ? oBodyPr.rot : 0) * AscFormat.cToRad);

			if (!AscFormat.fApproxEqual(_rot_angle, 0.0)) {
				global_MatrixTransformer.TranslateAppend(oMatrix, -XC, -YC);
				global_MatrixTransformer.RotateRadAppend(oMatrix, -_rot_angle);
				global_MatrixTransformer.TranslateAppend(oMatrix, XC, YC);
			}

			_t_x_lt = _shape_transform.TransformPointX(_l, _t);
			_t_y_lt = _shape_transform.TransformPointY(_l, _t);

			_t_x_rt = _shape_transform.TransformPointX(_r, _t);
			_t_y_rt = _shape_transform.TransformPointY(_r, _t);

			_t_x_lb = _shape_transform.TransformPointX(_l, _b);
			_t_y_lb = _shape_transform.TransformPointY(_l, _b);

			_t_x_rb = _shape_transform.TransformPointX(_r, _b);
			_t_y_rb = _shape_transform.TransformPointY(_r, _b);

			var _dx_t, _dy_t;
			_dx_t = _t_x_rt - _t_x_lt;
			_dy_t = _t_y_rt - _t_y_lt;

			var _dx_lt_rb, _dy_lt_rb;
			_dx_lt_rb = _t_x_rb - _t_x_lt;
			_dy_lt_rb = _t_y_rb - _t_y_lt;

			var _vertical_shift;
			var _text_rect_height = _b - _t;
			var _text_rect_width = _r - _l;
			var oClipRect;
			var Diff = 1.6;

			if (oForm) {
				if (oForm.IsMultiLineForm())
					_vertical_shift = 0;
				else
					_vertical_shift = (_text_rect_height - _content_height) * 0.5;

				global_MatrixTransformer.TranslateAppend(oMatrix, 0, _vertical_shift);
				if (_dx_lt_rb * _dy_t - _dy_lt_rb * _dx_t <= 0) {
					var alpha = Math.atan2(_dy_t, _dx_t);
					global_MatrixTransformer.RotateRadAppend(oMatrix, -alpha);
					global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_lt, _t_y_lt);
				} else {
					var alpha = Math.atan2(_dy_t, _dx_t);
					global_MatrixTransformer.RotateRadAppend(oMatrix, Math.PI - alpha);
					global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_rt, _t_y_rt);
				}
			} else if (!oBodyPr.upright) {
				if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
					if (bWordArtTransform) {
						_vertical_shift = 0;
					} else {
						if ((!this.bWordShape && oBodyPr.vertOverflow === AscFormat.nVOTOverflow) || _content_height < _text_rect_height) {
							switch (oBodyPr.anchor) {
								case 0: //b
								{ // (Text Anchor Enum ( Bottom ))
									_vertical_shift = _text_rect_height - _content_height;

									break;
								}
								case 1:    //ctr
								case 2:    //dist TODO: пока выравнивание  по центру. Переделать!
								case 3:    //just TODO: пока выравнивание  по центру. Переделать!
									// (Text Anchor Enum ( Center ))
									_vertical_shift = (_text_rect_height - _content_height) * 0.5;
									break;

								// case 2: //dist
								// {// (Text Anchor Enum ( Distributed )) TODO: пока выравнивание  по центру. Переделать!
								//     _vertical_shift = (_text_rect_height - _content_height) * 0.5;
								//     break;
								// }
								// case 3: //just
								// {// (Text Anchor Enum ( Justified )) TODO: пока выравнивание  по центру. Переделать!
								//     _vertical_shift = (_text_rect_height - _content_height) * 0.5;
								//     break;
								// }
								case 4: //t
								{//Top
									_vertical_shift = 0;
									break;
								}
							}
						} else {


							if ((!this.bWordShape && oBodyPr.vertOverflow === AscFormat.nVOTClip)
								&& oContent.Content[0] && oContent.Content[0].Lines[0] && oContent.Content[0].Lines[0].Bottom > _text_rect_height) {
								var _content_first_line = oContent.Content[0].Lines[0].Bottom;
								switch (oBodyPr.anchor) {
									case 0: //b
									{ // (Text Anchor Enum ( Bottom ))
										_vertical_shift = _text_rect_height - _content_first_line;
										break;
									}
									case 1:    //ctr
									{// (Text Anchor Enum ( Center ))
										_vertical_shift = (_text_rect_height - _content_first_line) * 0.5;
										break;
									}
									case 2: //dist
									{// (Text Anchor Enum ( Distributed ))
										_vertical_shift = (_text_rect_height - _content_first_line) * 0.5;
										break;
									}
									case 3: //just
									{// (Text Anchor Enum ( Justified ))
										_vertical_shift = (_text_rect_height - _content_first_line) * 0.5;
										break;
									}
									case 4: //t
									{//Top
										_vertical_shift = 0;
										break;
									}
								}
							} else {
								if (this.bWordShape) {
									_vertical_shift = 0;
								} else {
									if (oBodyPr.anchor === 0) {
										_vertical_shift = _text_rect_height - _content_height;
									} else {
										_vertical_shift = 0;
									}
								}
							}

						}
					}
					global_MatrixTransformer.TranslateAppend(oMatrix, 0, _vertical_shift);
					if (_dx_lt_rb * _dy_t - _dy_lt_rb * _dx_t <= 0) {
						var alpha = Math.atan2(_dy_t, _dx_t);
						global_MatrixTransformer.RotateRadAppend(oMatrix, -alpha);
						global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_lt, _t_y_lt);
					} else {
						alpha = Math.atan2(_dy_t, _dx_t);
						global_MatrixTransformer.RotateRadAppend(oMatrix, Math.PI - alpha);
						global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_rt, _t_y_rt);
					}
				} else {
					if (bWordArtTransform) {
						_vertical_shift = 0;
					} else {
						if ((!this.bWordShape && oBodyPr.vertOverflow === AscFormat.nVOTOverflow) || _content_height <= _text_rect_width) {
							switch (oBodyPr.anchor) {
								case 0: //b
								{ // (Text Anchor Enum ( Bottom ))
									_vertical_shift = _text_rect_width - _content_height;
									break;
								}
								case 1:    //ctr
								{// (Text Anchor Enum ( Center ))
									_vertical_shift = (_text_rect_width - _content_height) * 0.5;
									break;
								}
								case 2: //dist
								{// (Text Anchor Enum ( Distributed ))
									_vertical_shift = (_text_rect_width - _content_height) * 0.5;
									break;
								}
								case 3: //just
								{// (Text Anchor Enum ( Justified ))
									_vertical_shift = (_text_rect_width - _content_height) * 0.5;
									break;
								}
								case 4: //t
								{//Top
									_vertical_shift = 0;
									break;
								}
							}
						} else {

							if (this.bWordShape) {
								_vertical_shift = 0;
							} else {
								if (oBodyPr.anchor === 0) {
									_vertical_shift = _text_rect_width - _content_height;
								} else {
									_vertical_shift = 0;
								}
							}
						}
					}
					global_MatrixTransformer.TranslateAppend(oMatrix, 0, _vertical_shift);
					var _alpha;
					_alpha = Math.atan2(_dy_t, _dx_t);
					if (oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTeaVert) {
						if (_dx_lt_rb * _dy_t - _dy_lt_rb * _dx_t <= 0) {
							global_MatrixTransformer.RotateRadAppend(oMatrix, -_alpha - Math.PI * 0.5);
							global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_rt, _t_y_rt);
						} else {
							global_MatrixTransformer.RotateRadAppend(oMatrix, Math.PI * 0.5 - _alpha);
							global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_lt, _t_y_lt);
						}
					} else {
						if (_dx_lt_rb * _dy_t - _dy_lt_rb * _dx_t <= 0) {
							global_MatrixTransformer.RotateRadAppend(oMatrix, -_alpha - Math.PI * 1.5);
							global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_lb, _t_y_lb);
						} else {
							global_MatrixTransformer.RotateRadAppend(oMatrix, -Math.PI * 0.5 - _alpha);
							global_MatrixTransformer.TranslateAppend(oMatrix, _t_x_rb, _t_y_rb);
						}
					}
				}

				var rect = this.getTextRect && this.getTextRect();
				if (rect) {
					var clipW = rect.r - rect.l + Diff;
					if (clipW <= 0) {
						clipW = 0.01;
					}
					var clipH = rect.b - rect.t + Diff - b_ins - t_ins;
					if (clipH < 0) {
						clipH = 0.01;
					}
					oClipRect = {x: rect.l - Diff, y: rect.t - Diff + t_ins, w: clipW, h: clipH};
				} else {
					oClipRect = {x: -1.6, y: t_ins, w: this.extX + 3.2, h: this.extY - b_ins};
				}

			} else {
				var _full_rotate = this.getFullRotate();
				var _full_flip = this.getFullFlip();

				var _hc = this.extX * 0.5;
				var _vc = this.extY * 0.5;
				var _transformed_shape_xc = this.localTransform.TransformPointX(_hc, _vc);
				var _transformed_shape_yc = this.localTransform.TransformPointY(_hc, _vc);


				var _content_width, content_height2;
				if (checkNormalRotate(_full_rotate)) {
					if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
						_content_width = _r - _l;
						content_height2 = _b - _t;
					} else {
						_content_width = _b - _t;
						content_height2 = _r - _l;
					}
				} else {
					if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
						_content_width = _b - _t;
						content_height2 = _r - _l;

					} else {
						_content_width = _r - _l;
						content_height2 = _b - _t;
					}
				}

				if (bWordArtTransform) {
					_vertical_shift = 0;
				} else {
					if (!(this.bWordShape || this.worksheet) || _content_height < content_height2) {
						switch (oBodyPr.anchor) {
							case 0: //b
							{ // (Text Anchor Enum ( Bottom ))
								_vertical_shift = content_height2 - _content_height;
								break;
							}
							case 1:    //ctr
							{// (Text Anchor Enum ( Center ))
								_vertical_shift = (content_height2 - _content_height) * 0.5;
								break;
							}
							case 2: //dist
							{// (Text Anchor Enum ( Distributed ))
								_vertical_shift = (content_height2 - _content_height) * 0.5;
								break;
							}
							case 3: //just
							{// (Text Anchor Enum ( Justified ))
								_vertical_shift = (content_height2 - _content_height) * 0.5;
								break;
							}
							case 4: //t
							{//Top
								_vertical_shift = 0;
								break;
							}
						}
					} else {
						if (this.bWordShape) {
							_vertical_shift = 0;
						} else {
							if (oBodyPr.anchor === 0) {
								_vertical_shift = content_height2 - _content_height;
							} else {
								_vertical_shift = 0;
							}
						}
					}

				}

				var _text_rect_xc = _l + (_r - _l) * 0.5;
				var _text_rect_yc = _t + (_b - _t) * 0.5;

				var _vx = _text_rect_xc - _hc;
				var _vy = _text_rect_yc - _vc;

				var _transformed_text_xc, _transformed_text_yc;
				if (!_full_flip.flipH) {
					_transformed_text_xc = _transformed_shape_xc + _vx;
				} else {
					_transformed_text_xc = _transformed_shape_xc - _vx;
				}

				if (!_full_flip.flipV) {
					_transformed_text_yc = _transformed_shape_yc + _vy;
				} else {
					_transformed_text_yc = _transformed_shape_yc - _vy;
				}

				global_MatrixTransformer.TranslateAppend(oMatrix, 0, _vertical_shift);
				if (oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTeaVert) {
					global_MatrixTransformer.TranslateAppend(oMatrix, -_content_width * 0.5, -content_height2 * 0.5);
					global_MatrixTransformer.RotateRadAppend(oMatrix, -Math.PI * 0.5);
					global_MatrixTransformer.TranslateAppend(oMatrix, _content_width * 0.5, content_height2 * 0.5);

				}
				if (oBodyPr.vert === AscFormat.nVertTTvert270) {
					global_MatrixTransformer.TranslateAppend(oMatrix, -_content_width * 0.5, -content_height2 * 0.5);
					global_MatrixTransformer.RotateRadAppend(oMatrix, -Math.PI * 1.5);
					global_MatrixTransformer.TranslateAppend(oMatrix, _content_width * 0.5, content_height2 * 0.5);
				}
				global_MatrixTransformer.TranslateAppend(oMatrix, _transformed_text_xc - _content_width * 0.5, _transformed_text_yc - content_height2 * 0.5);

				if (this.bWordShape) {
					var DiffLeft = 0.01;
					var DiffRight = 0.01;
					var DiffLeft2, DiffRight2;
					var aContent = oContent.Content;
					for (var i = 0; i < aContent.length; ++i) {
						var oElement = aContent[i];
						if (!oElement.GetType) {
							continue;
						}
						var nElementType = oElement.GetType();
						if (nElementType === AscCommonWord.type_Paragraph) {
							var oCompiledParaPr = oElement.CompiledPr && oElement.CompiledPr.Pr && oElement.CompiledPr.Pr.ParaPr;
							if (oCompiledParaPr) {
								var oBorders = oCompiledParaPr.Brd;
								if (oBorders) {
									if (oBorders.Left && AscFormat.isRealNumber(oBorders.Left.Space) && AscFormat.isRealNumber(oBorders.Left.Size) && oBorders.Left.Size > 0.0) {
										DiffLeft2 = oBorders.Left.Space + oBorders.Left.Size;
										if (DiffLeft2 > DiffLeft) {
											DiffLeft = DiffLeft2;
										}
									}
									if (oBorders.Right && AscFormat.isRealNumber(oBorders.Right.Space) && AscFormat.isRealNumber(oBorders.Right.Size) && oBorders.Right.Size > 0.0) {
										DiffRight2 = oBorders.Right.Space + oBorders.Right.Size;
										if (oCompiledParaPr.Ind && AscFormat.isRealNumber(oCompiledParaPr.Ind.Right)) {
											DiffRight2 -= oCompiledParaPr.Ind.Right;
										}
										if (DiffRight2 > DiffRight) {
											DiffRight = DiffRight2;
										}
									}
								}
							}
						} else if (nElementType === AscCommonWord.type_Table) {
							DiffLeft2 = -oElement.GetTableOffsetCorrection();
							if (DiffLeft2 > DiffLeft) {
								DiffLeft = DiffLeft2;
							}
							DiffRight2 = oElement.GetRightTableOffsetCorrection();
							if (DiffRight2 > DiffRight) {
								DiffRight = DiffRight2;
							}
						} else if (nElementType === AscCommonWord.type_BlockLevelSdt) {

						}
					}


					var clipW = oRect.r - oRect.l + DiffLeft + DiffRight;
					if (clipW <= 0) {
						clipW = 0.01;
					}
					var clipH = oRect.b - oRect.t + Diff - b_ins - t_ins;
					if (clipH < 0) {
						clipH = 0.01;
					}
					oClipRect = {x: oRect.l - DiffLeft, y: oRect.t - Diff + t_ins, w: clipW, h: clipH};
				} else {
					var clipW = oRect.r - oRect.l + Diff - l_ins - r_ins;
					if (clipW <= 0) {
						clipW = 0.01;
					}
					var clipH = oRect.b - oRect.t + Diff - b_ins - t_ins;
					if (clipH < 0) {
						clipH = 0.01;
					}
					oClipRect = {x: oRect.l + l_ins - Diff, y: oRect.t - Diff + t_ins, w: clipW, h: clipH};

				}
			}
			return oClipRect;
		};

		CShape.prototype.setWordShape = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ShapeSetWordShape, this.bWordShape, pr));
			this.bWordShape = pr;
		};

		CShape.prototype.selectionCheck = function (X, Y, PageAbs, NearPos) {

			var content = this.getDocContent();
			if (content) {
				if (undefined !== NearPos)
					return content.CheckPosInSelection(X, Y, 0, NearPos);

				if (isRealObject(content) && this.hitInTextRect(X, Y) && this.invertTransformText) {
					var t_x = this.invertTransformText.TransformPointX(X, Y);
					var t_y = this.invertTransformText.TransformPointY(X, Y);
					return content.CheckPosInSelection(t_x, t_y, 0, NearPos);
				}
			}
			return false;
		};

		CShape.prototype.fillObject = function (copy, oPr) {
			if (this.nvSpPr)
				copy.setNvSpPr(this.nvSpPr.createDuplicate());
			if (this.spPr) {
				copy.setSpPr(this.spPr.createDuplicate());
				copy.spPr.setParent(copy);
			}
			if (this.style) {
				copy.setStyle(this.style.createDuplicate());
			}
			if (this.txBody) {
				copy.setTxBody(this.txBody.createDuplicate());
				copy.txBody.setParent(copy);
			}
			if (this.bodyPr) {
				copy.setBodyPr(this.bodyPr.createDuplicate());
			}
			if (this.textBoxContent) {
				copy.setTextBoxContent(this.textBoxContent.Copy(copy, oPr && oPr.drawingDocument, oPr && oPr.contentCopyPr));
			}
			if (this.signatureLine && copy.setSignature) {
				copy.setSignature(this.signatureLine.copy());
			}
			if (this.macro !== null) {
				copy.setMacro(this.macro);
			}
			if (this.textLink !== null) {
				copy.setTextLink(this.textLink);
			}
			if (this.clientData) {
				copy.setClientData(this.clientData.createDuplicate());
			}
			if (this.fLocksText !== null) {
				copy.setFLocksText(this.fLocksText);
			}
			copy.setWordShape(this.bWordShape);
			copy.setBDeleted(this.bDeleted);
			copy.setLocks(this.locks);
			if (!oPr || false !== oPr.cacheImage) {
				copy.cachedImage = this.getBase64Img();
				copy.cachedPixH = this.cachedPixH;
				copy.cachedPixW = this.cachedPixW;
			}
			if (this.txXfrm) {
				copy.setTxXfrm(this.txXfrm.createDuplicate());
			}
			copy.setModelId(this.modelId);
		};

		CShape.prototype.copy = function (oPr) {
			var copy = new CShape();
			this.fillObject(copy, oPr);
			return copy;
		};
		CShape.prototype.getProtectionLockText = function () {
			return this.fLocksText !== false;
		};

		CShape.prototype.canEditText = function () {
			let form = this.isForm && this.isForm() ? this.getInnerForm() : null;
			if (form && !form.CanPlaceCursorInside())
				return false;
			
			return this.superclass.prototype.canEditText.call(this);
		};
		CShape.prototype.canEditTextInSmartArt = function () {
			if (this.isObjectInSmartArt()) {
				var pointContent = this.getSmartArtPointContent();
				var shapePoint = this.getSmartArtShapePoint();
				return !!(pointContent && pointContent.length !== 0 && !shapePoint.isBlipFillPlaceholder());
			}
		}

		CShape.prototype.isPlaceholderInSmartArt = function () {
			if (this.isObjectInSmartArt()) {
				var pointContent = this.getSmartArtPointContent();
				if (pointContent && pointContent.length !== 0) {
					return pointContent.every(function (point) {
						return point && point.prSet && point.prSet.phldr;
					})
				}
			}
			return false;
		};

		CShape.prototype.getSmartArtDefaultTxFill = function () {
			if (this.isObjectInSmartArt()) {
				return this.group.group.getSmartArtDefaultTxFill(this);
			}
		};

		CShape.prototype.Get_Styles = function (level) {

			var _level = AscFormat.isRealNumber(level) ? level : 0;
			if (this.recalcInfo.recalculateTextStyles[_level]) {
				this.recalculateTextStyles(_level);
				this.recalcInfo.recalculateTextStyles[_level] = false;
			}
			this.recalcInfo.recalculateTextStyles[_level] = true;
			var ret = this.compiledStyles[_level];
			this.compiledStyles[_level] = undefined;
			return ret;
			//   return this.compiledStyles[_level];
		};


		CShape.prototype.recalculateTextStyles = function (level) {
			return AscFormat.ExecuteNoHistory(function () {
				var parent_objects = this.getParentObjects();
				var default_style = new CStyle("defaultStyle", null, null, null, true);
				default_style.ParaPr.Spacing.LineRule = Asc.linerule_Auto;
				default_style.ParaPr.Spacing.Line = 1;
				default_style.ParaPr.Spacing.Before = 0;
				default_style.ParaPr.Spacing.After = 0;
				default_style.ParaPr.DefaultTab = 25.4;
				default_style.ParaPr.Align = AscCommon.align_Center;
				if (parent_objects.theme) {
					default_style.TextPr.RFonts.SetFontStyle(AscFormat.fntStyleInd_minor);
				}
				if (!this.bCheckAutoFitFlag) {
					var oBodyPr = this.getBodyPr && this.getBodyPr();
					if (oBodyPr) {
						default_style.ParaPr.LnSpcReduction = oBodyPr.getLnSpcReduction();
						default_style.TextPr.FontScale = oBodyPr.getFontScale();
					}
				} else {
					if (this.tmpLnSpcReduction !== null && this.tmpLnSpcReduction !== undefined) {
						default_style.ParaPr.LnSpcReduction = this.tmpLnSpcReduction / 100000.0;
					}
					if (this.tmpFontScale !== null && this.tmpFontScale !== undefined) {
						default_style.TextPr.FontScale = this.tmpFontScale / 100000.0;
					}

				}
				if (this.getObjectType && this.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
					default_style.TextPr.FontSize = 18;
				}
				if (isRealObject(parent_objects.presentation) && isRealObject(parent_objects.presentation.defaultTextStyle)) {

					if (isRealObject(parent_objects.presentation.defaultTextStyle.levels[9])) {
						var default_ppt_style = parent_objects.presentation.defaultTextStyle.levels[9];
						default_style.ParaPr.Merge(default_ppt_style.Copy());
						default_ppt_style.DefaultRunPr && default_style.TextPr.Merge(default_ppt_style.DefaultRunPr.Copy());
					}
					if (!isRealObject(parent_objects.master) || !isRealObject(parent_objects.master.txStyles) || !this.isPlaceholder()) {
						if (isRealObject(parent_objects.presentation.defaultTextStyle.levels[level])) {
							var default_ppt_style = parent_objects.presentation.defaultTextStyle.levels[level];
							default_style.ParaPr.Merge(default_ppt_style.Copy());
							default_ppt_style.DefaultRunPr && default_style.TextPr.Merge(default_ppt_style.DefaultRunPr.Copy());
						}
					}
				}

				var master_style;
				if (isRealObject(parent_objects.master) && isRealObject(parent_objects.master.txStyles)) {
					var master_ppt_styles;
					master_style = new CStyle("masterStyle", null, null, null, true);
					if (parent_objects.master.kind === AscFormat.TYPE_KIND.NOTES_MASTER) {
						master_ppt_styles = parent_objects.master.txStyles;
					} else {
						if (this.isPlaceholder() && !(this instanceof AscFormat.CGraphicFrame)) {
							master_ppt_styles = parent_objects.master.txStyles.getStyleByPhType(this.getPlaceholderType());
						} else {
							master_ppt_styles = parent_objects.master.txStyles.otherStyle;
						}
					}

					if (isRealObject(master_ppt_styles) && isRealObject(master_ppt_styles.levels) && isRealObject(master_ppt_styles.levels[level])) {
						var master_ppt_style = master_ppt_styles.levels[level];
						master_style.ParaPr = master_ppt_style.Copy();
						if (master_ppt_style.DefaultRunPr) {
							master_style.TextPr = master_ppt_style.DefaultRunPr.Copy();
						}
					}
				}

				var hierarchy = this.getHierarchy(false);
				var hierarchy_styles = [];
				for (var i = 0; i < hierarchy.length; ++i) {
					var hierarchy_shape = hierarchy[i];
					if (isRealObject(hierarchy_shape)
						&& isRealObject(hierarchy_shape.txBody)
						&& isRealObject(hierarchy_shape.txBody.lstStyle)
						&& isRealObject(hierarchy_shape.txBody.lstStyle.levels)
						&& isRealObject(hierarchy_shape.txBody.lstStyle.levels[level])) {
						var hierarchy_ppt_style = hierarchy_shape.txBody.lstStyle.levels[level];
						var hierarchy_style = new CStyle("hierarchyStyle" + i, null, null, null, true);
						hierarchy_style.ParaPr = hierarchy_ppt_style.Copy();
						if (hierarchy_ppt_style.DefaultRunPr) {
							hierarchy_style.TextPr = hierarchy_ppt_style.DefaultRunPr.Copy();
						}
						hierarchy_styles.push(hierarchy_style);
					}
				}

				var ownStyle;
				if (isRealObject(this.txBody) && isRealObject(this.txBody.lstStyle) && isRealObject(this.txBody.lstStyle.levels[level])) {
					ownStyle = new CStyle("ownStyle", null, null, null, true);
					var own_ppt_style = this.txBody.lstStyle.levels[level];
					ownStyle.ParaPr = own_ppt_style.Copy();
					if (own_ppt_style.DefaultRunPr) {
						ownStyle.TextPr = own_ppt_style.DefaultRunPr.Copy();
					}
					hierarchy_styles.splice(0, 0, ownStyle);
				}
				var shape_text_style;
				var compiled_style = this.getCompiledStyle && this.getCompiledStyle();
				if (isRealObject(compiled_style) && isRealObject(compiled_style.fontRef)) {
					shape_text_style = new CStyle("shapeTextStyle", null, null, null, true);
					shape_text_style.TextPr.RFonts.SetFontStyle(compiled_style.fontRef.idx);
					if (compiled_style.fontRef.Color && compiled_style.fontRef.Color.isCorrect()) {
						shape_text_style.TextPr.Unifill = AscFormat.CreateUniFillByUniColor(compiled_style.fontRef.Color);
					}
					var smartArtTxFill = this.getSmartArtDefaultTxFill();
					if (smartArtTxFill) {
						shape_text_style.TextPr.Unifill = smartArtTxFill;
					}
				}
				var Styles = new CStyles(false);


				var last_style_id;
				var b_checked = false;
				var isPlaceholder = this.isPlaceholder();
				if (isPlaceholder || this.graphicObject instanceof CTable) {
					if (default_style) {
						//checkTextPr(default_style.TextPr);
						b_checked = true;
						Styles.Add(default_style);
						default_style.BasedOn = null;
						last_style_id = default_style.Id;
					}

					if (master_style) {
						//checkTextPr(master_style.TextPr);
						Styles.Add(master_style);
						master_style.BasedOn = last_style_id;
						last_style_id = master_style.Id;
					}
				} else {
					if (master_style) {
						// checkTextPr(master_style.TextPr);
						b_checked = true;
						Styles.Add(master_style);
						master_style.BasedOn = null;
						last_style_id = master_style.Id;
					}

					if (default_style) {
						//checkTextPr(default_style.TextPr);
						Styles.Add(default_style);
						default_style.BasedOn = last_style_id;
						last_style_id = default_style.Id;
					}
				}

				for (var i = hierarchy_styles.length - 1; i > -1; --i) {

					if (hierarchy_styles[i]) {
						//checkTextPr(hierarchy_styles[i].TextPr);
						Styles.Add(hierarchy_styles[i]);
						hierarchy_styles[i].BasedOn = last_style_id;
						last_style_id = hierarchy_styles[i].Id;
					}
				}


				if (shape_text_style) {
					//checkTextPr(shape_text_style.TextPr);
					Styles.Add(shape_text_style);
					shape_text_style.BasedOn = last_style_id;
					last_style_id = shape_text_style.Id;
				}

				this.compiledStyles[level] = {
					styles: Styles,
					lastId: last_style_id,
					shape: this,
					slide: parent_objects.slide,
					layout: parent_objects.layout,
					master: parent_objects.master,
					presentation: parent_objects.presentation,
					notes: parent_objects.notes
				};
				return this.compiledStyles[level];
			}, this, []);
		};

		CShape.prototype.recalculateBrush = function () {
			var compiled_style = this.getCompiledStyle();
			var RGBA = {R: 0, G: 0, B: 0, A: 255};
			var parents = this.getParentObjects();
			var oStyleBrush = null;
			var oLin;
			if (isRealObject(parents.theme) && isRealObject(compiled_style) && isRealObject(compiled_style.fillRef)) {
				this.brush = parents.theme.getFillStyle(compiled_style.fillRef.idx, compiled_style.fillRef.Color);
				if (this.brush) {
					oStyleBrush = this.brush.createDuplicate();
				}
			} else {
				this.brush = new AscFormat.CUniFill();
			}

			this.brush.merge(this.getCompiledFill());
			this.brush.transparent = this.getCompiledTransparent();
			this.brush.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
			if (this.brush.fill && this.brush.fill.type === Asc.c_oAscFill.FILL_TYPE_GRAD) {
				var oGradFill = this.brush.fill;
				if (!oGradFill.lin && !oGradFill.path) {
					if (oStyleBrush &&
						oStyleBrush.fill &&
						oStyleBrush.fill.type === Asc.c_oAscFill.FILL_TYPE_GRAD &&
						oStyleBrush.fill.lin) {
						oLin = oStyleBrush.fill.lin.createDuplicate();
					} else {
						oLin = new AscFormat.GradLin();
						oLin.setScale(false);
						oLin.setAngle(0);
						oGradFill.setLin(oLin);
					}
					oGradFill.setLin(oLin);
				}
			}
		};

		CShape.prototype.recalculatePen = function () {
			var compiled_style = this.getCompiledStyle();
			var RGBA = {R: 0, G: 0, B: 0, A: 255};
			var parents = this.getParentObjects();
			if (isRealObject(parents.theme) && isRealObject(compiled_style) && isRealObject(compiled_style.lnRef)) {
				//compiled_style.lnRef.Color.Calculate(parents.theme, parents.slide, parents.layout, parents.master, {R: 0, G: 0, B: 0, A:255});
				//RGBA = compiled_style.lnRef.Color.RGBA;
				this.pen = parents.theme.getLnStyle(compiled_style.lnRef.idx, compiled_style.lnRef.Color);
				//if (isRealObject(this.pen)) {
				//    if (isRealObject(compiled_style.lnRef.Color.color)
				//        && isRealObject(this.pen)
				//        && isRealObject(this.pen.Fill)
				//        && isRealObject(this.pen.Fill.fill)
				//        && this.pen.Fill.fill.type === FILL_TYPE_SOLID) {
				//        this.pen.Fill.fill.color = compiled_style.lnRef.Color.createDuplicate();
				//    }
				//}
				//else
				//{
				//    this.pen = new AscFormat.CLn();
				//}
			} else {
				this.pen = null;
			}

			var oCompiledLine = this.getCompiledLine();
			if (oCompiledLine) {
				if (!this.pen) {
					this.pen = new AscFormat.CLn();
				}
				this.pen.merge(oCompiledLine);
			}
			if (this.pen) {
				this.pen.calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
			}
		};

		CShape.prototype.Get_ParentTextTransform = function () {
			return this.transformText.CreateDublicate();
		};
		CShape.prototype.canAddButtonPlaceholder = function () {
			return (this.parent && (this.parent.getObjectType() === AscDFH.historyitem_type_Slide) ||
				this.isObjectInSmartArt());
		};
		CShape.prototype.isEmptyPlaceholder = function (bDefaultEmpty) {
			if (this.isObjectInSmartArt()) {
				if (this.isPlaceholderInSmartArt()) {
					if (this.txBody) {
						if (this.txBody.content) {
							return this.txBody.content.Is_Empty(bDefaultEmpty);
						}
						return true;
					}
				} else if (this.isActiveBlipFillPlaceholder()) {
					return true;
				}
			}
			if (this.isPlaceholder()) {
				var phldrType = this.getPhType();
				if (phldrType == AscFormat.phType_title
					|| phldrType == AscFormat.phType_ctrTitle
					|| phldrType == AscFormat.phType_body
					|| phldrType == AscFormat.phType_subTitle
					|| phldrType == null
					|| phldrType == AscFormat.phType_dt
					|| phldrType == AscFormat.phType_ftr
					|| phldrType == AscFormat.phType_hdr
					|| phldrType == AscFormat.phType_sldNum
					|| phldrType == AscFormat.phType_sldImg) {
					if (this.txBody) {
						if (this.txBody.content) {
							return this.txBody.content.Is_Empty();
						}
						return true;
					}
					return true;
				}
				if (phldrType == AscFormat.phType_chart
					|| phldrType == AscFormat.phType_media) {
					return true;
				}
				if (phldrType == AscFormat.phType_pic) {
					var _b_empty_text = true;
					if (this.txBody) {
						if (this.txBody.content) {
							_b_empty_text = this.txBody.content.Is_Empty();
						}
					}
					return (_b_empty_text /* && (this.brush == null || this.brush.fill == null)*/);
				}
			} else {
				return false;
			}
		};


		CShape.prototype.changeSize = function (kw, kh) {
			if (this.spPr && this.spPr.xfrm && this.spPr.xfrm.isNotNull()) {
				var xfrm = this.spPr.xfrm;
				// if(this.getNoChangeAspect()){
				//     var k = Math.min(kw, kh);
				//     var oldXC = xfrm.offX + xfrm.extX/2.0;
				//     var oldYC = xfrm.offY + xfrm.extY/2.0;
				//     xfrm.setExtX(xfrm.extX * k);
				//     xfrm.setExtY(xfrm.extY * k);
				//     xfrm.setOffX(oldXC * kw - xfrm.extX/2.0);
				//     xfrm.setOffY(oldYC * kh - xfrm.extY/2.0);
				// }
				// else
				{

					xfrm.setOffX(xfrm.offX * kw);
					xfrm.setOffY(xfrm.offY * kh);
					xfrm.setExtX(xfrm.extX * kw);
					xfrm.setExtY(xfrm.extY * kh);
				}
			}
			var txXfrm = this.txXfrm;
			if (txXfrm && txXfrm.isNotNull()) {

				txXfrm.setOffX(txXfrm.offX * kw);
				txXfrm.setOffY(txXfrm.offY * kh);
				txXfrm.setExtX(txXfrm.extX * kw);
				txXfrm.setExtY(txXfrm.extY * kh);

			}
			this.recalcTransform && this.recalcTransform();
		};

		CShape.prototype.recalculateTransform = function () {
			this.cachedImage = null;
			this.recalculateLocalTransform(this.transform);
			this.invertTransform = global_MatrixTransformer.Invert(this.transform);
			this.localTransform = this.transform.CreateDublicate();
		};


		CShape.prototype.checkAutofit = function (bIgnoreWordShape) {
			if (this.bWordShape || bIgnoreWordShape || this.bCheckAutoFitFlag) {
				var content = this.getDocContent();
				if (content) {
					var oBodyPr = this.getBodyPr();
					if (oBodyPr.textFit && oBodyPr.textFit.type === AscFormat.text_fit_Auto || oBodyPr.wrap === AscFormat.nTWTNone) {
						return true;
					}
				}
			}
			return false;
		};

		CShape.prototype.Check_AutoFit = function () {
			return this.checkAutofit(true) || this.checkContentWordArt(this.getDocContent()) || this.getBodyPr().prstTxWarp != null;
		};

		CShape.prototype.recalculateLocalTransform = function (transform) {
			AscFormat.ExecuteNoHistory(function () {
				var bNotesShape = false;

				let oParaDrawing = getParaDrawing(this);
				if (!isRealObject(this.group)) {
					var bUserShape = false;
					if (this.parent instanceof AscFormat.CRelSizeAnchor || this.parent instanceof AscFormat.CAbsSizeAnchor) {
						if (this.parent.parent instanceof AscFormat.CChartSpace) {
							this.x = this.parent.parent.extX * this.parent.fromX;
							this.y = this.parent.parent.extY * this.parent.fromY;
							if (this.parent instanceof AscFormat.CRelSizeAnchor) {
								this.extX = Math.max(0.0, this.parent.parent.extX * this.parent.toX - this.x);
								this.extY = Math.max(0.0, this.parent.parent.extY * this.parent.toY - this.y);
							} else {
								this.extX = Math.max(0.0, this.parent.toX);
								this.extY = Math.max(0.0, this.parent.toY);
							}
							var rot = 0;
							if (this.spPr && this.spPr.xfrm) {
								if (AscFormat.isRealNumber(this.spPr.xfrm.rot)) {
									rot = AscFormat.normalizeRotate(this.spPr.xfrm.rot);
								}
								this.flipH = this.spPr.xfrm.flipH === true;
								this.flipV = this.spPr.xfrm.flipV === true;
							}
							this.rot = rot;
							bUserShape = true;
						}
					}
					if (bUserShape) {
					} else if (this.drawingBase && !this.isCrop) {
						var metrics = this.drawingBase.getGraphicObjectMetrics();
						this.x = metrics.x;
						this.y = metrics.y;
						var rot = 0;
						if (this.spPr && this.spPr.xfrm) {
							if (AscFormat.isRealNumber(this.spPr.xfrm.rot)) {
								rot = AscFormat.normalizeRotate(this.spPr.xfrm.rot);
							}
							this.flipH = this.spPr.xfrm.flipH === true;
							this.flipV = this.spPr.xfrm.flipV === true;
						}
						this.rot = rot;

						var metricExtX, metricExtY;
						//  if(!(this instanceof AscFormat.CGroupShape))
						{
							metricExtX = metrics.extX;
							metricExtY = metrics.extY;
							if (checkNormalRotate(rot)) {
								this.extX = metrics.extX;
								this.extY = metrics.extY;
							} else {
								this.extX = metrics.extY;
								this.extY = metrics.extX;
							}
						}
						// else
						// {
						//     if(this.spPr && this.spPr.xfrm && AscFormat.isRealNumber(this.spPr.xfrm.extX) && AscFormat.isRealNumber(this.spPr.xfrm.extY))
						//     {
						//         this.extX = this.spPr.xfrm.extX;
						//         this.extY = this.spPr.xfrm.extY;
						//     }
						//     else
						//     {
						//         metricExtX = metrics.extX;
						//         metricExtY = metrics.extY;
						//     }
						// }

						if (checkNormalRotate(rot)) {
							this.x = metrics.x;
							this.y = metrics.y;
						} else {
							this.x = metrics.x + metricExtX / 2 - metricExtY / 2;
							this.y = metrics.y + metricExtY / 2 - metricExtX / 2;
						}
					} else if (typeof AscCommonSlide !== "undefined" && AscCommonSlide
						&& AscCommonSlide.CNotes && this.parent && this.parent instanceof AscCommonSlide.CNotes) {
						bNotesShape = true;
						this.x = 0;
						this.y = editor.WordControl.m_oLogicDocument.GetHeightMM();
						this.extX = this.parent.getWidth();
						this.extY = 2000;
						this.rot = 0;
						this.flipH = false;
						this.flipV = false;
					} else if (this.spPr && this.spPr.xfrm && this.spPr.xfrm.isNotNull()) {
						var xfrm = this.spPr.xfrm;
						var bAlign = false;
						if (this.parent) {
							if (this.parent.PositionH && this.parent.PositionH.Align
								|| this.parent.PositionV && this.parent.PositionV.Align) {
								bAlign = true;
							}
						}
						if (bAlign) {
							this.x = 0;
							this.y = 0;
						} else {
							this.x = xfrm.offX;
							this.y = xfrm.offY;
						}
						this.extX = xfrm.extX;
						this.extY = xfrm.extY;
						this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
						this.flipH = xfrm.flipH === true;
						this.flipV = xfrm.flipV === true;

						if (oParaDrawing) {
							if (oParaDrawing.Extent && AscFormat.isRealNumber(oParaDrawing.Extent.W) && AscFormat.isRealNumber(oParaDrawing.Extent.H)) {
								let dScaleCoefficient = this.getScaleCoefficient();
								this.extX = oParaDrawing.Extent.W * dScaleCoefficient;
								this.extY = oParaDrawing.Extent.H * dScaleCoefficient;
							}
							if (oParaDrawing.SizeRelH || oParaDrawing.SizeRelV) {
								this.m_oSectPr = null;
								var oParentParagraph = oParaDrawing.Get_ParentParagraph();
								if (oParentParagraph) {

									var oSectPr = oParentParagraph.Get_SectPr();
									if (oSectPr) {
										if (oParaDrawing.SizeRelH && oParaDrawing.SizeRelH.Percent > 0) {
											switch (oParaDrawing.SizeRelH.RelativeFrom) {
												case c_oAscSizeRelFromH.sizerelfromhMargin: {
													this.extX = oSectPr.GetContentFrameWidth();
													break;
												}
												case c_oAscSizeRelFromH.sizerelfromhPage: {
													this.extX = oSectPr.GetPageWidth();
													break;
												}
												case c_oAscSizeRelFromH.sizerelfromhLeftMargin: {
													this.extX = oSectPr.GetPageMarginLeft();
													break;
												}

												case c_oAscSizeRelFromH.sizerelfromhRightMargin: {
													this.extX = oSectPr.GetPageMarginRight();
													break;
												}
												default: {
													this.extX = oSectPr.GetPageMarginLeft();
													break;
												}
											}
											this.extX *= oParaDrawing.SizeRelH.Percent;
										}
										if (oParaDrawing.SizeRelV && oParaDrawing.SizeRelV.Percent > 0) {
											switch (oParaDrawing.SizeRelV.RelativeFrom) {
												case c_oAscSizeRelFromV.sizerelfromvMargin: {
													this.extY = oSectPr.GetContentFrameHeight();
													break;
												}
												case c_oAscSizeRelFromV.sizerelfromvPage: {
													this.extY = oSectPr.GetPageHeight();
													break;
												}
												case c_oAscSizeRelFromV.sizerelfromvTopMargin: {
													this.extY = oSectPr.GetPageMarginTop();
													break;
												}
												case c_oAscSizeRelFromV.sizerelfromvBottomMargin: {
													this.extY = oSectPr.GetPageMarginBottom();
													break;
												}
												default: {
													this.extY = oSectPr.GetPageMarginTop();
													break;
												}
											}
											this.extY *= oParaDrawing.SizeRelV.Percent;
										}
										this.m_oSectPr = new CSectionPr();
										this.m_oSectPr.Copy(oSectPr);
									}
								}
							}
						}
					} else {
						if (this.isPlaceholder()) {
							var hierarchy = this.getHierarchy();
							for (var i = 0; i < hierarchy.length; ++i) {
								var hierarchy_sp = hierarchy[i];
								if (isRealObject(hierarchy_sp) && hierarchy_sp.spPr.xfrm && hierarchy_sp.spPr.xfrm.isNotNull()) {
									var xfrm = hierarchy_sp.spPr.xfrm;
									this.x = xfrm.offX;
									this.y = xfrm.offY;
									this.extX = xfrm.extX;
									this.extY = xfrm.extY;
									this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
									this.flipH = xfrm.flipH === true;
									this.flipV = xfrm.flipV === true;
									break;
								}
							}
							if (i === hierarchy.length) {
								this.x = 0;
								this.y = 0;
								this.extX = 5;
								this.extY = 5;
								this.rot = 0;
								this.flipH = false;
								this.flipV = false;
							}
						} else {
							var extX, extY;
							if (oParaDrawing && oParaDrawing.Extent) {
								this.x = 0;
								this.y = 0;
								let dScaleCoefficient = oParaDrawing.GetScaleCoefficient();
								extX = oParaDrawing.Extent.W * dScaleCoefficient;
								extY = oParaDrawing.Extent.H * dScaleCoefficient;
							} else {
								this.x = 0;
								this.y = 0;
								extX = 5;
								extY = 5;
							}
							this.extX = extX;
							this.extY = extY;
							this.rot = 0;
							this.flipH = false;
							this.flipV = false;
						}
					}
				} else {
					var xfrm;
					if (this.spPr && this.spPr.xfrm && this.spPr.xfrm.isNotNull()) {
						xfrm = this.spPr.xfrm;
					} else {
						if (this.isPlaceholder()) {
							var hierarchy = this.getHierarchy();
							for (var i = 0; i < hierarchy.length; ++i) {
								var hierarchy_sp = hierarchy[i];
								if (isRealObject(hierarchy_sp) && hierarchy_sp.spPr.xfrm.isNotNull()) {
									xfrm = hierarchy_sp.spPr.xfrm;
									break;
								}
							}
							if (i === hierarchy.length) {
								xfrm = new AscFormat.CXfrm();
								xfrm.offX = 0;
								xfrm.offX = 0;
								xfrm.extX = 5;
								xfrm.extY = 5;
							}
						} else {
							xfrm = new AscFormat.CXfrm();
							xfrm.offX = 0;
							xfrm.offY = 0;
							xfrm.extX = 5;
							xfrm.extY = 5;
						}
					}


					var scale_scale_coefficients = this.group.getResultScaleCoefficients();
					this.x = scale_scale_coefficients.cx * (xfrm.offX - this.group.spPr.xfrm.chOffX);
					this.y = scale_scale_coefficients.cy * (xfrm.offY - this.group.spPr.xfrm.chOffY);
					this.extX = scale_scale_coefficients.cx * xfrm.extX;
					this.extY = scale_scale_coefficients.cy * xfrm.extY;
					this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
					this.flipH = xfrm.flipH === true;
					this.flipV = xfrm.flipV === true;
				}


				if (this.checkAutofit && this.checkAutofit() && (!this.bWordShape || !this.group || this.bCheckAutoFitFlag) && !bNotesShape) {
					var oBodyPr = this.getBodyPr();
					if (this.bWordShape) {
						if (this.recalcInfo.recalculateTxBoxContent) {
							this.recalcInfo.oContentMetrics = this.recalculateTxBoxContent();
							//this.recalcInfo.recalculateTxBoxContent = false;
							this.recalcInfo.AllDrawings = [];
							var oContent = this.getDocContent();
							if (oContent) {
								oContent.GetAllDrawingObjects(this.recalcInfo.AllDrawings);
							}
						}
					} else {
						if (this.recalcInfo.recalculateContent) {
							this.recalcInfo.oContentMetrics = this.recalculateContent();
							this.recalcInfo.recalculateContent = false;
						}
					}
					var oContentMetrics = this.recalcInfo.oContentMetrics;

					var l_ins, t_ins, r_ins, b_ins;
					if (oBodyPr) {
						l_ins = AscFormat.isRealNumber(oBodyPr.lIns) ? oBodyPr.lIns : 2.54;
						r_ins = AscFormat.isRealNumber(oBodyPr.rIns) ? oBodyPr.rIns : 2.54;
						t_ins = AscFormat.isRealNumber(oBodyPr.tIns) ? oBodyPr.tIns : 1.27;
						b_ins = AscFormat.isRealNumber(oBodyPr.bIns) ? oBodyPr.bIns : 1.27;
					} else {
						l_ins = 2.54;
						r_ins = 2.54;
						t_ins = 1.27;
						b_ins = 1.27;
					}
					var oGeometry = this.spPr && this.spPr.geometry, oWH;
					var dOldExtX = this.extX, dOldExtY = this.extY, dDeltaX = 0, dDeltaY = 0;


					var bAutoFit = AscCommon.isRealObject(oBodyPr.textFit) && oBodyPr.textFit.type === AscFormat.text_fit_Auto;
					if (oBodyPr.wrap === AscFormat.nTWTNone) {
						if (!oBodyPr.upright) {
							if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
								if (oGeometry) {
									oWH = oGeometry.getNewWHByTextRect(oContentMetrics.w + l_ins + r_ins, oContentMetrics.contentH + t_ins + b_ins, undefined, bAutoFit ? undefined : this.extY);
									if (!oWH.bError) {
										this.extX = oWH.W;
										this.extY = oWH.H;
									}
								} else {
									this.extX = oContentMetrics.w + l_ins + r_ins;
									this.extY = bAutoFit ? oContentMetrics.contentH + t_ins + b_ins : this.extY;
								}

							} else {
								if (oGeometry) {
									oWH = oGeometry.getNewWHByTextRect(oContentMetrics.contentH + l_ins + r_ins, oContentMetrics.w + t_ins + b_ins, bAutoFit ? undefined : this.extX);
									if (!oWH.bError) {
										this.extX = oWH.W;
										this.extY = oWH.H;
									}
								} else {
									this.extY = oContentMetrics.w + t_ins + b_ins;
									this.extX = bAutoFit ? oContentMetrics.contentH + l_ins + r_ins : this.extX;
								}
							}
						} else {
							var _full_rotate = this.getFullRotate();
							if (checkNormalRotate(_full_rotate)) {
								if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {

									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(oContentMetrics.w + l_ins + r_ins, oContentMetrics.contentH + t_ins + b_ins, undefined, bAutoFit ? undefined : this.extY);
										if (!oWH.bError) {
											this.extX = oWH.W;
											this.extY = oWH.H;
										}
									} else {
										this.extX = oContentMetrics.w + l_ins + r_ins;
										this.extY = bAutoFit ? oContentMetrics.contentH + t_ins + b_ins : this.extY;
									}
								} else {
									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(oContentMetrics.contentH + l_ins + r_ins, oContentMetrics.w + t_ins + b_ins, bAutoFit ? undefined : this.extX);
										if (!oWH.bError) {
											this.extX = oWH.W;
											this.extY = oWH.H;
										}
									} else {
										this.extY = oContentMetrics.w + t_ins + b_ins;
										this.extX = bAutoFit ? oContentMetrics.contentH + l_ins + r_ins : this.extX;
									}

								}
							} else {
								if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(oContentMetrics.w + l_ins + r_ins, oContentMetrics.contentH + t_ins + b_ins, undefined, bAutoFit ? undefined : this.extY);
										if (!oWH.bError) {
											this.extX = oWH.W;
											this.extY = oWH.H;
										}
									} else {
										this.extX = oContentMetrics.w + l_ins + r_ins;
										this.extY = bAutoFit ? oContentMetrics.contentH + t_ins + b_ins : this.extY;
									}
								} else {
									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(oContentMetrics.contentH + l_ins + r_ins, oContentMetrics.w + t_ins + b_ins, bAutoFit ? undefined : this.extX);
										if (!oWH.bError) {
											this.extX = oWH.W;
											this.extY = oWH.H;
										}
									} else {
										this.extY = oContentMetrics.w + t_ins + b_ins;
										this.extX = bAutoFit ? oContentMetrics.contentH + l_ins + r_ins : this.extX;
									}
								}
							}
						}
					} else {
						if (!oBodyPr.upright) {
							if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
								if (oGeometry) {
									oWH = oGeometry.getNewWHByTextRect(undefined, oContentMetrics.contentH + t_ins + b_ins, this.extX, undefined);
									if (!oWH.bError) {
										this.extY = oWH.H;
									}
								} else {
									this.extY = oContentMetrics.contentH + t_ins + b_ins;
								}
							} else {
								if (oGeometry) {
									oWH = oGeometry.getNewWHByTextRect(oContentMetrics.contentH + l_ins + b_ins, undefined, undefined, this.extY);
									if (!oWH.bError) {
										this.extX = oWH.W;
									}
								} else {
									this.extX = oContentMetrics.contentH + l_ins + r_ins;
								}
							}
						} else {
							var _full_rotate = this.getFullRotate();
							if (checkNormalRotate(_full_rotate)) {
								if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(undefined, oContentMetrics.contentH + t_ins + b_ins, this.extX, undefined);
										if (!oWH.bError) {
											this.extY = oWH.H;
										}
									} else {
										this.extY = oContentMetrics.contentH + t_ins + b_ins;
									}
								} else {
									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(oContentMetrics.contentH + l_ins + r_ins, undefined, undefined, this.extY);
										if (!oWH.bError) {
											this.extX = oWH.W;
										}
									} else {
										this.extX = oContentMetrics.contentH + l_ins + r_ins;
									}
								}
							} else {
								if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(oContentMetrics.contentH + l_ins + r_ins, undefined, undefined, this.extY);
										if (!oWH.bError) {
											this.extX = oWH.W;
										}
									} else {
										this.extX = oContentMetrics.contentH + l_ins + r_ins;
									}
								} else {
									if (oGeometry) {
										oWH = oGeometry.getNewWHByTextRect(undefined, oContentMetrics.contentH + t_ins + b_ins, this.extX, undefined);
										if (!oWH.bError) {
											this.extY = oWH.H;
										}
									} else {
										this.extY = oContentMetrics.contentH + t_ins + b_ins;
									}
								}
							}
						}
					}

					if (!this.bWordShape || this.group)//в презентациях и в таблицах изменям позицию: по горизонтали - в зависимости от выравнивания первого параграфа в контенте,
						// по вертикали - в зависимости от вертикального выравнивания контента.
					{
						var dSin = Math.sin(this.rot), dCos = Math.cos(this.rot);
						var oContent = this.getDocContent();
						var nJc, nAnchor;
						if (AscFormat.isRealNumber(this.rot) && !AscFormat.fApproxEqual(this.rot, 0)) {
							nJc = AscCommon.align_Center;
							nAnchor = 1;
						} else {
							nJc = oContent.Content[0].CompiledPr.Pr.ParaPr.Jc;
							nAnchor = oBodyPr.anchor;
						}
						var FreezePointX, FreezePointY;
						switch (nJc) {
							case AscCommon.align_Right: {
								dDeltaX = dOldExtX - this.extX;
								FreezePointX = oContent.XLimit;
								break;
							}
							case AscCommon.align_Left: {
								dDeltaX = 0;
								FreezePointX = 0;
								break;
							}
							default: {
								dDeltaX = (dOldExtX - this.extX) / 2;
								FreezePointX = oContent.XLimit / 2.0;
								break;
							}
						}

						switch (nAnchor) {
							case 0: //b
							{
								dDeltaY = dOldExtY - this.extY;
								FreezePointY = oContent.GetSummaryHeight();
								break;
							}
							case 1:    //ctr
							case 2: //dist
							case 3: //just
							{// (Text Anchor Enum ( Center ))
								dDeltaY = (dOldExtY - this.extY) / 2;
								FreezePointY = oContent.GetSummaryHeight() / 2.0;
								break;
							}
							default: {
								FreezePointY = 0.0;
								break;
							}
						}
						// var tx1 = this.localTransformText.TransformPointX(FreezePointX, FreezePointY);
						// var ty1 = this.localTransformText.TransformPointY(FreezePointX, FreezePointY);
						// this.recalculateTransformText();
						// var tx2 = this.localTransformText.TransformPointX(FreezePointX, FreezePointY);
						// var ty2 = this.localTransformText.TransformPointY(FreezePointX, FreezePointY);


						var dTrDeltaX, dTrDeltaY;
						// if(this.invertTransform)
						// {
						//     var oInvMatrix = this.invertTransform.CreateDublicate();
						//     oInvMatrix.tx = 0.0;
						//     oInvMatrix.ty = 0.0;
						//     dTrDeltaX = oInvMatrix.TransformPointX(dDeltaX, dDeltaY);
						//     dTrDeltaY = oInvMatrix.TransformPointY(dDeltaX, dDeltaY);
						// }
						// else
						{
							dTrDeltaX = dDeltaX;
							dTrDeltaY = dDeltaY;
						}
						this.x += dTrDeltaX;
						this.y += dTrDeltaY;
					}
				}
				this.localX = this.x;
				this.localY = this.y;
				transform.Reset();
				var hc = this.extX * 0.5;
				var vc = this.extY * 0.5;
				global_MatrixTransformer.TranslateAppend(transform, -hc, -vc);
				if (this.flipH)
					global_MatrixTransformer.ScaleAppend(transform, -1, 1);
				if (this.flipV)
					global_MatrixTransformer.ScaleAppend(transform, 1, -1);
				global_MatrixTransformer.RotateRadAppend(transform, -this.rot);
				global_MatrixTransformer.TranslateAppend(transform, this.x + hc, this.y + vc);
				if (isRealObject(this.group)) {
					global_MatrixTransformer.MultiplyAppend(transform, this.group.getLocalTransform());
				}
				if (this.parent instanceof AscFormat.CRelSizeAnchor || this.parent instanceof AscFormat.CAbsSizeAnchor) {
					if (this.parent.parent instanceof AscFormat.CChartSpace) {
						if (this.parent.parent.recalcInfo.recalculateTransform) {
							this.parent.parent.recalculateTransform();
							this.parent.parent.rectGeometry.Recalculate(this.parent.parent.extX, this.parent.parent.extY);
							this.parent.parent.recalcInfo.recalculateTransform = false;
						}
						global_MatrixTransformer.MultiplyAppend(transform, this.parent.parent.localTransform);
					}
				}
				if (oParaDrawing) {
					this.m_oSectPr = null;
					var oParentParagraph = oParaDrawing.Get_ParentParagraph();
					if (oParentParagraph) {
						var oSectPr = oParentParagraph.Get_SectPr();
						if (oSectPr) {
							this.m_oSectPr = new CSectionPr();
							this.m_oSectPr.Copy(oSectPr);
						}
					}
				}
				this.localTransform = transform;
				this.transform = transform;
			}, this, []);
		};

		CShape.prototype.CheckNeedRecalcAutoFit = function (oSectPr) {
			var Width, Height, Width2, Height2;
			var bRet = false;
			var oParaDrawing = getParaDrawing(this);
			var bSizRel = (oParaDrawing && (oParaDrawing.SizeRelH || oParaDrawing.SizeRelV));
			if (this.checkAutofit() || bSizRel) {
				if (oSectPr) {
					if (!this.m_oSectPr) {
						this.recalcBounds();
						this.recalcText();
						this.recalcGeometry();
						if (bSizRel) {
							this.recalcTransform();
						}
						bRet = true;
					} else {
						Width = oSectPr.GetContentFrameWidth();
						Height = oSectPr.GetContentFrameHeight();

						Width2 = this.m_oSectPr.GetContentFrameWidth();
						Height2 = this.m_oSectPr.GetContentFrameHeight();

						bRet = (Math.abs(Width - Width2) > 0.001 || Math.abs(Height - Height2) > 0.001);
						if (bRet) {
							this.recalcBounds();
							this.recalcText();
							this.recalcGeometry();
							if (bSizRel) {
								this.recalcTransform();
							}
						}
						return bRet;
					}
				} else {
					if (this.m_oSectPr) {
						this.recalcBounds();
						this.recalcText();
						this.recalcGeometry();
						bRet = true;
					}
				}
			}
			return bRet;
		};


		CShape.prototype.recalculateDocContent = function (oDocContent, oBodyPr) {
			var nStartPage = this.Get_AbsolutePage ? this.Get_AbsolutePage() : 0;
			var oRet = {w: 0, h: 0, contentH: 0};
			var l_ins, t_ins, r_ins, b_ins;
			if (oBodyPr) {
				l_ins = AscFormat.isRealNumber(oBodyPr.lIns) ? oBodyPr.lIns : 2.54;
				r_ins = AscFormat.isRealNumber(oBodyPr.rIns) ? oBodyPr.rIns : 2.54;
				t_ins = AscFormat.isRealNumber(oBodyPr.tIns) ? oBodyPr.tIns : 1.27;
				b_ins = AscFormat.isRealNumber(oBodyPr.bIns) ? oBodyPr.bIns : 1.27;
			} else {
				l_ins = 2.54;
				r_ins = 2.54;
				t_ins = 1.27;
				b_ins = 1.27;
			}
			if (this.bWordShape) {
				var oPen = this.pen;
				if (oPen) {
					var penW = (oPen.w == null) ? 12700 : parseInt(oPen.w);
					penW /= 36000.0;
					switch (oPen.algn) {
						case 1: {
							break;
						}
						default: {
							penW /= 2.0;
							break;
						}
					}

					l_ins += penW;
					r_ins += penW;
					t_ins += penW;
					b_ins += penW;
				}
			}

			let oForm = this.isForm && this.isForm() ? this.getInnerForm() : null;
			if (oForm) {
				let nFormHorPadding = this.getFormHorPadding();
				l_ins = nFormHorPadding;
				r_ins = nFormHorPadding;
				t_ins = 0;
				b_ins = 0;
			}

			var oRect = this.getTextRect();
			var w, h;
			w = oRect.r - oRect.l - (l_ins + r_ins);
			h = oRect.b - oRect.t - (t_ins + b_ins);
			if (oBodyPr.wrap === AscFormat.nTWTNone) {
				var dMaxWidth = 100000;
				if (this.bWordShape) {
					this.m_oSectPr = null;
					var oParaDrawing = getParaDrawing(this);
					if (oParaDrawing) {
						var oParentParagraph = oParaDrawing.Get_ParentParagraph();
						if (oParentParagraph) {
							var oSectPr = oParentParagraph.Get_SectPr();
							if (oSectPr) {
								if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
									dMaxWidth = oSectPr.GetContentFrameWidth() - l_ins - r_ins;
								} else {
									dMaxWidth = oSectPr.GetContentFrameHeight();
								}
								this.m_oSectPr = new CSectionPr();
								this.m_oSectPr.Copy(oSectPr);
							}
						}
					}
				}
				var dMaxWidthRec = RecalculateDocContentByMaxLine(oDocContent, dMaxWidth, this.bWordShape);
				if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
					if (dMaxWidthRec < w && (!this.bWordShape && !this.bCheckAutoFitFlag)) {
						oRet.w = w + TEXT_RECT_ERROR;
						oDocContent.RecalculateContent(oRet.w, h, nStartPage);
						oRet.contentH = oDocContent.GetSummaryHeight();
						oRet.h = oRet.contentH;
					} else {
						oRet.w = dMaxWidthRec + TEXT_RECT_ERROR;
						oDocContent.RecalculateContent(oRet.w, h, nStartPage);
						oRet.contentH = oDocContent.GetSummaryHeight();
						oRet.h = oRet.contentH;
					}
					oRet.correctW = l_ins + r_ins;
					oRet.correctH = t_ins + b_ins;
					oRet.textRectW = w;
					oRet.textRectH = h;
				} else {
					if (dMaxWidthRec < h && !this.bWordShape) {
						oRet.w = h + TEXT_RECT_ERROR;
						oDocContent.RecalculateContent(oRet.w, h, nStartPage);
						oRet.contentH = oDocContent.GetSummaryHeight();
						oRet.h = oRet.contentH;
					} else {
						oRet.w = dMaxWidthRec + TEXT_RECT_ERROR;
						oDocContent.RecalculateContent(oRet.w, h, nStartPage);
						oRet.contentH = oDocContent.GetSummaryHeight();
						oRet.h = oRet.contentH;
					}
					oRet.correctW = t_ins + b_ins;
					oRet.correctH = l_ins + r_ins;
					oRet.textRectW = h;
					oRet.textRectH = w;
				}
			} else//AscFormat.nTWTSquare
			{
				if (!oBodyPr.upright) {
					if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
						oRet.w = w + TEXT_RECT_ERROR;
						oRet.h = h + TEXT_RECT_ERROR;
						oRet.correctW = l_ins + r_ins;
						oRet.correctH = t_ins + b_ins;
					} else {
						oRet.w = h + TEXT_RECT_ERROR;
						oRet.h = w + TEXT_RECT_ERROR;
						oRet.correctW = t_ins + b_ins;
						oRet.correctH = l_ins + r_ins;
					}
				} else {
					var _full_rotate = this.getFullRotate();
					if (checkNormalRotate(_full_rotate)) {
						if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
							oRet.w = w + TEXT_RECT_ERROR;
							oRet.h = h + TEXT_RECT_ERROR;
							oRet.correctW = l_ins + r_ins;
							oRet.correctH = t_ins + b_ins;
						} else {
							oRet.w = h + TEXT_RECT_ERROR;
							oRet.h = w + TEXT_RECT_ERROR;
							oRet.correctW = t_ins + b_ins;
							oRet.correctH = l_ins + r_ins;
						}
					} else {
						if (!(oBodyPr.vert === AscFormat.nVertTTvert || oBodyPr.vert === AscFormat.nVertTTvert270 || oBodyPr.vert === AscFormat.nVertTTeaVert)) {
							oRet.w = h + TEXT_RECT_ERROR;
							oRet.h = w + TEXT_RECT_ERROR;
							oRet.correctW = t_ins + b_ins;
							oRet.correctH = l_ins + r_ins;
						} else {
							oRet.w = w + TEXT_RECT_ERROR;
							oRet.h = h + TEXT_RECT_ERROR;
							oRet.correctW = l_ins + r_ins;
							oRet.correctH = t_ins + b_ins;
						}
					}
				}
				oRet.textRectW = oRet.w;
				oRet.textRectH = oRet.h;

				//oDocContent.Set_StartPage(0);
				/*oDocContent.Reset(0, 0, oRet.w, 20000);
        var CurPage = 0;
        var RecalcResult = recalcresult2_NextPage;
        while ( recalcresult2_End !== RecalcResult  )
            RecalcResult = oDocContent.Recalculate_Page( CurPage++, true );*/

				var oContentW = oRet.w;

				if (oForm && !oForm.IsMultiLineForm())
					oDocContent.SetUseXLimit(false);
				else
					oDocContent.SetUseXLimit(true);

				oDocContent.RecalculateContent(oContentW, oRet.h, nStartPage);
				oRet.contentH = oDocContent.GetSummaryHeight();

				if (this.bWordShape) {
					this.m_oSectPr = null;
					var oParaDrawing = getParaDrawing(this);
					if (oParaDrawing) {
						var oParentParagraph = oParaDrawing.Get_ParentParagraph();
						if (oParentParagraph) {
							var oSectPr = oParentParagraph.Get_SectPr();
							if (oSectPr) {
								this.m_oSectPr = new CSectionPr();
								this.m_oSectPr.Copy(oSectPr);
							}
						}
					}
				}
			}
			return oRet;
		};

		CShape.prototype.getSmartArtPointContent = function () {
			if (this.isObjectInSmartArt()) {
				return this.getSmartArtInfo() && this.getSmartArtInfo().contentPoint;
			}
		}

		CShape.prototype.getSmartArtShapePoint = function () {
			return this.getSmartArtInfo() && this.getSmartArtInfo().shapePoint;
		}
		CShape.prototype.recalculateContent2 = function () {
			if (this.txBody) {
				var pointContent = this.getSmartArtPointContent();
				if (this.isPlaceholder() || this.isPlaceholderInSmartArt()) {
					if (!this.isEmptyPlaceholder()) {
						return;
					}
					let aHierarchy = this.getHierarchy();
					for (let nPlaceholder = 0; nPlaceholder < aHierarchy.length; ++nPlaceholder) {
						let oHSp = aHierarchy[nPlaceholder];
						if (isRealObject(oHSp) && oHSp.hasCustomPrompt()) {
							if (oHSp.txBody && oHSp.txBody.content) {
								this.txBody.content2 = oHSp.txBody.content.Copy(this.txBody, this.getDrawingDocument(), {});
							}
						}
					}


					if (!this.txBody.content2) {
						var text;
						if (typeof AscCommonSlide !== "undefined" && AscCommonSlide.CNotes && this.parent instanceof AscCommonSlide.CNotes && this.nvSpPr.nvPr.ph.type === AscFormat.phType_body) {
							text = AscCommon.translateManager.getValue("Click to add notes");
						} else if (this.isObjectInSmartArt()) {
							text = AscCommon.translateManager.getValue(pointContent[0].prSet.phldrT || '');
						} else {
							text = this.getPlaceholderName();
						}
						this.txBody.content2 = AscFormat.CreateDocContentFromString(text, this.getDrawingDocument(), this.txBody);
						if (this.txBody.content && this.isObjectInSmartArt()) {
							var oContent = this.txBody.content;
							var oContent2 = this.txBody.content2;
							var contentLength = oContent.Content.length;
							var phldrParagraph = oContent2.Content[0];
							for (var i = 1; i < contentLength; i += 1) {
								var oCopy = phldrParagraph.Copy(oContent2, this.getDrawingDocument());
								oContent2.Internal_Content_Add(i, oCopy, false);
							}
						}
					} else {
						this.txBody.content2.Recalc_AllParagraphs_CompiledPr();
					}

					var content = this.txBody.content2;
					if (content) {
						var w, h;
						var l_ins, t_ins, r_ins, b_ins;
						var body_pr = this.getBodyPr();
						if (body_pr) {
							l_ins = AscFormat.isRealNumber(body_pr.lIns) ? body_pr.lIns : 2.54;
							r_ins = AscFormat.isRealNumber(body_pr.rIns) ? body_pr.rIns : 2.54;
							t_ins = AscFormat.isRealNumber(body_pr.tIns) ? body_pr.tIns : 1.27;
							b_ins = AscFormat.isRealNumber(body_pr.bIns) ? body_pr.bIns : 1.27;
						} else {
							l_ins = 2.54;
							r_ins = 2.54;
							t_ins = 1.27;
							b_ins = 1.27;
						}
						var rect = this.getTextRect && this.getTextRect();
						if (AscFormat.isRealNumber(rect.l) && AscFormat.isRealNumber(rect.t)
							&& AscFormat.isRealNumber(rect.r) && AscFormat.isRealNumber(rect.r)) {
							w = rect.r - rect.l - (l_ins + r_ins);
							h = rect.b - rect.t - (t_ins + b_ins);
						} else {
							w = this.extX - (l_ins + r_ins);
							h = this.extY - (t_ins + b_ins);
						}

						if (!body_pr.upright) {
							if (!(body_pr.vert === AscFormat.nVertTTvert || body_pr.vert === AscFormat.nVertTTvert270 || body_pr.vert === AscFormat.nVertTTeaVert)) {
								this.txBody.contentWidth2 = w;
								this.txBody.contentHeight2 = h;
							} else {
								this.txBody.contentWidth2 = h;
								this.txBody.contentHeight2 = w;
							}

						} else {
							var _full_rotate = this.getFullRotate();
							if (AscFormat.checkNormalRotate(_full_rotate)) {
								if (!(body_pr.vert === AscFormat.nVertTTvert || body_pr.vert === AscFormat.nVertTTvert270 || body_pr.vert === AscFormat.nVertTTeaVert)) {

									this.txBody.contentWidth2 = w;
									this.txBody.contentHeight2 = h;
								} else {
									this.txBody.contentWidth2 = h;
									this.txBody.contentHeight2 = w;
								}
							} else {
								if (!(body_pr.vert === AscFormat.nVertTTvert || body_pr.vert === AscFormat.nVertTTvert270 || body_pr.vert === AscFormat.nVertTTeaVert)) {

									this.txBody.contentWidth2 = h;
									this.txBody.contentHeight2 = w;
								} else {
									this.txBody.contentWidth2 = w;
									this.txBody.contentHeight2 = h;
								}
							}
						}


					}
					this.contentWidth2 = this.txBody.contentWidth2;
					this.contentHeight2 = this.txBody.contentHeight2;


					var content_ = this.getDocContent();
					if (content_) {
						for (i = 0; i < content.Content.length; i += 1) {
							if (content_.Content[i]) {
								content.Content[i].Pr = content_.Content[i].Pr.Copy();
								if (!content.Content[i].Pr.DefaultRunPr) {
									content.Content[i].Pr.DefaultRunPr = new AscCommonWord.CTextPr();
								}
								content.Content[i].Pr.DefaultRunPr.Merge(content_.Content[i].GetFirstRunPr());
							}
						}
					}
					this.bCheckAutoFitFlag = true;
					this.tmpFontScale = undefined;
					this.tmpLnSpcReduction = undefined;
					content.Set_StartPage(0);
					content.Reset(0, 0, w, 20000);
					content.RecalculateContent(this.txBody.contentWidth2, this.txBody.contentHeight2, 0);
					var oTextWarpContent = this.checkTextWarp(content, body_pr, this.txBody.contentWidth2, this.txBody.contentHeight2, false, true);
					this.txWarpStructParamarks2 = oTextWarpContent.oTxWarpStructParamarks;
					this.txWarpStruct2 = oTextWarpContent.oTxWarpStruct;
					this.bCheckAutoFitFlag = false;
				} else {
					this.txBody.content2 = null;
					this.txWarpStructParamarks2 = null;
					this.txWarpStruct2 = null;
				}
			} else {
				this.txWarpStructParamarks2 = null;
				this.txWarpStruct2 = null;
			}
		};

		CShape.prototype.isShape = function () {
			return true;
		};

		CShape.prototype.isInk = function () {
			const oGeometry = this.spPr && this.spPr.geometry;
			if(!oGeometry) {
				return false;
			}
			return oGeometry.isInk();
		};

		var aScales = [25000, 30000, 35000, 40000, 45000, 50000, 55000, 60000, 65000, 70000, 75000, 80000, 85000, 90000, 95000, 10000];


		CShape.prototype.recalculateContentWitCompiledPr = function () {
			var oContent = this.getDocContent && this.getDocContent();
			if (oContent) {
				oContent.Recalc_AllParagraphs_CompiledPr();
				this.recalcInfo.recalculateContent = true;
				if (this.isPlaceholderInSmartArt()) {
					this.recalcInfo.recalculateContent2 = true;
				}
				this.recalcInfo.recalculateTransformText = true;
				this.recalculate();
			}
		};

		CShape.prototype.getFirstFontSize = function () {
			let currentFontSize;
			if (this.txBody && this.txBody.content) {
				this.txBody.content.CheckRunContent(function (paraRun) {
					if (!currentFontSize) {
						currentFontSize = paraRun.Get_FontSize();
						return true;
					}
				});
				if (!currentFontSize) {
					this.txBody.content.Content.forEach(function (paragraph) {
						if (!currentFontSize) {
							currentFontSize = paragraph.TextPr.Value.FontSize;
							return true;
						}
					});
				}
			}
			return currentFontSize;
		}

		CShape.prototype.setFontSizeInSmartArt = function (fontSize, bSkipRecalculateContent2) {
			const oContent = this.txBody && this.txBody.content;
			if (this.txBody && oContent) {
				const currentFontSize = this.getFirstFontSize();
				const oBodyPr = this.txBody.getBodyPr();
				if (oBodyPr) {
					const paddings = {};
					const pointContent = this.getSmartArtPointContent();
					const point = pointContent && pointContent[0];
					if (point) {
						const isRecalculateInsets = point.isRecalculateInsets();
						if (isRecalculateInsets.Top) {
							const tInsetPerPt = oBodyPr.tIns / currentFontSize;
							paddings.Top = tInsetPerPt * fontSize;
						}
						if (isRecalculateInsets.Bottom) {
							const bInsetPerPt = oBodyPr.bIns / currentFontSize;
							paddings.Bottom = bInsetPerPt * fontSize;
						}
						if (isRecalculateInsets.Left) {
							const lInsetPerPt = oBodyPr.lIns / currentFontSize;
							paddings.Left = lInsetPerPt * fontSize;
						}
						if (isRecalculateInsets.Right) {
							const rInsetPerPt = oBodyPr.rIns / currentFontSize;
							paddings.Right = rInsetPerPt * fontSize;
						}
					}
					// In files layout.xml insets depend on font size.
					// While there is no recalculation, we consider new insets as a dependency on the previous font size.
					this.setPaddings(paddings, {bNotCopyToPoints: true});
				}
				const bOldApplyToAll = oContent.ApplyToAll;
				oContent.ApplyToAll = true;
				oContent.AddToParagraph(new AscCommonWord.ParaTextPr({FontSize: (Math.min(fontSize, 300))}), false);
				oContent.ApplyToAll = bOldApplyToAll;
				if (!bSkipRecalculateContent2) {
					this.recalculateContent2();
				}
			}
		};
		CShape.prototype.resetSmartArtMaxFontSize = function () {
			const oSmartArtInfo = this.getSmartArtInfo();
			if (oSmartArtInfo) {
				delete oSmartArtInfo.maxFontSize;
			}
		};
		CShape.prototype.correctSmartArtUndo = function () {
			if (this.isObjectInSmartArt()) {
				this.group.group.fitFontSize();
				const oSmartArtInfo = this.getSmartArtInfo();
				if (oSmartArtInfo) {
					if (!this.isEmptyPlaceholder(true)) {
						if (this.isPlaceholderInSmartArt()) {
							const oContent = this.getDocContent && this.getDocContent();
							const arrPointContent = this.getSmartArtPointContent();
							const bIsNotEmptyShape = oContent.Content.some(function (paragraph) {
								return !paragraph.Is_Empty({SkipEnd: true, SkipPlcHldr: false});
							});
							if (bIsNotEmptyShape) {
								arrPointContent.forEach(function (point) {
									point.prSet.setPhldr(false);
								})
								this.txBody.content2 = null;
							}
						}
					}
				}
			}

		};

		CShape.prototype.getInsets = function (properties) {
			const oBodyPr = properties.bodyPr || this.getBodyPr && this.getBodyPr();
			properties = properties || {};
			let lIns = 0, tIns = 0, rIns = 0, bIns = 0;
			if (!properties.bIgnoreInsets) {
				lIns = (AscFormat.isRealNumber(oBodyPr.lIns) ? oBodyPr.lIns : 2.54);
				tIns = (AscFormat.isRealNumber(oBodyPr.tIns) ? oBodyPr.tIns : 1.27);
				rIns = (AscFormat.isRealNumber(oBodyPr.rIns) ? oBodyPr.rIns : 2.54);
				bIns = (AscFormat.isRealNumber(oBodyPr.bIns) ? oBodyPr.bIns : 1.27);
			}

			if (this.bWordShape) {
				const oPen = this.pen;
				if (oPen) {
					let penW = (oPen.w == null) ? 12700 : parseInt(oPen.w);
					penW /= 36000;
					switch (oPen.algn) {
						case 1: {
							break;
						}
						default: {
							penW /= 2;
							break;
						}
					}
					lIns += penW;
					rIns += penW;
					tIns += penW;
					bIns += penW;
				}
			}

			const oForm = this.isForm && this.isForm() ? this.getInnerForm() : null;
			if (oForm) {
				const nFormHorPadding = this.getFormHorPadding();
				lIns = nFormHorPadding;
				rIns = nFormHorPadding;
				tIns = 0;
				bIns = 0;
			}
			return {
				lIns: lIns,
				tIns: tIns,
				rIns: rIns,
				bIns: bIns
			};
		};

		CShape.prototype.getTextRectBoundsWithInsets = function (properties) {
			properties = properties || {};
			const result = {l: 0, t: 0, r: 0, b: 0};
			const oBodyPr = properties.bodyPr || this.getBodyPr && this.getBodyPr();
			if (oBodyPr) {
				const insets = this.getInsets(properties);
				const oRect = this.getTextRect();

				let _l = oRect.l + insets.lIns;
				let _t = oRect.t + insets.tIns;
				let _r = oRect.r - insets.rIns;
				let _b = oRect.b - insets.bIns;

				if (_l >= _r) {
					const _c = (_l + _r) * 0.5;
					_l = _c - 0.01;
					_r = _c + 0.01;
				}

				if (_t >= _b) {
					const _c = (_t + _b) * 0.5;
					_t = _c - 0.01;
					_b = _c + 0.01;
				}

				result.r = _r;
				result.l = _l;
				result.t = _t;
				result.b = _b;
			}
			return result;
		};

		CShape.prototype.getTextRectContentHW = function (properties) {
			const bounds = this.getTextRectBoundsWithInsets(properties);
			return {
				height: bounds.b - bounds.t,
				width: bounds.r - bounds.l
			};
		};

		CShape.prototype.compareHeightOfBoundsTextInSmartArt = function () {
			const sizesOfTextRectContent = this.getTextRectContentHW();
			const vert = this.txBody && this.txBody.bodyPr.vert;


			let nContentHeight;
			if (this.txBody && this.txBody.content2) {
				this.recalculateContent();
				nContentHeight = this.contentHeight;
			} else {
				this.recalculateContent();
				nContentHeight = this.contentHeight;
			}
			if (vert === AscFormat.nVertTTvert270 || vert === AscFormat.nVertTTvert || vert === AscFormat.nVertTTeaVert) {
				return nContentHeight >= sizesOfTextRectContent.width;
			}
			return nContentHeight >= sizesOfTextRectContent.height;
		};
		CShape.prototype.compareWidthOfBoundsTextInSmartArt = function (bMax) {
			const oContent = this.getCurrentDocContentInSmartArt();
			const sizesOfTextRectContent = this.getTextRectContentHW();
			const vert = this.txBody && this.txBody.bodyPr.vert;
			let widthOfContent = oContent.RecalculateMinMaxContentWidth();
			if (bMax) {
				widthOfContent = widthOfContent.Max;
			} else {
				widthOfContent = widthOfContent.Min;
			}
			if (vert === AscFormat.nVertTTvert270 || vert === AscFormat.nVertTTvert || vert === AscFormat.nVertTTeaVert) {
				return widthOfContent > sizesOfTextRectContent.height;
			}
			return widthOfContent > sizesOfTextRectContent.width;
		};
		CShape.prototype.checkFitContentForSmartArt = function () {
			// почему-то у майков не подбирается шрифт для ширины, если вставлено только уравнение
			const oContent = this.getCurrentDocContentInSmartArt();
			if (oContent) {
				for (let i = 0; i < oContent.Content.length; i += 1) {
					const oParagraph = oContent.Content[i];
					for (let j = 0; j < oParagraph.Content.length; j += 1) {
						const oElement = oParagraph.Content[j];
						if (oElement.Type !== para_Math && !oElement.Is_Empty({SkipEnd: true})) {
							return true;
						}
					}
				}
			}
			return false;
		};
		CShape.prototype.findFitFontSize = function (nMinFontSize, nMaxFontSize, bMax) {
			if (nMinFontSize > nMaxFontSize) {
				return null;
			}
			if (nMinFontSize === nMaxFontSize) {
				return nMaxFontSize;
			}
			if (nMinFontSize)
			return AscFormat.ExecuteNoHistory(function () {
				const MAX_FONT_SIZE = nMaxFontSize || 65;
				const content = this.getCurrentDocContentInSmartArt();
				if (content) {
					const nOldFontSize = this.getFirstFontSize();
					const scalesForSmartArt = Array(Math.max(1, Math.trunc(MAX_FONT_SIZE - (nMinFontSize - 1)))).fill(0).map(function (e, ind) {
						return ind + nMinFontSize;
					});
					let a = 0;
					let b = scalesForSmartArt.length - 1;
					let averageAmount = Math.floor((a + b) / 2);
					while (a !== averageAmount && b !== averageAmount) {
						this.setFontSizeInSmartArt(scalesForSmartArt[averageAmount]);
						let bCheck = this.compareWidthOfBoundsTextInSmartArt(bMax) || this.compareHeightOfBoundsTextInSmartArt();

						if (bCheck) {
							b = averageAmount;
						} else {
							a = averageAmount;
						}
						averageAmount = Math.floor((a + b) / 2);
					}
					this.setFontSizeInSmartArt(nOldFontSize);
					this.recalculateContent();
					return scalesForSmartArt[averageAmount];
				}
				return MAX_FONT_SIZE;
			}, this, []);
		};
		CShape.prototype.findFitFontSizeForSmartArt = function (bMax) {
			return AscFormat.ExecuteNoHistory(function () {
				const MAX_FONT_SIZE = 65;
				const content = this.getCurrentDocContentInSmartArt();
				if (content) {
					const nOldFontSize = this.getFirstFontSize();
					const scalesForSmartArt = Array((MAX_FONT_SIZE - 4) > 0 ? MAX_FONT_SIZE - 4 : 1).fill(0).map(function (e, ind) {
						return ind + 5;
					});
					let a = 0;
					let b = scalesForSmartArt.length - 1;
					let averageAmount = Math.floor((a + b) / 2);
					const bNeedCheckWidth = this.checkFitContentForSmartArt();
					while (a !== averageAmount && b !== averageAmount) {
						this.setFontSizeInSmartArt(scalesForSmartArt[averageAmount]);
						let bCheck;
						if (bNeedCheckWidth) {
							bCheck = this.compareWidthOfBoundsTextInSmartArt(bMax) || this.compareHeightOfBoundsTextInSmartArt();
						} else {
							bCheck = this.compareHeightOfBoundsTextInSmartArt();
						}

						if (bCheck) {
							b = averageAmount;
						} else {
							a = averageAmount;
						}
						averageAmount = Math.floor((a + b) / 2);
					}
					this.setFontSizeInSmartArt(nOldFontSize);
					this.recalculateContent();
					return scalesForSmartArt[averageAmount];
				}
				return MAX_FONT_SIZE;
			}, this, []);
		};

		CShape.prototype.getShapesForFitText = function () {
			return this.isObjectInSmartArt() ? this.group.group.getShapesForFitText(this) : [];
		};

		CShape.prototype.setTruthFontSizeInSmartArt = function () {
			const arrMainContentPoints = this.getSmartArtPointContent();
			if (!arrMainContentPoints) return;
			const bIsFitText = arrMainContentPoints.every(function (point) {
				return point && point.prSet && (typeof point.prSet.phldrT === "string") && !point.prSet.custT && !point.prSet.phldr;
			});
			let bIsPlaceholder = arrMainContentPoints.every(function (point) {
				return point && point.prSet && (typeof point.prSet.phldrT === "string") && !point.prSet.custT && point.prSet.phldr;
			});

			if (!bIsFitText && !bIsPlaceholder) {
				return;
			}
			const oSmartArtInfo = this.getSmartArtInfo();
			if (oSmartArtInfo) {
				oSmartArtInfo.setMaxFontSize(this.findFitFontSizeForSmartArt());
			}
			const arrShapes = this.getShapesForFitText();
			const arrPlaceholders = [];
			const arrFitText = [];
			for (let i = 0; i < arrShapes.length; i += 1) {
				const oShape = arrShapes[i];
				var contentPoints = oShape.getSmartArtPointContent();
				const isPlaceholder = contentPoints.every(function (point) {
					return point && point.prSet && (typeof point.prSet.phldrT === "string") && !point.prSet.custT && point.prSet.phldr;
				});
				const isNotPlaceholder = contentPoints.every(function (point) {
					return point && point.prSet && (typeof point.prSet.phldrT === "string") && !point.prSet.custT && !point.prSet.phldr;
				});
				if (isPlaceholder) {
					arrPlaceholders.push(oShape);
				} else if (isNotPlaceholder) {
					arrFitText.push(oShape);
				}
			}

			let nFitFontSize = 65;
			for (let i = 0; i < arrFitText.length; i += 1) {
				const oShape = arrFitText[i];
				const oShapeSmartArtInfo = oShape.getSmartArtInfo();
				if (oShapeSmartArtInfo) {
					if (!AscFormat.isRealNumber(oShapeSmartArtInfo.maxFontSize)) {
						oShapeSmartArtInfo.setMaxFontSize(oShape.findFitFontSizeForSmartArt());
					}
					if (oShapeSmartArtInfo.maxFontSize < nFitFontSize) {
						nFitFontSize = oShapeSmartArtInfo.maxFontSize;
					}
				}
			}
			for (let i = 0; i < arrFitText.length; i += 1) {
				const oShape = arrFitText[i];
				const nCurrentFontSize = oShape.getFirstFontSize();
				if (nCurrentFontSize !== nFitFontSize) {
					oShape.setFontSizeInSmartArt(nFitFontSize, true);
				}
			}

			for (let i = 0; i < arrPlaceholders.length; i += 1) {
				const oShape = arrPlaceholders[i];
				const nCurrentFontSize = oShape.getFirstFontSize();
				const oPlaceholderSmartArtInfo = oShape.getSmartArtInfo();
				if (oPlaceholderSmartArtInfo) {
					if (!AscFormat.isRealNumber(oPlaceholderSmartArtInfo.maxFontSize)) {
						oPlaceholderSmartArtInfo.setMaxFontSize(oShape.findFitFontSizeForSmartArt());
					}
					const nPlaceholderFontSize = Math.min(oPlaceholderSmartArtInfo.maxFontSize, nFitFontSize);
					if (nCurrentFontSize !== nPlaceholderFontSize) {
						oShape.setFontSizeInSmartArt(nPlaceholderFontSize, true);
					}
				} else if (nCurrentFontSize !== nFitFontSize) {
					oShape.setFontSizeInSmartArt(nFitFontSize, true);
				}
			}
		};

		CShape.prototype.checkExtentsByDocContent = function (bForce, bNeedRecalc) {
			if ((!this.bWordShape || this.group || bForce) && this.checkAutofit(true)) {
				var oMainGroup = this.getMainGroup();
				if (oMainGroup && !(bNeedRecalc === false)) {
					oMainGroup.normalize();
				}

				this.tmpFontScale = undefined;
				this.tmpLnSpcReduction = undefined;
				this.bCheckAutoFitFlag = true;
				var oOldRecalcTitle = this.recalcInfo.recalcTitle;
				var bOldRecalcTitle = this.recalcInfo.bRecalculatedTitle;
				this.handleUpdateExtents();
				this.recalcInfo.bRecalculatedTitle = false;
				this.recalcInfo.recalcTitle = this;
				this.recalculate();
				this.bCheckAutoFitFlag = false;
				this.recalcInfo.recalcTitle = oOldRecalcTitle;
				this.recalcInfo.bRecalculatedTitle = bOldRecalcTitle;
				AscFormat.CheckSpPrXfrm(this, true);
				this.spPr.xfrm.setExtX(this.extX + TEXT_RECT_ERROR);
				this.spPr.xfrm.setExtY(this.extY + TEXT_RECT_ERROR);
				if (!this.bWordShape || this.group) {
					this.spPr.xfrm.setOffX(this.x);
					this.spPr.xfrm.setOffY(this.y);
					if (this.drawingBase) {
						CheckExcelDrawingXfrm(this.spPr.xfrm);
					}
				}
				if (!(bNeedRecalc === false)) {
					if (oMainGroup) {
						oMainGroup.updateCoordinatesAfterInternalResize();
						if (oMainGroup.parent && oMainGroup.parent.CheckWH) {
							oMainGroup.parent.CheckWH();
							if (this.bWordShape) {
								editor.WordControl.m_oLogicDocument.Recalculate();
							}
						}
					} else {
						this.checkDrawingBaseCoords();
					}
				}
				return true;
			} else {
				var oBodyPr = this.getBodyPr && this.getBodyPr();
				var oContent = this.getDocContent && this.getDocContent();
				var pointContent = this.getSmartArtPointContent();
				if (oBodyPr && oContent && this.clipRect) {
					var oTextFit = oBodyPr.textFit;
					if (oTextFit && oTextFit.type === AscFormat.text_fit_NormAuto) {
						var dOldContentHeight = this.contentHeight;
						var dOldClipW = this.clipRect.w;
						var dOldClipH = this.clipRect.h;
						this.recalcInfo.recalculateContent = true;
						this.recalculate();
						if (!AscFormat.isRealNumber(oTextFit.fontScale) && !AscFormat.isRealNumber(oTextFit.lnSpcReduction)
							&& this.contentHeight <= this.clipRect.h) {
							return;
						}
						if (!bForce && AscFormat.isRealNumber(dOldClipW) && AscFormat.isRealNumber(dOldClipH)
							&& AscFormat.fApproxEqual(dOldClipW, this.clipRect.w) && AscFormat.fApproxEqual(dOldClipH, this.clipRect.h)) {
							if (AscFormat.isRealNumber(dOldContentHeight) && AscFormat.fApproxEqual(dOldContentHeight, this.contentHeight)) {
								return;
							}
							if (dOldContentHeight < this.contentHeight && oTextFit.fontScale === aScales[0]) {
								return;
							}
						}

						this.bCheckAutoFitFlag = true;

						this.tmpFontScale = undefined;
						this.tmpLnSpcReduction = undefined;
						this.recalculateContentWitCompiledPr();
						if (this.contentHeight <= this.clipRect.h) {
							oBodyPr = oBodyPr.createDuplicate();
							oBodyPr.textFit.lnSpcReduction = this.tmpLnSpcReduction;
							oBodyPr.textFit.fontScale = this.tmpFontScale;
							if (this.bWordShape) {
								this.setBodyPr(oBodyPr);
							} else {
								if (this.txBody) {
									this.txBody.setBodyPr(oBodyPr);
								}
							}
							this.bCheckAutoFitFlag = false;
							this.tmpFontScale = undefined;
							this.tmpLnSpcReduction = undefined;
							return;
						}


						var dReductionScale = 0.2;

						var nCurIndex = aScales.length - 1;
						var nCurShift = -((aScales.length) / 2);
						while (true) {
							nCurIndex += nCurShift;
							if (nCurIndex - 1 >= 0) {
								this.tmpFontScale = aScales[nCurIndex - 1];
								this.tmpLnSpcReduction = dReductionScale * (100000 - this.tmpFontScale) >> 0;
								this.recalculateContentWitCompiledPr();


								if (this.contentHeight <= this.clipRect.h) {
									this.tmpFontScale = aScales[nCurIndex];
									this.tmpLnSpcReduction = dReductionScale * (100000 - this.tmpFontScale) >> 0;
									this.recalculateContentWitCompiledPr();
									if (this.contentHeight >= this.clipRect.h) {
										this.tmpFontScale = aScales[nCurIndex - 1];
										this.tmpLnSpcReduction = dReductionScale * (100000 - this.tmpFontScale) >> 0;
										break;
									} else {
										nCurShift = Math.abs(nCurShift) / 2;
									}
								} else {
									nCurShift = -Math.abs(nCurShift) / 2;
								}
								if (Math.abs(nCurShift) < 1) {
									break;
								}
							} else {
								this.tmpFontScale = aScales[0];
								this.tmpLnSpcReduction = dReductionScale * (100000 - this.tmpFontScale) >> 0;
								break;
							}
						}
						if (AscFormat.isRealNumber(this.tmpFontScale) && this.tmpFontScale < 90000) {
							if (this.isPlaceholder()) {
								var nType = this.getPlaceholderType();
								if (nType === AscFormat.phType_title || nType === AscFormat.phType_ctrTitle) {
									this.tmpFontScale = 90000;
									this.tmpLnSpcReduction = dReductionScale * (100000 - this.tmpFontScale) >> 0;
								}
							}
						}

						if (oBodyPr.textFit.lnSpcReduction !== this.tmpLnSpcReduction
							|| oBodyPr.textFit.fontScale !== this.tmpFontScale) {
							oBodyPr = oBodyPr.createDuplicate();
							oBodyPr.textFit.lnSpcReduction = this.tmpLnSpcReduction;
							oBodyPr.textFit.fontScale = this.tmpFontScale;
							if (this.bWordShape) {
								this.setBodyPr(oBodyPr);
							} else {
								if (this.txBody) {
									this.txBody.setBodyPr(oBodyPr);
								}
							}
						}
						this.bCheckAutoFitFlag = false;
						this.tmpFontScale = undefined;
						this.tmpLnSpcReduction = undefined;
						this.recalculateContentWitCompiledPr();
					} else {
						var oForm;
						if (this.isForm && this.isForm() && (oForm = this.getInnerForm()) && oForm.IsAutoFitContent() && oForm.GetLogicDocument() && oForm.GetLogicDocument().CheckFormAutoFit)
							oForm.GetLogicDocument().CheckFormAutoFit(oForm);

						if (bForce) {
							this.recalculateContentWitCompiledPr();
						}
					}
					if (this.isPlaceholderInSmartArt()) {
						var isNotEmptyShape = oContent.Content.some(function (paragraph) {
							return !paragraph.Is_Empty({SkipEnd: true, SkipPlcHldr: false});
						});
						if (isNotEmptyShape) {
							pointContent.forEach(function (point) {
								point.prSet.setPhldr(false);
							})
							this.txBody.content2 = null;
						}
					}
					if (this.isObjectInSmartArt()) {
						this.copyTextInfoFromShapeToPoint();
						this.setTruthFontSizeInSmartArt();
					}
				}
			}
			return false;
		};


		CShape.prototype.getTransformMatrix = function () {
			return this.transform;
		};

		CShape.prototype.getTransform = function () {

			return {
				x: this.x,
				y: this.y,
				extX: this.extX,
				extY: this.extY,
				rot: this.rot,
				flipH: this.flipH,
				flipV: this.flipV
			};
		};

		CShape.prototype.getAngle = function (x, y) {
			var px = this.invertTransform.TransformPointX(x, y);
			var py = this.invertTransform.TransformPointY(x, y);
			return Math.PI * 0.5 + Math.atan2(px - this.extX * 0.5, py - this.extY * 0.5);
		};


		CShape.prototype.recalculateGeometry = function () {
			if (this.spPr && isRealObject(this.spPr.geometry)) {
				var transform = this.getTransform();
				this.spPr.geometry.Recalculate(transform.extX, transform.extY);
			}
		};
		CShape.prototype.drawAdjustments = function (drawingDocument) {
			if (this.spPr && isRealObject(this.spPr.geometry) && this.canChangeAdjustments()) {
				this.spPr.geometry.drawAdjustments(drawingDocument, this.transform, false);
			}
			if (this.recalcInfo.warpGeometry) {
				this.recalcInfo.warpGeometry.drawAdjustments(drawingDocument, this.transformTextWordArt, true);
			}
		};

		CShape.prototype.getHandlePosByIndex = function (numHandle) {
			var t = this.transform;
			switch (numHandle) {
				case 0:
					return {x: t.TransformPointX(0, 0), y: t.TransformPointY(0, 0)};
				case 1:
					return {x: t.TransformPointX(this.extX / 2.0, 0), y: t.TransformPointY(this.extX / 2.0, 0)};
				case 2:
					return {x: t.TransformPointX(this.extX, 0), y: t.TransformPointY(this.extX, 0)};
				case 3:
					return {
						x: t.TransformPointX(this.extX, this.extY / 2.0),
						y: t.TransformPointY(this.extX, this.extY / 2.0)
					};
				case 4:
					return {x: t.TransformPointX(this.extX, this.extY), y: t.TransformPointY(this.extX, this.extY)};
				case 5:
					return {
						x: t.TransformPointX(this.extX / 2.0, this.extY),
						y: t.TransformPointY(this.extX / 2.0, this.extY)
					};
				case 6:
					return {x: t.TransformPointX(0, this.extY), y: t.TransformPointY(0, this.extY)};
				case 7:
					return {x: t.TransformPointX(0, this.extY / 2.0), y: t.TransformPointY(0, 0)};
				default:
			}
		};


		CShape.prototype.getGroupHierarchy = function () {
			if (this.recalcInfo.recalculateGroupHierarchy) {
				this.groupHierarchy = [];
				if (isRealObject(this.group)) {
					var parent_group_hierarchy = this.group.getGroupHierarchy();
					for (var i = 0; i < parent_group_hierarchy.length; ++i) {
						this.groupHierarchy.push(parent_group_hierarchy[i]);
					}
					this.groupHierarchy.push(this.group);
				}
				this.recalcInfo.recalculateGroupHierarchy = false;
			}
			return this.groupHierarchy;
		};


		CShape.prototype.hitInTextRectWord = function (x, y) {
			if (!AscFormat.canSelectDrawing(this)) {
				return false;
			}
			var content = this.getDocContent && this.getDocContent();
			if (content) {
				var w, h, x_, y_;
				if (this.isObjectInSmartArt() && this.invertTransformText) {
					x_ = 0;
					y_ = 0;
					w = this.contentWidth;
					h = this.contentHeight;
					return AscFormat.HitToRect(x, y, this.invertTransformText, x_, y_, w, h);
				} else if (this.invertTransform) {
					var rect = this.getTextRect && this.getTextRect();
					if (rect && AscFormat.isRealNumber(rect.l) && AscFormat.isRealNumber(rect.t)
						&& AscFormat.isRealNumber(rect.r) && AscFormat.isRealNumber(rect.r)) {
						x_ = rect.l;
						y_ = rect.t;
						w = rect.r - rect.l;
						h = rect.b - rect.t;
					} else {
						x_ = 0;
						y_ = 0;
						w = this.extX;
						h = this.extY;
					}
					return AscFormat.HitToRect(x, y, this.invertTransform, x_, y_, w, h);
				}
			}
			return false;
		};


		CShape.prototype.hitInTextRect = function (x, y) {
			var oController = this.getDrawingObjectsController && this.getDrawingObjectsController();
			if (this.parent && this.parent.kind === AscFormat.TYPE_KIND.NOTES) {
				return true;
			}


			if (this.checkHiddenInAnimation && this.checkHiddenInAnimation()) {
				return false;
			}
			if (!AscFormat.canSelectDrawing(this)) {
				return false;
			}

			var bForceWord = ((this.isEmptyPlaceholder && this.isEmptyPlaceholder()) || (this.isPlaceholder && this.isPlaceholder() && oController && (AscFormat.getTargetTextObject(oController) === this)));
			if (bForceWord) {
				if (this.hitInTextRectWord(x, y)) {
					return true;
				}
			}
			if (!this.txWarpStruct || !this.recalcInfo.warpGeometry ||
				this.recalcInfo.warpGeometry.preset === "textNoShape" ||
				oController && (AscFormat.getTargetTextObject(oController) === this || (oController.curState.startTargetTextObject === this))) {
				var content = this.getDocContent && this.getDocContent();
				if (content && this.invertTransformText) {
					return AscFormat.HitToRect(x, y, this.invertTransformText, 0, 0, this.contentWidth, this.contentHeight);
				}
			} else {
				return this.hitInTextRectWord(x, y);
			}

			return false;
		};


		CShape.prototype.updateCursorType = function (x, y, e) {
			if (this.invertTransformText) {
				var tx = this.invertTransformText.TransformPointX(x, y);
				var ty = this.invertTransformText.TransformPointY(x, y);
				this.txBody.content.UpdateCursorType(tx, ty, 0);
			}
		};

		CShape.prototype.selectionSetStart = function (e, x, y, slideIndex) {
			if (this.isProtectedText && this.isProtectedText()) {
				return;
			}
			var content = this.getDocContent();
			if (isRealObject(content)) {
				var tx, ty;
				tx = this.invertTransformText.TransformPointX(x, y);
				ty = this.invertTransformText.TransformPointY(x, y);
				if (e.Button === AscCommon.g_mouse_button_right) {
					if (content.CheckPosInSelection(tx, ty, 0)) {
						this.rightButtonFlag = true;
						return;
					}
				}
				if (!(/*content.IsTextSelectionUse() && */e.ShiftKey))
					content.Selection_SetStart(tx, ty, slideIndex - content.Get_StartPage_Relative(), e);
				else {
					if (!content.IsTextSelectionUse()) {
						content.StartSelectionFromCurPos();
					}
					content.Selection_SetEnd(tx, ty, slideIndex - content.Get_StartPage_Relative(), e);
				}
			}
		};

		CShape.prototype.selectionSetEnd = function (e, x, y, slideIndex) {
			if (this.isProtectedText && this.isProtectedText()) {
				return;
			}
			var content = this.getDocContent();
			if (isRealObject(content)) {
				var tx, ty;
				tx = this.invertTransformText.TransformPointX(x, y);
				ty = this.invertTransformText.TransformPointY(x, y);
				if (!(e.Type === AscCommon.g_mouse_event_type_up && this.rightButtonFlag)) {
					content.Selection_SetEnd(tx, ty, slideIndex - content.Get_StartPage_Relative(), e);
				}
			}
			delete this.rightButtonFlag;
		};

		CShape.prototype.Get_Theme = function () {
			return this.getParentObjects().theme;
		};

		CShape.prototype.updateSelectionState = function () {
			var drawing_document = this.getDrawingDocument();
			if (drawing_document) {
				var content = this.getDocContent();
				if (content) {
					var oLogicDocument = content.GetLogicDocument ? content.GetLogicDocument() : null;

					var oMatrix = null;
					if (this.transformText) {
						oMatrix = this.transformText.CreateDublicate();
					}
					drawing_document.UpdateTargetTransform(oMatrix);
					if (true === content.IsSelectionUse()) {
						// Выделение нумерации
						if (selectionflag_Numbering == content.Selection.Flag) {
							drawing_document.TargetEnd();
							drawing_document.SelectEnabled(true);
							drawing_document.SelectClear();
							drawing_document.SelectShow();
						}
						// Обрабатываем движение границы у таблиц
						else if (null != content.Selection.Data && true === content.Selection.Data.TableBorder && type_Table == content.Content[content.Selection.Data.Pos].GetType()) {
							// Убираем курсор, если он был
							drawing_document.TargetEnd();
						} else {
							if (false === content.IsSelectionEmpty()) {
								drawing_document.TargetEnd();
								drawing_document.SelectEnabled(true);
								drawing_document.SelectClear();
								drawing_document.SelectShow();
							} else {
								if (true !== content.Selection.Start) {
									content.RemoveSelection();
								}
								drawing_document.SelectEnabled(false);
								content.RecalculateCurPos();

								drawing_document.TargetStart();
								drawing_document.TargetShow();

								if (oLogicDocument && oLogicDocument.IsFillingFormMode && oLogicDocument.IsFillingFormMode()) {
									var oContentControl = oLogicDocument.GetContentControl();
									if (oContentControl && oContentControl.IsCheckBox())
										drawing_document.TargetEnd();
								}
							}
						}
					} else {
						drawing_document.SelectEnabled(false);
						content.RecalculateCurPos();

						drawing_document.TargetStart();
						drawing_document.TargetShow();

						if (oLogicDocument && oLogicDocument.IsFillingFormMode && oLogicDocument.IsFillingFormMode()) {
							var oContentControl = oLogicDocument.GetContentControl();
							if (oContentControl && oContentControl.IsCheckBox())
								drawing_document.TargetEnd();
						}
					}
				} else {
					drawing_document.UpdateTargetTransform(new CMatrix());
					drawing_document.TargetEnd();
					drawing_document.SelectEnabled(false);
					drawing_document.SelectClear();
					drawing_document.SelectShow();
				}
			}
		};

		CShape.prototype.check_bounds = function (checker) {
			if (this.spPr && this.spPr.geometry) {
				this.spPr.geometry.check_bounds(checker);
			} else {
				checker._s();
				checker._m(0, 0);
				checker._l(this.extX, 0);
				checker._l(this.extX, this.extY);
				checker._l(0, this.extY);
				checker._z();
				checker._e();
			}
		};


		CShape.prototype.haveSelectedDrawingInContent = function () {
			if (this.bWordShape) {
				var aAllDrawings = this.recalcInfo.AllDrawings;
				for (var i = 0; i < aAllDrawings.length; ++i) {
					if (aAllDrawings[i] && aAllDrawings[i].GraphicObj && aAllDrawings[i].GraphicObj.selected) {
						return true;
					}
				}
			}
			return false;
		};


		CShape.prototype.clipTextRect = function (graphics, transform, transformText, pageIndex) {
			if (this.clipRect) {
				var transform_ = transform ? transform : this.transform;
				var transformText_ = transformText ? transformText : this.transformText;
				var clip_rect = this.clipRect;
				var oBodyPr = this.getBodyPr();
				if (!this.bWordShape) {
					if (oBodyPr.vertOverflow === AscFormat.nVOTOverflow) {
						return;
					}
				}
				if (!oBodyPr || !oBodyPr.upright) {
					graphics.transform3(transform_);
					graphics.AddClipRect(clip_rect.x, clip_rect.y, clip_rect.w, clip_rect.h);

					graphics.SetIntegerGrid(false);
					graphics.transform3(transformText_, true);
				} else {
					var oTransform = new CMatrix();
					var cX = transform_.TransformPointX(this.extX / 2, this.extY / 2);
					var cY = transform_.TransformPointY(this.extX / 2, this.extY / 2);

					if (checkNormalRotate(this.rot)) {
						oTransform.tx = cX - this.extX / 2;
						oTransform.ty = cY - this.extY / 2;
					} else {
						global_MatrixTransformer.TranslateAppend(oTransform, -this.extX / 2, -this.extY / 2);
						global_MatrixTransformer.RotateRadAppend(oTransform, Math.PI / 2);
						global_MatrixTransformer.TranslateAppend(oTransform, cX, cY);
					}
					graphics.transform3(oTransform, true);
					graphics.AddClipRect(clip_rect.x, clip_rect.y, clip_rect.w, clip_rect.h);

					graphics.SetIntegerGrid(false);
					graphics.transform3(transformText_, true);
				}
			}
		};

		CShape.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {

			if (this.checkNeedRecalculate && this.checkNeedRecalculate()) {
				return;
			}

			var oUR = graphics.updatedRect;
			if (oUR && this.bounds) {
				if (!oUR.isIntersectOther(this.bounds)) {
					return;
				}
			}
			if (graphics.animationDrawer) {
				graphics.animationDrawer.drawObject(this, graphics);
				return;
			}
			var options = opt || {};
			var _transform = transform ? transform : this.transform;
			var _transform_text = transformText ? transformText : this.transformText;
			var _transform_text2 = options.transformText2 || this.transformText2;
			var geometry = this.calcGeometry || this.spPr && this.spPr.geometry;

			if (graphics.IsSlideBoundsCheckerType === true) {

				this.drawShdw && this.drawShdw(graphics);
				graphics.transform3(_transform);
				if (!this.spPr || null == geometry || geometry.isEmpty() || !graphics.IsShapeNeedBounds(geometry.preset)) {
					graphics._s();
					graphics._m(0, 0);
					graphics._l(this.extX, 0);
					graphics._l(this.extX, this.extY);
					graphics._l(0, this.extY);
					graphics._e();
				} else {
					geometry.check_bounds(graphics, this);
				}

				if (this.txBody) {
					graphics.SetIntegerGrid(false);

					var transform_text;
					if ((!this.txBody.content || this.txBody.content.Is_Empty()) && this.txBody.content2 != null && !this.txBody.checkCurrentPlaceholder() && (this.isEmptyPlaceholder ? this.isEmptyPlaceholder() : false) && this.transformText2) {
						transform_text = _transform_text2;
					} else if (this.txBody.content) {
						transform_text = _transform_text;
					}

					graphics.transform3(transform_text);

					if (graphics.CheckUseFonts2 !== undefined)
						graphics.CheckUseFonts2(transform_text);
					this.txBody.draw(graphics);
					if (graphics.UncheckUseFonts2 !== undefined)
						graphics.UncheckUseFonts2(transform_text);
					graphics.SetIntegerGrid(true);
				}

				graphics.reset();
				return;
			}

			if (graphics.StartDrawShape) {
				graphics.StartDrawShape(undefined, this.isForm && this.isForm() ? true : false);
			}
			var oClipRect;
			if (!graphics.IsSlideBoundsCheckerType && this.getClipRect) {
				oClipRect = this.getClipRect();
			}
			if (oClipRect) {
				graphics.SaveGrState();
				graphics.AddClipRect(oClipRect.x, oClipRect.y, oClipRect.w, oClipRect.h);
			}
			this.drawShdw && this.drawShdw(graphics);
			var _oldBrush = this.brush;
			if (this.signatureLine) {
				var sSignatureUrl = null;

				var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
				if (_editor) {
					sSignatureUrl = _editor.asc_getSignatureImage(this.signatureLine.id);
				}
				if (typeof sSignatureUrl === "string" && sSignatureUrl.length > 0) {
					this.brush = AscFormat.CreateBlipFillUniFillFromUrl(sSignatureUrl);
				}
			}

			if ((geometry || (this.getObjectType && (this.getObjectType() === AscDFH.historyitem_type_DLbl || this.getObjectType() === AscDFH.historyitem_type_Title || this.getObjectType() === AscDFH.historyitem_type_Legend))) && (this.style || (this.brush && this.brush.fill) || (this.pen && this.pen.Fill && this.pen.Fill.fill))) {
				graphics.SetIntegerGrid(false);
				graphics.transform3(_transform, false);

				var shape_drawer = new AscCommon.CShapeDrawer();
				shape_drawer.fromShape2(this, graphics, geometry);
				shape_drawer.draw(geometry);
			}

			if (!graphics.isSmartArtPreviewDrawer && !graphics.RENDERER_PDF_FLAG && !this.bWordShape && this.isEmptyPlaceholder() && !(this.parent && this.parent.kind === AscFormat.TYPE_KIND.NOTES) && !(this.pen && this.pen.Fill && this.pen.Fill.fill && !(this.pen.Fill.fill instanceof AscFormat.CNoFill)) && graphics.IsNoDrawingEmptyPlaceholder !== true && !AscCommon.IsShapeToImageConverter) {
				var drawingObjects = this.getDrawingObjectsController();
				if (typeof editor !== "undefined" && editor && graphics.m_oContext !== undefined && graphics.m_oContext !== null && graphics.IsTrack === undefined && (!drawingObjects || AscFormat.getTargetTextObject(drawingObjects) !== this)) {
					var angle = _transform.GetRotation();
					if (AscFormat.fApproxEqual(angle, 0.0, 0.0) ||
						AscFormat.fApproxEqual(angle, 90.0, 0.0) ||
						AscFormat.fApproxEqual(angle, 180.0, 0.0) ||
						AscFormat.fApproxEqual(angle, 270.0, 0.0)) {
						graphics.transform3(_transform, false);
						var tr = graphics.m_oFullTransform;
						graphics.SetIntegerGrid(true);

						var _x = tr.TransformPointX(0, 0);
						var _y = tr.TransformPointY(0, 0);
						var _r = tr.TransformPointX(this.extX, this.extY);
						var _b = tr.TransformPointY(this.extX, this.extY);

						var __x = Math.min(_x, _r);
						var __y = Math.min(_y, _b);
						var __r = Math.max(_x, _r);
						var __b = Math.max(_y, _b);
						graphics.m_oContext.lineWidth = 1;
						graphics.p_color(127, 127, 127, 255);

						graphics._s();
						editor.WordControl.m_oDrawingDocument.AutoShapesTrack.AddRectDashClever(graphics.m_oContext, __x >> 0, __y >> 0, __r >> 0, __b >> 0, 2, 2, true);
						graphics._s();
					} else {
						graphics.transform3(_transform, false);
						var tr = graphics.m_oFullTransform;
						graphics.SetIntegerGrid(true);

						var _r = this.extX;
						var _b = this.extY;

						var x1 = tr.TransformPointX(0, 0) >> 0;
						var y1 = tr.TransformPointY(0, 0) >> 0;

						var x2 = tr.TransformPointX(_r, 0) >> 0;
						var y2 = tr.TransformPointY(_r, 0) >> 0;

						var x3 = tr.TransformPointX(0, _b) >> 0;
						var y3 = tr.TransformPointY(0, _b) >> 0;

						var x4 = tr.TransformPointX(_r, _b) >> 0;
						var y4 = tr.TransformPointY(_r, _b) >> 0;

						graphics.m_oContext.lineWidth = 1;
						graphics.p_color(127, 127, 127, 255);

						graphics._s();
						editor.WordControl.m_oDrawingDocument.AutoShapesTrack.AddRectDash(graphics.m_oContext, x1, y1, x2, y2, x3, y3, x4, y4, 3, 1, true);
						graphics._s();
					}
				} else {
					graphics.SetIntegerGrid(false);
					graphics.p_width(70);
					graphics.transform3(_transform, false);
					graphics.p_color(0, 0, 0, 255);
					graphics._s();
					graphics._m(0, 0);
					graphics._l(this.extX, 0);
					graphics._l(this.extX, this.extY);
					graphics._l(0, this.extY);
					graphics._z();
					graphics.ds();

					graphics.SetIntegerGrid(true);
				}
			}
			this.brush = _oldBrush;
			var oController = this.getDrawingObjectsController && this.getDrawingObjectsController();

			if (!this.cropObject) {
				if (!this.txWarpStruct && !this.txWarpStructParamarksNoTransform || (!this.txWarpStructParamarksNoTransform && oController && (AscFormat.getTargetTextObject(oController) === this) || (!this.txBody && !this.textBoxContent)) /*|| this.haveSelectedDrawingInContent()*/) {
					if (this.txBody) {
						graphics.SaveGrState();
						graphics.SetIntegerGrid(false);
						var transform_text;
						if ((!this.txBody.content || this.txBody.content.Is_Empty()) && !AscCommon.IsShapeToImageConverter && this.txBody.content2 != null && !this.txBody.checkCurrentPlaceholder() && (this.isEmptyPlaceholder ? this.isEmptyPlaceholder() : false) && this.transformText2) {
							transform_text = this.transformText2;
						} else if (this.txBody.content) {
							transform_text = _transform_text;
						}

						if (this instanceof CShape) {
							if (!(oController && (AscFormat.getTargetTextObject(oController) === this)))
								this.clipTextRect(graphics, transform, transformText, pageIndex);
						}
						graphics.transform3(transform_text, true);
						if (graphics.CheckUseFonts2 !== undefined)
							graphics.CheckUseFonts2(transform_text);

						graphics.SetIntegerGrid(true);
						this.txBody.draw(graphics);
						if (graphics.UncheckUseFonts2 !== undefined)
							graphics.UncheckUseFonts2(transform_text);
						graphics.RestoreGrState();
					}

					if (this.textBoxContent && !graphics.IsNoSupportTextDraw && this.transformText) {
						var old_start_page = this.textBoxContent.Get_StartPage_Relative();
						this.textBoxContent.Set_StartPage(pageIndex);

						graphics.SaveGrState();
						graphics.SetIntegerGrid(false);
						this.clipTextRect(graphics, transform, transformText, pageIndex);
						var result_page_index = AscFormat.isRealNumber(graphics.shapePageIndex) ? graphics.shapePageIndex : old_start_page;

						if (graphics.CheckUseFonts2 !== undefined)
							graphics.CheckUseFonts2(this.transformText, this.isForm && this.isForm() ? true : false);

						if (AscCommon.IsShapeToImageConverter) {
							this.textBoxContent.Set_StartPage(0);
							result_page_index = 0;
						}


						this.textBoxContent.Set_StartPage(result_page_index);
						this.textBoxContent.Draw(result_page_index, graphics);

						if (graphics.UncheckUseFonts2 !== undefined)
							graphics.UncheckUseFonts2();

						this.textBoxContent.Set_StartPage(old_start_page);
						graphics.RestoreGrState();
					}
				} else {
					var oTheme = this.getParentObjects().theme;
					var oColorMap = this.Get_ColorMap();
					if (!this.bWordShape && (!this.txBody.content || this.txBody.content.Is_Empty()) && !AscCommon.IsShapeToImageConverter && this.txBody.content2 != null && !this.txBody.checkCurrentPlaceholder() && (this.isEmptyPlaceholder ? this.isEmptyPlaceholder() : false)) {
						if (graphics.IsNoDrawingEmptyPlaceholder !== true && graphics.IsNoDrawingEmptyPlaceholderText !== true && !graphics.RENDERER_PDF_FLAG) {
							if (editor && editor.ShowParaMarks) {
								this.txWarpStructParamarks2.draw(graphics, this.transformTextWordArt2, oTheme, oColorMap);
							} else {
								if (this.txWarpStruct2)
									this.txWarpStruct2.draw(graphics, this.transformTextWordArt2, oTheme, oColorMap);
							}
						}
					} else {

						var oContent = this.getDocContent();
						var result_page_index = AscFormat.isRealNumber(graphics.shapePageIndex) ? graphics.shapePageIndex : (oContent ? oContent.Get_StartPage_Relative() : 0);
						graphics.PageNum = result_page_index;
						var bNeedRestoreState = false;
						var bEditTextArt = isRealObject(oController) && (AscFormat.getTargetTextObject(oController) === this);
						if (this.bWordShape && this.clipRect /*&& (!this.bodyPr.prstTxWarp || this.bodyPr.prstTxWarp.preset === "textNoShape" || bEditTextArt)*/) {
							bNeedRestoreState = true;
							var clip_rect = this.clipRect;
							if (!this.bodyPr.upright) {
								graphics.SaveGrState();
								graphics.SetIntegerGrid(false);
								graphics.transform3(this.transform);
								graphics.AddClipRect(clip_rect.x, clip_rect.y, clip_rect.w, clip_rect.h);
							} else {
								graphics.SaveGrState();
								graphics.SetIntegerGrid(false);
								graphics.transform3(this.transformText, true);
								graphics.AddClipRect(clip_rect.x, clip_rect.y, clip_rect.w, clip_rect.h);
							}
						}

						var oTransform = this.transformTextWordArt;
						if (editor && editor.ShowParaMarks) {
							if (bEditTextArt && this.txWarpStructParamarksNoTransform) {
								this.txWarpStructParamarksNoTransform.draw(graphics, this.transformText, oTheme, oColorMap);
							} else if (this.txWarpStructParamarks) {
								this.txWarpStructParamarks.draw(graphics, oTransform, oTheme, oColorMap);
								if (this.checkNeedRecalcDocContentForTxWarp(this.bodyPr)) {
									if (this.txWarpStructParamarksNoTransform) {
										this.txWarpStructParamarksNoTransform.drawComments(graphics, undefined, oTransform);
									}
								}
							}
						} else {
							if (bEditTextArt && this.txWarpStructNoTransform) {
								this.txWarpStructNoTransform.draw(graphics, this.transformText, oTheme, oColorMap);
							} else if (this.txWarpStruct) {
								this.txWarpStruct.draw(graphics, oTransform, oTheme, oColorMap);
								if (this.checkNeedRecalcDocContentForTxWarp(this.bodyPr)) {
									if (this.txWarpStructNoTransform) {
										this.txWarpStructNoTransform.drawComments(graphics, undefined, oTransform);
									}
								}
							}
						}
						delete graphics.PageNum;
						if (bNeedRestoreState) {
							graphics.RestoreGrState();
						}
					}
				}
			}
			this.drawLocks && this.drawLocks(_transform, graphics);
			if (oClipRect) {
				graphics.RestoreGrState();
			}
			//if(this.txXfrm && this.group) {
			//    graphics.SetIntegerGrid(false);
			//    _transform = new AscCommon.CMatrix();
			//    _transform.tx = this.txXfrm.offX + this.txXfrm.extX / 2 + this.group.transform.tx;
			//    _transform.ty = this.txXfrm.offY + this.txXfrm.extY / 2 + this.group.transform.ty;
			//    graphics.transform3(_transform, false);
			//    graphics.b_color1(255, 0, 0, 255);
			//    graphics.rect(-2, -2, 4, 4);
			//    graphics.df();
			//
			//    graphics.p_color(255, 0, 0, 255);
			//    graphics.rect(-this.txXfrm.extX / 2, -this.txXfrm.extY / 2, this.txXfrm.extX, this.txXfrm.extY);
			//    graphics.ds();
			//}
			graphics.SetIntegerGrid(true);
			graphics.reset();
			if (graphics.EndDrawShape) {
				graphics.EndDrawShape();
			}
		};

		CShape.prototype.recalculateGeometry = function () {
			this.calcGeometry = null;
			if (isRealObject(this.spPr && this.spPr.geometry)) {
				this.calcGeometry = this.spPr.geometry;
			} else {
				if (this.getHierarchy) {
					var hierarchy = this.getHierarchy();
					for (var i = 0; i < hierarchy.length; ++i) {
						if (hierarchy[i] && hierarchy[i].spPr && hierarchy[i].spPr.geometry) {
							var _g = hierarchy[i].spPr.geometry;
							this.calcGeometry = AscFormat.ExecuteNoHistory(function () {
								var _r = _g.createDuplicate();
								_r.setParent(this);
								return _r;
							}, this, []);
							break;
						}
					}
				}
			}
			if (isRealObject(this.calcGeometry)) {
				var transform = this.getTransform();
				this.calcGeometry.Recalculate(transform.extX, transform.extY);
			}
		};

		CShape.prototype.getRotateAngle = function (x, y) {
			var transform = this.getTransformMatrix();
			var rotate_distance = this.convertPixToMM(AscCommon.TRACK_DISTANCE_ROTATE);
			var hc = this.extX * 0.5;
			var vc = this.extY * 0.5;
			var xc_t = transform.TransformPointX(hc, vc);
			var yc_t = transform.TransformPointY(hc, vc);
			var rot_x_t = transform.TransformPointX(hc, -rotate_distance);
			var rot_y_t = transform.TransformPointY(hc, -rotate_distance);

			var invert_transform = this.getInvertTransform();
			if (!invert_transform) {
				return 0.0;
			}
			var rel_x = invert_transform.TransformPointX(x, y);

			var v1_x, v1_y, v2_x, v2_y;
			v1_x = x - xc_t;
			v1_y = y - yc_t;

			v2_x = rot_x_t - xc_t;
			v2_y = rot_y_t - yc_t;

			var flip_h = this.getFullFlipH();
			var flip_v = this.getFullFlipV();
			var same_flip = flip_h && flip_v || !flip_h && !flip_v;
			var angle = rel_x > this.extX * 0.5 ? Math.atan2(Math.abs(v1_x * v2_y - v1_y * v2_x), v1_x * v2_x + v1_y * v2_y) : -Math.atan2(Math.abs(v1_x * v2_y - v1_y * v2_x), v1_x * v2_x + v1_y * v2_y);
			return same_flip ? angle : -angle;
		};

		CShape.prototype.getInvertTransform = function () {
			return this.invertTransform ? this.invertTransform : new CMatrix();
		};


		CShape.prototype.getFullOffset = function () {
			if (!isRealObject(this.group))
				return {offX: this.x, offY: this.y};
			var group_offset = this.group.getFullOffset();
			return {offX: this.x + group_offset.offX, offY: this.y + group_offset.offY};
		};


		CShape.prototype.getTextArtProperties = function () {
			var oContent = this.getDocContent(), oTextPr, oRet = null;
			if (oContent) {
				oRet = {Fill: undefined, Line: undefined, Form: undefined};
				var oController = this.getDrawingObjectsController();
				// if(oController)
				{
					//var oTargetDocContent = oController.getTargetDocContent();
					//if(oTargetDocContent === oContent)
					//{
					//    oTextPr = oContent.GetCalculatedTextPr();
					//}
					//else
					//{
					//    oContent.SetApplyToAll(true);
					//    oTextPr = oContent.GetCalculatedTextPr();
					//    oContent.SetApplyToAll(false);
					//}
					//if(oTextPr.TextFill)
					//{
					//    oRet.Fill = oTextPr.TextFill;
					//}
					//else if(oTextPr.Unifill)
					//{
					//    oRet.Fill = oTextPr.Unifill;
					//}
					//else if(oTextPr.Color)
					//{
					//    oRet.Fill = CreateUnfilFromRGB(oTextPr.Color.r, oTextPr.Color.g, oTextPr.Color.b);
					//}
					//oRet.Line = oTextPr.TextOutline;
					var oBodyPr = this.getBodyPr();
					if (oBodyPr && oBodyPr.prstTxWarp) {
						oRet.Form = oBodyPr.prstTxWarp.preset;
					} else {
						oRet.Form = "textNoShape";
					}
				}
			}
			return oRet;
		};

		CShape.prototype.applyTextArtForm = function (sPreset) {
			var oBodyPr = this.getBodyPr().createDuplicate();
			oBodyPr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry(sPreset);
			if (this.bWordShape) {
				this.setBodyPr(oBodyPr);
			} else {
				if (this.txBody) {
					this.txBody.setBodyPr(oBodyPr);
				}
			}
		};

		CShape.prototype.getParagraphParaPr = function () {
			if (this.txBody && this.txBody.content) {
				var _result;
				this.txBody.content.SetApplyToAll(true);
				_result = this.txBody.content.GetCalculatedParaPr();
				this.txBody.content.SetApplyToAll(false);
				return _result;
			}
			return null;
		};

		CShape.prototype.getParagraphTextPr = function () {
			if (this.txBody && this.txBody.content) {
				var _result;
				this.txBody.content.SetApplyToAll(true);
				_result = this.txBody.content.GetCalculatedTextPr();
				this.txBody.content.SetApplyToAll(false);
				return _result;
			}
			return null;
		};

		CShape.prototype.getImageFromBulletsMap = function (oImages) {
			const oContent = this.getDocContent();
			if (!oContent) {
				return;
			}
			const aParagraphs = oContent.Content;
			for (let nPar = 0; nPar < aParagraphs.length; ++nPar) {
				var oPr = aParagraphs[nPar].Pr;
				if (oPr.Bullet) {
					const sImageId = oPr.Bullet.getImageBulletURL();
					if (sImageId) {
						oImages[sImageId] = true;
					}
				}
			}
		};

		CShape.prototype.getDocContentsWithImageBullets = function (arrContents) {
			const oContent = this.getDocContent();
			if (!oContent) {
				return;
			}
			const aParagraphs = oContent.Content;
			for (let nPar = 0; nPar < aParagraphs.length; ++nPar) {
				var oPr = aParagraphs[nPar].Pr;
				if (oPr.Bullet) {
					const sImageId = oPr.Bullet.getImageBulletURL();
					if (sImageId) {
						arrContents.push(oContent);
						break;
					}
				}
			}
		}

		CShape.prototype.getAllRasterImages = function (images) {
			if (this.spPr && this.spPr.Fill && this.spPr.Fill.fill && typeof (this.spPr.Fill.fill.RasterImageId) === "string" && this.spPr.Fill.fill.RasterImageId.length > 0)
				images.push(this.spPr.Fill.fill.RasterImageId);


			var compiled_style = this.getCompiledStyle();
			var parents = this.getParentObjects();
			if (isRealObject(parents.theme) && isRealObject(compiled_style) && isRealObject(compiled_style.fillRef)) {
				var brush = parents.theme.getFillStyle(compiled_style.fillRef.idx, compiled_style.fillRef.Color);
				if (brush && brush.fill && typeof (brush.fill.RasterImageId) === "string" && brush.fill.RasterImageId.length > 0) {
					images.push(brush.fill.RasterImageId);
				}
			}
			var oContent = this.getDocContent();
			if (oContent) {
				if (this.bWordShape) {
					var drawings = oContent.GetAllDrawingObjects();
					for (var i = 0; i < drawings.length; ++i) {
						drawings[i].GraphicObj && drawings[i].GraphicObj.getAllRasterImages && drawings[i].GraphicObj.getAllRasterImages(images);
					}
				} else {
					oContent.getBulletImages(images);
				}
				var fCallback = function (oRun) {
					var oTextPr = oRun && oRun.Pr;
					if (oTextPr && oTextPr.Unifill && oTextPr.Unifill.fill && oTextPr.Unifill.fill.type == c_oAscFill.FILL_TYPE_BLIP) {
						images.push(oTextPr.Unifill.fill.RasterImageId);
					}
					return false;
				};

				oContent.CheckRunContent(fCallback);
			}
		};
		CShape.prototype.getAllDocContents = function (aDocContents) {
			if (this.textBoxContent) {
				aDocContents.push(this.textBoxContent);
			}
		};

		CShape.prototype.checkRunContent = function (fCallback) {
			let oContent = this.getDocContent();
			if (oContent) {
				oContent.CheckRunContent(fCallback);
			}
		};

		CShape.prototype.changePositionInSmartArt = function (newX, newY) {
			if (this.isObjectInSmartArt()) {
				var point = this.getSmartArtShapePoint();
				if (point) {
					var prSet = point.getPrSet();
					if (prSet) {
						var originalPosX;
						var originalPosY;
						var defaultExtX;
						var defaultExtY;
						var isNormalRotate = AscFormat.checkNormalRotate(this.getDefaultRotSA());
						if (isNormalRotate) {
							originalPosX = this.x;
							originalPosY = this.y;
							defaultExtX = this.extX;
							defaultExtY = this.extY;
						} else {
							originalPosX = this.x + (this.extX - this.extY) / 2;
							originalPosY = this.y + (this.extY - this.extX) / 2;
							defaultExtX = this.extY;
							defaultExtY = this.extX;
						}


						if (prSet) {
							if (prSet.custScaleX) {
								defaultExtX /= prSet.custScaleX;
							}
							if (prSet.custScaleY) {
								defaultExtY /= prSet.custScaleY;
							}
							if (prSet.custLinFactNeighborX) {
								originalPosX -= (prSet.custLinFactNeighborX) * defaultExtX;
							}
							if (prSet.custLinFactNeighborY) {
								originalPosY -= (prSet.custLinFactNeighborY) * defaultExtY;
							}
							if (prSet.custLinFactX) {
								originalPosX -= (prSet.custLinFactX) * defaultExtX;
							}
							if (prSet.custLinFactY) {
								originalPosY -= (prSet.custLinFactY) * defaultExtY;
							}
							if (this.x !== newX) {
								if (prSet.custLinFactNeighborX) {
									prSet.setCustLinFactNeighborX(null);
								}
								prSet.setCustLinFactX(((newX - originalPosX) / defaultExtX));
							}
							if (this.y !== newY) {
								if (prSet.custLinFactNeighborY) {
									prSet.setCustLinFactNeighborY(null);
								}
								prSet.setCustLinFactY(((newY - originalPosY) / defaultExtY));
							}
						}
					}
				}
			}
		};

		CShape.prototype.changePresetGeom = function (sPreset) {


			if (sPreset === "textRect") {
				if (this.isObjectInSmartArt()) {
					return;
				}
				this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

				if (this.bWordShape) {
					if (this.style) {
						this.setStyle(null);
					}
				} else {
					this.setStyle(AscFormat.CreateDefaultTextRectStyle());
				}

				var fill = new AscFormat.CUniFill();
				fill.setFill(new AscFormat.CSolidFill());
				fill.fill.setColor(new AscFormat.CUniColor());
				fill.fill.color.setColor(new AscFormat.CSchemeColor());
				fill.fill.color.color.setId(12);
				this.spPr.setFill(fill);

				var ln = new AscFormat.CLn();
				ln.setW(6350);
				ln.setFill(new AscFormat.CUniFill());
				ln.Fill.setFill(new AscFormat.CSolidFill());
				ln.Fill.fill.setColor(new AscFormat.CUniColor());
				ln.Fill.fill.color.setColor(new AscFormat.CPrstColor());
				ln.Fill.fill.color.color.setId("black");
				this.spPr.setLn(ln);
				if (this.bWordShape) {
					if (!this.textBoxContent) {
						this.setTextBoxContent(new CDocumentContent(this, this.getDrawingDocument(), 0, 0, 0, 0, false, false, false));
						var body_pr = new AscFormat.CBodyPr();
						body_pr.setDefault();
						this.setBodyPr(body_pr);
					}
				} else {
					if (!this.txBody) {
						this.setTxBody(new AscFormat.CTextBody());
						var content = new AscFormat.CDrawingDocContent(this.txBody, this.getDrawingDocument(), 0, 0, 0, 0, false, false, true);
						this.txBody.setParent(this);
						this.txBody.setContent(content);
						var body_pr = new AscFormat.CBodyPr();
						body_pr.setDefault();
						this.txBody.setBodyPr(body_pr);
					}
				}
				return;
			}
			var _final_preset;
			var _old_line;
			var _new_line;


			if (this.spPr.ln == null) {
				_old_line = null;
			} else {
				_old_line = this.spPr.ln.createDuplicate();
			}
			switch (sPreset) {
				case "lineWithArrow": {
					_final_preset = "line";
					if (_old_line == null) {
						_new_line = new AscFormat.CLn();
					} else {
						_new_line = this.spPr.ln.createDuplicate();
					}
					_new_line.tailEnd = new AscFormat.EndArrow();
					_new_line.tailEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.tailEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.tailEnd.w = AscFormat.LineEndSize.Mid;
					break;
				}
				case "lineWithTwoArrows": {
					_final_preset = "line";
					if (_old_line == null) {
						_new_line = new AscFormat.CLn();

					} else {
						_new_line = this.spPr.ln.createDuplicate();
					}
					_new_line.tailEnd = new AscFormat.EndArrow();
					_new_line.tailEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.tailEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.tailEnd.w = AscFormat.LineEndSize.Mid;

					_new_line.headEnd = new AscFormat.EndArrow();
					_new_line.headEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.headEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.headEnd.w = AscFormat.LineEndSize.Mid;
					break;
				}
				case "bentConnector5WithArrow": {
					_final_preset = "bentConnector5";
					if (_old_line == null) {
						_new_line = new AscFormat.CLn();

					} else {
						_new_line = this.spPr.ln.createDuplicate();
					}
					_new_line.tailEnd = new AscFormat.EndArrow();
					_new_line.tailEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.tailEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.tailEnd.w = AscFormat.LineEndSize.Mid;
					break;
				}
				case "bentConnector5WithTwoArrows": {
					_final_preset = "bentConnector5";
					if (_old_line == null) {
						_new_line = new AscFormat.CLn();
					} else {
						_new_line = this.spPr.ln.createDuplicate();
					}
					_new_line.tailEnd = new AscFormat.EndArrow();
					_new_line.tailEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.tailEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.tailEnd.w = AscFormat.LineEndSize.Mid;

					_new_line.headEnd = new AscFormat.EndArrow();
					_new_line.headEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.headEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.headEnd.w = AscFormat.LineEndSize.Mid;
					break;
				}
				case "curvedConnector3WithArrow": {
					_final_preset = "curvedConnector3";
					if (_old_line == null) {
						_new_line = new AscFormat.CLn();
					} else {
						_new_line = this.spPr.ln.createDuplicate();
					}
					_new_line.tailEnd = new AscFormat.EndArrow();
					_new_line.tailEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.tailEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.tailEnd.w = AscFormat.LineEndSize.Mid;
					break;
				}
				case "curvedConnector3WithTwoArrows": {
					_final_preset = "curvedConnector3";
					if (_old_line == null) {
						_new_line = new AscFormat.CLn();

					} else {
						_new_line = this.spPr.ln.createDuplicate();
					}
					_new_line.tailEnd = new AscFormat.EndArrow();
					_new_line.tailEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.tailEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.tailEnd.w = AscFormat.LineEndSize.Mid;

					_new_line.headEnd = new AscFormat.EndArrow();
					_new_line.headEnd.type = AscFormat.LineEndType.Arrow;
					_new_line.headEnd.len = AscFormat.LineEndSize.Mid;
					_new_line.headEnd.w = AscFormat.LineEndSize.Mid;
					break;
				}
				default: {
					_final_preset = sPreset;
					if (_old_line == null) {
						_new_line = new AscFormat.CLn();
					} else {
						_new_line = this.spPr.ln.createDuplicate();
					}
					_new_line.tailEnd = null;

					_new_line.headEnd = null;
					break;
				}
			}
			var point = this.getSmartArtShapePoint();
			if (_final_preset != null) {
				this.spPr.setGeometry(AscFormat.CreateGeometry(_final_preset));
				point && point.setGeometry(AscFormat.CreateGeometry(_final_preset));
			} else {
				this.spPr.setGeometry(null);
				point && point.setGeometry(null);
			}
			if (!this.bWordShape) {
				this.checkExtentsByDocContent();
			}
			var setLine = _new_line;
			if ((!this.brush || !this.brush.fill) && (!this.pen || !this.pen.Fill || !this.pen.Fill.fill)) {
				var new_line2 = new AscFormat.CLn();
				new_line2.Fill = new AscFormat.CUniFill();
				new_line2.Fill.fill = new AscFormat.CSolidFill();
				new_line2.Fill.fill.color = new AscFormat.CUniColor();
				new_line2.Fill.fill.color.color = new AscFormat.CSchemeColor();
				new_line2.Fill.fill.color.color.id = 0;
				if (isRealObject(_new_line)) {
					new_line2.merge(_new_line);
				}
				setLine = new_line2;
			}
			this.spPr.setLn(setLine);
			point && point.setLine(setLine);
		};

		CShape.prototype.changeFill = function (unifill) {

			if (this.recalcInfo.recalculateBrush) {
				this.recalculateBrush();
			}
			var unifill2 = AscFormat.CorrectUniFill(unifill, this.brush, this.getEditorType());
			unifill2.convertToPPTXMods();
			this.setFill(unifill2);
			var point = this.getSmartArtShapePoint();
			if (point) {
				point.setUniFill(unifill2);
			}
		};
		CShape.prototype.changeShadow = function (oShadow) {

			this.spPr && this.spPr.changeShadow(oShadow);

			var point = this.getSmartArtShapePoint();
			point && point.changeShadow(oShadow);
		};
		CShape.prototype.setFill = function (fill) {

			this.spPr.setFill(fill);
		};

		CShape.prototype.changeLine = function (line) {
			if (this.recalcInfo.recalculatePen) {
				this.recalculatePen();
			}
			var stroke = AscFormat.CorrectUniStroke(line, this.pen);
			if (stroke.Fill) {
				stroke.Fill.convertToPPTXMods();
			}
			this.spPr.setLn(stroke);

			var point = this.getSmartArtShapePoint();
			point && point.setLine(stroke);
		};

		CShape.prototype.hitToAdjustment = function (x, y) {
			if (!AscFormat.canSelectDrawing(this)) {
				return false;
			}
			var oApi = Asc.editor || editor;
			var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;

			if (isDrawHandles && this.isForm && this.isForm() && this.getInnerForm() && this.getInnerForm().IsFormLocked())
				isDrawHandles = false;

			if (isDrawHandles === false) {
				return {hit: false, adjPolarFlag: null, adjNum: null, warp: false};
			}
			var invert_transform;
			var t_x, t_y, ret;
			var _calcGeom = this.calcGeometry || (this.spPr && this.spPr.geometry);
			var _dist;
			if (global_mouseEvent && global_mouseEvent.AscHitToHandlesEpsilon) {
				_dist = global_mouseEvent.AscHitToHandlesEpsilon;
			} else {
				_dist = this.convertPixToMM(global_mouseEvent.KoefPixToMM * AscCommon.TRACK_CIRCLE_RADIUS);
			}
			if (_calcGeom) {
				invert_transform = this.getInvertTransform();
				if (!invert_transform) {
					return {hit: false, adjPolarFlag: null, adjNum: null, warp: false};
				}
				t_x = invert_transform.TransformPointX(x, y);
				t_y = invert_transform.TransformPointY(x, y);
				ret = _calcGeom.hitToAdj(t_x, t_y, _dist);
				if (ret.hit) {
					ret.warp = false;
					return ret;
				}
			}
			if (this.recalcInfo.warpGeometry && this.invertTransformTextWordArt) {
				invert_transform = this.invertTransformTextWordArt;
				t_x = invert_transform.TransformPointX(x, y);
				t_y = invert_transform.TransformPointY(x, y);
				ret = this.recalcInfo.warpGeometry.hitToAdj(t_x, t_y, _dist);
				ret.warp = true;
				return ret;
			}

			return {hit: false, adjPolarFlag: null, adjNum: null, warp: false};
		};

		CShape.prototype.hit = function (x, y) {
			return this.hitInInnerArea(x, y) || this.hitInPath(x, y) || this.hitInTextRect(x, y);
		};

		CShape.prototype.hitInPath = function (x, y) {

			var oInnerForm = null;
			if (this.isForm && this.isForm() && (oInnerForm = this.getInnerForm()) && oInnerForm.CanPlaceCursorInside()) {
				var oApi = Asc.editor || editor;
				var oLogicDocument = oApi && oApi.WordControl && oApi.WordControl.m_oLogicDocument ? oApi.WordControl.m_oLogicDocument : null;
				if (oLogicDocument && oLogicDocument.IsDocumentEditor() && oLogicDocument.IsFillingFormMode())
					return false;
			}

			if (!this.checkHitToBounds(x, y))
				return false;
			var invert_transform = this.getInvertTransform();
			if (!invert_transform) {
				return false;
			}
			var x_t = invert_transform.TransformPointX(x, y);
			var y_t = invert_transform.TransformPointY(x, y);
			var oGeometry = this.spPr && this.spPr.geometry || this.calcGeometry;
			if (oGeometry) {
				const dOldDIst = AscFormat.DIST_HIT_IN_LINE;
				if(this.pen) {
					const nW = this.pen.w || 12700;
					const dWidth = nW / 36000;
					const oAPI = Asc.editor;
					if(oAPI.isEraseInkMode() && !oGeometry.preset) {
						AscFormat.DIST_HIT_IN_LINE = dWidth;
					}
					else {
						AscFormat.DIST_HIT_IN_LINE = Math.max(dOldDIst, dWidth);
					}
				}
				let bResult = oGeometry.hitInPath(this.getCanvasContext(), x_t, y_t);
				AscFormat.DIST_HIT_IN_LINE = dOldDIst;
				return bResult;
			}
			else
				return this.hitInBoundingRect(x, y);
			return false;
		};

		CShape.prototype.hitInInnerArea = function (x, y) {
			if ((this.getObjectType && this.getObjectType() === AscDFH.historyitem_type_ChartSpace || this.getObjectType() === AscDFH.historyitem_type_Title) ||
				(this.brush != null && this.brush.isVisible() || this.blipFill || (this.isTextBox && this.isTextBox())) && this.checkHitToBounds(x, y)) {
				var invert_transform = this.getInvertTransform();
				if (!invert_transform) {
					return false;
				}
				var x_t = invert_transform.TransformPointX(x, y);
				var y_t = invert_transform.TransformPointY(x, y);
				var oGeometry = this.spPr && this.spPr.geometry || this.calcGeometry;
				if (isRealObject(oGeometry) && oGeometry.pathLst.length > 0 && !(this.getObjectType && this.getObjectType() === AscDFH.historyitem_type_ChartSpace))
					return oGeometry.hitInInnerArea(this.getCanvasContext(), x_t, y_t);
				if (this.getObjectType() === AscDFH.historyitem_type_Shape) {
					return false;
				}
				return x_t > 0 && x_t < this.extX && y_t > 0 && y_t < this.extY;
			}
			return false;
		};

		CShape.prototype.hitInBoundingRect = function (x, y) {
			if (!AscFormat.canSelectDrawing(this)) {
				return false;
			}
			var oInnerForm = null;
			if (this.isForm && this.isForm() && (oInnerForm = this.getInnerForm()) && oInnerForm.CanPlaceCursorInside()) {
				var oApi = Asc.editor || editor;
				var oLogicDocument = oApi && oApi.WordControl && oApi.WordControl.m_oLogicDocument ? oApi.WordControl.m_oLogicDocument : null;
				if (oLogicDocument && oLogicDocument.IsDocumentEditor() && oLogicDocument.IsFillingFormMode())
					return false;
			}

			if (this.parent && this.parent.kind === AscFormat.TYPE_KIND.NOTES) {
				return false;
			}
			var invert_transform = this.getInvertTransform();
			if (!invert_transform) {
				return false;
			}
			var x_t = invert_transform.TransformPointX(x, y);
			var y_t = invert_transform.TransformPointY(x, y);

			var _hit_context = this.getCanvasContext();

			return !(CheckObjectLine(this)) && (HitInLine(_hit_context, x_t, y_t, 0, 0, this.extX, 0) ||
				HitInLine(_hit_context, x_t, y_t, this.extX, 0, this.extX, this.extY) ||
				HitInLine(_hit_context, x_t, y_t, this.extX, this.extY, 0, this.extY) ||
				HitInLine(_hit_context, x_t, y_t, 0, this.extY, 0, 0) ||
				(this.canRotate && this.canRotate() && HitInLine(_hit_context, x_t, y_t, this.extX * 0.5, 0, this.extX * 0.5, -this.convertPixToMM(AscCommon.TRACK_DISTANCE_ROTATE))));
		};

		CShape.prototype.canRotate = function () {
			if (this.cropObject) {
				return false;
			}
			if (this.signatureLine) {
				return false;
			}
			return AscFormat.CGraphicObjectBase.prototype.canRotate.call(this);
		};


		CShape.prototype.canGroup = function () {
			if (this.isPlaceholder()) {
				return false;
			}
			if(this.isForm()) {
				return false;
			}
			if (this.signatureLine) {
				return false;
			}
			return AscFormat.CGraphicObjectBase.prototype.canGroup.call(this);
		};


		CShape.prototype.createRotateTrack = function () {
			return new AscFormat.RotateTrackShapeImage(this);
		};

		CShape.prototype.createResizeTrack = function (cardDirection, oController) {
			return new AscFormat.ResizeTrackShapeImage(this, cardDirection, oController);
		};

		CShape.prototype.setResizeHeightConstr = function (height) {
			AscFormat.CheckSpPrXfrm(this);
			if (AscFormat.isRealNumber(height)) {
				this.spPr.xfrm.setExtY(height);
			}
		}

		CShape.prototype.setResizeWidthConstr = function (width) {
			AscFormat.CheckSpPrXfrm(this);
			if (AscFormat.isRealNumber(width)) {
				this.spPr.xfrm.setExtX(width);
			}
		}

		CShape.prototype.createMoveTrack = function () {
			return new AscFormat.MoveShapeImageTrack(this);
		};


		CShape.prototype.remove = function (Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord) {
			if (this.txBody) {
				this.txBody.content.Remove(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord);
				this.recalcInfo.recalculateContent = true;
				this.recalcInfo.recalculateTransformText = true;
			}
		};

		CShape.prototype.getWatermarkProps = function () {
			var oProps = new Asc.CAscWatermarkProperties(), oTextPr, oRGBAColor, oInterfaceTextPr, oContent;
			oContent = this.getDocContent();
			oProps.put_Type(Asc.c_oAscWatermarkType.Text);
			oProps.put_IsDiagonal(!AscFormat.fApproxEqual(this.rot, 0.0));
			oContent.SetApplyToAll(true);
			oProps.put_Text(oContent.GetSelectedText(true, {NewLineParagraph: false, NewLine: false}));
			oTextPr = oContent.GetCalculatedTextPr();
			oProps.put_Opacity(255);
			if (!AscFormat.isRealNumber(oTextPr.FontSize) ||
				oTextPr.FontSize < 36 ||
				oTextPr.FontSize - (oTextPr.FontSize >> 0) > 0) {
				oTextPr.FontSize = -1;
			}
			oInterfaceTextPr = new Asc.CTextProp(oTextPr);
			if (oTextPr.TextFill) {
				oTextPr.TextFill.check(this.Get_Theme(), this.Get_ColorMap());
				if (oTextPr.TextFill.fill && oTextPr.TextFill.fill.type === c_oAscFill.FILL_TYPE_SOLID && oTextPr.TextFill.fill.color) {
					oInterfaceTextPr.put_Color(AscCommon.CreateAscColor(oTextPr.TextFill.fill.color));
				} else {
					oRGBAColor = oTextPr.TextFill.getRGBAColor();
					oInterfaceTextPr.put_Color(AscCommon.CreateAscColorCustom(oRGBAColor.R, oRGBAColor.G, oRGBAColor.B, false));
				}
				oProps.put_Opacity(AscFormat.isRealNumber(oTextPr.TextFill.transparent) ? oTextPr.TextFill.transparent : 255);
			}
			oProps.put_TextPr(oInterfaceTextPr);
			oContent.SetApplyToAll(false);
			return oProps;
		};

		CShape.prototype.RestartSpellCheck = function () {
			this.recalcInfo.recalculateShapeStyleForParagraph = true;
			var content = this.getDocContent();
			content && content.RestartSpellCheck();
		};

		CShape.prototype.Refresh_RecalcData = function (data) {
			switch (data.Type) {
				case AscDFH.historyitem_AutoShapes_SetDrawingBaseCoors: {
					break;
				}
				case AscDFH.historyitem_AutoShapes_RemoveFromDrawingObjects: {
					break;
				}

				case AscDFH.historyitem_AutoShapes_AddToDrawingObjects: {
					break;
				}
				case AscDFH.historyitem_AutoShapes_SetWorksheet: {
					break;
				}
				case AscDFH.historyitem_ShapeSetBDeleted: {
					break;
				}
				case AscDFH.historyitem_ShapeSetNvSpPr: {
					break;
				}
				case AscDFH.historyitem_ShapeSetSpPr: {
					break;
				}
				case AscDFH.historyitem_ShapeSetStyle: {
					break;
				}
				case AscDFH.historyitem_ShapeSetTxBody: {
					this.Refresh_RecalcData2();
					break;
				}
				case AscDFH.historyitem_ShapeSetTextBoxContent: {
					this.Refresh_RecalcData2();
					break;
				}
				case AscDFH.historyitem_ShapeSetParent: {
					break;
				}
				case AscDFH.historyitem_ShapeSetGroup: {
					break;
				}
				case AscDFH.historyitem_ShapeSetBodyPr: {
					this.Refresh_RecalcData2();
					break;
				}
				case AscDFH.historyitem_ShapeSetWordShape: {
					break;
				}
				default: {
					this.Refresh_RecalcData2();
				}
			}
		};

		CShape.prototype.Refresh_RecalcData2 = function (pageIndex/*для текста*/) {
			this.recalcContent();
			this.recalcContent2 && this.recalcContent2();
			this.recalcTransformText();
			this.addToRecalculate();
			var oController = this.getDrawingObjectsController();
			if (oController && AscFormat.getTargetTextObject(oController) === this) {
				this.recalcInfo.recalcTitle = this.getDocContent();
				this.recalcInfo.bRecalculatedTitle = true;
			}
			if (this.parent && this.parent.getObjectType && this.parent.getObjectType() === AscDFH.historyitem_type_Notes) {
				if (this.parent.slide && this.parent.slide.addToRecalculate) {
					this.parent.slide.addToRecalculate();
				}
			}
		};


		CShape.prototype.Load_LinkData = function (linkData) {
		};

		CShape.prototype.Get_PageContentStartPos = function (pageNum) {
			if (this.textBoxContent) {
				if (this.getTextRect) {
					var rect = this.getTextRect();
					return {X: 0, Y: 0, XLimit: rect.r - rect.l, YLimit: 20000};
				} else {
					return {X: 0, Y: 0, XLimit: this.extX, YLimit: 20000};
				}
			}
			return null;
		};

		CShape.prototype.OnContentRecalculate = function () {
		};

		CShape.prototype.recalculateBounds = function () {
			var boundsChecker = new AscFormat.CSlideBoundsChecker();
			this.draw(boundsChecker, this.localTransform, this.localTransformText, undefined, {transformText2: this.localTransformText2});
			this.bounds.l = boundsChecker.Bounds.min_x;
			this.bounds.t = boundsChecker.Bounds.min_y;
			this.bounds.r = boundsChecker.Bounds.max_x;
			this.bounds.b = boundsChecker.Bounds.max_y;

			this.bounds.x = this.bounds.l;
			this.bounds.y = this.bounds.t;
			this.bounds.w = this.bounds.r - this.bounds.l;
			this.bounds.h = this.bounds.b - this.bounds.t;
		};

		CShape.prototype.checkContentWordArt = function (oContent) {
			if (!oContent) {
				return false;
			}
			return oContent.CheckRunContent(CheckWordArtTextPr);
		};

		CShape.prototype.checkNeedRecalcDocContentForTxWarp = function (oBodyPr) {
			return oBodyPr && oBodyPr.prstTxWarp && (oBodyPr.prstTxWarp.pathLst.length / 2 - ((oBodyPr.prstTxWarp.pathLst.length / 2) >> 0) > 0);
		};

		CShape.prototype.chekBodyPrTransform = function (oBodyPr) {
			return isRealObject(oBodyPr) && isRealObject(oBodyPr.prstTxWarp) && oBodyPr.prstTxWarp.preset !== "textNoShape";
		};

		CShape.prototype.checkTextWarp = function (oContent, oBodyPr, dWidth, dHeight, bNeedNoTransform, bNeedWarp) {
			return AscFormat.ExecuteNoHistory(function () {
				var oRet = {
					oTxWarpStruct: null,
					oTxWarpStructParamarks: null,
					oTxWarpStructNoTransform: null,
					oTxWarpStructParamarksNoTransform: null
				};
				//return oRet;
				var bTransform = this.chekBodyPrTransform(oBodyPr) && bNeedWarp;
				var warpGeometry = oBodyPr.prstTxWarp;
				warpGeometry && warpGeometry.Recalculate(dWidth, dHeight);
				this.recalcInfo.warpGeometry = warpGeometry;
				var bCheckWordArtContent = this.checkContentWordArt(oContent);
				var bColumns = oContent.Get_ColumnsCount() > 1;
				var bContentRecalculated = false;
				if (bTransform || bCheckWordArtContent) {
					var bNeedRecalc = this.checkNeedRecalcDocContentForTxWarp(oBodyPr), dOneLineWidth,
						dMinPolygonLength = 0, dKoeff = 1;
					var oTheme = this.Get_Theme(), oColorMap = this.Get_ColorMap();
					var oTextDrawer = new AscFormat.CTextDrawer(dWidth, dHeight, true, oTheme, bNeedRecalc);
					oTextDrawer.bCheckLines = bTransform && bNeedWarp;
					var oContentToDraw = oContent;
					if (bNeedRecalc && bNeedWarp) {
						oContentToDraw = oContent.Copy(oContent.Parent, oContent.DrawingDocument);
						var bNeedTurnOn = false;
						if (this.bWordShape && editor && editor.WordControl.m_oLogicDocument) {
							if (!editor.WordControl.m_oLogicDocument.TurnOffRecalc) {
								bNeedTurnOn = true;
								editor.WordControl.m_oLogicDocument.TurnOff_Recalculate();
							}
						}
						oContentToDraw.SetApplyToAll(true);
						oContentToDraw.SetParagraphSpacing({Before: 0, After: 0});
						oContentToDraw.SetApplyToAll(false);
						if (bNeedTurnOn) {
							editor.WordControl.m_oLogicDocument.TurnOn_Recalculate(false);
						}
						dMinPolygonLength = warpGeometry.getMinPathPolygonLength();
						dOneLineWidth = AscFormat.GetRectContentWidth(oContentToDraw);
						if (dOneLineWidth > dMinPolygonLength) {
							dKoeff = dMinPolygonLength / dOneLineWidth;
							oContentToDraw.Reset(0, 0, dOneLineWidth, 20000);
						} else {
							oContentToDraw.Reset(0, 0, dMinPolygonLength, 20000);
						}
						oContentToDraw.Recalculate_Page(0, true);
					} else if (bTransform && bColumns) {
						oContentToDraw = oContent.Copy(oContent.Parent, oContent.DrawingDocument);
						oContentToDraw.Reset(0, 0, oContent.XLimit, 20000);
						oContentToDraw.Recalculate_Page(0, true);
					}
					var dContentHeight = oContentToDraw.GetSummaryHeight();
					var OldShowParaMarks, width_ = dWidth * dKoeff, height_ = dHeight * dKoeff;
					if (isRealObject(editor)) {
						OldShowParaMarks = editor.ShowParaMarks;
						editor.ShowParaMarks = true;
					}
					if (bNeedWarp) {
						oContentToDraw.Draw(oContentToDraw.StartPage, oTextDrawer);
						oRet.oTxWarpStructParamarks = oTextDrawer.m_oDocContentStructure;
						oRet.oTxWarpStructParamarks.Recalculate(oTheme, oColorMap, width_, height_, this);
						if (bTransform) {
							oRet.oTxWarpStructParamarks.checkByWarpStruct(warpGeometry, dWidth, dHeight, oTheme, oColorMap, this, dOneLineWidth, oContentToDraw.XLimit, dContentHeight, dKoeff);
							if (bNeedNoTransform && bCheckWordArtContent) {
								if (oRet.oTxWarpStructParamarks.m_aComments.length > 0) {
									oContent.Recalculate_Page(0, true);
									bContentRecalculated = true;
								}
								oContent.Draw(oContent.StartPage, oTextDrawer);
								oRet.oTxWarpStructParamarksNoTransform = oTextDrawer.m_oDocContentStructure;
								oRet.oTxWarpStructParamarksNoTransform.Recalculate(oTheme, oColorMap, dWidth, dHeight, this);
								oRet.oTxWarpStructParamarksNoTransform.checkUnionPaths();
							}
						} else {
							oRet.oTxWarpStructParamarks.checkUnionPaths();
							if (bNeedNoTransform && bCheckWordArtContent) {
								oRet.oTxWarpStructParamarksNoTransform = oRet.oTxWarpStructParamarks;
							}
						}
					} else {
						if (bNeedNoTransform && bCheckWordArtContent) {
							oContent.Draw(oContent.StartPage, oTextDrawer);
							oRet.oTxWarpStructParamarksNoTransform = oTextDrawer.m_oDocContentStructure;
							oRet.oTxWarpStructParamarksNoTransform.Recalculate(oTheme, oColorMap, dWidth, dHeight, this);
							oRet.oTxWarpStructParamarksNoTransform.checkUnionPaths();
						}
					}

					if (isRealObject(editor)) {
						editor.ShowParaMarks = false;
					}
					if (bNeedWarp) {
						oContentToDraw.Draw(oContentToDraw.StartPage, oTextDrawer);
						oRet.oTxWarpStruct = oTextDrawer.m_oDocContentStructure;
						oRet.oTxWarpStruct.Recalculate(oTheme, oColorMap, width_, height_, this);
						if (bTransform) {
							oRet.oTxWarpStruct.checkByWarpStruct(warpGeometry, dWidth, dHeight, oTheme, oColorMap, this, dOneLineWidth, oContentToDraw.XLimit, dContentHeight, dKoeff);
							if (bNeedNoTransform && bCheckWordArtContent) {
								if (oRet.oTxWarpStruct.m_aComments.length > 0 && !bContentRecalculated) {
									oContent.Recalculate_Page(0, true);
								}
								oContent.Draw(oContent.StartPage, oTextDrawer);
								oRet.oTxWarpStructNoTransform = oTextDrawer.m_oDocContentStructure;
								oRet.oTxWarpStructNoTransform.Recalculate(oTheme, oColorMap, dWidth, dHeight, this);
								oRet.oTxWarpStructNoTransform.checkUnionPaths();
							}
						} else {
							oRet.oTxWarpStruct.checkUnionPaths();
							if (bNeedNoTransform && bCheckWordArtContent) {
								oRet.oTxWarpStructNoTransform = oRet.oTxWarpStruct;
							}
						}
					} else {
						if (bNeedNoTransform && bCheckWordArtContent) {
							oContent.Draw(oContent.StartPage, oTextDrawer);
							oRet.oTxWarpStructNoTransform = oTextDrawer.m_oDocContentStructure;
							oRet.oTxWarpStructNoTransform.Recalculate(oTheme, oColorMap, dWidth, dHeight, this);
							oRet.oTxWarpStructNoTransform.checkUnionPaths();
						}
					}

					if (isRealObject(editor)) {
						editor.ShowParaMarks = OldShowParaMarks;
					}
				}
				return oRet;
			}, this, []);
		};

		CShape.prototype.checkTypeCorrect = function () {
			if (!this.spPr) {
				return false;
			}
			return true;
		};


		CShape.prototype.getColumnNumber = function () {
			if (this.bWordShape) {
				return 1;
			}
			var oBodyPr = this.getBodyPr();
			if (AscFormat.isRealNumber(oBodyPr.numCol)) {
				return oBodyPr.numCol;
			}
			return 1;
		};

		CShape.prototype.getColumnSpace = function () {
			if (this.bWordShape) {
				return 0;
			}
			var oBodyPr = this.getBodyPr();
			if (AscFormat.isRealNumber(oBodyPr.spcCol)) {
				return oBodyPr.spcCol;
			}
			return 0;
		};


		CShape.prototype.getTextFitType = function () {
			var oBodyPr = this.getBodyPr();
			if (AscCommon.isRealObject(oBodyPr.textFit) && AscFormat.isRealNumber(oBodyPr.textFit.type)) {
				return oBodyPr.textFit.type;
			}
			return AscFormat.text_fit_No;
		};

		CShape.prototype.getVertOverflowType = function () {
			var oBodyPr = this.getBodyPr();
			if (AscFormat.isRealNumber(oBodyPr.vertOverflow)) {
				return oBodyPr.vertOverflow;
			}
			return AscFormat.nVOTOverflow;
		};


		CShape.prototype.checkWrap = function () {
			if (!this.txBody) {
				return;
			}
			var new_body_pr = this.getBodyPr();
			if (new_body_pr) {
				if (new_body_pr.numCol > 1) {
					if (new_body_pr.wrap === AscFormat.nTWTNone) {
						new_body_pr = new_body_pr.createDuplicate();
						new_body_pr.wrap = AscFormat.nTWTSquare;
						this.txBody.setBodyPr(new_body_pr);
					}
				}
			}
		};


		CShape.prototype.setColumnNumber = function (num) {
			if (!this.bWordShape && !CheckObjectLine(this)) {
				var new_body_pr = this.getBodyPr();
				if (new_body_pr) {
					new_body_pr = new_body_pr.createDuplicate();
					new_body_pr.numCol = (num >> 0);

					if (!this.txBody) {
						this.createTextBody();
					}
					if (this.txBody) {
						this.txBody.setBodyPr(new_body_pr);
					}
					this.checkWrap();
				}
			}

		};

		CShape.prototype.setColumnSpace = function (spcCol) {
			if (!this.bWordShape && !CheckObjectLine(this)) {
				var new_body_pr = this.getBodyPr();
				if (new_body_pr) {
					new_body_pr = new_body_pr.createDuplicate();
					new_body_pr.spcCol = spcCol;
					if (!this.txBody) {
						this.createTextBody();
					}
					if (this.txBody) {
						this.txBody.setBodyPr(new_body_pr);
					}
					this.checkWrap();
				}
			}
		};
		CShape.prototype.GetAllContentControls = function (arrContentControls) {
			var oContent = this.getDocContent();
			if (oContent) {
				oContent.GetAllContentControls(arrContentControls);
			}
		};

		CShape.prototype.getCopyWithSourceFormatting = function () {
			var oCopy = this.copy(undefined);
			if (this.pen || this.brush) {
				if (!oCopy.spPr) {
					oCopy.setSpPr(AscFormat.CSpPr());
					oCopy.spPr.setParent(oCopy);
				}
				if (this.brush) {
					oCopy.spPr.setFill(this.brush.saveSourceFormatting());
				}
				if (this.pen) {
					oCopy.spPr.setLn(this.pen.createDuplicate(true));
				}
			}
			if (oCopy.txBody && oCopy.txBody.content) {
				var oTheme = this.Get_Theme();
				var oColorMap = this.Get_ColorMap();
				if (this.txBody && this.txBody.content) {
					SaveContentSourceFormatting(this.txBody.content.Content, oCopy.txBody.content.Content, oTheme, oColorMap)
				}
			}
			if (oCopy.isPlaceholder() && !this.recalcInfo.recalculateTransform) {
				var oXfrm = oCopy.spPr.xfrm;
				if (!oXfrm || !oXfrm.isNotNull()) {
					oCopy.x = this.x;
					oCopy.y = this.y;
					oCopy.extX = this.extX;
					oCopy.extY = this.extY;
					AscFormat.CheckSpPrXfrm(oCopy, true);
				}
			}
			if (this.txXfrm) {
				oCopy.setTxXfrm(this.txXfrm.createDuplicate());
				oCopy.convertFromSmartArt();
			}
			return oCopy;
		};
		CShape.prototype.getSignatureLineGuid = function () {
			if (this.signatureLine) {
				return this.signatureLine.id;
			}
			return null;
		};

		CShape.prototype.GetAllFields = function (isUseSelection, arrFields) {
			var oContent = this.getDocContent();
			if (oContent) {
				return oContent.GetAllFields(isUseSelection, arrFields)
			}
			return arrFields ? arrFields : [];
		};
		CShape.prototype.GetAllSeqFieldsByType = function (sType, aFields) {
			var oContent = this.getDocContent();
			if (oContent) {
				return oContent.GetAllSeqFieldsByType(sType, aFields)
			}
		};

		CShape.prototype.Get_TextBackGroundColor = function () {
			if (!this.brush) {
				return undefined;
			}
			var oTheme = this.Get_Theme && this.Get_Theme();
			var oColorMap = this.Get_ColorMap && this.Get_ColorMap();
			if (oTheme && oColorMap) {
				this.brush.check(oTheme, oColorMap);
			}
			return this.brush.Get_TextBackGroundColor();
		};

		CShape.prototype.checkResetAutoFit = function (bCheckMinVal) {
			if (this.txBody) {
				var oCompiledBodyPr = this.getBodyPr();
				var oNewBodyPr;
				var oTextFit = oCompiledBodyPr.textFit;
				if (oTextFit) {
					if (oTextFit.type === AscFormat.text_fit_NormAuto) {
						if (AscFormat.isRealNumber(oTextFit.fontScale)) {
							if (oTextFit.fontScale < 100000) {
								var bReset = false;
								if (bCheckMinVal) {
									if (oTextFit.fontScale <= 25000) {
										bReset = true;
										var oContent = this.txBody.content;
										if (oContent) {
											oContent.CheckRunContent(function (oRun) {
												var oTextPr = oRun.Pr;
												var oCompiledPr = oRun.CompiledPr;
												if (AscFormat.isRealNumber(oTextPr.FontSize) &&
													AscFormat.isRealNumber(oCompiledPr.FontSize) &&
													oTextPr.FontSize !== oCompiledPr.FontSize) {
													oRun.SetFontSize(oCompiledPr.FontSize);
													if (oRun.IsParaEndRun()) {
														var oParagraph = oRun.Paragraph;
														if (oParagraph) {
															oParagraph.TextPr.Apply_TextPr(oTextPr);
														}
													}
												}
											});
										}
									}
								} else {
									bReset = true;
								}
								if (bReset) {
									oNewBodyPr = oCompiledBodyPr.createDuplicate();
									oNewBodyPr.textFit = new AscFormat.CTextFit();
									oNewBodyPr.textFit.type = AscFormat.text_fit_No;
									this.txBody.setBodyPr(oNewBodyPr);
								}
							}
						}
					}
				}
			}
		};
		CShape.prototype.getInnerForm = function () {
			return this.textBoxContent ? this.textBoxContent.GetInnerForm() : null;
		};
		CShape.prototype.getFormHorPadding = function () {
			let oInnerForm;
			if (this.isForm
				&& this.isForm()
				&& (oInnerForm = this.getInnerForm())
				&& !oInnerForm.IsPictureForm()
				&& !oInnerForm.IsCheckBox()
				&& (!oInnerForm.IsTextForm() || !oInnerForm.GetTextFormPr().IsComb()))
				return 2 * 25.4 / 72; // 2pt

			return 0;
		};

		//for bug 52775. remove in the next version
		CShape.prototype.applySmartArtTextStyle = function () {
			if (this.textBoxContent) {
				if (this.style && this.style.fontRef) {
					if (this.style.fontRef.Color) {
						var oUnifill = AscFormat.CreateUniFillByUniColorCopy(this.style.fontRef.Color);
						this.textBoxContent.CheckRunContent(function (oRun) {
							if (oRun instanceof AscCommonWord.ParaRun) {
								if (!oRun.Pr.Unifill && !oRun.Pr.TextFill) {
									oRun.Set_Unifill(oUnifill);
								}
							}
							return false;
						});
					}
				}
			}
		};

		CShape.prototype.getTypeName = function () {
			if (this.isPlaceholder()) {
				return this.getPlaceholderName();
			}
			var sPreset = this.getPresetGeom();
			if (typeof sPreset === "string" && sPreset.length > 0) {
				var oApi = Asc.editor || editor;
				return oApi.getShapeName(sPreset);
			}
			return AscCommon.translateManager.getValue("Shape");
		};

		CShape.prototype.applyImagePlaceholderCallback = function (aImages, oPlaceholder) {
			var _image = aImages[0];
			if (this.isObjectInSmartArt()) {
				if (this.spPr) {
					const imageWidth = _image.Image.width;
					const imageHeight = _image.Image.height;
					const shapeWidth = this.extX;
					const shapeHeight = this.extY;
					const srcRect = new AscFormat.CSrcRect();
					srcRect.setValueForFitBlipFill(shapeWidth, shapeHeight, imageWidth, imageHeight);
					const oBlipFillUniFill = AscFormat.CreateBlipFillUniFillFromUrl(_image.src);
					oBlipFillUniFill.fill.setSrcRect(srcRect);
					this.changeFill(oBlipFillUniFill);
				}
			}
			return true;
		};

		CShape.prototype.pasteFormatting = function (oFormatData) {
			if(!oFormatData)
				return;
			let oDrawing = oFormatData.Drawing;
			if(oDrawing) {
				this.pasteDrawingFormatting(oFormatData.Drawing);
			}
			if(oFormatData.ParaPr || oFormatData.TextPr) {
				let oContent = this.getDocContent();
				if(oContent) {
					let bApplyToAll = true;
					let oController = this.getDrawingObjectsController && this.getDrawingObjectsController();
					if(oController) {
						if(AscFormat.getTargetTextObject(oController) === this) {
							bApplyToAll = false;
						}
					}
					if(bApplyToAll) {
						oContent.SetApplyToAll(true);
					}
					let fDocContentMethod = AscCommonWord.CDocumentContent.prototype.PasteFormatting;
					let fTableMethod = AscCommonWord.CTable.prototype.PasteFormatting;
					this.applyTextFunction(fDocContentMethod, fTableMethod, [oFormatData]);
					if(bApplyToAll) {
						oContent.SetApplyToAll(false);
					}
				}
			}
		};


		function CreateBinaryReader(szSrc, offset, srcLen) {
			var memoryData = AscCommon.Base64.decode(szSrc, true, srcLen, offset);
			return new AscCommon.FT_Stream2(memoryData, memoryData.length);
		}

		function getParaDrawing(oDrawing) {
			var oCurDrawing = oDrawing;
			while (oCurDrawing.group) {
				oCurDrawing = oCurDrawing.group;
			}
			if (oCurDrawing.parent instanceof AscCommonWord.ParaDrawing) {
				return oCurDrawing.parent;
			}
			return null;
		}

		function checkDrawingsTransformBeforePaste(oEndContent, oSourceContent, oTempParent) {
			var i, j;
			for (i = 0; i < oEndContent.Drawings.length; ++i) {
				var shape = oEndContent.Drawings[i].Drawing;
				if (shape.isPlaceholder && shape.isPlaceholder() && (!shape.spPr || !shape.spPr.xfrm || !shape.spPr.xfrm.isNotNull())) {
					var oOldParent = shape.parent;
					shape.parent = oTempParent;
					var hierarchy = shape.getHierarchy();
					for (j = 0; j < hierarchy.length; ++j) {
						if (hierarchy[j] && hierarchy[j].spPr && hierarchy[j].spPr.xfrm && hierarchy[j].spPr.xfrm.isNotNull()) {
							break;
						}
					}
					if (j === hierarchy.length) {
						if (oSourceContent.Drawings[i] && oSourceContent.Drawings[i].Drawing) {
							var oSourceShape = oSourceContent.Drawings[i].Drawing;
							if (oSourceShape && oSourceShape.spPr && oSourceShape.spPr.xfrm && oSourceShape.spPr.xfrm.isNotNull()) {
								shape.x = oSourceShape.spPr.xfrm.offX;
								shape.y = oSourceShape.spPr.xfrm.offY;
								shape.extX = oSourceShape.spPr.xfrm.extX;
								shape.extY = oSourceShape.spPr.xfrm.extY;
								AscFormat.CheckSpPrXfrm(shape);
							}
						}
					}
					shape.parent = oOldParent;
				}
			}
		}


		function SaveSourceFormattingTextPr(oTextPr, oTheme, oColorMap) {
			oTextPr.ReplaceThemeFonts(oTheme.themeElements.fontScheme);
			if (oTextPr.Unifill) {
				oTextPr.Unifill.check(oTheme, oColorMap);
				oTextPr.Unifill = oTextPr.Unifill.saveSourceFormatting();
			}
			if (oTextPr.TextOutline && oTextPr.TextOutline.Fill) {
				oTextPr.TextOutline.Fill.check(oTheme, oColorMap);
				oTextPr.TextOutline.Fill = oTextPr.TextOutline.Fill.saveSourceFormatting();
			}
			return oTextPr;
		}

		function SaveContentSourceFormatting(aSourceContent, aCopyContent, oTheme, oColorMap) {
			if (aCopyContent.length === aSourceContent.length) {
				var bMergeRunPr = (aCopyContent === aSourceContent);
				var oElem;
				for (var i = 0; i < aSourceContent.length; ++i) {
					oElem = aSourceContent[i];
					if (oElem.CompiledPr.Pr) {
						var oPr = oElem.CompiledPr.Pr.ParaPr.Copy();
						oPr.DefaultRunPr = SaveSourceFormattingTextPr(oElem.CompiledPr.Pr.TextPr.Copy(), oTheme, oColorMap);
						aCopyContent[i].Set_Pr(oPr);
						SaveRunsFormatting(oElem.Content, aCopyContent[i].Content, oTheme, oColorMap, oPr);
					} else {
						if (aCopyContent[i].Pr && aCopyContent[i].Pr.DefaultRunPr && aCopyContent[i].Pr.DefaultRunPr.Unifill) {
							var oPr = aCopyContent[i].Pr.Copy();
							oPr.DefaultRunPr.Unifill.check(oTheme, oColorMap);
							oPr.DefaultRunPr.Unifill = oPr.DefaultRunPr.Unifill.saveSourceFormatting();
							oPr.DefaultRunPr = SaveSourceFormattingTextPr(oPr.DefaultRunPr.Copy(), oTheme, oColorMap);
							aCopyContent[i].Set_Pr(oPr);
						}
					}
				}
			}
		}

		function SaveRunsFormatting(aSourceContent, aCopyContent, oTheme, oColorMap, oPr) {
			var bMergeRunPr = (aCopyContent === aSourceContent);
			if(aCopyContent.length !== aSourceContent.length) {
				return;
			}
			for (var i = 0; i < aCopyContent.length; ++i) {
				if (aCopyContent[i] instanceof ParaRun && aCopyContent[i].Pr) {
					if (bMergeRunPr) {
						var oCoprPr = oPr.DefaultRunPr.Copy();
						oCoprPr.Merge(SaveSourceFormattingTextPr(aCopyContent[i].Pr.Copy(), oTheme, oColorMap));
						aCopyContent[i].Set_Pr(oCoprPr)
					} else {
						aCopyContent[i].Set_Pr(SaveSourceFormattingTextPr(aCopyContent[i].Pr.Copy(), oTheme, oColorMap));
					}

				} else if (aSourceContent[i].Content) {
					var oElem = aSourceContent[i];
					SaveRunsFormatting(oElem.Content, aCopyContent[i].Content, oTheme, oColorMap, oPr);
					if (oElem.Get_CompiledCtrPrp && aCopyContent[i].setCtrPrp) {
						var oCtrPr = oElem.Get_CompiledCtrPrp();
						aCopyContent[i].setCtrPrp(oCtrPr);
					}
				} else if (aSourceContent[i] instanceof AscCommonWord.ParaMath && aSourceContent[i].Root && aSourceContent[i].Root.Content) {
					SaveRunsFormatting(aSourceContent[i].Root.Content, aCopyContent[i].Root.Content, oTheme, oColorMap, oPr);

				}
			}
		}


		AscFormat.checkPlaceholdersText = function () {
			if (AscFonts.IsCheckSymbols) {
				for (var i = AscFormat.pHText.length - 1; i >= 0; i--)
					AscFonts.FontPickerByCharacter.getFontsByString(AscCommon.translateManager.getValue(AscFormat.pHText[i]));
			}
		};
		//--------------------------------------------------------export----------------------------------------------------
		window['AscFormat'] = window['AscFormat'] || {};
		window['AscFormat'].CheckObjectLine = CheckObjectLine;
		window['AscFormat'].CreateUniFillByUniColorCopy = CreateUniFillByUniColorCopy;
		window['AscFormat'].CreateUniFillByUniColor = CreateUniFillByUniColor;
		window['AscFormat'].ConvertParagraphToPPTX = ConvertParagraphToPPTX;
		window['AscFormat'].ConvertElementsToPPTX = ConvertElementsToPPTX;
		window['AscFormat'].ConvertParagraphToWord = ConvertParagraphToWord;
		window['AscFormat'].SetXfrmFromMetrics = SetXfrmFromMetrics;
		window['AscFormat'].CShape = CShape;
		window['AscFormat'].CreateBinaryReader = CreateBinaryReader;
		window['AscFormat'].getParaDrawing = getParaDrawing;
		window['AscFormat'].ConvertGraphicFrameToWordTable = ConvertGraphicFrameToWordTable;
		window['AscFormat'].ConvertTableToGraphicFrame = ConvertTableToGraphicFrame;
		window['AscFormat'].CSignatureLine = CSignatureLine;
		window['AscFormat'].checkDrawingsTransformBeforePaste = checkDrawingsTransformBeforePaste;
		window['AscFormat'].SaveContentSourceFormatting = SaveContentSourceFormatting;
		window['AscFormat'].hitToHandles = hitToHandles;
		window['AscFormat'].pHText = pHText;
	})(window);
