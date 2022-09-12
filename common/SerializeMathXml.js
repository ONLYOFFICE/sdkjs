/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(window, undefined) {
	function CT_StringM() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_StringM, AscCommon.CT_String, 0);
	CT_StringM.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeStringEncode("m:val", val);
	};
	function CT_BoolM() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_BoolM, AscCommon.CT_Bool, 0);
	CT_BoolM.prototype.writeAttrVal = function(writer, val) {
		if (!val) {
			writer.WriteXmlNullableAttributeBool("m:val", val);
		}
	};

	function CT_IntM() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_IntM, AscCommon.CT_Int, 0);
	CT_IntM.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeInt("m:val", this.val);
	};

	function CT_UIntM() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_UIntM, AscCommon.CT_UInt, 0);
	CT_UIntM.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeUInt("m:val", this.val);
	};

	function CT_DoubleM() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_DoubleM, AscCommon.CT_Double, 0);
	CT_DoubleM.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeDouble("m:val", this.val);
	};


	function CT_OMathParaPr() {
		this.jc = null;
		return this;
	}
	CT_OMathParaPr.prototype.fromXml = function (reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "jc" : {
					this.jc = fromXml_ST_Jc(CT_StringM.prototype.toVal(reader, this.jc), this.jc);
					break;
				}
			}
		}
	};
	CT_OMathParaPr.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_Jc(this.jc)), "m:jc");
		writer.WriteXmlNodeEnd(name);
	};

	function CT_OMathPara() {
		this.OMathParaPr = null;
		return this;
	}
	CT_OMathPara.prototype.fromXml = function (reader, paragraphContent) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "oMathParaPr" : {
					this.OMathParaPr = new CT_OMathParaPr();
					this.OMathParaPr.fromXml(reader);
					break;
				}
				case "oMath" : {
					if (!paragraphContent) {
						break;
					}
					let paraMath = new ParaMath();
					paragraphContent.AddToContentToEnd(paraMath);
					paraMath.fromXml(reader);
					if (this.OMathParaPr) {
						paraMath.Set_Align(this.OMathParaPr.jc);
					}
					break;
				}
			}
		}
	};
	CT_OMathPara.prototype.toXml = function (writer, name, oMath) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.OMathParaPr, "m:oMathParaPr");
		writer.WriteXmlNullable(oMath, "m:oMath");
		writer.WriteXmlNodeEnd(name);
	};
	CT_OMathPara.prototype.initMathParaPr = function (oMath) {
		if (oMath && undefined !== oMath.Jc) {
			this.OMathParaPr = new CT_OMathParaPr();
			this.OMathParaPr.jc = oMath.Jc;
		}
	};

	ParaMath.prototype.fromXml = function(reader) {
		this.Root.fromXml(reader, this.GetParent());
		this.Root.Correct_Content(true);
	};
	ParaMath.prototype.toXml = function(writer, name) {
		this.Root.toXml(writer, name);
	};
	CMathContent.prototype.fromXml = function(reader, opt_paragraphContent) {
		let elem, depth = reader.GetDepth();
		let oReadResult = reader.context.oReadResult;
		while (reader.ReadNextSiblingNode(depth)) {
			elem = null;
			let name = reader.GetNameNoNS();
			switch (name) {
				case "argPr" : {
					this.ArgSize.fromXml(reader);
					break;
				}
				case "acc" : {
					elem = new CAccent(null);
					elem.fromXml(reader);
					break;
				}
				case "bar" : {
					elem = new CBar(null);
					elem.fromXml(reader);
					break;
				}
				case "box" : {
					elem = new CBox(null);
					elem.fromXml(reader);
					break;
				}
				case "borderBox" : {
					elem = new CBorderBox(null);
					elem.fromXml(reader);
					break;
				}
				case "d" : {
					elem = new CDelimiter();
					elem.fromXml(reader);
					break;
				}
				case "eqArr" : {
					elem = new CEqArray();
					elem.fromXml(reader);
					//todo
					break;
				}
				case "f" : {
					elem = new CFraction();
					elem.fromXml(reader);
					break;
				}
				case "func" : {
					elem = new CMathFunc(null);
					elem.fromXml(reader);
					break;
				}
				case "groupChr" : {
					elem = new CGroupCharacter(null);
					elem.fromXml(reader);
					break;
				}
				case "limLow" : {
					elem = new CLimit(null);
					elem.fromXml(reader);
					break;
				}
				case "limUpp" : {
					elem = new CLimit(null);
					elem.fromXml(reader);
					break;
				}
				case "m" : {
					elem = new CMathMatrix(null);
					elem.fromXml(reader);
					break;
				}
				case "nary" : {
					elem = new CNary(null);
					elem.fromXml(reader);
					break;
				}
				case "phant" : {
					elem = new CPhantom(null);
					elem.fromXml(reader);
					break;
				}
				case "rad" : {
					elem = new CRadical(null);
					elem.fromXml(reader);
					break;
				}
				case "sPre" : {
					elem = new CDegreeSubSup(null);
					elem.Pr.type = DEGREE_PreSubSup;
					elem.fromXml(reader);
					break;
				}
				case "sSub" : {
					elem = new CDegree(null);
					elem.Pr.type = DEGREE_SUBSCRIPT;
					elem.fromXml(reader);
					break;
				}
				case "sSubSup" : {
					elem = new CDegreeSubSup(null);
					elem.Pr.type = DEGREE_SubSup;
					elem.fromXml(reader);
					break;
				}
				case "sSup" : {
					elem = new CDegree(null);
					elem.Pr.type = DEGREE_SUPERSCRIPT;
					elem.fromXml(reader);
					break;
				}
				case "ctrlPr" : {
					this.setCtrPrp(CMathBase.prototype.fromXmlCtrlPr.call(this, reader, null));
					break;
				}
				case "r" : {
					elem = new ParaRun(this.Paragraph, true);
					elem.fromXml(reader, opt_paragraphContent);
					break;
				}
				case "del": {
					if (oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.paragraphContent = this;
						let startPos = this.GetElementsCount();
						trackChange.fromXml(reader);
						let endPos = this.GetElementsCount();
						for (let i = startPos; i < endPos; ++i) {
							oReadResult.setNestedReviewType(this.GetElement(i), reviewtype_Remove, trackChange.ReviewInfo);
						}
					}
					break;
				}
				case "ins": {
					let trackChange = new CT_TrackChange();
					trackChange.paragraphContent = this;
					let startPos = this.GetElementsCount();
					trackChange.fromXml(reader);
					let endPos = this.GetElementsCount();
					if (oReadResult.checkReadRevisions()) {
						for (let i = startPos; i < endPos; ++i) {
							oReadResult.setNestedReviewType(this.GetElement(i), reviewtype_Add, trackChange.ReviewInfo);
						}
					}
					break;
				}
				// case "bookmarkStart" : {
				// 	elem = new CParagraphBookmark(true);
				// 	elem.fromXml(reader);
				// 	oReadResult.addBookmarkStart(this, elem, true);
				// 	break;
				// }
				// case "bookmarkEnd" : {
				// 	elem = new CParagraphBookmark(false);
				// 	elem.fromXml(reader);
				// 	oReadResult.addBookmarkEnd(this, elem, true);
				// 	break;
				// }
				// case "moveFromRangeStart" : {
				// 	oReadResult.readMoveRangeStartXml(oReadResult, reader, this, true);
				// 	break;
				// }
				// case "moveFromRangeEnd" : {
				// 	oReadResult.readMoveRangeEndXml(oReadResult, reader, this, true);
				// 	break;
				// }
				// case "moveToRangeStart" : {
				// 	oReadResult.readMoveRangeStartXml(oReadResult, reader, this, false);
				// 	break;
				// }
				// case "moveToRangeEnd" : {
				// 	oReadResult.readMoveRangeEndXml(oReadResult, reader, this, true);
				// 	break;
				// }
				//todo
				// case "customXml" : {
				// 	elem = new CT_CustomXmlRun();
				// 	elem.fromXml(reader);
				// 	this.customXml.push(elem);
				// 	break;
				// }
				// case "fldSimple" : {
				// 	elem = new CT_SimpleField();
				// 	elem.fromXml(reader);
				// 	this.fldSimple.push(elem);
				// 	break;
				// }
				// case "hyperlink" : {
				// 	elem = new CT_Hyperlink();
				// 	elem.fromXml(reader);
				// 	this.hyperlink.push(elem);
				// 	break;
				// }
				// case "smartTag" : {
				// 	elem = new CT_SmartTagRun();
				// 	elem.fromXml(reader);
				// 	this.smartTag.push(elem);
				// 	break;
				// }
				// case "sdt" : {
				// 	elem = new CT_SdtRun();
				// 	elem.fromXml(reader);
				// 	this.sdt.push(elem);
				// 	break;
				// }
				// case "proofErr" : {
				// 	elem = new CT_ProofErr();
				// 	elem.fromXml(reader);
				// 	this.proofErr.push(elem);
				// 	break;
				// }
				// case "permStart" : {
				// 	elem = new CT_PermStart();
				// 	elem.fromXml(reader);
				// 	this.permStart.push(elem);
				// 	break;
				// }
				// case "permEnd" : {
				// 	elem = new CT_Perm();
				// 	elem.fromXml(reader);
				// 	this.permEnd.push(elem);
				// 	break;
				// }
				// case "commentRangeStart" : {
				// 	elem = new CT_MarkupRange();
				// 	elem.fromXml(reader);
				// 	this.commentRangeStart.push(elem);
				// 	break;
				// }
				// case "commentRangeEnd" : {
				// 	elem = new CT_MarkupRange();
				// 	elem.fromXml(reader);
				// 	this.commentRangeEnd.push(elem);
				// 	break;
				// }
				// case "customXmlInsRangeStart" : {
				// 	elem = new CT_TrackChange();
				// 	elem.fromXml(reader);
				// 	this.customXmlInsRangeStart.push(elem);
				// 	break;
				// }
				// case "customXmlInsRangeEnd" : {
				// 	elem = new CT_Markup();
				// 	elem.fromXml(reader);
				// 	this.customXmlInsRangeEnd.push(elem);
				// 	break;
				// }
				// case "customXmlDelRangeStart" : {
				// 	elem = new CT_TrackChange();
				// 	elem.fromXml(reader);
				// 	this.customXmlDelRangeStart.push(elem);
				// 	break;
				// }
				// case "customXmlDelRangeEnd" : {
				// 	elem = new CT_Markup();
				// 	elem.fromXml(reader);
				// 	this.customXmlDelRangeEnd.push(elem);
				// 	break;
				// }
				// case "customXmlMoveFromRangeStart" : {
				// 	elem = new CT_TrackChange();
				// 	elem.fromXml(reader);
				// 	this.customXmlMoveFromRangeStart.push(elem);
				// 	break;
				// }
				// case "customXmlMoveFromRangeEnd" : {
				// 	elem = new CT_Markup();
				// 	elem.fromXml(reader);
				// 	this.customXmlMoveFromRangeEnd.push(elem);
				// 	break;
				// }
				// case "customXmlMoveToRangeStart" : {
				// 	elem = new CT_TrackChange();
				// 	elem.fromXml(reader);
				// 	this.customXmlMoveToRangeStart.push(elem);
				// 	break;
				// }
				// case "customXmlMoveToRangeEnd" : {
				// 	elem = new CT_Markup();
				// 	elem.fromXml(reader);
				// 	this.customXmlMoveToRangeEnd.push(elem);
				// 	break;
				// }
				// case "moveFrom" : {
				// 	elem = new CT_RunTrackChange();
				// 	elem.fromXml(reader);
				// 	this.moveFrom.push(elem);
				// 	break;
				// }
				// case "moveTo" : {
				// 	elem = new CT_RunTrackChange();
				// 	elem.fromXml(reader);
				// 	this.moveTo.push(elem);
				// 	break;
				// }
				// case "oMathPara" : {
				// 	elem = new CT_OMathPara();
				// 	elem.fromXml(reader);
				// 	this.oMathPara.push(elem);
				// 	break;
				// }
				// case "oMath" : {
				// 	elem = new CT_OMath();
				// 	elem.fromXml(reader);
				// 	this.oMath.push(elem);
				// 	break;
				// }
			}
			if (elem) {
				this.addElementToContent(elem);
				elem.Paragraph = this.Paragraph;
			}
		}
	};
	CMathContent.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		if (this.ArgSize.Is_Init())
			this.ArgSize.toXml(writer, "m:argPr");
		var isSingle = 1 === this.Content.length;
		for(let i = 0; i <= this.Content.length - 1; i++)
		{
			let item = this.Content[i];
			switch (item.Type) {
				case para_Math_Composition: {
					switch (item.kind) {
						case MATH_ACCENT            :
							item.toXml(writer, "m:acc");
							break;
						case MATH_BAR                :
							item.toXml(writer, "m:bar");
							break;
						case MATH_BOX                :
							item.toXml(writer, "m:box");
							break;
						case MATH_BORDER_BOX        :
							item.toXml(writer, "m:borderBox");
							break;
						case MATH_DELIMITER            :
							item.toXml(writer, "m:d");
							break;
						case MATH_EQ_ARRAY            :
							item.toXml(writer, "m:eqArr");
							break;
						case MATH_FRACTION            :
							item.toXml(writer, "m:f");
							break;
						case MATH_FUNCTION            :
							item.toXml(writer, "m:func");
							break;
						case MATH_GROUP_CHARACTER    :
							item.toXml(writer, "m:groupChr");
							break;
						case MATH_LIMIT                :
							if (LIMIT_LOW === item.Pr.type)
								item.toXml(writer, "m:limLow");
							else if (LIMIT_UP === item.Pr.type)
								item.toXml(writer, "m:limUpp");
							break;
						case MATH_MATRIX            :
							item.toXml(writer, "m:m");
							break;
						case MATH_NARY            :
							item.toXml(writer, "m:nary");
							break;
						case MATH_PHANTOM        :
							item.toXml(writer, "m:phant");
							break;
						case MATH_RADICAL        :
							item.toXml(writer, "m:rad");
							break;
						case MATH_DEGREESubSup    :
							if (DEGREE_PreSubSup === item.Pr.type)
								item.toXml(writer, "m:sPre");
							else if (DEGREE_SubSup === item.Pr.type)
								item.toXml(writer, "m:sSubSup");
							break;
						case MATH_DEGREE        :
							if (DEGREE_SUBSCRIPT === item.Pr.type)
								item.toXml(writer, "m:sSub");
							else if (DEGREE_SUPERSCRIPT === item.Pr.type)
								item.toXml(writer, "m:sSup");
							break;
						case MATH_RUN            :
							if (!item.Is_Empty() || isSingle) {
								item.toXml(writer, "m:r");
							}
							break;
					}
					break;
				}
				case para_Math_Run:
					if (!item.Is_Empty() || isSingle) {
						item.toXml(writer, "m:r");
					}
					break;
			}
		}
		writer.WriteXmlNodeEnd(name);
	};
	CMathBase.prototype.fromXmlCtrlPr = function(reader, mathElem) {
		let elem, rPr, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			//todo if another ns
			switch (reader.GetName()) {
				case "a:rPr" : {
					rPr = new CTextPr();
					rPr.fromDrawingML(reader);
					break;
				}
				case "w:rPr" : {
					rPr = new CTextPr();
					rPr.fromXml(reader);
					break;
				}
				case "w:ins" : {
					if(mathElem) {
						elem = new CT_TrackChange();
						elem.fromXml(reader);
						mathElem.SetReviewTypeWithInfo(reviewtype_Add, elem.ReviewInfo, false);
					}
					break;
				}
				case "w:del" : {
					if(mathElem) {
						elem = new CT_TrackChange();
						elem.fromXml(reader);
						if (reader.context.oReadResult.checkReadRevisions()) {
							mathElem.SetReviewTypeWithInfo(reviewtype_Remove, elem.ReviewInfo, false);
						}
					}
					break;
				}
			}
		}
		return rPr;
	};
	CMathBase.prototype.toXmlCtrlPr = function (writer, name, mathElem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		if (mathElem.Is_FromDocument()) {
			writer.WriteXmlNullable(mathElem.CtrPrp, "w:rPr");
		} else {
			//todo
			// writer.WriteXmlNullable(mathElem.CtrPrp, "a:rPr");
		}
		if (mathElem.ReviewInfo && reviewtype_Common !== mathElem.GetReviewType()) {
			let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, mathElem.ReviewInfo);
			let ReviewType = mathElem.GetReviewType();
			trackChange.toXml(writer, reviewtype_Remove === mathElem.GetReviewType() ? "w:del" : "w:ins")
		}
		writer.WriteXmlNodeEnd(name);
	};


	CMPrp.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "lit" : {
					this.lit = CT_BoolM.prototype.toVal(reader, this.lit);
					break;
				}
				case "nor" : {
					this.nor = CT_BoolM.prototype.toVal(reader, this.nor);
					break;
				}
				case "scr" : {
					this.scr = fromXml_ST_Script(CT_StringM.prototype.toVal(reader, this.scr), this.scr);
					break;
				}
				case "sty" : {
					this.sty = fromXml_ST_Style(CT_StringM.prototype.toVal(reader, this.sty), this.sty);
					break;
				}
				case "brk" : {
					this.brk = new CMathBreak();
					this.brk.fromXml(reader);
					break;
				}
				case "aln" : {
					this.aln = CT_BoolM.prototype.toVal(reader, this.aln);
					break;
				}
			}
		}
	};
	CMPrp.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.lit), "m:lit");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.nor), "m:nor");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_Script(this.scr)), "m:scr");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_Style(this.sty)), "m:sty");
		writer.WriteXmlNullable(this.brk, "m:brk");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.aln), "m:aln");
		writer.WriteXmlNodeEnd(name);
	};
	CMathAccentPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "chr" : {
					this.chr = CMathBase.prototype.ConvertStrToOperator(CT_StringM.prototype.toVal(reader, ""));
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathAccentPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(CMathBase.prototype.ConvertOperatorToStr(this.chr)), "m:chr");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CAccent.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathAccentPr();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "accPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content = [elem];
					break;
				}
			}
		}
		this.init(props);
	};
	CAccent.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:accPr", this);
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathBarPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "pos" : {
					this.pos = fromXml_ST_TopBot(CT_StringM.prototype.toVal(reader, this.pos), this.pos);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathBarPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_TopBot(this.pos)), "m:pos");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CBar.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathBarPr();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "barPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content = [elem];
					break;
				}
			}
		}
		this.init(props);
	};
	CBar.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:barPr", this);
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathBreak.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "alnAt": {
					this.alnAt = Math.min(255, Math.max(1, reader.GetValueUInt(1)));
					break;
				}
			}
		}
	};
	CMathBreak.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CMathBreak.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUInt("alnAt", this.alnAt);
		writer.WriteXmlAttributesEnd(true);
	};
	CMathBoxPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "opEmu" : {
					this.opEmu = CT_BoolM.prototype.toVal(reader, this.OpEmu);
					break;
				}
				case "noBreak" : {
					this.noBreak = CT_BoolM.prototype.toVal(reader, this.noBreak);
					break;
				}
				case "diff" : {
					this.diff = CT_BoolM.prototype.toVal(reader, this.diff);
					break;
				}
				case "brk" : {
					this.brk = new CMathBreak();
					this.brk.fromXml(reader);
					break;
				}
				case "aln" : {
					this.aln = CT_BoolM.prototype.toVal(reader, this.aln);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathBoxPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.opEmu), "m:opEmu");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.noBreak), "m:noBreak");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.diff), "m:diff");
		writer.WriteXmlNullable(this.brk, "m:brk");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.aln), "m:aln");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CBox.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathBoxPr();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "boxPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content = [elem];
					break;
				}
			}
		}
		this.init(props);
	};
	CBox.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:boxPr", this);
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathBorderBoxPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "hideTop" : {
					this.hideTop = CT_BoolM.prototype.toVal(reader, this.hideTop);
					break;
				}
				case "hideBot" : {
					this.hideBot = CT_BoolM.prototype.toVal(reader, this.hideBot);
					break;
				}
				case "hideLeft" : {
					this.hideLeft = CT_BoolM.prototype.toVal(reader, this.hideLeft);
					break;
				}
				case "hideRight" : {
					this.hideRight = CT_BoolM.prototype.toVal(reader, this.hideRight);
					break;
				}
				case "strikeH" : {
					this.strikeH = CT_BoolM.prototype.toVal(reader, this.strikeH);
					break;
				}
				case "strikeV" : {
					this.strikeV = CT_BoolM.prototype.toVal(reader, this.strikeV);
					break;
				}
				case "strikeBLTR" : {
					this.strikeBLTR = CT_BoolM.prototype.toVal(reader, this.strikeBLTR);
					break;
				}
				case "strikeTLBR" : {
					this.strikeTLBR = CT_BoolM.prototype.toVal(reader, this.strikeTLBR);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathBorderBoxPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.hideTop), "m:hideTop");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.hideBot), "m:hideBot");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.hideLeft), "m:hideLeft");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.hideRight), "m:hideRight");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.strikeH), "m:strikeH");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.strikeV), "m:strikeV");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.strikeBLTR), "m:strikeBLTR");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.strikeTLBR), "m:strikeTLBR");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CBorderBox.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathBorderBoxPr();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "borderBoxPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content = [elem];
					break;
				}
			}
		}
		this.init(props);
	};
	CBorderBox.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:borderBoxPr", this);
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathDelimiterPr.prototype.fromXml = function (reader, mathElem) {
		let  depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "begChr" : {
					this.begChr = CMathBase.prototype.ConvertStrToOperator(CT_StringM.prototype.toVal(reader, ""));
					break;
				}
				case "sepChr" : {
					this.sepChr = CMathBase.prototype.ConvertStrToOperator(CT_StringM.prototype.toVal(reader, ""));
					break;
				}
				case "endChr" : {
					this.endChr = CMathBase.prototype.ConvertStrToOperator(CT_StringM.prototype.toVal(reader, ""));
					break;
				}
				case "grow" : {
					this.grow = CT_BoolM.prototype.toVal(reader, this.grow);
					break;
				}
				case "shp" : {
					this.shp = fromXml_ST_Shp(CT_StringM.prototype.toVal(reader, this.shp), this.shp);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathDelimiterPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(CMathBase.prototype.ConvertOperatorToStr(this.begChr)), "m:begChr");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(CMathBase.prototype.ConvertOperatorToStr(this.sepChr)), "m:sepChr");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(CMathBase.prototype.ConvertOperatorToStr(this.endChr)), "m:endChr");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.grow), "m:grow");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_Shp(this.shp)), "m:shp");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CDelimiter.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathDelimiterPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "dPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content.push(elem);
					break;
				}
			}
		}
		this.init(props);
	};
	CDelimiter.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:dPr", this);
		for (let i = 0; i < this.nCol; i++) {
			writer.WriteXmlNullable(this.getBase(i), "m:e");
		}
		writer.WriteXmlNodeEnd(name);
	};
	CMathEqArrPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "baseJc" : {
					this.baseJc = fromXml_ST_YAlignMath(CT_StringM.prototype.toVal(reader, this.baseJc), this.baseJc);
					break;
				}
				case "maxDist" : {
					this.maxDist = CT_BoolM.prototype.toVal(reader, this.maxDist);
					break;
				}
				case "objDist" : {
					this.objDist = CT_BoolM.prototype.toVal(reader, this.objDist);
					break;
				}
				case "rSpRule" : {
					this.rSpRule = Math.min(4, CT_UIntM.prototype.toVal(reader, this.rSpRule));
					break;
				}
				case "rSp" : {
					this.rSp = CT_UIntM.prototype.toVal(reader, this.rSp);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathEqArrPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_YAlignMath(this.baseJc)), "m:baseJc");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.maxDist), "m:maxDist");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.objDist), "m:objDist");
		writer.WriteXmlNullable(CT_UIntM.prototype.fromVal(this.rSpRule), "m:rSpRule");
		writer.WriteXmlNullable(CT_UIntM.prototype.fromVal(this.rSp), "m:rSp");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CEqArray.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathEqArrPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "eqArrPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content.push(elem);
					break;
				}
			}
		}
		this.init(props);
		this.initPostOpen(props.mcJc);
	};
	CEqArray.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:eqArrPr", this);
		for (let i = 0; i < this.elements.length; i++) {
			writer.WriteXmlNullable(this.getElement(i), "m:e");
		}
		writer.WriteXmlNodeEnd(name);
	};
	CMathFractionPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "type" : {
					this.type = fromXml_ST_FType(CT_StringM.prototype.toVal(reader, this.type), this.type);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathFractionPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_FType(this.type)), "m:type");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CFraction.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathFractionPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "fPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "num" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
				case "den" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CFraction.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:fPr", this);
		writer.WriteXmlNullable(this.getNumerator(), "m:num");
		writer.WriteXmlNullable(this.getDenominator(), "m:den");
		writer.WriteXmlNodeEnd(name);
	};
	CMathBasePr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathBasePr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CMathFunc.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathBasePr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "funcPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "fName" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CMathFunc.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:funcPr", this);
		writer.WriteXmlNullable(this.getFName(), "m:fName");
		writer.WriteXmlNullable(this.getArgument(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathGroupChrPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "chr" : {
					this.chr = CMathBase.prototype.ConvertStrToOperator(CT_StringM.prototype.toVal(reader, ""));
					break;
				}
				case "pos" : {
					this.pos = fromXml_ST_TopBot(CT_StringM.prototype.toVal(reader, this.pos), this.pos);
					break;
				}
				case "vertJc" : {
					this.vertJc = fromXml_ST_TopBot(CT_StringM.prototype.toVal(reader, this.vertJc), this.vertJc);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathGroupChrPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(CMathBase.prototype.ConvertOperatorToStr(this.chr)), "m:chr");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_TopBot(this.pos)), "m:pos");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_TopBot(this.vertJc)), "m:vertJc");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CGroupCharacter.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathGroupChrPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "groupChrPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CGroupCharacter.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:groupChrPr", this);
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathLimitPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathLimitPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CLimit.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathLimitPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "limLowPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
				case "lim" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CLimit.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:limLowPr", this);
		writer.WriteXmlNullable(this.getFName(), "m:e");
		writer.WriteXmlNullable(this.getIterator(), "m:lim");
		writer.WriteXmlNodeEnd(name);
	};
	CMathMatrixPr.prototype.fromXml = function (reader, mathElem) {
		let t = this;
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "baseJc" : {
					this.baseJc = fromXml_ST_YAlignMath(CT_StringM.prototype.toVal(reader, this.baseJc), this.baseJc);
					break;
				}
				case "plcHide" : {
					this.plcHide = CT_BoolM.prototype.toVal(reader, this.plcHide);
					break;
				}
				case "rSpRule" : {
					this.rSpRule = Math.min(4, CT_UIntM.prototype.toVal(reader, this.rSpRule));
					break;
				}
				case "cGpRule" : {
					this.cGpRule = Math.min(4, CT_UIntM.prototype.toVal(reader, this.cGpRule));
					break;
				}
				case "rSp" : {
					this.rSp = CT_UIntM.prototype.toVal(reader, this.rSp);
					break;
				}
				case "cSp" : {
					this.cSp = CT_UIntM.prototype.toVal(reader, this.cSp);
					break;
				}
				case "cGp" : {
					this.cGp = CT_UIntM.prototype.toVal(reader, this.cGp);
					break;
				}
				case "mcs" : {
					reader.readXmlArray("mc", function() {
						let elem = new CT_XmlNode();
						elem.fromXml(reader);
						if (elem.members["mcPr"]) {
							let count = 0;
							let mcJc = MCJC_CENTER;
							if (elem.members["mcPr"].members["count"]) {
								count = AscCommon.StaxParser.prototype.GetUInt(elem.members["mcPr"].members["count"].attributes["val"], count);
							}
							if (elem.members["mcPr"].members["mcJc"]) {
								mcJc = fromXml_ST_XAlignMath(elem.members["mcPr"].members["mcJc"].attributes["val"], mcJc);
							}
							t.mcs.push({count: count, mcJc: mcJc});
						}
					});
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathMatrixPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_YAlignMath(this.baseJc)), "m:baseJc");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.plcHide), "m:plcHide");
		writer.WriteXmlNullable(CT_UIntM.prototype.fromVal(this.rSpRule), "m:rSpRule");
		writer.WriteXmlNullable(CT_UIntM.prototype.fromVal(this.cGpRule), "m:cGpRule");
		writer.WriteXmlNullable(CT_UIntM.prototype.fromVal(this.rSp), "m:rSp");
		writer.WriteXmlNullable(CT_UIntM.prototype.fromVal(this.cSp), "m:cSp");
		writer.WriteXmlNullable(CT_UIntM.prototype.fromVal(this.cGp), "m:cGp");
		if(this.mcs.length > 0) {
			writer.WriteXmlNodeStart("m:mcs");
			writer.WriteXmlAttributesEnd();
			this.mcs.forEach(function(mc) {
				let mcPr = new CT_XmlNode();
				if (null != mc.mcJc) {
					let mcJc = new CT_XmlNode();
					mcJc.attributes["m:val"] = toXml_ST_XAlignMath(mc.mcJc);
					mcPr.members["m:mcJc"] = mcJc;
				}
				if (null != mc.count) {
					let count = new CT_XmlNode();
					count.attributes["m:val"] = mc.count.toString();
					mcPr.members["m:count"] = count;
				}
				let elem = new CT_XmlNode();
				elem.members["m:mcPr"] = mcPr;
				writer.WriteXmlNullable(elem, "m:mc");
			});
			writer.WriteXmlNodeEnd("m:mcs");
		}
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CMathMatrix.prototype.fromXml = function (reader) {
		let props = new CMathMatrixPr(), depth = reader.GetDepth();
		props.mrs = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "mPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "mr" : {
					let row = [];
					reader.readXmlArray("e", function() {
						let e = new CMathContent();
						e.fromXml(reader);
						row.push(e);
					});
					props.mrs.push(row);
					break;
				}
			}
		}
		this.init(props);
	};
	CMathMatrix.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:mPr", this);
		for (let i = 0; i < this.nRow; ++i) {
			writer.WriteXmlNodeStart("m:mr");
			writer.WriteXmlAttributesEnd();
			for (let j = 0; j < this.nCol; ++j) {
				writer.WriteXmlNullable(this.getElement(i,j), "m:e");
			}
			writer.WriteXmlNodeEnd("m:mr");
		}
		writer.WriteXmlNodeEnd(name);
	};
	CMathNaryPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "chr" : {
					this.chr = CMathBase.prototype.ConvertStrToOperator(CT_StringM.prototype.toVal(reader, ""));
					break;
				}
				case "limLoc" : {
					this.limLoc = fromXml_ST_LimLoc(CT_StringM.prototype.toVal(reader, this.limLoc), this.limLoc);
					break;
				}
				case "grow" : {
					this.grow = CT_BoolM.prototype.toVal(reader, this.grow);
					break;
				}
				case "subHide" : {
					this.subHide = CT_BoolM.prototype.toVal(reader, this.subHide);
					break;
				}
				case "supHide" : {
					this.supHide = CT_BoolM.prototype.toVal(reader, this.supHide);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathNaryPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(CMathBase.prototype.ConvertOperatorToStr(this.chr)), "m:chr");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_LimLoc(this.limLoc)), "m:limLoc");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.grow), "m:grow");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.subHide), "m:subHide");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.supHide), "m:supHide");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CNary.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathNaryPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "naryPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "sub" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
				case "sup" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[2] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CNary.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:naryPr", this);
		writer.WriteXmlNullable(this.getLowerIterator(), "m:sub");
		writer.WriteXmlNullable(this.getUpperIterator(), "m:sup");
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathPhantomPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "show" : {
					this.show = CT_BoolM.prototype.toVal(reader, this.show);
					break;
				}
				case "zeroWid" : {
					this.zeroWid = CT_BoolM.prototype.toVal(reader, this.zeroWid);
					break;
				}
				case "zeroAsc" : {
					this.zeroAsc = CT_BoolM.prototype.toVal(reader, this.zeroAsc);
					break;
				}
				case "zeroDesc" : {
					this.zeroDesc = CT_BoolM.prototype.toVal(reader, this.zeroDesc);
					break;
				}
				case "transp" : {
					this.transp = CT_BoolM.prototype.toVal(reader, this.transp);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathPhantomPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.show), "m:show");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.zeroWid), "m:zeroWid");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.zeroAsc), "m:zeroAsc");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.zeroDesc), "m:zeroDesc");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.transp), "m:transp");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CPhantom.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathPhantomPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "phantPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CPhantom.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:phantPr", this);
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathRadicalPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "degHide" : {
					this.degHide = CT_BoolM.prototype.toVal(reader, this.degHide);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathRadicalPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.degHide), "m:degHide");
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CRadical.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathRadicalPr();
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "radPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "deg" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CRadical.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		this.Pr.toXml(writer, "m:radPr", this);
		writer.WriteXmlNullable(this.getDegree(), "m:deg");
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNodeEnd(name);
	};
	CMathDegreeSubSupPr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "alnScr" : {
					this.alnScr = CT_BoolM.prototype.toVal(reader, this.alnScr);
					break;
				}
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathDegreeSubSupPr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		if (DEGREE_SubSup == this.type) {
			writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.alnScr), "m:alnScr");
		}
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CDegreeSubSup.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathDegreeSubSupPr();
		props.type = this.Pr.type;
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "sPrePr" :
				case "sSubSupPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
				case "sub" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[2] = elem;
					break;
				}
				case "sup" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CDegreeSubSup.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		if (DEGREE_PreSubSup == this.Pr.type) {
			this.Pr.toXml(writer, "m:sPrePr", this);
		}
		else if (DEGREE_SubSup == this.Pr.type) {
			this.Pr.toXml(writer, "m:sSubSupPr", this);
		}
		writer.WriteXmlNullable(this.getBase(), "m:e");
		writer.WriteXmlNullable(this.getLowerIterator(), "m:sub");
		writer.WriteXmlNullable(this.getUpperIterator(), "m:sup");
		writer.WriteXmlNodeEnd(name);
	};
	CMathDegreePr.prototype.fromXml = function (reader, mathElem) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ctrlPr" : {
					this.ctrPrp = CMathBase.prototype.fromXmlCtrlPr.call(this, reader, mathElem);
					break;
				}
			}
		}
	};
	CMathDegreePr.prototype.toXml = function (writer, name, elem) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		CMathBase.prototype.toXmlCtrlPr(writer, "m:ctrlPr", elem);
		writer.WriteXmlNodeEnd(name);
	};
	CDegree.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		let props = new CMathDegreePr();
		props.type = this.Pr.type;
		props.content = [];
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "sSubPr" :
				case "sSupPr" : {
					props.fromXml(reader, this);
					break;
				}
				case "e" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[0] = elem;
					break;
				}
				case "sub" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
				case "sup" : {
					elem = new CMathContent();
					elem.fromXml(reader);
					props.content[1] = elem;
					break;
				}
			}
		}
		this.init(props);
	};
	CDegree.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		if (DEGREE_SUBSCRIPT == this.Pr.type) {
			this.Pr.toXml(writer, "m:sSubPr", this);
			writer.WriteXmlNullable(this.getBase(), "m:e");
			writer.WriteXmlNullable(this.getLowerIterator(), "m:sub");
		}
		else if (DEGREE_SUPERSCRIPT == this.Pr.type) {
			this.Pr.toXml(writer, "m:sSupPr", this);
			writer.WriteXmlNullable(this.getBase(), "m:e");
			writer.WriteXmlNullable(this.getUpperIterator(), "m:sup");
		}
		writer.WriteXmlNodeEnd(name);
	};
	CMathArgSize.prototype.fromXml = function (reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "argSz" : {
					this.value = CT_IntM.prototype.toVal(reader, this.value);
					break;
				}
			}
		}
	};
	CMathArgSize.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.value), "m:argSz");
		writer.WriteXmlNodeEnd(name);
	};
	CMathPropertiesSettings.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "mathFont" : {
					this.mathFont = CT_StringM.prototype.toVal(reader, this.mathFont);
					break;
				}
				case "brkBin" : {
					this.brkBin = fromXml_ST_BreakBin(CT_StringM.prototype.toVal(reader, this.brkBin), this.brkBin);
					break;
				}
				case "brkBinSub" : {
					this.brkBinSub = fromXml_ST_BreakBinSub(CT_StringM.prototype.toVal(reader, this.brkBinSub), this.brkBinSub);
					break;
				}
				case "smallFrac" : {
					this.smallFrac = CT_BoolM.prototype.toVal(reader, this.smallFrac);
					break;
				}
				case "dispDef" : {
					this.dispDef = CT_BoolM.prototype.toVal(reader, this.dispDef);
					break;
				}
				case "lMargin" : {
					this.lMargin = AscCommon.universalMeasureToMm(CT_StringM.prototype.toVal(reader, this.lMargin), AscCommonWord.g_dKoef_twips_to_mm, this.lMargin);
					break;
				}
				case "rMargin" : {
					this.rMargin = AscCommon.universalMeasureToMm(CT_StringM.prototype.toVal(reader, this.rMargin), AscCommonWord.g_dKoef_twips_to_mm, this.rMargin);
					break;
				}
				case "defJc" : {
					this.defJc = fromXml_ST_Jc(CT_StringM.prototype.toVal(reader, this.defJc), this.defJc);
					break;
				}
				case "preSp" : {
					this.preSp = AscCommon.universalMeasureToMm(CT_StringM.prototype.toVal(reader, this.preSp), AscCommonWord.g_dKoef_twips_to_mm, this.preSp);
					break;
				}
				case "postSp" : {
					this.postSp = AscCommon.universalMeasureToMm(CT_StringM.prototype.toVal(reader, this.postSp), AscCommonWord.g_dKoef_twips_to_mm, this.postSp);
					break;
				}
				case "interSp" : {
					this.interSp = AscCommon.universalMeasureToMm(CT_StringM.prototype.toVal(reader, this.interSp), AscCommonWord.g_dKoef_twips_to_mm, this.interSp);
					break;
				}
				case "intraSp" : {
					this.intraSp = AscCommon.universalMeasureToMm(CT_StringM.prototype.toVal(reader, this.intraSp), AscCommonWord.g_dKoef_twips_to_mm, this.intraSp);
					break;
				}
				case "wrapIndent" : {
					this.wrapIndent = AscCommon.universalMeasureToMm(CT_StringM.prototype.toVal(reader, this.wrapIndent), AscCommonWord.g_dKoef_twips_to_mm, this.wrapIndent);
					break;
				}
				case "wrapRight" : {
					this.wrapRight = CT_BoolM.prototype.toVal(reader, this.wrapRight);
					break;
				}
				case "intLim" : {
					this.intLim = fromXml_ST_LimLoc(CT_StringM.prototype.toVal(reader, this.intLim), this.intLim);
					break;
				}
				case "naryLim" : {
					this.naryLim = fromXml_ST_LimLoc(CT_StringM.prototype.toVal(reader, this.naryLim), this.naryLim);
					break;
				}
			}
		}
	};
	CMathPropertiesSettings.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(this.mathFont), "m:mathFont");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_BreakBin(this.brkBin)), "m:brkBin");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_BreakBinSub(this.brkBinSub)), "m:brkBinSub");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.smallFrac), "m:smallFrac");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.dispDef), "m:dispDef");
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.lMargin, g_dKoef_mm_to_twips), "m:lMargin");
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.rMargin, g_dKoef_mm_to_twips), "m:rMargin");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_Jc(this.defJc)), "m:defJc");
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.preSp, g_dKoef_mm_to_twips), "m:preSp");
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.postSp, g_dKoef_mm_to_twips), "m:postSp");
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.interSp, g_dKoef_mm_to_twips), "m:interSp");
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.intraSp, g_dKoef_mm_to_twips), "m:intraSp");
		writer.WriteXmlNullable(CT_IntM.prototype.fromVal(this.wrapIndent, g_dKoef_mm_to_twips), "m:wrapIndent");
		writer.WriteXmlNullable(CT_BoolM.prototype.fromVal(this.wrapRight), "m:wrapRight");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_LimLoc(this.intLim)), "m:intLim");
		writer.WriteXmlNullable(CT_StringM.prototype.fromVal(toXml_ST_LimLoc(this.naryLim)), "m:naryLim");
		writer.WriteXmlNodeEnd(name);
	};
	
	function CT_TrackChange(id, ReviewInfo) {
		this.id = id;
		this.ReviewInfo = ReviewInfo || new CReviewInfo();
		this.paragraphContent = undefined;
		this.run = undefined;
		this.writeCallback = undefined;
		this.runContent = undefined;
		this.pPrChange = undefined;
		this.rPrChange = undefined;
		this.tblPrChange = undefined;
		this.tblGridChange = undefined;
		this.trPrChange = undefined;
		this.tcPrChange = undefined;
		this.sectPrChange = undefined;
		this.VMerge = undefined;
		this.VMergeOrigin = undefined;
		return this;
	}
	CT_TrackChange.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "id": {
					this.id = reader.GetValueInt(this.id);
					break;
				}
				case "author": {
					this.ReviewInfo.UserName = reader.GetValueDecodeXml();
					break;
				}
				case "date": {
					let dateStr = reader.GetValueDecodeXml();
					let dateMs = AscCommon.getTimeISO8601(dateStr);
					if (isNaN(dateMs)) {
						dateMs = new Date().getTime();
					}
					this.ReviewInfo.DateTime = dateMs;
					break;
				}
				case "oouserid": {
					this.ReviewInfo.UserId = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	CT_TrackChange.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		if (this.run) {
			this.run.fromXml(reader);
			return;
		}
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			let name = reader.GetNameNoNS();
			switch (name) {
				case "rPr": {
					this.rPrChange = new CTextPr();
					this.rPrChange.fromXml(reader);
					break;
				}
				case "pPr": {
					this.pPrChange = new CParaPr();
					break;
				}
				case "tcPr": {
					this.tcPrChange = new CTableCellPr();
					this.tcPrChange.fromXml(reader);
					break;
				}
				case "trPr": {
					this.trPrChange = new CTableRowPr();
					this.trPrChange.fromXml(reader);
					break;
				}
				case "tblGrid": {
					this.tblGridChange = new AscCommonWord.CT_TblGrid();
					this.tblGridChange.fromXml(reader);
					//todo
					break;
				}
				case "tblPr": {
					this.tblPrChange = new CTablePr();
					this.tblPrChange.fromXml(reader);
					break;
				}
				default:
					if (this.paragraphContent) {
						CParagraphContentWithParagraphLikeContent.prototype.fromXmlElem.call(this.paragraphContent, reader, name);
					}
					break;
			}
		}
	};
	CT_TrackChange.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:id", this.id);
		if (this.ReviewInfo) {
			let dateUtc = this.ReviewInfo.DateTime ? new Date(this.ReviewInfo.DateTime).toISOString().slice(0, 19) + 'Z' : null;
			writer.WriteXmlNonEmptyAttributeStringEncode("w:author", this.ReviewInfo.UserName);
			writer.WriteXmlNullableAttributeString("w:date", dateUtc);
			writer.WriteXmlNonEmptyAttributeStringEncode("oouserid", this.ReviewInfo.UserId);
		}
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.rPrChange, "w:rPr");
		writer.WriteXmlNullable(this.pPrChange, "w:pPr");
		writer.WriteXmlNullable(this.tcPrChange, "w:tcPr");
		writer.WriteXmlNullable(this.trPrChange, "w:trPr");
		writer.WriteXmlNullable(this.tblGridChange, "w:tblGrid");
		writer.WriteXmlNullable(this.tblPrChange, "w:tblPr");
		if (this.writeCallback) {
			this.writeCallback();
		}
		writer.WriteXmlNodeEnd(name);
	};

	function fromXml_ST_Script(val, def) {
		switch (val) {
			case "roman":
				return TXT_ROMAN;
			case "script":
				return TXT_SCRIPT;
			case "fraktur":
				return TXT_FRAKTUR;
			case "double-struck":
				return TXT_DOUBLE_STRUCK;
			case "sans-serif":
				return TXT_SANS_SERIF;
			case "monospace":
				return TXT_MONOSPACE;
		}
		return def;
	}
	function toXml_ST_Script(val) {
		switch (val) {
			case TXT_ROMAN:
				return "roman";
			case TXT_SCRIPT:
				return "script";
			case TXT_FRAKTUR:
				return "fraktur";
			case TXT_DOUBLE_STRUCK:
				return "double-struck";
			case TXT_SANS_SERIF:
				return "sans-serif";
			case TXT_MONOSPACE:
				return "monospace";
		}
		return null;
	}

	function fromXml_ST_Style(val, def) {
		switch (val) {
			case "p":
				return STY_PLAIN;
			case "b":
				return STY_BOLD;
			case "i":
				return STY_ITALIC;
			case "bi":
				return STY_BI;
		}
		return def;
	}
	function toXml_ST_Style(val) {
		switch (val) {
			case STY_PLAIN:
				return "p";
			case STY_BOLD:
				return "b";
			case STY_ITALIC:
				return "i";
			case STY_BI:
				return "bi";
		}
		return null;
	}

	function fromXml_ST_Jc(val, def) {
		switch (val) {
			case "left":
				return align_Left;
			case "right":
				return align_Right;
			case "center":
				return align_Center;
			case "centerGroup":
				return align_Justify;
		}
		return def;
	}
	function toXml_ST_Jc(val) {
		switch (val) {
			case align_Left:
				return "left";
			case align_Right:
				return "right";
			case align_Center:
				return "center";
			case align_Justify:
				return "centerGroup";
		}
		return null;
	}

	function fromXml_ST_TopBot(val, def) {
		switch (val) {
			case "top":
				return LOCATION_TOP;
			case "bot":
				return LOCATION_BOT;
		}
		return def;
	}
	function toXml_ST_TopBot(val) {
		switch (val) {
			case LOCATION_TOP:
				return "top";
			case LOCATION_BOT:
				return "bot";
		}
		return null;
	}

	function fromXml_ST_Shp(val, def) {
		switch (val) {
			case "centered":
				return DELIMITER_SHAPE_CENTERED;
			case "match":
				return DELIMITER_SHAPE_MATCH;
		}
		return def;
	}

	function toXml_ST_Shp(val) {
		switch (val) {
			case DELIMITER_SHAPE_CENTERED:
				return "centered";
			case DELIMITER_SHAPE_MATCH:
				return "match";
		}
		return null;
	}

	function fromXml_ST_YAlignMath(val) {
		//todo
		switch (val) {
			case "inline":
				return BASEJC_TOP;
			case "top":
				return BASEJC_TOP;
			case "center":
				return BASEJC_CENTER;
			case "bottom":
				return BASEJC_BOTTOM;
			case "inside":
				return BASEJC_TOP;
			case "outside":
				return BASEJC_TOP;
		}
		return undefined;
	}
	function toXml_ST_YAlignMath(val) {
		switch (val) {
			case BASEJC_INLINE:
				return "inline";
			case BASEJC_TOP:
				return "top";
			case BASEJC_CENTER:
				return "center";
			case BASEJC_BOTTOM:
				return "bottom";
			case BASEJC_INSIDE:
				return "inside";
			case BASEJC_OUTSIDE:
				return "outside";
		}
		return null;
	}

	function fromXml_ST_XAlignMath(val, def) {
		switch (val) {
			case "left":
				return MCJC_LEFT;
			case "center":
				return MCJC_CENTER;
			case "right":
				return MCJC_RIGHT;
			case "inside":
				return MCJC_INSIDE;
			case "outside":
				return MCJC_OUTSIDE;
		}
		return def;
	}
	function toXml_ST_XAlignMath(val) {
		switch (val) {
			case MCJC_LEFT:
				return "left";
			case MCJC_CENTER:
				return "center";
			case MCJC_RIGHT:
				return "right";
			case MCJC_INSIDE:
				return "inside";
			case MCJC_OUTSIDE:
				return "outside";
		}
		return null;
	}

	function fromXml_ST_FType(val, def) {
		switch (val) {
			case "bar":
				return BAR_FRACTION;
			case "skw":
				return SKEWED_FRACTION;
			case "lin":
				return LINEAR_FRACTION;
			case "noBar":
				return NO_BAR_FRACTION;
		}
		return def;
	}
	function toXml_ST_FType(val) {
		switch (val) {
			case BAR_FRACTION:
				return "bar";
			case SKEWED_FRACTION:
				return "skw";
			case LINEAR_FRACTION:
				return "lin";
			case NO_BAR_FRACTION:
				return "noBar";
		}
		return null;
	}

	function fromXml_ST_LimLoc(val, def) {
		switch (val) {
			case "undOvr":
				return NARY_UndOvr;
			case "subSup":
				return NARY_SubSup;
		}
		return def;
	}
	function toXml_ST_LimLoc(val) {
		switch (val) {
			case NARY_UndOvr:
				return "undOvr";
			case NARY_SubSup:
				return "subSup";
		}
		return null;
	}
	function fromXml_ST_BreakBin(val, def) {
		switch (val) {
			case "before":
				return BREAK_BEFORE;
			case "after":
				return BREAK_AFTER;
			case "repeat":
				return BREAK_REPEAT;
		}
		return def;
	}
	function toXml_ST_BreakBin(val) {
		switch (val) {
			case BREAK_BEFORE:
				return "before";
			case BREAK_AFTER:
				return "after";
			case BREAK_REPEAT:
				return "repeat";
		}
		return null;
	}
	function fromXml_ST_BreakBinSub(val, def) {
		switch (val) {
			case "--":
				return BREAK_MIN_MIN;
			case "-+":
				return BREAK_MIN_PLUS;
			case "+-":
				return BREAK_PLUS_MIN;
		}
		return def;
	}
	function toXml_ST_BreakBinSub(val) {
		switch (val) {
			case BREAK_MIN_MIN:
				return "--";
			case BREAK_MIN_PLUS:
				return "-+";
			case BREAK_PLUS_MIN:
				return "+-";
		}
		return null;
	}

	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CT_OMathPara = CT_OMathPara;
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CT_TrackChange = CT_TrackChange;
})(window);
