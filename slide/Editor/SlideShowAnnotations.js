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
function (window, undefined) {



	function CAnnotations(oSlide) {
		this.inks = [];
		this.slide = oSlide;
	}

	CAnnotations.prototype.addInk = function (oInk) {
		oInk.recalculate();
		this.inks.push(oInk);
		this.onUpdate();
	};
	CAnnotations.prototype.eraseInk = function (nIdx) {
		this.inks.splice(nIdx, 1);
		this.onUpdate();
	};
	CAnnotations.prototype.draw = function (oGraphics) {
		for(let nIdx = 0; nIdx < this.inks.length; ++nIdx) {
			this.inks[nIdx].draw(oGraphics);
		}
	};
	CAnnotations.prototype.onUpdate = function () {
		let oManager = Asc.editor.getDemoManager();
		if(oManager) {
			if(oManager.GetCurrentSlide() === this.slide) {
				oManager.Redraw();
			}
		}
	};
	CAnnotations.prototype.isEmpty = function() {
		return this.inks.length === 0;
	};
	CAnnotations.prototype.saveAnnotations = function () {
		for(let nIdx = 0; nIdx < this.inks.length; ++nIdx) {
			this.slide.addToSpTreeToPos(undefined, this.inks[nIdx].copy());
		}
	};
	CAnnotations.prototype.clear = function () {
		this.inks.length = 0;
		this.onUpdate();
	};

	function CSlideShowAnnotations() {
		this.annotations = {};
		this.track = null;
		this.locked = false;

		this.binaryWriter = new AscCommon.CMemory(true);
		this.binaryWriter.Init(1024*1024);
	}
	CSlideShowAnnotations.prototype.clear = function () {
		this.annotations = {};
		this.track = null;
		this.locked = false;
		this.binaryWriter.Seek(0);
	};
	CSlideShowAnnotations.prototype.clearTrack = function () {
		this.track = null;
		let oManager = Asc.editor.getDemoManager();
		if(oManager) {
			oManager.Redraw();
		}
	};
	CSlideShowAnnotations.prototype.getInks = function (oSlide) {
		if(!this.annotations[oSlide.Id]) {
			return [];
		}
		return this.annotations[oSlide.Id].inks;
	};
	CSlideShowAnnotations.prototype.getPresentation = function () {
		return Asc.editor.getLogicDocument();
	};
	CSlideShowAnnotations.prototype.onCreatePolylineTrack = function (track, slide) {
		this.track = track;
		this.binaryWriter.Seek(0);
		track.pen.Write_ToBinary(this.binaryWriter)
		let sPenBin = this.binaryWriter.GetBase64Memory();
		this.sendData("create_track", slide.Id + ";" + sPenBin);
	};
	CSlideShowAnnotations.prototype.onPolylineTrackChanged = function (track) {
		this.binaryWriter.Seek(0);
		track.serialize(this.binaryWriter);
		this.sendData("change_track", this.binaryWriter.GetBase64Memory());

		if(this.track) {
			let oManager = Asc.editor.getDemoManager();
			if(oManager) {
				if(oManager.GetCurrentSlide() === this.track.drawingObjects.drawingObjects) {
					oManager.Redraw();
				}
			}
		}
	};
	CSlideShowAnnotations.prototype.addInkInternal = function (oInk) {
		let oSlide = oInk.parent;
		if(!this.annotations[oSlide.Id]) {
			this.annotations[oSlide.Id] = new CAnnotations(oSlide);
		}
		let oAnnots = this.annotations[oSlide.Id];
		oAnnots.addInk(oInk);

	};
	CSlideShowAnnotations.prototype.addInk = function (oInk) {
		let oSlide = oInk.parent;
		this.track = null;
		this.addInkInternal(oInk);
		let pptx_writer = new AscCommon.CBinaryFileWriter();
		pptx_writer.WriteShape(oInk);
		let sValue = oSlide.Id + ";" + pptx_writer.pos + ";" + pptx_writer.GetBase64Memory();

		this.sendData("add_ink", sValue);
	};
	CSlideShowAnnotations.prototype.eraseInk = function (oSlide, nIdx) {
		let oAnnots = this.annotations[oSlide.Id];
		if(!oAnnots) {
			return;
		}
		oAnnots.eraseInk(nIdx);
		this.sendData("erase_ink", oSlide.Id + ";" + nIdx);
	};
	CSlideShowAnnotations.prototype.clearOnSlide = function (oSlide) {
		let oAnnots = this.annotations[oSlide.Id];
		if(!oAnnots) {
			return;
		}
		oAnnots.clear();
		this.sendData("erase_ink_all", oSlide.Id);
	};
	CSlideShowAnnotations.prototype.sendData = function(sType, sValue) {
		let oData = {
			"type": sType,
			"value": sValue
		}
		this.sendMessage(JSON.stringify(oData));
	};
	CSlideShowAnnotations.prototype.sendMessage = function (value) {
		if(Asc.editor.isReporterMode) {
			Asc.editor.sendFromReporter("{ \"reporter_command\" : \"annotation\", \"value\" : " + value + " }");
		}
		else {
			Asc.editor.sendToReporter("{ \"main_command\" : true, \"annotation\" : " + value + " }");
		}
	};
	CSlideShowAnnotations.prototype.handleMessage = function(oData) {
		let sType = oData["type"];
		let sVal = oData["value"]

		switch (sType) {
			case "create_track": {
				let aParts = sVal.split(";");
				let oSlide = AscCommon.g_oTableId.Get_ById(aParts[0]);
				let memoryData = AscCommon.Base64.decode(aParts[1], true, undefined, 0);
				let r = new AscCommon.FT_Stream2(memoryData, memoryData.length);
				Asc.editor.getAnnotations = function () {return null;};
				this.track = new AscFormat.PolyLine(oSlide.graphicObjects, oSlide.getTheme(), null, null, null, oSlide.num);

				delete Asc.editor.getAnnotations;
				//this.track.pen = new AscFormat.CLn();
				this.track.pen.Read_FromBinary(r);
				this.track.pen.Fill.calculate(oSlide.getTheme(), oSlide, null, null,  {R: 0, G: 0, B: 0, A: 255})
				break;
			}
			case "change_track": {
				if(this.track) {
					let memoryData = AscCommon.Base64.decode(sVal, true, undefined, 0);
					let r = new AscCommon.FT_Stream2(memoryData, memoryData.length);
					this.track.deserialize(r);
					let oManager = Asc.editor.getDemoManager();
					if(oManager) {
						if(oManager.GetCurrentSlide() === this.track.drawingObjects.drawingObjects) {
							oManager.Redraw();
						}
					}
				}
				break;
			}
			case "add_ink": {
				let aParts = sVal.split(";");
				let sSlideId = aParts[0];
				let oSlide = AscCommon.g_oTableId.Get_ById(sSlideId);
				let sBinary = aParts[1]  + ";" + aParts[2];
				let oBinaryReader = AscFormat.CreatePPTYLoader(sBinary, 0, sBinary.length);
				oBinaryReader.TempMainObject = oSlide;
				this.track = null;
				AscFormat.ExecuteNoHistory(function () {
					let oShape = oBinaryReader.ReadGraphicObject();
					oShape.setParent(oSlide);
					this.addInkInternal(oShape);
				}, this, []);
				break;
			}
			case "erase_ink": {
				let aParts = sVal.split(";");
				let nIdx = parseInt(aParts[1]);
				let oAnnots = this.annotations[aParts[0]];
				if(!oAnnots) {
					return;
				}
				this.track = null;
				oAnnots.eraseInk(nIdx);
				break;
			}
			case "erase_ink_all": {
				let oAnnots = this.annotations[sVal];
				if(!oAnnots) {
					return;
				}
				this.track = null;
				oAnnots.clear();
				break;
			}
		}

	};
	CSlideShowAnnotations.prototype.draw = function(oGraphics, oSlide) {
		let oAnnots = this.annotations[oSlide.Id];
		let bDrawTrack = (this.track && this.track.drawingObjects.drawingObjects === oSlide);

		if(oAnnots && oAnnots.inks.length > 0 || bDrawTrack) {
			oGraphics.SaveGrState();
			oGraphics.AddClipRect(0, 0, oSlide.Width, oSlide.Height);

			if(oAnnots) {
				oAnnots.draw(oGraphics);
			}
			if(bDrawTrack) {
				this.track.draw(oGraphics);
			}
			oGraphics.RestoreGrState();
		}
	};
	CSlideShowAnnotations.prototype.isEmpty = function() {
		for(let sKey in this.annotations) {
			if(this.annotations.hasOwnProperty(sKey)) {
				if(!this.annotations[sKey].isEmpty()) {
					return false;
				}
			}
		}
		return true;
	};
	CSlideShowAnnotations.prototype.canSaveAnnotations = function() {
		if(!Asc.editor.canEdit()) {
			return false;
		}
		return !this.isEmpty();
	};
	CSlideShowAnnotations.prototype.saveAnnotations = function () {
		if(!this.canSaveAnnotations()) {
			this.clear();
			return;
		}
		let oPresentation = this.getPresentation();
		oPresentation.StartAction(AscDFH.historydescription_Presentation_SaveAnnotations);
		for(let sKey in this.annotations) {
			if(this.annotations.hasOwnProperty(sKey)) {
				this.annotations[sKey].saveAnnotations();
			}
		}
		oPresentation.FinalizeAction(false);
		this.clear();
	};

	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonSlide'].CSlideShowAnnotations = CSlideShowAnnotations;

})(window);
