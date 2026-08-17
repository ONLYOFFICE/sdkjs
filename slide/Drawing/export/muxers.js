/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function (window, undefined) {

	window["AscCommonSlide"] = window['AscCommonSlide'] || {};

	function concat(parts) {
		let length = 0;
		for (let index = 0; index < parts.length; ++index) {
			length += parts[index].length;
		}
		const result = new Uint8Array(length);
		let position = 0;
		for (let index = 0; index < parts.length; ++index) {
			result.set(parts[index], position);
			position += parts[index].length;
		}
		return result;
	}

	// ---------------------------------------------------------------------------
	// WebP: RIFF container with animation chunks
	// ---------------------------------------------------------------------------
	const RIFF_HEADER_SIZE = 12;
	const CHUNK_HEADER_SIZE = 8;
	const VP8X_FLAG_ANIMATION = 0x02;
	const VP8X_FLAG_ALPHA = 0x10;
	const ANMF_FLAG_NO_BLEND = 0x02;
	const ANMF_HEADER_SIZE = 16;
	const MAX_DURATION = 0xFFFFFF;

	function readTag(bytes, position) {
		return String.fromCharCode(bytes[position], bytes[position + 1], bytes[position + 2], bytes[position + 3]);
	}
	function readUint32(bytes, position) {
		return (bytes[position] | (bytes[position + 1] << 8) | (bytes[position + 2] << 16)) +
			bytes[position + 3] * 16777216;
	}

	function extractFrameChunks(bytes) {
		const result = {alpha: null, bitstream: null, hasAlpha: false};
		if (bytes.length < RIFF_HEADER_SIZE || "RIFF" !== readTag(bytes, 0) || "WEBP" !== readTag(bytes, 8)) {
			return result;
		}
		let position = RIFF_HEADER_SIZE;
		while (position + CHUNK_HEADER_SIZE <= bytes.length) {
			const tag = readTag(bytes, position);
			const size = readUint32(bytes, position + 4);
			const chunkSize = CHUNK_HEADER_SIZE + size + (size & 1);
			if (position + chunkSize > bytes.length) {
				break;
			}
			if ("ALPH" === tag) {
				result.alpha = bytes.subarray(position, position + chunkSize);
				result.hasAlpha = true;
			}
			else if ("VP8 " === tag || "VP8L" === tag) {
				result.bitstream = bytes.subarray(position, position + chunkSize);
				if ("VP8L" === tag) {
					result.hasAlpha = true;
				}
			}
			position += chunkSize;
		}
		return result;
	}

	function CWebPAnimEncoder(width, height, loopCount) {
		this.width = width;
		this.height = height;
		this.loopCount = undefined === loopCount ? 0 : loopCount;
		this.frames = [];
		this.hasAlpha = false;
	}
	CWebPAnimEncoder.prototype.getFrameCount = function () {
		return this.frames.length;
	};
	CWebPAnimEncoder.prototype.addFrame = function (webpBytes, rect, durationMs, isOverwrite) {
		const chunks = extractFrameChunks(webpBytes);
		if (!chunks.bitstream) {
			return false;
		}
		if (chunks.hasAlpha) {
			this.hasAlpha = true;
		}
		this.frames.push({
			x: rect.x,
			y: rect.y,
			w: rect.w,
			h: rect.h,
			duration: Math.min(MAX_DURATION, durationMs),
			isOverwrite: false !== isOverwrite,
			alpha: chunks.alpha,
			bitstream: chunks.bitstream
		});
		return true;
	};
	CWebPAnimEncoder.prototype.addDuration = function (durationMs) {
		if (0 === this.frames.length) {
			return false;
		}
		const frame = this.frames[this.frames.length - 1];
		frame.duration = Math.min(MAX_DURATION, frame.duration + durationMs);
		return true;
	};
	CWebPAnimEncoder.prototype.getDurationMs = function () {
		let total = 0;
		for (let index = 0; index < this.frames.length; ++index) {
			total += this.frames[index].duration;
		}
		return total;
	};
	CWebPAnimEncoder.prototype.getSize = function () {
		let size = RIFF_HEADER_SIZE;
		size += CHUNK_HEADER_SIZE + 10;
		size += CHUNK_HEADER_SIZE + 6;
		for (let index = 0; index < this.frames.length; ++index) {
			const frame = this.frames[index];
			const payload = ANMF_HEADER_SIZE + (frame.alpha ? frame.alpha.length : 0) + frame.bitstream.length;
			size += CHUNK_HEADER_SIZE + payload + (payload & 1);
		}
		return size;
	};
	CWebPAnimEncoder.prototype.finish = function () {
		const data = new Uint8Array(this.getSize());
		let position = 0;

		const writeTag = function (tag) {
			for (let index = 0; index < 4; ++index) {
				data[position++] = tag.charCodeAt(index);
			}
		};
		const writeUint32 = function (value) {
			data[position++] = value & 0xFF;
			data[position++] = (value >>> 8) & 0xFF;
			data[position++] = (value >>> 16) & 0xFF;
			data[position++] = (value >>> 24) & 0xFF;
		};
		const writeUint24 = function (value) {
			data[position++] = value & 0xFF;
			data[position++] = (value >>> 8) & 0xFF;
			data[position++] = (value >>> 16) & 0xFF;
		};
		const writeUint16 = function (value) {
			data[position++] = value & 0xFF;
			data[position++] = (value >>> 8) & 0xFF;
		};
		const writeBytes = function (bytes) {
			data.set(bytes, position);
			position += bytes.length;
		};

		writeTag("RIFF");
		writeUint32(data.length - CHUNK_HEADER_SIZE);
		writeTag("WEBP");

		writeTag("VP8X");
		writeUint32(10);
		data[position++] = VP8X_FLAG_ANIMATION | (this.hasAlpha ? VP8X_FLAG_ALPHA : 0);
		writeUint24(0);
		writeUint24(this.width - 1);
		writeUint24(this.height - 1);

		writeTag("ANIM");
		writeUint32(6);
		writeUint32(0xFFFFFFFF);
		writeUint16(this.loopCount);

		for (let index = 0; index < this.frames.length; ++index) {
			const frame = this.frames[index];
			const alphaLength = frame.alpha ? frame.alpha.length : 0;
			const payloadSize = ANMF_HEADER_SIZE + alphaLength + frame.bitstream.length;

			writeTag("ANMF");
			writeUint32(payloadSize);
			writeUint24(frame.x >> 1);
			writeUint24(frame.y >> 1);
			writeUint24(frame.w - 1);
			writeUint24(frame.h - 1);
			writeUint24(frame.duration);
			data[position++] = frame.isOverwrite ? ANMF_FLAG_NO_BLEND : 0;
			if (frame.alpha) {
				writeBytes(frame.alpha);
			}
			writeBytes(frame.bitstream);
			if (payloadSize & 1) {
				data[position++] = 0;
			}
		}

		return data;
	};

	window['AscCommonSlide'].CWebPAnimEncoder = CWebPAnimEncoder;
	window['AscCommonSlide'].extractWebPFrameChunks = extractFrameChunks;

	// ---------------------------------------------------------------------------
	// WebM: EBML container
	// ---------------------------------------------------------------------------
	const ID_EBML = [0x1A, 0x45, 0xDF, 0xA3];
	const ID_EBML_VERSION = [0x42, 0x86];
	const ID_EBML_READ_VERSION = [0x42, 0xF7];
	const ID_EBML_MAX_ID_LENGTH = [0x42, 0xF2];
	const ID_EBML_MAX_SIZE_LENGTH = [0x42, 0xF3];
	const ID_DOC_TYPE = [0x42, 0x82];
	const ID_DOC_TYPE_VERSION = [0x42, 0x87];
	const ID_DOC_TYPE_READ_VERSION = [0x42, 0x85];
	const ID_CUES = [0x1C, 0x53, 0xBB, 0x6B];
	const ID_CUE_POINT = [0xBB];
	const ID_CUE_TIME = [0xB3];
	const ID_CUE_TRACK_POSITIONS = [0xB7];
	const ID_CUE_TRACK = [0xF7];
	const ID_CUE_CLUSTER_POSITION = [0xF1];
	const ID_SEGMENT = [0x18, 0x53, 0x80, 0x67];
	const ID_INFO = [0x15, 0x49, 0xA9, 0x66];
	const ID_TIMECODE_SCALE = [0x2A, 0xD7, 0xB1];
	const ID_DURATION = [0x44, 0x89];
	const ID_MUXING_APP = [0x4D, 0x80];
	const ID_WRITING_APP = [0x57, 0x41];
	const ID_TRACKS = [0x16, 0x54, 0xAE, 0x6B];
	const ID_TRACK_ENTRY = [0xAE];
	const ID_TRACK_NUMBER = [0xD7];
	const ID_TRACK_UID = [0x73, 0xC5];
	const ID_TRACK_TYPE = [0x83];
	const ID_FLAG_LACING = [0x9C];
	const ID_CODEC_ID = [0x86];
	const ID_VIDEO = [0xE0];
	const ID_PIXEL_WIDTH = [0xB0];
	const ID_PIXEL_HEIGHT = [0xBA];
	const ID_CLUSTER = [0x1F, 0x43, 0xB6, 0x75];
	const ID_TIMECODE = [0xE7];
	const ID_SIMPLE_BLOCK = [0xA3];

	const TRACK_NUMBER = 1;
	const TRACK_TYPE_VIDEO = 1;
	const TIMECODE_SCALE_NS = 1000000;
	const MAX_CLUSTER_DURATION_MS = 5000;
	const KEYFRAME_FLAG = 0x80;

	function encodeVint(value) {
		let length = 1;
		while (length < 8 && value >= Math.pow(2, 7 * length) - 1) {
			++length;
		}
		const bytes = new Uint8Array(length);
		let rest = value;
		for (let index = length - 1; index >= 0; --index) {
			bytes[index] = rest & 0xFF;
			rest = Math.floor(rest / 256);
		}
		bytes[0] |= 1 << (8 - length);
		return bytes;
	}
	function encodeUint(value) {
		const bytes = [];
		let rest = value;
		do {
			bytes.unshift(rest & 0xFF);
			rest = Math.floor(rest / 256);
		} while (rest > 0);
		return new Uint8Array(bytes);
	}
	function encodeFloat64(value) {
		const bytes = new Uint8Array(8);
		new DataView(bytes.buffer).setFloat64(0, value, false);
		return bytes;
	}
	function encodeString(text) {
		const bytes = new Uint8Array(text.length);
		for (let index = 0; index < text.length; ++index) {
			bytes[index] = text.charCodeAt(index) & 0xFF;
		}
		return bytes;
	}
	function element(id, payload) {
		return concat([new Uint8Array(id), encodeVint(payload.length), payload]);
	}
	function uintElement(id, value) {
		return element(id, encodeUint(value));
	}
	function stringElement(id, text) {
		return element(id, encodeString(text));
	}

	function CWebMMuxer(width, height, codecId) {
		this.width = width;
		this.height = height;
		this.codecId = codecId || "V_VP8";
		this.clusters = [];
		this.clusterTimes = [];
		this.currentBlocks = null;
		this.clusterTimeMs = 0;
		this.durationMs = 0;
		this.frameCount = 0;
	}
	CWebMMuxer.prototype.getFrameCount = function () {
		return this.frameCount;
	};
	CWebMMuxer.prototype.getDurationMs = function () {
		return this.durationMs;
	};
	CWebMMuxer.prototype.addFrame = function (data, timeMs, isKeyFrame) {
		const isNewCluster = !this.currentBlocks || isKeyFrame ||
			timeMs - this.clusterTimeMs >= MAX_CLUSTER_DURATION_MS;
		if (isNewCluster) {
			this.flushCluster();
			this.currentBlocks = [];
			this.clusterTimeMs = timeMs;
		}

		const relative = timeMs - this.clusterTimeMs;
		const header = new Uint8Array(4);
		header[0] = 0x80 | TRACK_NUMBER;
		header[1] = (relative >> 8) & 0xFF;
		header[2] = relative & 0xFF;
		header[3] = isKeyFrame ? KEYFRAME_FLAG : 0;

		this.currentBlocks.push(element(ID_SIMPLE_BLOCK, concat([header, data])));
		++this.frameCount;
		if (timeMs > this.durationMs) {
			this.durationMs = timeMs;
		}
	};
	CWebMMuxer.prototype.setDurationMs = function (durationMs) {
		if (durationMs > this.durationMs) {
			this.durationMs = durationMs;
		}
	};
	CWebMMuxer.prototype.flushCluster = function () {
		if (!this.currentBlocks || 0 === this.currentBlocks.length) {
			this.currentBlocks = null;
			return;
		}
		const payload = concat([uintElement(ID_TIMECODE, this.clusterTimeMs), concat(this.currentBlocks)]);
		this.clusters.push(element(ID_CLUSTER, payload));
		this.clusterTimes.push(this.clusterTimeMs);
		this.currentBlocks = null;
	};
	CWebMMuxer.prototype.getHeader = function () {
		return element(ID_EBML, concat([
			uintElement(ID_EBML_VERSION, 1),
			uintElement(ID_EBML_READ_VERSION, 1),
			uintElement(ID_EBML_MAX_ID_LENGTH, 4),
			uintElement(ID_EBML_MAX_SIZE_LENGTH, 8),
			stringElement(ID_DOC_TYPE, "webm"),
			uintElement(ID_DOC_TYPE_VERSION, 2),
			uintElement(ID_DOC_TYPE_READ_VERSION, 2)
		]));
	};
	CWebMMuxer.prototype.getInfo = function () {
		return element(ID_INFO, concat([
			uintElement(ID_TIMECODE_SCALE, TIMECODE_SCALE_NS),
			element(ID_DURATION, encodeFloat64(this.durationMs)),
			stringElement(ID_MUXING_APP, "ONLYOFFICE"),
			stringElement(ID_WRITING_APP, "ONLYOFFICE")
		]));
	};
	CWebMMuxer.prototype.getTracks = function () {
		const video = element(ID_VIDEO, concat([
			uintElement(ID_PIXEL_WIDTH, this.width),
			uintElement(ID_PIXEL_HEIGHT, this.height)
		]));
		const track = element(ID_TRACK_ENTRY, concat([
			uintElement(ID_TRACK_NUMBER, TRACK_NUMBER),
			uintElement(ID_TRACK_UID, TRACK_NUMBER),
			uintElement(ID_TRACK_TYPE, TRACK_TYPE_VIDEO),
			uintElement(ID_FLAG_LACING, 0),
			stringElement(ID_CODEC_ID, this.codecId),
			video
		]));
		return element(ID_TRACKS, track);
	};
	CWebMMuxer.prototype.getCues = function (clustersOffset) {
		const points = [];
		let position = clustersOffset;
		for (let index = 0; index < this.clusters.length; ++index) {
			points.push(element(ID_CUE_POINT, concat([
				uintElement(ID_CUE_TIME, this.clusterTimes[index]),
				element(ID_CUE_TRACK_POSITIONS, concat([
					uintElement(ID_CUE_TRACK, TRACK_NUMBER),
					uintElement(ID_CUE_CLUSTER_POSITION, position)
				]))
			])));
			position += this.clusters[index].length;
		}
		return element(ID_CUES, concat(points));
	};
	CWebMMuxer.prototype.finish = function () {
		this.flushCluster();
		const info = this.getInfo();
		const tracks = this.getTracks();
		const clusters = concat(this.clusters);
		const cues = this.getCues(info.length + tracks.length);
		const segment = element(ID_SEGMENT, concat([info, tracks, clusters, cues]));
		return concat([this.getHeader(), segment]);
	};

	window['AscCommonSlide'].CWebMMuxer = CWebMMuxer;

	// ---------------------------------------------------------------------------
	// MP4: ISO base media container
	// ---------------------------------------------------------------------------
	const TIMESCALE = 1000;
	const TRACK_ID = 1;
	const LANGUAGE_UNDEFINED = 0x55C4;
	const HEADER_SIZE = 8;
	const MATRIX_UNITY = [0x00010000, 0, 0, 0, 0x00010000, 0, 0, 0, 0x40000000];

	function uint32(value) {
		return new Uint8Array([(value >>> 24) & 0xFF, (value >>> 16) & 0xFF, (value >>> 8) & 0xFF, value & 0xFF]);
	}
	function uint16(value) {
		return new Uint8Array([(value >>> 8) & 0xFF, value & 0xFF]);
	}
	function zeros(count) {
		return new Uint8Array(count);
	}
	function ascii(text) {
		const bytes = new Uint8Array(text.length);
		for (let index = 0; index < text.length; ++index) {
			bytes[index] = text.charCodeAt(index) & 0xFF;
		}
		return bytes;
	}
	function box(type, parts) {
		const payload = concat(parts);
		return concat([uint32(HEADER_SIZE + payload.length), ascii(type), payload]);
	}
	function fullBox(type, version, flags, parts) {
		const head = new Uint8Array([version, (flags >>> 16) & 0xFF, (flags >>> 8) & 0xFF, flags & 0xFF]);
		return box(type, [head].concat(parts));
	}
	function matrix() {
		const parts = [];
		for (let index = 0; index < MATRIX_UNITY.length; ++index) {
			parts.push(uint32(MATRIX_UNITY[index]));
		}
		return concat(parts);
	}

	function CMp4Muxer(width, height) {
		this.width = width;
		this.height = height;
		this.description = null;
		this.samples = [];
		this.durationMs = 0;
	}
	CMp4Muxer.prototype.setDescription = function (description) {
		if (!description || this.description) {
			return;
		}
		this.description = new Uint8Array(description.buffer ? description.buffer : description,
			description.byteOffset || 0, description.byteLength || description.length);
	};
	CMp4Muxer.prototype.hasDescription = function () {
		return !!this.description;
	};
	CMp4Muxer.prototype.addSample = function (data, durationMs, isKeyFrame) {
		this.samples.push({data: data, duration: Math.max(1, Math.round(durationMs)), isKeyFrame: !!isKeyFrame});
		this.durationMs += Math.max(1, Math.round(durationMs));
	};
	CMp4Muxer.prototype.addDurationToLastSample = function (durationMs) {
		if (0 === this.samples.length) {
			return false;
		}
		const delta = Math.round(durationMs);
		this.samples[this.samples.length - 1].duration += delta;
		this.durationMs += delta;
		return true;
	};
	CMp4Muxer.prototype.getSampleCount = function () {
		return this.samples.length;
	};
	CMp4Muxer.prototype.getDurationMs = function () {
		return this.durationMs;
	};
	CMp4Muxer.prototype.getFtyp = function () {
		return box("ftyp", [ascii("isom"), uint32(0x200), ascii("isom"), ascii("iso2"), ascii("avc1"), ascii("mp41")]);
	};
	CMp4Muxer.prototype.getMdat = function () {
		const parts = [];
		for (let index = 0; index < this.samples.length; ++index) {
			parts.push(this.samples[index].data);
		}
		return box("mdat", parts);
	};
	CMp4Muxer.prototype.getMvhd = function () {
		return fullBox("mvhd", 0, 0, [
			uint32(0),
			uint32(0),
			uint32(TIMESCALE),
			uint32(this.durationMs),
			uint32(0x00010000),
			uint16(0x0100),
			zeros(10),
			matrix(),
			zeros(24),
			uint32(TRACK_ID + 1)
		]);
	};
	CMp4Muxer.prototype.getTkhd = function () {
		return fullBox("tkhd", 0, 0x000007, [
			uint32(0),
			uint32(0),
			uint32(TRACK_ID),
			uint32(0),
			uint32(this.durationMs),
			zeros(8),
			uint16(0),
			uint16(0),
			uint16(0),
			uint16(0),
			matrix(),
			uint32(this.width << 16),
			uint32(this.height << 16)
		]);
	};
	CMp4Muxer.prototype.getMdhd = function () {
		return fullBox("mdhd", 0, 0, [
			uint32(0),
			uint32(0),
			uint32(TIMESCALE),
			uint32(this.durationMs),
			uint16(LANGUAGE_UNDEFINED),
			uint16(0)
		]);
	};
	CMp4Muxer.prototype.getHdlr = function () {
		return fullBox("hdlr", 0, 0, [
			uint32(0),
			ascii("vide"),
			zeros(12),
			ascii("VideoHandler"),
			zeros(1)
		]);
	};
	CMp4Muxer.prototype.getAvc1 = function () {
		const avcC = box("avcC", [this.description]);
		const compressorName = new Uint8Array(32);
		return box("avc1", [
			zeros(6),
			uint16(1),
			zeros(16),
			uint16(this.width),
			uint16(this.height),
			uint32(0x00480000),
			uint32(0x00480000),
			uint32(0),
			uint16(1),
			compressorName,
			uint16(0x0018),
			new Uint8Array([0xFF, 0xFF]),
			avcC
		]);
	};
	CMp4Muxer.prototype.getStts = function () {
		const entries = [];
		for (let index = 0; index < this.samples.length; ++index) {
			const duration = this.samples[index].duration;
			if (entries.length > 0 && entries[entries.length - 1].duration === duration) {
				++entries[entries.length - 1].count;
				continue;
			}
			entries.push({count: 1, duration: duration});
		}
		const parts = [uint32(entries.length)];
		for (let index = 0; index < entries.length; ++index) {
			parts.push(uint32(entries[index].count));
			parts.push(uint32(entries[index].duration));
		}
		return fullBox("stts", 0, 0, parts);
	};
	CMp4Muxer.prototype.getStss = function () {
		const keys = [];
		for (let index = 0; index < this.samples.length; ++index) {
			if (this.samples[index].isKeyFrame) {
				keys.push(index + 1);
			}
		}
		const parts = [uint32(keys.length)];
		for (let index = 0; index < keys.length; ++index) {
			parts.push(uint32(keys[index]));
		}
		return fullBox("stss", 0, 0, parts);
	};
	CMp4Muxer.prototype.getStsc = function () {
		return fullBox("stsc", 0, 0, [uint32(1), uint32(1), uint32(this.samples.length), uint32(1)]);
	};
	CMp4Muxer.prototype.getStsz = function () {
		const parts = [uint32(0), uint32(this.samples.length)];
		for (let index = 0; index < this.samples.length; ++index) {
			parts.push(uint32(this.samples[index].data.length));
		}
		return fullBox("stsz", 0, 0, parts);
	};
	CMp4Muxer.prototype.getStco = function (mdatDataOffset) {
		return fullBox("stco", 0, 0, [uint32(1), uint32(mdatDataOffset)]);
	};
	CMp4Muxer.prototype.getStbl = function (mdatDataOffset) {
		return box("stbl", [
			fullBox("stsd", 0, 0, [uint32(1), this.getAvc1()]),
			this.getStts(),
			this.getStss(),
			this.getStsc(),
			this.getStsz(),
			this.getStco(mdatDataOffset)
		]);
	};
	CMp4Muxer.prototype.getMinf = function (mdatDataOffset) {
		const vmhd = fullBox("vmhd", 0, 1, [uint16(0), zeros(6)]);
		const dref = fullBox("dref", 0, 0, [uint32(1), fullBox("url ", 0, 1, [])]);
		return box("minf", [vmhd, box("dinf", [dref]), this.getStbl(mdatDataOffset)]);
	};
	CMp4Muxer.prototype.getMoov = function (mdatDataOffset) {
		const mdia = box("mdia", [this.getMdhd(), this.getHdlr(), this.getMinf(mdatDataOffset)]);
		return box("moov", [this.getMvhd(), box("trak", [this.getTkhd(), mdia])]);
	};
	CMp4Muxer.prototype.finish = function () {
		if (!this.description || 0 === this.samples.length) {
			return null;
		}
		const ftyp = this.getFtyp();
		const mdat = this.getMdat();
		const moov = this.getMoov(ftyp.length + HEADER_SIZE);
		return concat([ftyp, mdat, moov]);
	};

	window['AscCommonSlide'].CMp4Muxer = CMp4Muxer;

})(window);
