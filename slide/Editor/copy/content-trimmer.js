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

(function () {
	function ContentTrimmer(docContent) {
		this.docContent = docContent;
	}
	ContentTrimmer.prototype.trim = function () {
		new StartContentTrimmer(this.docContent).trim();
		new EndContentTrimmer(this.docContent).trim();
	};
	function StartContentTrimmer(docContent) {
		this.docContent = docContent;
	}
	StartContentTrimmer.prototype.trim = function () {
		const content = this.docContent.Content;
		for (let i = 0; i < content.length; ++i) {
			const checkParagraph = content[i];
			if (!checkParagraph.IsEmpty()) {
				const runs = checkParagraph.Content;
				this.trimRuns(runs);
			}
			if (!checkParagraph.IsEmpty()) {
				return;
			}
			this.docContent.Internal_Content_Remove(i, 1, false);
			i--;
		}
	};
	StartContentTrimmer.prototype.trimRuns = function (runs) {
		if (runs.length > 1) {
			for (let i = 0; i < runs.length - 1; ++i) {
				const run = runs[i];
				if (run.Type === para_Run) {
					this.trimRun(run);
					if (run.Content.length !== 0) {
						return;
					}
					runs.splice(i, 1);
					i--;
				}
			}
		}
	};
	StartContentTrimmer.prototype.trimRun = function (run) {
		for (let i = 0; i < run.Content.length; ++i) {
			const runElement = run.Content[i];
			if (runElement.Type !== para_NewLine) {
				return;
			}
			run.Content.splice(i, 1);
			i--;
		}
	};

	function EndContentTrimmer(docContent) {
		this.docContent = docContent;
	}
	EndContentTrimmer.prototype.trim = function () {
		const content = this.docContent.Content;
		for (let i = content.length - 1; i > -1; --i) {
			const checkParagraph = content[i];
			if (!checkParagraph.IsEmpty()) {
				const runs = checkParagraph.Content;
				this.trimRuns(runs);
			}
			if (!checkParagraph.IsEmpty()) {
				return;
			}
			this.docContent.Internal_Content_Remove(i, 1, false);
		}
	};
	EndContentTrimmer.prototype.trimRuns = function (runs) {
		if (runs.length > 1) {
			for (let i = runs.length - 2; i > -1; --i) {
				const run = runs[i];
				if (run.Type === para_Run) {
					this.trimRun(run);
					if (run.Content.length !== 0) {
						return;
					}
					runs.splice(i, 1);
				}
			}
		}
	};
	EndContentTrimmer.prototype.trimRun = function (run) {
		for (let i = run.Content.length - 1; i > -1; --i) {
			const runElement = run.Content[i];
			if (runElement.Type !== para_NewLine) {
				return;
			}
			run.Content.splice(i, 1);
		}
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].ContentTrimmer = ContentTrimmer;
})();
