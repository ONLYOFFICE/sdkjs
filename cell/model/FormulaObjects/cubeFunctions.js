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

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {
	var cBaseFunction = AscCommonExcel.cBaseFunction;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup;
	var argType = Asc.c_oAscFormulaArgumentType;

	cFormulaFunctionGroup['Cube'] = cFormulaFunctionGroup['Cube'] || [];
	cFormulaFunctionGroup['Cube'].push(cCUBEKPIMEMBER, cCUBEMEMBER, cCUBEMEMBERPROPERTY, cCUBERANKEDMEMBER, cCUBESET,
		cCUBESETCOUNT, cCUBEVALUE);

	cFormulaFunctionGroup['NotRealised'] = cFormulaFunctionGroup['NotRealised'] || [];
	cFormulaFunctionGroup['NotRealised'].push(cCUBEKPIMEMBER, cCUBEMEMBER, cCUBEMEMBERPROPERTY, cCUBERANKEDMEMBER,
		cCUBESET, cCUBESETCOUNT, cCUBEVALUE);

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEKPIMEMBER() {
	}

	cCUBEKPIMEMBER.prototype = Object.create(cBaseFunction.prototype);
	cCUBEKPIMEMBER.prototype.constructor = cCUBEKPIMEMBER;
	cCUBEKPIMEMBER.prototype.name = 'CUBEKPIMEMBER';
	cCUBEKPIMEMBER.prototype.argumentsType = [argType.text, argType.text, argType.number, argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEMEMBER() {
	}

	cCUBEMEMBER.prototype = Object.create(cBaseFunction.prototype);
	cCUBEMEMBER.prototype.constructor = cCUBEMEMBER;
	cCUBEMEMBER.prototype.name = 'CUBEMEMBER';
	cCUBEMEMBER.prototype.argumentsType = [argType.text, argType.text, argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEMEMBERPROPERTY() {
	}

	cCUBEMEMBERPROPERTY.prototype = Object.create(cBaseFunction.prototype);
	cCUBEMEMBERPROPERTY.prototype.constructor = cCUBEMEMBERPROPERTY;
	cCUBEMEMBERPROPERTY.prototype.name = 'CUBEMEMBERPROPERTY';
	cCUBEMEMBERPROPERTY.prototype.argumentsType = [argType.text, argType.text, argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBERANKEDMEMBER() {
	}

	cCUBERANKEDMEMBER.prototype = Object.create(cBaseFunction.prototype);
	cCUBERANKEDMEMBER.prototype.constructor = cCUBERANKEDMEMBER;
	cCUBERANKEDMEMBER.prototype.name = 'CUBERANKEDMEMBER';
	cCUBERANKEDMEMBER.prototype.argumentsType = [argType.text, argType.text, argType.number, argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBESET() {
	}

	cCUBESET.prototype = Object.create(cBaseFunction.prototype);
	cCUBESET.prototype.constructor = cCUBESET;
	cCUBESET.prototype.name = 'CUBESET';
	cCUBESET.prototype.argumentsType = [argType.text, argType.text, argType.text, argType.number, argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBESETCOUNT() {
	}

	cCUBESETCOUNT.prototype = Object.create(cBaseFunction.prototype);
	cCUBESETCOUNT.prototype.constructor = cCUBESETCOUNT;
	cCUBESETCOUNT.prototype.name = 'CUBESETCOUNT';
	cCUBESETCOUNT.prototype.argumentsType = [argType.text];

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEVALUE() {
	}

	cCUBEVALUE.prototype = Object.create(cBaseFunction.prototype);
	cCUBEVALUE.prototype.constructor = cCUBEVALUE;
	cCUBEVALUE.prototype.name = 'CUBEVALUE';
	cCUBEVALUE.prototype.argumentsType = [argType.text, [argType.text]];
})(window);
