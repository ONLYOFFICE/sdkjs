/* jshint -W040 */
/*
 * Copyright (c) 2017 Nexedi SA and Contributors. All Rights Reserved.
 * Author: Boris Kocherov
 *
 * This extension was developed by Nexedi as part of
 * OpenPaaS::NG PSPC collaborative R&D project financed by BPI France
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact jp@nexedi.com.
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
 * @param {Object} RSVP
 * @param {Xmla} Xmla
 * @param {console} console
 * @param {undefined} undefined
 */
	function (window, RSVP, Xmla, console, undefined) {
	var cBaseFunction = AscCommonExcel.cBaseFunction;
	var cFormulaFunctionGroup = AscCommonExcel.cFormulaFunctionGroup,
		cElementType = AscCommonExcel.cElementType,
		cNumber = AscCommonExcel.cNumber,
		cString = AscCommonExcel.cString,
		cBool = AscCommonExcel.cBool,
		cError = AscCommonExcel.cError,
		cErrorType = AscCommonExcel.cErrorType,
		cArea = AscCommonExcel.cArea,
		cArea3D = AscCommonExcel.cArea3D,
		cRef = AscCommonExcel.cRef,
		cRef3D = AscCommonExcel.cRef3D,
		cEmpty = AscCommonExcel.cEmpty,
		cArray = AscCommonExcel.cArray,
		cubeScheme = {},
		cubeExecutionScheme = {};

	cFormulaFunctionGroup.Cube = cFormulaFunctionGroup.Cube || [];
	cFormulaFunctionGroup.Cube.push(cCUBEKPIMEMBER, cCUBEMEMBER, cCUBEMEMBERPROPERTY, cCUBERANKEDMEMBER, cCUBESET,
		cCUBESETCOUNT, cCUBEVALUE);

	cFormulaFunctionGroup.NotRealised = cFormulaFunctionGroup.NotRealised || [];
	cFormulaFunctionGroup.NotRealised.push(cCUBEKPIMEMBER, cCUBEMEMBERPROPERTY, cCUBERANKEDMEMBER,
		cCUBESET, cCUBESETCOUNT);

	var xmla = new Xmla({
// 		listeners: {
// 			events: Xmla.EVENT_ERROR,
// 			handler: function (eventName, eventData, xmla) {
// 				console.log(eventData.exception);
// //        alert(
// //          "Snap, an error occurred: " + eventData.exception.message + " (" + eventData.exception.code + ")" +
// //          (eventData.exception.code === Xmla.Exception.HTTP_ERROR_CDE
// //            ? "\nstatus: " + eventData.exception.data.status + "; statusText: " + eventData.exception.data.statusText
// //            : "")
// //        );
// 			}
// 		},
		async: true
	});

	function xmla_request(func, prop) {
		var xmla = new Xmla({async: true});
		// return function () {
		return new RSVP.Queue()
			.push(function () {
				return new RSVP.Promise(function (resolve, reject) {
					prop.success = function (xmla, options, response) {
						resolve(response);
					};
					prop.error = function (xmla, options, response) {
						reject(response);
					};
					xmla[func](prop);
				});
			});
	}

	function xmla_request_retry(func, prop) {
		return xmla_request(func, prop)
			.push(undefined, function (response) {
				// fix mondrian Internal and Sql errors
				if (response) {
					switch (response["code"]) {
						case "SOAP-ENV:Server.00HSBE02":
						case "SOAP-ENV:00UE001.Internal Error":
							// rarely server error, so try again
							return xmla_request(func, prop);
					}
				}
				throw response;
			});
	}

	function discover_hierarchies(connection) {
		var settings = getProperties(connection),
			prop = settings.prop;
		prop.restrictions = {
//      'CATALOG_NAME': 'FoodMart',
// 			'HIERARCHY_NAME': hierarchy_name,
// 			'HIERARCHY_UNIQUE_NAME': hierarchy_name,
			'CUBE_NAME': settings["cube"]
		};
		return xmla_request_retry("discoverMDHierarchies", prop)
			.push(function (response) {
				var hierarchies = {},
					hierarchy,
					uname,
					caption,
					all_member,
					dimension_uname,
					dimension,
					dimensions = {};
				while (response.hasMoreRows()) {
					uname = response["getHierarchyUniqueName"]();
					caption = response["getHierarchyCaption"]();
					all_member = response["getAllMember"]();
					dimension_uname = response["getDimensionUniqueName"]();
					dimension = dimensions[dimension_uname];
					if (!dimension) {
						dimension = {
							"uname": dimension_uname,
							"all_member": all_member
						};
						dimensions[dimension_uname] = dimension;
					}
					if (!dimension.all_member && all_member) {
						dimension.all_member = all_member;
					}
					hierarchy = {
						"uname": uname,
						"caption": caption,
						"all_member": all_member,
						"dimension_uname": dimension_uname,
						"dimension": dimension
					};
					hierarchies[uname] = hierarchy;
					hierarchies[caption] = hierarchy;
					response.nextRow();
				}
				return {
					"hierarchies": hierarchies,
					"dimensions": dimensions
				};
			});
	}

	function getProperties(connection) {
		var connections = {
			"xmla": {
				"prop": {
					"url": "https://d1.erp5.ru/saiku/xmla",
					"properties": {
						"DataSourceInfo": "FoodMart",
						"Catalog": "FoodMart"
					}
				},
				"cube": "Sales"
			},
			"olapy": {
				"prop": {
					"url": "https://d1.erp5.ru/olapy/xmla",
					"properties": {
						"DataSourceInfo": "-",
						"Catalog": "sales"
					}
				},
				"cube": "Sales"
			}
		};
		connection = connections[connection];
		if (!connection) {
			throw "connection not exist";
		}
		connection = JSON.parse(JSON.stringify(connection));
		return connection;
	}

	function getScheme(connection) {
		var scheme = cubeScheme[connection],
			queue = new RSVP.Queue();
		if (scheme) {
			return queue.push(function () {
				return scheme;
			});
		}
		cubeScheme[connection] = queue;
		return queue
			.push(function () {
				return discover_hierarchies(connection);
			})
			.push(function (arg) {
				scheme = {
					members: {},
					hierarchies: arg.hierarchies,
					dimensions: arg.dimensions
				};
				cubeScheme[connection] = scheme;
				return scheme;
			});
	}

	function getExecutionScheme(connection) {
		var scheme = cubeExecutionScheme[connection];
		if (scheme) {
			return scheme;
		} else {
			scheme = {
				members: {},
				hierarchies: {},
				levels: {}
			};
			cubeExecutionScheme[connection] = scheme;
			return scheme;
		}
	}

	function getCell(arg0) {
		if (arg0 instanceof cArray) {
			arg0 = arg0.getElement(0);
			// } else if (arg0 instanceof cArea || arg0 instanceof cArea3D) {
			// 	arg0 = arg0.cross(arguments[1].bbox);
		} else if (arg0 instanceof cRef || arg0 instanceof cRef3D) {
			arg0 = arg0.getValue();
		}
		return arg0;
	}

	function parseArgs(mdx_array) {
		return function () {
			var members = [];

			function stringForge(value) {
				var array;
				if (value.cube_value) {
					array = value.cube_value;
				} else {
					array = value.value.split(',');
				}
				if (array.length > 0) {
					// filter members already existed
					members = members.filter(function (i) {
						return array.indexOf(i) === -1;
					});
					members = members.concat(array);
				}
			}

			function cellForge(cell) {
				if (cell) {
					if (cell.type === cElementType.error) {
						// debugger;
						throw "referenced cell contain error";
					}
					if (cell.formulaParsed && cell.formulaParsed.value) {
						stringForge(cell.formulaParsed.value);
					} else {
						stringForge({value: cell.getValue()});
					}
				}
			}

			mdx_array.forEach(function (element) {
				if (element instanceof cArea || element instanceof cArea3D ||
					element instanceof cRef || element instanceof cRef3D) {
					element.getRange()._foreach(cellForge);
				} else {
					stringForge(element);
				}
			});
			return members;
		};
	}

	var AddCubeValueCalculate = (function () {
		var deferred = RSVP.defer(),
			cells = [];
		return function (cell_id) {
			if (cells.indexOf(cell_id) === -1) {
				cells.push(cell_id);
			}
			// console.log('+ ' + cells);
			return function () {
				var i = cells.indexOf(cell_id);
				if (i !== -1) {
					cells.splice(i, 1);
				}
				// console.log('-' + cells);
				if (cells.length === 0) {
					deferred.resolve();
					deferred = RSVP.defer();
					return {};
				}
				return deferred.promise;
			};
		};
	})();

	function execute(connection) {
		var execution_scheme = getExecutionScheme(connection),
			scheme;
		if (!execution_scheme.execute) {
			execution_scheme.execute = RSVP.defer();
			return getScheme(connection)
				.push(function (s) {
					var settings = getProperties(connection),
						prop = settings.prop,
						hierarchies = execution_scheme.hierarchies,
						hierarchy,
						mdx = [],
						tuple_str,
						all_member;
					scheme = s;
					for (hierarchy in hierarchies) {
						tuple_str = hierarchies[hierarchy].join(",");
						all_member = scheme.hierarchies[hierarchy]["all_member"];
						if (all_member) {
							tuple_str = tuple_str + ',' + all_member;
						}
						mdx.push("{" + tuple_str + "}");
					}
					prop.statement = "SELECT " + mdx.join("*") +
						" ON 0 FROM [" + settings["cube"] + "]";
					return xmla_request("execute", prop);
				})
				.push(function (dataset) {
					var cellset = dataset.getCellset(),
						axis_count = dataset.axisCount(),
						axis_array = [],
						axis_id,
						cube = {
							"axes": {"length": axis_count},
							"members": {},
							"hierarchies": {"length": 0},
							"hierarchies_info": scheme.hierarchies,
							"cells": []
						};


					for (axis_id = 0; axis_id < axis_count; axis_id++) {
						axis_array.push(dataset.getAxis(axis_id));
					}

					axis_array.forEach(function (axis, axis_id) {
						cube.axes[axis_id] = {
							tuples: {},
							length: 0
						};
						axis.eachTuple(function (tuple) {
							var coordinate_tuple = [];
							axis.eachHierarchy(function () {
								var member = this.member();
								if (!cube.members.hasOwnProperty(member["UName"])) {
									cube.members[member["UName"]] = member;
								}
								coordinate_tuple.push(member["UName"]);
							});
							cube.axes[axis_id].tuples[coordinate_tuple.join(',')] = tuple.index;
							cube.axes[axis_id].length++;
						});
						axis.eachHierarchy(function (hierarchy) {
							cube.hierarchies[hierarchy.name] = {
								axis_id: axis_id, tuple_id: hierarchy.index, name: hierarchy.name
							};
							cube.hierarchies[cube.hierarchies.length] = cube.hierarchies[hierarchy.name];
							cube.hierarchies['' + axis_id + ',' + hierarchy.index] = cube.hierarchies[hierarchy.name];
							cube.hierarchies.length++;
						});
					});

					do {
						cube.cells[cellset.cellOrdinal()] = cellset["cellValue"]();
					} while (cellset.nextCell() > 0);
					execution_scheme.cube = cube;
					execution_scheme.execute.resolve(cube);
					execution_scheme.execute = null;
					execution_scheme.hierarchies = [];
					return cube;
				})
				.push(undefined, function (error) {
					console.error(error);
					execution_scheme.execute = null;
					execution_scheme.hierarchies = [];
				});
		}
		return execution_scheme.execute.promise;
	}

	function discover_members(connection, opt) {
		return new RSVP.Queue()
			.push(function () {
				var settings = getProperties(connection),
					prop = settings.prop,
					cached_member,
					scheme = getExecutionScheme(connection);
				prop.restrictions = {
//      'CATALOG_NAME': 'FoodMart',
					'CUBE_NAME': settings["cube"]
				};
				if (!opt) {
					opt = {};
				}
				if (opt.member_uname) {
					prop.restrictions["MEMBER_UNIQUE_NAME"] = opt.member_uname;
					cached_member = scheme.members[opt.member_uname];
				}
				if (opt.level_uname) {
					prop.restrictions["LEVEL_UNIQUE_NAME"] = opt.level_uname;
				}
				if (cached_member) {
					return [cached_member];
				} else {
					return xmla_request_retry("discoverMDMembers", prop)
						.push(function (r) {
							var ret = [],
								uname,
								level,
								cached_member;
							while (r.hasMoreRows()) {
								uname = r["getMemberUniqueName"]();
								level = r["getLevelUniqueName"]();
								// we can check cache twice because fist check
								// only if discover by member_uname
								if (!scheme.members.hasOwnProperty(uname)) {
									cached_member = {
										uname: uname,
										h: r["getHierarchyUniqueName"](),
										level: r["getLevelUniqueName"](),
										caption: r["getMemberCaption"](),
										type: r["getMemberType"]()
									};
									scheme.members[uname] = cached_member;
								} else {
									cached_member = scheme.members[uname];
								}
								ret.push(cached_member);
								r.nextRow();
								if (!scheme.levels.hasOwnProperty(level)) {
									scheme.levels[level] = discover_level(connection, scheme, level);
								}
							}
							return ret;
						});
				}
			});
	}

	function discover_level(connection, scheme, level) {
		return discover_members(connection, {
			level_uname: level
		})
			.push(function (members) {
				var i;

				function compare(a, b) {
					if (a.uname < b.uname)
						return -1;
					if (a.uname > b.uname)
						return 1;
					return 0;
				}

				members.sort(compare);
				for (i = 0; i < members.length; i++) {
					members[i].level_index = i;
				}
				scheme.levels[level] = members;
			});
	}

	function discover_members_for_arguments(connection, members) {
		var promises = [],
			hierarchies = {};

		function check_interseption(hierarchy) {
			if (hierarchies.hasOwnProperty(hierarchy)) {
				throw  "The tuple is invalid because there is no intersection for the specified values.";
			} else {
				hierarchies[hierarchy] = 1;
			}
		}

		members.forEach(function (member) {
			if (member) {
				promises
					.push(
						discover_members(connection, {
							member_uname: member
						})
							.push(function (members) {
								var member;
								if (members.length > 0) {
									member = members[0];
									check_interseption(member.h);
									return member;
								} else {
									throw "member not found";
								}
							})
					);
			}
		});
		return RSVP.all(promises);
	}

	function error_handler(current_cell_id) {
		return function (error) {
			console.error(current_cell_id, error);
			var ret;
			if (error === "referenced cell contain error") {
				ret = new cError(cErrorType.wrong_value_type);
			} else if (error === "connection not exist" ||
				error instanceof Xmla.Exception) {
				ret = new cError(cErrorType.wrong_name);
			} else {
				ret = new cError(cErrorType.not_available);
			}
			return ret;
		};
	}

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEKPIMEMBER() {
	}

	cCUBEKPIMEMBER.prototype = Object.create(cBaseFunction.prototype);
	cCUBEKPIMEMBER.prototype.constructor = cCUBEKPIMEMBER;
	cCUBEKPIMEMBER.prototype.name = 'CUBEKPIMEMBER';

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEMEMBER() {
	}

	cCUBEMEMBER.prototype = Object.create(cBaseFunction.prototype);
	cCUBEMEMBER.prototype.constructor = cCUBEMEMBER;
	cCUBEMEMBER.prototype.name = 'CUBEMEMBER';
	cCUBEMEMBER.prototype.argumentsMin = 2;
	cCUBEMEMBER.prototype.argumentsMax = 3;
	cCUBEMEMBER.prototype.ca = true;
	cCUBEMEMBER.prototype.CalculateLazy = function (queue, bbox, isDefName, ws) {
		var connection,
			current_cell_id = [ws.getId(),bbox.r1,bbox.c2].join(";"),
			caption;
		return queue
			.push(function (arg) {
				connection = getCell(arg[0]);
				caption = getCell(arg[2]);
				if (caption) {
					caption = caption.getValue();
				}
				return parseArgs([arg[1]])();
			})
			.push(function (members) {
				return discover_members_for_arguments(connection, members);
			})
			.push(function (members) {
				var last_id = members.length - 1,
					ret;
				if (!caption) {
					caption = members[last_id].caption;
				}
				ret = new cString(caption);
				ret.cube_value = [];
				members.forEach(function (member) {
					ret.cube_value.push(member.uname);
				});
				return ret;
			})
			.push(undefined, error_handler(current_cell_id));
	};
	cCUBEMEMBER.prototype.changeOffsetElem = function (arg, offset) {
		var connection = getCell(arg[0]),
			scheme = getExecutionScheme(connection),
			i,
			elem,
			member,
			new_member,
			level;
		for (i = 0; i < arg.length; i++) {
			elem = arg[i];
			if (cElementType.string === elem.type) {
				member = scheme.members[elem.value];
				if (member && (member.level_index >= 0)) {
					level = scheme.levels[member.level];
					new_member = level[member.level_index + offset.offsetCol + offset.offsetRow];
					if (new_member) {
						elem.value = new_member.uname;
					} else {
						elem.value = "";
					}
				}
			}
		}
	};
	cCUBEMEMBER.prototype.getInfo = function () {
		return {
			name: this.name, args: "( connection, members, caption )"
		};
	};

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEMEMBERPROPERTY() {
	}

	cCUBEMEMBERPROPERTY.prototype = Object.create(cBaseFunction.prototype);
	cCUBEMEMBERPROPERTY.prototype.constructor = cCUBEMEMBERPROPERTY;
	cCUBEMEMBERPROPERTY.prototype.name = 'CUBEMEMBERPROPERTY';

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBERANKEDMEMBER() {
	}

	cCUBERANKEDMEMBER.prototype = Object.create(cBaseFunction.prototype);
	cCUBERANKEDMEMBER.prototype.constructor = cCUBERANKEDMEMBER;
	cCUBERANKEDMEMBER.prototype.name = 'CUBERANKEDMEMBER';

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBESET() {
	}

	cCUBESET.prototype = Object.create(cBaseFunction.prototype);
	cCUBESET.prototype.constructor = cCUBESET;
	cCUBESET.prototype.name = 'CUBESET';

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBESETCOUNT() {
	}

	cCUBESETCOUNT.prototype = Object.create(cBaseFunction.prototype);
	cCUBESETCOUNT.prototype.constructor = cCUBESETCOUNT;
	cCUBESETCOUNT.prototype.name = 'CUBESETCOUNT';

	/**
	 * @constructor
	 * @extends {AscCommonExcel.cBaseFunction}
	 */
	function cCUBEVALUE() {
	}

	cCUBEVALUE.prototype = Object.create(cBaseFunction.prototype);
	cCUBEVALUE.prototype.constructor = cCUBEVALUE;
	cCUBEVALUE.prototype.name = 'CUBEVALUE';
	cCUBEVALUE.prototype.argumentsMin = 2;
	cCUBEVALUE.prototype.argumentsMax = 5;
	cCUBEVALUE.prototype.ca = true;
	cCUBEVALUE.prototype.CalculateLazy = function (queue, bbox, isDefName, ws) {
		var scheme,
			connection,
			members = [],
			current_cell_id = [ws.getId(),bbox.r1,bbox.c2].join(";"),
			waiter = AddCubeValueCalculate(current_cell_id);
		return queue
			.push(function (arg) {
				connection = getCell(arg[0]);
				scheme = getExecutionScheme(connection);
				return parseArgs(arg.slice(1))();
			})
			.push(function (members) {
				return discover_members_for_arguments(connection, members);
			})
			.push(function (m) {
				var member_uname,
					member,
					h,
					hierarchy;
				for (member_uname in m) {
					if (m.hasOwnProperty(member_uname)) {
						member = m[member_uname];
						hierarchy = member.h;
						h = scheme.hierarchies[hierarchy];
						if (!h) {
							h = [];
							scheme.hierarchies[hierarchy] = h;
						}
						if (h.indexOf(member.uname) === -1) {
							h.push(member.uname);
						}
						members.push(member.uname);
					}
				}
				return waiter();
			})
			.push(function () {
				return execute(connection);
			})
			.push(function (cube) {
				var cell_id = 0,
					p_d = 1,
					h,
					member_path,
					coordinate = [],
					i,
					ret;

				function getHierarchyByMember(member_path) {
					var h;
					h = cube.members[member_path];
					if (h === undefined) {
						throw "query result not contain data for member:" +
						member_path;
					}
					h = h.hierarchy;
					h = cube.hierarchies[h];
					return h;
				}

				for (i = 0; i < cube.hierarchies.length; i++) {
					h = cube.hierarchies[i];
					if (!coordinate[h.axis_id]) {
						coordinate[h.axis_id] = [];
					}
					coordinate[h.axis_id][h.tuple_id] = null;
				}
				for (i = 0; i < members.length; i++) {
					member_path = members[i];
					h = getHierarchyByMember(members[i]);
					coordinate[h.axis_id][h.tuple_id] = member_path;
				}
				coordinate = coordinate.map(function (axis, axis_id) {
					return axis.map(function (h, h_id) {
						var hierarchy_name,
							all_member;
						if (!h) {
							hierarchy_name = cube.hierarchies[axis_id + ',' + h_id].name;
							all_member = cube["hierarchies_info"][hierarchy_name]["all_member"];
							if (all_member) {
								h = getHierarchyByMember(all_member);
								if (h) {
									return all_member;
								}
							}
							throw "Axis:" + axis_id + " hierarchy:" +
							cube.hierarchies[axis_id + ',' + h_id].name +
							" not determinated";
						}
						return h;
					}).join(',');
				});
				coordinate.forEach(function (tuple, axis_id) {
					var axis = cube.axes[axis_id];
					cell_id = p_d * axis.tuples[tuple] + cell_id;
					p_d = p_d * axis.length;
				});
				ret = new cNumber(cube.cells[cell_id]);
				return ret;
			})
			.push(undefined, function (error) {
				// issue in one cell(cubevalue) not stop calculation in other
				return new RSVP.Queue()
					.push(function () {
						return waiter();
					})
					.push(function () {
						return error_handler(current_cell_id)(error);
					});
			});
	};
	cCUBEVALUE.prototype.changeOffsetElem = cCUBEMEMBER.prototype.changeOffsetElem;
	cCUBEVALUE.prototype.getInfo = function () {
		return {
			name: this.name, args: "( connection, member1, member2, .. )"
		};
	};
})(window, RSVP, Xmla, console);
