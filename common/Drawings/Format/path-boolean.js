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

(function (window) {

	function InitClassWithStatics(fClass, fBase) {
		fClass.prototype = Object.create(fBase.prototype);
		fClass.prototype.superclass = fBase;
		fClass.prototype.constructor = fClass;
		Object.getOwnPropertyNames(fBase).forEach(function (prop) {
			if (['prototype', 'name', 'length'].includes(prop) || Function.prototype.hasOwnProperty(prop)) { return; }
			Object.defineProperty(fClass, prop, Object.getOwnPropertyDescriptor(fBase, prop));
		});
		fClass.prototype.initialize = fClass;
	}

	const OPERATIONS = {
		divide: 0,
		unite: 1,
		intersect: 2,
		subtract: 3,
		exclude: 4,
	};

	const CURVE_TYPES = {
		serpentine: 0,
		loop: 1,
		arc: 2,
		line: 3,
		quadratic: 4,
		cusp: 5,
	};

	const CollisionDetection = {
		findItemBoundsCollisions: function (items1, items2, tolerance) {
			function getBounds(items) {
				const bounds = new Array(items.length);
				for (let i = 0; i < items.length; i++) {
					const rect = items[i].getBounds();
					bounds[i] = [rect.getLeft(), rect.getTop(), rect.getRight(), rect.getBottom()];
				}
				return bounds;
			}

			const bounds1 = getBounds(items1);
			const bounds2 = (!items2 || items2 === items1) ? bounds1 : getBounds(items2);
			return this.findBoundsCollisions(bounds1, bounds2, tolerance);
		},

		findCurveBoundsCollisions: function (curves1, curves2, tolerance, bothAxis) {
			function getBounds(curves) {
				const bounds = new Array(curves.length);
				for (let i = 0; i < curves.length; i++) {
					const v = curves[i];
					bounds[i] = [
						Math.min(v[0], v[2], v[4], v[6]),
						Math.min(v[1], v[3], v[5], v[7]),
						Math.max(v[0], v[2], v[4], v[6]),
						Math.max(v[1], v[3], v[5], v[7])
					];
				}
				return bounds;
			}

			const bounds1 = getBounds(curves1);
			const bounds2 = (!curves2 || curves2 === curves1) ? bounds1 : getBounds(curves2);
			if (bothAxis) {
				const hor = this.findBoundsCollisions(bounds1, bounds2, tolerance || 0, false, true);
				const ver = this.findBoundsCollisions(bounds1, bounds2, tolerance || 0, true, true);
				const list = [];
				for (let i = 0, l = hor.length; i < l; i++) {
					list[i] = { hor: hor[i], ver: ver[i] };
				}
				return list;
			}
			return this.findBoundsCollisions(bounds1, bounds2, tolerance || 0);
		},

		findBoundsCollisions: function (boundsA, boundsB, tolerance, sweepVertical, onlySweepAxisCollisions) {
			const self = !boundsB || boundsA === boundsB;
			const allBounds = self ? boundsA : boundsA.concat(boundsB);

			function binarySearch(indices, coord, value) {
				let lo = 0;
				let hi = indices.length;
				while (lo < hi) {
					const mid = (hi + lo) >>> 1;
					allBounds[indices[mid]][coord] < value ? lo = mid + 1 : hi = mid;
				}
				return lo - 1;
			}

			const pri0 = sweepVertical ? 1 : 0;
			const pri1 = pri0 + 2;
			const sec0 = sweepVertical ? 0 : 1;
			const sec1 = sec0 + 2;

			const allIndicesByPri0 = new Array(allBounds.length);
			for (let i = 0; i < allBounds.length; i++) {
				allIndicesByPri0[i] = i;
			}
			allIndicesByPri0.sort(function (i1, i2) {
				return allBounds[i1][pri0] - allBounds[i2][pri0];
			});

			const activeIndicesByPri1 = [];
			const allCollisions = new Array(boundsA.length);
			for (let i = 0; i < allBounds.length; i++) {
				const curIndex = allIndicesByPri0[i];
				const curBounds = allBounds[curIndex];
				const origIndex = self ? curIndex : curIndex - boundsA.length;
				const isCurrentA = curIndex < boundsA.length;
				const isCurrentB = self || !isCurrentA;
				let curCollisions = isCurrentA ? [] : null;

				if (activeIndicesByPri1.length) {
					const pruneCount = binarySearch(activeIndicesByPri1, pri1, curBounds[pri0] - tolerance) + 1;
					activeIndicesByPri1.splice(0, pruneCount);
					if (self && onlySweepAxisCollisions) {
						curCollisions = curCollisions.concat(activeIndicesByPri1);
						for (let j = 0; j < activeIndicesByPri1.length; j++) {
							const activeIndex = activeIndicesByPri1[j];
							allCollisions[activeIndex].push(origIndex);
						}
					} else {
						for (let j = 0; j < activeIndicesByPri1.length; j++) {
							const activeIndex = activeIndicesByPri1[j];
							const activeBounds = allBounds[activeIndex];
							const isActiveA = activeIndex < boundsA.length;
							const isActiveB = self || activeIndex >= boundsA.length;
							const isMatchingPairA = isCurrentA && isActiveB;
							const isMatchingPairB = isCurrentB && isActiveA;
							const hasBoundaryOverlap = (
								curBounds[sec1] >= activeBounds[sec0] - tolerance &&
								curBounds[sec0] <= activeBounds[sec1] + tolerance
							);
							const shouldCheckCollision = onlySweepAxisCollisions || (isMatchingPairA || isMatchingPairB) && hasBoundaryOverlap;
							if (shouldCheckCollision) {
								if (isMatchingPairA) {
									curCollisions.push(self ? activeIndex : activeIndex - boundsA.length);
								}
								if (isMatchingPairB) {
									allCollisions[activeIndex].push(origIndex);
								}
							}
						}
					}
				}
				if (isCurrentA) {
					if (boundsA === boundsB) {
						curCollisions.push(curIndex);
					}
					allCollisions[curIndex] = curCollisions;
				}
				if (activeIndicesByPri1.length) {
					const curPri1 = curBounds[pri1];
					const index = binarySearch(activeIndicesByPri1, pri1, curPri1);
					activeIndicesByPri1.splice(index + 1, 0, curIndex);
				} else {
					activeIndicesByPri1.push(curIndex);
				}
			}
			for (let i = 0; i < allCollisions.length; i++) {
				const collisions = allCollisions[i];
				if (collisions) {
					collisions.sort(function (i1, i2) { return i1 - i2; });
				}
			}
			return allCollisions;
		},
	};

	const Numerical = new function () {

		const abscissas = [
			[0.5773502691896257645091488],
			[0, 0.7745966692414833770358531],
			[0.3399810435848562648026658, 0.8611363115940525752239465],
			[0, 0.5384693101056830910363144, 0.9061798459386639927976269],
			[0.2386191860831969086305017, 0.6612093864662645136613996, 0.9324695142031520278123016],
			[0, 0.4058451513773971669066064, 0.7415311855993944398638648, 0.9491079123427585245261897],
			[0.1834346424956498049394761, 0.5255324099163289858177390, 0.7966664774136267395915539, 0.9602898564975362316835609],
			[0, 0.3242534234038089290385380, 0.6133714327005903973087020, 0.8360311073266357942994298, 0.9681602395076260898355762],
			[0.1488743389816312108848260, 0.4333953941292471907992659, 0.6794095682990244062343274, 0.8650633666889845107320967, 0.9739065285171717200779640],
			[0, 0.2695431559523449723315320, 0.5190961292068118159257257, 0.7301520055740493240934163, 0.8870625997680952990751578, 0.9782286581460569928039380],
			[0.1252334085114689154724414, 0.3678314989981801937526915, 0.5873179542866174472967024, 0.7699026741943046870368938, 0.9041172563704748566784659, 0.9815606342467192506905491],
			[0, 0.2304583159551347940655281, 0.4484927510364468528779129, 0.6423493394403402206439846, 0.8015780907333099127942065, 0.9175983992229779652065478, 0.9841830547185881494728294],
			[0.1080549487073436620662447, 0.3191123689278897604356718, 0.5152486363581540919652907, 0.6872929048116854701480198, 0.8272013150697649931897947, 0.9284348836635735173363911, 0.9862838086968123388415973],
			[0, 0.2011940939974345223006283, 0.3941513470775633698972074, 0.5709721726085388475372267, 0.7244177313601700474161861, 0.8482065834104272162006483, 0.9372733924007059043077589, 0.9879925180204854284895657],
			[0.0950125098376374401853193, 0.2816035507792589132304605, 0.4580167776572273863424194, 0.6178762444026437484466718, 0.7554044083550030338951012, 0.8656312023878317438804679, 0.9445750230732325760779884, 0.9894009349916499325961542]
		];

		const weights = [
			[1],
			[0.8888888888888888888888889, 0.5555555555555555555555556],
			[0.6521451548625461426269361, 0.3478548451374538573730639],
			[0.5688888888888888888888889, 0.4786286704993664680412915, 0.2369268850561890875142640],
			[0.4679139345726910473898703, 0.3607615730481386075698335, 0.1713244923791703450402961],
			[0.4179591836734693877551020, 0.3818300505051189449503698, 0.2797053914892766679014678, 0.1294849661688696932706114],
			[0.3626837833783619829651504, 0.3137066458778872873379622, 0.2223810344533744705443560, 0.1012285362903762591525314],
			[0.3302393550012597631645251, 0.3123470770400028400686304, 0.2606106964029354623187429, 0.1806481606948574040584720, 0.0812743883615744119718922],
			[0.2955242247147528701738930, 0.2692667193099963550912269, 0.2190863625159820439955349, 0.1494513491505805931457763, 0.0666713443086881375935688],
			[0.2729250867779006307144835, 0.2628045445102466621806889, 0.2331937645919904799185237, 0.1862902109277342514260976, 0.1255803694649046246346943, 0.0556685671161736664827537],
			[0.2491470458134027850005624, 0.2334925365383548087608499, 0.2031674267230659217490645, 0.1600783285433462263346525, 0.1069393259953184309602547, 0.0471753363865118271946160],
			[0.2325515532308739101945895, 0.2262831802628972384120902, 0.2078160475368885023125232, 0.1781459807619457382800467, 0.1388735102197872384636018, 0.0921214998377284479144218, 0.0404840047653158795200216],
			[0.2152638534631577901958764, 0.2051984637212956039659241, 0.1855383974779378137417166, 0.1572031671581935345696019, 0.1215185706879031846894148, 0.0801580871597602098056333, 0.0351194603317518630318329],
			[0.2025782419255612728806202, 0.1984314853271115764561183, 0.1861610000155622110268006, 0.1662692058169939335532009, 0.1395706779261543144478048, 0.1071592204671719350118695, 0.0703660474881081247092674, 0.0307532419961172683546284],
			[0.1894506104550684962853967, 0.1826034150449235888667637, 0.1691565193950025381893121, 0.1495959888165767320815017, 0.1246289712555338720524763, 0.0951585116824927848099251, 0.0622535239386478928628438, 0.0271524594117540948517806]
		];

		const EPSILON = 1e-12;
		const MACHINE_EPSILON = 1.12e-16;

		const log2 = function (x) { return Math.log(x) * Math.LOG2E; }
		const clamp = function (value, min, max) { return value < min ? min : value > max ? max : value; }

		function getDiscriminant(a, b, c) {
			function split(v) {
				const x = v * 134217729;
				const y = v - x;
				const hi = y + x;
				const lo = v - hi;
				return [hi, lo];
			}

			let D = b * b - a * c;
			const E = b * b + a * c;
			if (Math.abs(D) * 3 < E) {
				const ad = split(a);
				const bd = split(b);
				const cd = split(c);
				const p = b * b;
				const dp = (bd[0] * bd[0] - p + 2 * bd[0] * bd[1]) + bd[1] * bd[1];
				const q = a * c;
				const dq = (ad[0] * cd[0] - q + ad[0] * cd[1] + ad[1] * cd[0]) + ad[1] * cd[1];
				D = (p - q) + (dp - dq);
			}
			return D;
		}

		function getNormalizationFactor() {
			const norm = Math.max.apply(Math, arguments);
			return norm && (norm < 1e-8 || norm > 1e8) ? Math.pow(2, -Math.round(log2(norm))) : 0;
		}

		return {
			EPSILON: EPSILON,
			MACHINE_EPSILON: MACHINE_EPSILON,
			CURVETIME_EPSILON: 1e-8,
			GEOMETRIC_EPSILON: 1e-7,

			isZero: function (val) { return val >= -EPSILON && val <= EPSILON; },

			isMachineZero: function (val) { return val >= -MACHINE_EPSILON && val <= MACHINE_EPSILON; },

			clamp: clamp,

			integrate: function (f, a, b, n) {
				const x = abscissas[n - 2];
				const w = weights[n - 2];
				const A = (b - a) * 0.5;
				const B = A + a;
				const m = (n + 1) >> 1;
				let i = 0;
				let sum = n & 1 ? w[i++] * f(B) : 0;
				while (i < m) {
					const Ax = A * x[i];
					sum += w[i++] * (f(B + Ax) + f(B - Ax));
				}
				return A * sum;
			},

			findRoot: function (f, df, x, a, b, n, tolerance) {
				for (let i = 0; i < n; i++) {
					const fx = f(x);
					const dx = fx / df(x);
					const nx = x - dx;
					if (Math.abs(dx) < tolerance) { x = nx; break; }
					if (fx > 0) {
						b = x; x = nx <= a ? (a + b) * 0.5 : nx;
					} else { a = x; x = nx >= b ? (a + b) * 0.5 : nx; }
				}
				return clamp(x, a, b);
			},

			solveQuadratic: function (a, b, c, roots, min, max) {
				let x1;
				let x2 = Infinity;
				if (Math.abs(a) < EPSILON) {
					if (Math.abs(b) < EPSILON)
						return Math.abs(c) < EPSILON ? -1 : 0;
					x1 = -c / b;
				} else {
					b *= -0.5;
					let D = getDiscriminant(a, b, c);
					if (D && Math.abs(D) < MACHINE_EPSILON) {
						const f = getNormalizationFactor(Math.abs(a), Math.abs(b), Math.abs(c));
						if (f) {
							a *= f;
							b *= f;
							c *= f;
							D = getDiscriminant(a, b, c);
						}
					}
					if (D >= -MACHINE_EPSILON) {
						const Q = D < 0 ? 0 : Math.sqrt(D);
						const R = b + (b < 0 ? -Q : Q);
						x1 = (R === 0) ? c / a : R / a;
						x2 = (R === 0) ? -x1 : c / R;
					}
				}
				let count = 0;
				const boundless = min == null;
				const minB = min - EPSILON;
				const maxB = max + EPSILON;
				if (isFinite(x1) && (boundless || x1 > minB && x1 < maxB))
					roots[count++] = boundless ? x1 : clamp(x1, min, max);
				if (x2 !== x1 && isFinite(x2) && (boundless || x2 > minB && x2 < maxB))
					roots[count++] = boundless ? x2 : clamp(x2, min, max);
				return count;
			},

			solveCubic: function (a, b, c, d, roots, min, max) {
				const f = getNormalizationFactor(Math.abs(a), Math.abs(b), Math.abs(c), Math.abs(d));
				if (f) {
					a *= f;
					b *= f;
					c *= f;
					d *= f;
				}

				let x, b1, c2, qd, q;

				function evaluate(x0) {
					x = x0;
					const tmp = a * x;
					b1 = tmp + b;
					c2 = b1 * x + c;
					qd = (tmp + b1) * x + c2;
					q = c2 * x + d;
				}

				if (Math.abs(a) < EPSILON) {
					a = b;
					b1 = c;
					c2 = d;
					x = Infinity;
				} else if (Math.abs(d) < EPSILON) {
					b1 = b;
					c2 = c;
					x = 0;
				} else {
					evaluate(-(b / a) / 3);
					const t = q / a;
					const r = Math.pow(Math.abs(t), 1 / 3);
					const s = t < 0 ? -1 : 1;
					const td = -qd / a;
					const rd = td > 0 ? 1.324717957244746 * Math.max(r, Math.sqrt(td)) : r;
					let x0 = x - s * rd;
					if (x0 !== x) {
						do {
							evaluate(x0);
							x0 = qd === 0 ? x : x - q / qd / (1 + MACHINE_EPSILON);
						} while (s * x0 > s * x);
						if (Math.abs(a) * x * x > Math.abs(d / x)) {
							c2 = -d / x;
							b1 = (c2 - c) / x;
						}
					}
				}
				let count = Numerical.solveQuadratic(a, b1, c2, roots, min, max);
				const boundless = min == null;
				const isUniqueRoot = (count === 0 || count > 0 && x !== roots[0] && x !== roots[1]);
				const isWithinBounds = boundless || x > min - EPSILON && x < max + EPSILON;
				if (isFinite(x) && isUniqueRoot && isWithinBounds)
					roots[count++] = boundless ? x : clamp(x, min, max);
				return count;
			},
		};
	};

	const UID = {
		id: 1,
		pools: {},

		get: function (name) {
			if (name) {
				let pool = this.pools[name];
				if (!pool)
					pool = this.pools[name] = { _id: 1 };
				return pool._id++;
			} else {
				return this.id++;
			}
		}
	};

	const Base = function () { };

	Base.each = function (obj, iter, bind) {
		if (obj) {
			const descriptor = Object.getOwnPropertyDescriptor(obj, 'length');
			const forIn = function (iter, bind) {
				for (let i in this) {
					if (this.hasOwnProperty(i))
						iter.call(bind, this[i], i, this);
				}
			};
			const iterFunction = descriptor && typeof descriptor.value === 'number' ? Array.prototype.forEach : forIn;
			iterFunction.call(obj, iter, bind = bind || obj);
		}
		return bind;
	};
	Base.isPlainObject = function (obj) {
		const ctor = obj != null && obj.constructor;
		return ctor && (ctor === Object || ctor === Base || ctor.name === 'Object');
	};
	Base.pick = function (a, b) {
		return a !== undefined ? a : b;
	};
	Base.slice = function (list, begin, end) {
		return Array.prototype.slice.call(list, begin, end);
	};
	Base.equals = function (obj1, obj2) {
		if (obj1 === obj2)
			return true;
		if (obj1 && obj1.equals)
			return obj1.equals(obj2);
		if (obj2 && obj2.equals)
			return obj2.equals(obj1);
		if (obj1 && obj2
			&& typeof obj1 === 'object' && typeof obj2 === 'object') {
			if (Array.isArray(obj1) && Array.isArray(obj2)) {
				let length = obj1.length;
				if (length !== obj2.length)
					return false;
				while (length--) {
					if (!Base.equals(obj1[length], obj2[length]))
						return false;
				}
			} else {
				const keys = Object.keys(obj1);
				let length = keys.length;
				if (length !== Object.keys(obj2).length)
					return false;
				while (length--) {
					const key = keys[length];
					if (!(obj2.hasOwnProperty(key) && Base.equals(obj1[key], obj2[key])))
						return false;
				}
			}
			return true;
		}
		return false;
	};
	Base.read = function (list, start, options, amount) {
		if (this === Base) {
			const value = list[list.__index = start || list.__index || 0];
			list.__index++;
			return value;
		}
		const readIndex = this === Point || this === Rectangle;
		const begin = start || readIndex && list.__index || 0;
		let obj = list[begin];
		amount = amount || list.length - begin;
		if (obj instanceof this || options && options.readNull && obj == null && amount <= 1) {
			if (readIndex) { list.__index = begin + 1; }
			return obj && options && options.clone ? obj.clone() : obj;
		}
		obj = Object.create(this.prototype);
		if (readIndex)
			obj.__read = true;
		obj = obj.initialize.apply(obj, begin > 0 || begin + amount < list.length
			? Base.slice(list, begin, begin + amount)
			: list) || obj;
		if (readIndex) {
			list.__index = begin + obj.__read;
			obj.__read = undefined;
		}
		return obj;
	};
	Base.readList = function (list, start, options, amount) {
		const res = [];
		const begin = start || 0;
		const end = amount ? begin + amount : list.length;
		for (let i = begin; i < end; i++) {
			const entry = list[i];
			res.push(Array.isArray(entry) ? this.read(entry, 0, options) : this.read(list, i, options, 1));
		}
		return res;
	};
	Base.filter = function (dest, source, exclude, prioritize) {
		let processed;

		function handleKey(key) {
			if (!(exclude && key in exclude) && !(processed && key in processed)) {
				const value = source[key];
				if (value !== undefined) dest[key] = value;
			}
		}

		if (prioritize) {
			const keys = {};
			for (let i = 0, l = prioritize.length; i < l; i++) {
				const key = prioritize[i];
				if (key in source) {
					handleKey(key);
					keys[key] = true;
				}
			}
			processed = keys;
		}

		Object.keys(source).forEach(handleKey);
		return dest;
	};
	Base.isPlainValue = function (obj, asString) {
		return Base.isPlainObject(obj) || Array.isArray(obj) || asString && typeof obj === 'string';
	};
	Base.splice = function (list, items, index, remove) {
		const amount = items && items.length;
		const append = index === undefined;
		index = append ? list.length : index;
		if (index > list.length)
			index = list.length;
		for (let i = 0; i < amount; i++)
			items[i]._index = index + i;
		if (append) {
			list.push.apply(list, items);
			return [];
		} else {
			const args = [index, remove];
			if (items)
				args.push.apply(args, items);
			const removed = list.splice.apply(list, args);
			for (let i = 0, l = removed.length; i < l; i++)
				removed[i]._index = undefined;
			for (let i = index + amount, l = list.length; i < l; i++)
				list[i]._index = i;
			return removed;
		}
	};
	function isRealNumber(n) {
		return typeof n === "number" && !isNaN(n) && isFinite(n);
	}

	const Point = function (arg0, arg1, owner) {
		const type = typeof arg0;
		const isReading = this.__read;
		let readCount = 0;

		if (type === 'number') {
			const hasY = typeof arg1 === 'number';
			this._set(arg0, hasY ? arg1 : arg0);
			if (isReading) { readCount = hasY ? 2 : 1; }
		} else if (type === 'undefined' || arg0 === null) {
			this._set(0, 0);
			if (isReading) { readCount = arg0 === null ? 1 : 0; }
		} else {
			readCount = 1;
			if (Array.isArray(arg0)) {
				this._set(+arg0[0], +(arg0.length > 1 ? arg0[1] : arg0[0]));
			} else if (isRealNumber(arg0.x)) {
				this._set(arg0.x || 0, arg0.y || 0);
			} else {
				this._set(0, 0);
				readCount = 0;
			}
		}

		if (isReading) { this.__read = readCount; }
		if (owner) this._owner = owner;
		return this;
	};
	InitClassWithStatics(Point, Base);

	Point.prototype.set = Point;
	Point.prototype._set = function (x, y) {
		this.x = x;
		this.y = y;
		if (this._owner) this._owner._changed(this);
		return this;
	};
	Point.prototype.getX = function () {
		return this.x;
	};
	Point.prototype.getY = function () {
		return this.y;
	};
	Point.prototype.equals = function (point) {
		return this === point || point && (
			this.x === point.x && this.y === point.y ||
			Array.isArray(point) && this.x === point[0] && this.y === point[1]
		);
	};
	Point.prototype.clone = function () {
		return new Point(this.x, this.y);
	};
	Point.prototype.getLength = function () {
		return Math.sqrt(this.x * this.x + this.y * this.y);
	};
	Point.prototype.getAngle = function () {
		return this.getAngleInRadians.apply(this, arguments) * 180 / Math.PI;
	};
	Point.prototype.getAngleInRadians = function () {
		if (!arguments.length) {
			return this.isZero()
				? this._angle || 0
				: this._angle = Math.atan2(this.y, this.x);
		} else {
			const point = Point.read(arguments);
			const div = this.getLength() * point.getLength();
			if (Numerical.isZero(div)) {
				return NaN;
			} else {
				const a = this.dot(point) / div;
				return Math.acos(a < -1 ? -1 : a > 1 ? 1 : a);
			}
		}
	};
	Point.prototype.getDistance = function () {
		const point = Point.read(arguments);
		const x = point.x - this.x;
		const y = point.y - this.y;
		const d = x * x + y * y;
		const squared = Base.read(arguments);
		return squared ? d : Math.sqrt(d);
	};
	Point.prototype.normalize = function (length) {
		if (length === undefined)
			length = 1;
		const current = this.getLength();
		const scale = current !== 0 ? length / current : 0;
		const point = new Point(this.x * scale, this.y * scale);
		if (scale >= 0)
			point._angle = this._angle;
		return point;
	};
	Point.prototype.rotate = function (angle, center) {
		if (angle === 0)
			return this.clone();
		angle = angle * Math.PI / 180;
		let point = center ? this.subtract(center) : this;
		const sin = Math.sin(angle);
		const cos = Math.cos(angle);
		point = new Point(
			point.x * cos - point.y * sin,
			point.x * sin + point.y * cos
		);
		return center ? point.add(center) : point;
	};
	Point.prototype.transform = function (matrix) {
		return matrix ? matrix._transformPoint(this) : this;
	};
	Point.prototype.add = function () {
		const point = Point.read(arguments);
		return new Point(this.x + point.x, this.y + point.y);
	};
	Point.prototype.subtract = function () {
		const point = Point.read(arguments);
		return new Point(this.x - point.x, this.y - point.y);
	};
	Point.prototype.multiply = function () {
		const point = Point.read(arguments);
		return new Point(this.x * point.x, this.y * point.y);
	};
	Point.prototype.divide = function () {
		const point = Point.read(arguments);
		return new Point(this.x / point.x, this.y / point.y);
	};
	Point.prototype.negate = function () {
		return new Point(-this.x, -this.y);
	};
	Point.prototype.isInside = function () {
		return Rectangle.read(arguments).contains(this);
	};
	Point.prototype.isClose = function () {
		const point = Point.read(arguments);
		const tolerance = Base.read(arguments);
		return this.getDistance(point) <= tolerance;
	};
	Point.prototype.isCollinear = function () {
		const point = Point.read(arguments);
		return Point.isCollinear(this.x, this.y, point.x, point.y);
	};
	Point.prototype.isOrthogonal = function () {
		const point = Point.read(arguments);
		return Point.isOrthogonal(this.x, this.y, point.x, point.y);
	};
	Point.prototype.isZero = function () {
		return Numerical.isZero(this.x) && Numerical.isZero(this.y);
	};
	Point.prototype.isNaN = function () {
		return isNaN(this.x) || isNaN(this.y);
	};
	Point.prototype.dot = function () {
		const point = Point.read(arguments);
		return this.x * point.x + this.y * point.y;
	};
	Point.prototype.cross = function () {
		const point = Point.read(arguments);
		return this.x * point.y - this.y * point.x;
	};

	Point.isCollinear = function (x1, y1, x2, y2) {
		return Math.abs(x1 * y2 - y1 * x2)
			<= Math.sqrt((x1 * x1 + y1 * y1) * (x2 * x2 + y2 * y2)) * 1e-8;
	};
	Point.isOrthogonal = function (x1, y1, x2, y2) {
		return Math.abs(x1 * x2 + y1 * y2)
			<= Math.sqrt((x1 * x1 + y1 * y1) * (x2 * x2 + y2 * y2)) * 1e-8;
	};

	const Rectangle = function (arg0, arg1, arg2, arg3) {
		const type = typeof arg0;
		let read;
		if (type === 'number') {
			this._set(arg0, arg1, arg2, arg3);
			read = 4;
		} else if (type === 'undefined' || arg0 === null) {
			this._set(0, 0, 0, 0);
			read = arg0 === null ? 1 : 0;
		}
		if (this.__read) { this.__read = read; }
		return this;
	};
	InitClassWithStatics(Rectangle, Base);

	Rectangle.prototype._set = function (x, y, width, height) {
		this.x = x;
		this.y = y;
		this.width = width;
		this.height = height;
		return this;
	};
	Rectangle.prototype.clone = function () {
		return new Rectangle(this.x, this.y, this.width, this.height);
	};
	Rectangle.prototype.equals = function (rect) {
		const rt = Base.isPlainValue(rect) ? Rectangle.read(arguments) : rect;
		return rt === this || rt
			&& this.x === rt.x
			&& this.y === rt.y
			&& this.width === rt.width
			&& this.height === rt.height;
	};
	Rectangle.prototype.getPoint = function (_dontLink) {
		return new Point(this.x, this.y, this);
	};
	Rectangle.prototype.getLeft = Rectangle.prototype.getX = function () {
		return this.x;
	};
	Rectangle.prototype.getTop = Rectangle.prototype.getY = function () {
		return this.y;
	};
	Rectangle.prototype.getRight = function () {
		return this.x + this.width;
	};
	Rectangle.prototype.getBottom = function () {
		return this.y + this.height;
	};
	Rectangle.prototype.getWidth = function () {
		return this.width;
	};
	Rectangle.prototype.getHeight = function () {
		return this.height;
	};
	Rectangle.prototype.getCenterX = function () {
		return this.getLeft() + this.getWidth() / 2;
	};
	Rectangle.prototype.getCenterY = function () {
		return this.getTop() + this.getHeight() / 2;
	};
	Rectangle.prototype.getCenter = function (_dontLink) {
		return new Point(this.getCenterX(), this.getCenterY());
	};
	Rectangle.prototype.getTopLeft = function (_dontLink) {
		return new Point(this.getLeft(), this.getTop());
	};
	Rectangle.prototype.getArea = function () {
		return this.width * this.height;
	};
	Rectangle.prototype.isEmpty = function () {
		return this.width === 0 || this.height === 0;
	};
	Rectangle.prototype.contains = function (arg) {
		return arg && arg.width !== undefined || (Array.isArray(arg) ? arg : arguments).length === 4
			? this._containsRectangle(Rectangle.read(arguments))
			: this._containsPoint(Point.read(arguments));
	};
	Rectangle.prototype._containsPoint = function (point) {
		const x = point.x;
		const y = point.y;
		return x >= this.x && y >= this.y
			&& x <= this.x + this.width
			&& y <= this.y + this.height;
	};
	Rectangle.prototype._containsRectangle = function (rect) {
		const x = rect.x;
		const y = rect.y;
		return x >= this.x && y >= this.y
			&& x + rect.width <= this.x + this.width
			&& y + rect.height <= this.y + this.height;
	};
	Rectangle.prototype.intersects = function () {
		const rect = Rectangle.read(arguments);
		const epsilon = Base.read(arguments) || 0;
		return rect.x + rect.width > this.x - epsilon
			&& rect.y + rect.height > this.y - epsilon
			&& rect.x < this.x + this.width + epsilon
			&& rect.y < this.y + this.height + epsilon;
	};

	const Matrix = function (arg, _dontNotify) {
		const count = arguments.length;
		if (count >= 6) {
			this._set.apply(this, arguments);
		} else if (count === 1 || count === 2) {
			if (arg instanceof Matrix) {
				this._set(arg._a, arg._b, arg._c, arg._d, arg._tx, arg._ty, _dontNotify);
			} else if (Array.isArray(arg)) {
				this._set.apply(this, _dontNotify ? arg.concat([_dontNotify]) : arg);
			}
		} else if (!count) {
			this.reset();
		}
		return this;
	};
	InitClassWithStatics(Matrix, Base);

	Matrix.prototype.set = Matrix;
	Matrix.prototype._set = function (a, b, c, d, tx, ty, _dontNotify) {
		this._a = a;
		this._b = b;
		this._c = c;
		this._d = d;
		this._tx = tx;
		this._ty = ty;
		if (!_dontNotify)
			this._changed();
		return this;
	};
	Matrix.prototype._changed = function () {
		if (this._owner) {
			if (this._owner._applyMatrix) {
				this._owner.transform(null, true);
			} else {
				this._owner._changed(25);
			}
		}
	};
	Matrix.prototype.clone = function () {
		return new Matrix(this._a, this._b, this._c, this._d, this._tx, this._ty);
	};
	Matrix.prototype.equals = function (mx) {
		return mx === this || mx
			&& this._a === mx._a
			&& this._b === mx._b
			&& this._c === mx._c
			&& this._d === mx._d
			&& this._tx === mx._tx
			&& this._ty === mx._ty;
	};
	Matrix.prototype.reset = function (_dontNotify) {
		this._a = this._d = 1;
		this._b = this._c = this._tx = this._ty = 0;
		if (!_dontNotify)
			this._changed();
		return this;
	};
	Matrix.prototype.apply = function (recursively, _setApplyMatrix) {
		if (!this._owner) { return false; }
		this._owner.transform(null, Base.pick(recursively, true), _setApplyMatrix);
		return this.isIdentity();
	};
	Matrix.prototype.translate = function () {
		const point = Point.read(arguments);
		this._tx += point.x * this._a + point.y * this._c;
		this._ty += point.x * this._b + point.y * this._d;
		this._changed();
		return this;
	};
	Matrix.prototype.scale = function () {
		const scale = Point.read(arguments);
		const center = Point.read(arguments, 0, { readNull: true });
		if (center) { this.translate(center); }
		this._a *= scale.x;
		this._b *= scale.x;
		this._c *= scale.y;
		this._d *= scale.y;
		if (center) { this.translate(center.negate()); }
		this._changed();
		return this;
	};
	Matrix.prototype.rotate = function (angle) {
		angle *= Math.PI / 180;
		const center = Point.read(arguments, 1);
		const cos = Math.cos(angle);
		const sin = Math.sin(angle);
		const tx = center.x - center.x * cos + center.y * sin;
		const ty = center.y - center.x * sin - center.y * cos;
		const a = this._a, b = this._b, c = this._c, d = this._d;
		this._a = cos * a + sin * c;
		this._b = cos * b + sin * d;
		this._c = -sin * a + cos * c;
		this._d = -sin * b + cos * d;
		this._tx += tx * a + ty * c;
		this._ty += tx * b + ty * d;
		this._changed();
		return this;
	};
	Matrix.prototype.shear = function () {
		const shear = Point.read(arguments);
		const center = Point.read(arguments, 0, { readNull: true });
		if (center) { this.translate(center); }
		const a = this._a, b = this._b;
		this._a += shear.y * this._c;
		this._b += shear.y * this._d;
		this._c += shear.x * a;
		this._d += shear.x * b;
		if (center) { this.translate(center.negate()); }
		this._changed();
		return this;
	};
	Matrix.prototype.skew = function () {
		const skew = Point.read(arguments);
		const center = Point.read(arguments, 0, { readNull: true });
		const toRadians = Math.PI / 180;
		const shear = new Point(Math.tan(skew.x * toRadians), Math.tan(skew.y * toRadians));
		return this.shear(shear, center);
	};
	Matrix.prototype.append = function (mx, _dontNotify) {
		if (mx) {
			const a1 = this._a, b1 = this._b, c1 = this._c, d1 = this._d;
			const a2 = mx._a, b2 = mx._c, c2 = mx._b, d2 = mx._d;
			const tx2 = mx._tx, ty2 = mx._ty;
			this._a = a2 * a1 + c2 * c1;
			this._c = b2 * a1 + d2 * c1;
			this._b = a2 * b1 + c2 * d1;
			this._d = b2 * b1 + d2 * d1;
			this._tx += tx2 * a1 + ty2 * c1;
			this._ty += tx2 * b1 + ty2 * d1;
			if (!_dontNotify)
				this._changed();
		}
		return this;
	};
	Matrix.prototype.prepend = function (mx, _dontNotify) {
		if (mx) {
			const a1 = this._a, b1 = this._b, c1 = this._c, d1 = this._d;
			const a2 = mx._a, b2 = mx._c, c2 = mx._b, d2 = mx._d;
			const tx1 = this._tx, ty1 = this._ty;
			const tx2 = mx._tx, ty2 = mx._ty;
			this._a = a2 * a1 + b2 * b1;
			this._c = a2 * c1 + b2 * d1;
			this._b = c2 * a1 + d2 * b1;
			this._d = c2 * c1 + d2 * d1;
			this._tx = a2 * tx1 + b2 * ty1 + tx2;
			this._ty = c2 * tx1 + d2 * ty1 + ty2;
			if (!_dontNotify)
				this._changed();
		}
		return this;
	};
	Matrix.prototype.appended = function (mx) {
		return this.clone().append(mx);
	};
	Matrix.prototype.prepended = function (mx) {
		return this.clone().prepend(mx);
	};
	Matrix.prototype._orNullIfIdentity = function () {
		return this.isIdentity() ? null : this;
	};
	Matrix.prototype.isIdentity = function () {
		return this._a === 1 && this._b === 0 && this._c === 0 && this._d === 1
			&& this._tx === 0 && this._ty === 0;
	};
	Matrix.prototype.isInvertible = function () {
		const det = this._a * this._d - this._c * this._b;
		return det && !isNaN(det) && isFinite(this._tx) && isFinite(this._ty);
	};
	Matrix.prototype.transform = function (src, dst, count) {
		return arguments.length < 3
			? this._transformPoint(Point.read(arguments))
			: this._transformCoordinates(src, dst, count);
	};
	Matrix.prototype._transformPoint = function (point, dest, _dontNotify) {
		if (!dest)
			dest = new Point();
		return dest._set(
			point.x * this._a + point.y * this._c + this._tx,
			point.x * this._b + point.y * this._d + this._ty,
			_dontNotify
		);
	};
	Matrix.prototype._transformCoordinates = function (src, dst, count) {
		for (let i = 0, max = 2 * count; i < max; i += 2) {
			const x = src[i];
			const y = src[i + 1];
			dst[i] = x * this._a + y * this._c + this._tx;
			dst[i + 1] = x * this._b + y * this._d + this._ty;
		}
		return dst;
	};
	Matrix.prototype._transformCorners = function (rect) {
		const x1 = rect.x;
		const y1 = rect.y;
		const x2 = x1 + rect.width;
		const y2 = y1 + rect.height;
		const coords = [x1, y1, x2, y1, x2, y2, x1, y2];
		return this._transformCoordinates(coords, coords, 4);
	};
	Matrix.prototype._transformBounds = function (bounds, dest, _dontNotify) {
		const coords = this._transformCorners(bounds);
		const min = coords.slice(0, 2);
		const max = min.slice();
		for (let i = 2; i < 8; i++) {
			const val = coords[i];
			const j = i & 1;
			if (val < min[j]) {
				min[j] = val;
			} else if (val > max[j]) {
				max[j] = val;
			}
		}
		if (!dest)
			dest = new Rectangle();
		return dest._set(min[0], min[1], max[0] - min[0], max[1] - min[1], _dontNotify);
	};
	Matrix.prototype._inverseTransform = function (point, dest, _dontNotify) {
		const a = this._a, b = this._b, c = this._c, d = this._d;
		const tx = this._tx, ty = this._ty;
		const det = a * d - b * c;
		let res = null;
		if (det && !isNaN(det) && isFinite(tx) && isFinite(ty)) {
			const x = point.x - this._tx;
			const y = point.y - this._ty;
			if (!dest)
				dest = new Point();
			res = dest._set(
				(x * d - y * c) / det,
				(y * a - x * b) / det,
				_dontNotify
			);
		}
		return res;
	};
	Matrix.prototype.decompose = function () {
		const a = this._a, b = this._b, c = this._c, d = this._d;
		const det = a * d - b * c;
		const degrees = 180 / Math.PI;
		let rotate, scale, skew;
		if (a !== 0 || b !== 0) {
			const r = Math.sqrt(a * a + b * b);
			rotate = Math.acos(a / r) * (b > 0 ? 1 : -1);
			scale = [r, det / r];
			skew = [Math.atan2(a * c + b * d, r * r), 0];
		} else if (c !== 0 || d !== 0) {
			const s = Math.sqrt(c * c + d * d);
			rotate = Math.asin(c / s) * (d > 0 ? 1 : -1);
			scale = [det / s, s];
			skew = [0, Math.atan2(a * c + b * d, s * s)];
		} else {
			rotate = 0;
			skew = scale = [0, 0];
		}
		return {
			translation: this.getTranslation(),
			rotation: rotate * degrees,
			scaling: new Point(scale),
			skewing: new Point(skew[0] * degrees, skew[1] * degrees)
		};
	};
	Matrix.prototype.getValues = function () {
		return [this._a, this._b, this._c, this._d, this._tx, this._ty];
	};
	Matrix.prototype.getTranslation = function () {
		return new Point(this._tx, this._ty);
	};
	Matrix.prototype.getRotation = function () {
		return this.decompose().rotation;
	};

	const Line = function Line(arg0, arg1, arg2, arg3, arg4) {
		let asVector = false;
		if (arguments.length >= 4) {
			this._px = arg0;
			this._py = arg1;
			this._vx = arg2;
			this._vy = arg3;
			asVector = arg4;
		} else {
			this._px = arg0.x;
			this._py = arg0.y;
			this._vx = arg1.x;
			this._vy = arg1.y;
			asVector = arg2;
		}
		if (!asVector) {
			this._vx -= this._px;
			this._vy -= this._py;
		}
	};
	InitClassWithStatics(Line, Base);

	Line.prototype.getPoint = function () {
		return new Point(this._px, this._py);
	};
	Line.prototype.getVector = function () {
		return new Point(this._vx, this._vy);
	};
	Line.prototype.getLength = function () {
		return this.getVector().getLength();
	};
	Line.prototype.intersect = function (line, isInfinite) {
		return Line.intersect(
			this._px, this._py, this._vx, this._vy,
			line._px, line._py, line._vx, line._vy,
			true, isInfinite);
	};
	Line.prototype.getSide = function (point, isInfinite) {
		return Line.getSide(
			this._px, this._py, this._vx, this._vy,
			point.x, point.y, true, isInfinite);
	};
	Line.prototype.getDistance = function (point) {
		return Math.abs(this.getSignedDistance(point));
	};
	Line.prototype.getSignedDistance = function (point) {
		return Line.getSignedDistance(this._px, this._py, this._vx, this._vy,
			point.x, point.y, true);
	};
	Line.prototype.isCollinear = function (line) {
		return Point.isCollinear(this._vx, this._vy, line._vx, line._vy);
	};
	Line.prototype.isOrthogonal = function (line) {
		return Point.isOrthogonal(this._vx, this._vy, line._vx, line._vy);
	};

	Line.intersect = function (p1x, p1y, v1x, v1y, p2x, p2y, v2x, v2y, asVector,
		isInfinite) {
		if (!asVector) {
			v1x -= p1x;
			v1y -= p1y;
			v2x -= p2x;
			v2y -= p2y;
		}
		const cross = v1x * v2y - v1y * v2x;
		if (!Numerical.isMachineZero(cross)) {
			const dx = p1x - p2x;
			const dy = p1y - p2y;
			let u1 = (v2x * dy - v2y * dx) / cross;
			let u2 = (v1x * dy - v1y * dx) / cross;
			const uMin = -Numerical.EPSILON;
			const uMax = 1 + Numerical.EPSILON;
			if (isInfinite || uMin < u1 && u1 < uMax && uMin < u2 && u2 < uMax) {
				if (!isInfinite) {
					u1 = u1 <= 0 ? 0 : u1 >= 1 ? 1 : u1;
				}
				return new Point(p1x + u1 * v1x, p1y + u1 * v1y);
			}
		}
	};
	Line.getSide = function (px, py, vx, vy, x, y, asVector, isInfinite) {
		if (!asVector) {
			vx -= px;
			vy -= py;
		}
		const v2x = x - px;
		const v2y = y - py;
		let ccw = v2x * vy - v2y * vx;
		if (!isInfinite && Numerical.isMachineZero(ccw)) {
			ccw = (v2x * vx + v2x * vx) / (vx * vx + vy * vy);
			if (ccw >= 0 && ccw <= 1) {
				ccw = 0;
			}
		}
		return ccw < 0 ? -1 : ccw > 0 ? 1 : 0;
	};
	Line.getSignedDistance = function (px, py, vx, vy, x, y, asVector) {
		if (!asVector) {
			vx -= px;
			vy -= py;
		}
		return vx === 0 ? (vy > 0 ? x - px : px - x)
			: vy === 0 ? (vx < 0 ? y - py : py - y)
				: ((x - px) * vy - (y - py) * vx) / (
					vy > vx
						? vy * Math.sqrt(1 + (vx * vx) / (vy * vy))
						: vx * Math.sqrt(1 + (vy * vy) / (vx * vx))
				);
	};
	Line.getDistance = function (px, py, vx, vy, x, y, asVector) {
		return Math.abs(Line.getSignedDistance(px, py, vx, vy, x, y, asVector));
	};

	const Item = function () { };
	InitClassWithStatics(Item, Base);

	Item.prototype._applyMatrix = true;
	Item.prototype._canApplyMatrix = true;
	Item.prototype._pivot = null;
	Item.prototype._initialize = function (props, point) {
		const hasProps = props && Base.isPlainObject(props);
		const isInternal = hasProps && props.internal === true;
		const matrix = this._matrix = new Matrix();
		const settings = {
			applyMatrix: true,
		};
		this._id = isInternal ? null : UID.get();
		this._parent = this._index = null;
		this._applyMatrix = this._canApplyMatrix && settings.applyMatrix;
		if (point)
			matrix.translate(point);
		matrix._owner = this;
		return hasProps;
	};
	Item.prototype._changed = function (flags) {
		if (flags & 8) {
			this._bounds = this._position = this._decomposed = undefined;
		}
		if (flags & 16) {
			this._globalMatrix = undefined;
		}
		if (this._parent && (flags & 72)) {
			Item._clearBoundsCache(this._parent);
		}
		if (flags & 2) {
			Item._clearBoundsCache(this);
		}
	};
	Item.prototype.getPosition = function (_dontLink) {
		const position = this._position || (this._position = this._getPositionFromBounds());
		return new Point(position.x, position.y, this);
	};
	Item.prototype.setPosition = function () {
		this.translate(Point.read(arguments).subtract(this.getPosition(true)));
	};
	Item.prototype._getPositionFromBounds = function (bounds) {
		return this._pivot
			? this._matrix._transformPoint(this._pivot)
			: (bounds || this.getBounds()).getCenter(true);
	};
	Item.prototype.setPivot = function () {
		this._pivot = Point.read(arguments, 0, { clone: true, readNull: true });
		this._position = undefined;
	};
	Item.prototype.getBounds = function (matrix) {
		const opts = Object.assign({}, matrix);
		opts.cacheItem = this;
		const rect = this._getCachedBounds(false, opts).rect;
		return !!arguments.length
			? rect
			: new Rectangle(rect.x, rect.y, rect.width, rect.height);
	};
	Item.prototype.setBounds = function () {
		const rect = Rectangle.read(arguments);
		let bounds = this.getBounds();
		const matrix = new Matrix();
		let center = rect.getCenter();
		matrix.translate(center);
		if (rect.width != bounds.width || rect.height != bounds.height) {
			if (!this._matrix.isInvertible()) {
				this._matrix.set(this._matrix._backup || new Matrix().translate(this._matrix.getTranslation()));
				bounds = this.getBounds();
			}
			matrix.scale(
				bounds.width !== 0 ? rect.width / bounds.width : 0,
				bounds.height !== 0 ? rect.height / bounds.height : 0);
		}
		center = bounds.getCenter();
		matrix.translate(-center.x, -center.y);
		this.transform(matrix);
	};
	Item.prototype._getBounds = function (matrix, options) {
		if (!this._children || !this._children.length) {
			return new Rectangle();
		}
		Item._updateBoundsCache(this, options.cacheItem);
		return Item._getBounds(this._children, matrix, options);
	};
	Item.prototype._getBoundsCacheKey = function (options, internal) {
		return [
			options.stroke ? 1 : 0,
			options.handle ? 1 : 0,
			internal ? 1 : 0
		].join('');
	};
	Item.prototype._getCachedBounds = function (matrix, options, noInternal) {
		matrix = matrix && matrix._orNullIfIdentity();
		const isInternal = options.internal && !noInternal;
		const cacheItem = options.cacheItem;
		const _matrix = isInternal ? null : this._matrix._orNullIfIdentity();
		const cacheKey = cacheItem && (!matrix || matrix.equals(_matrix)) && this._getBoundsCacheKey(options, isInternal);
		Item._updateBoundsCache(this._parent, cacheItem);
		let cached;
		if (cacheKey && this._bounds && cacheKey in this._bounds) {
			cached = this._bounds[cacheKey];
			return {
				rect: cached.rect.clone(),
				nonscaling: cached.nonscaling
			};
		}
		const res = this._getBounds(matrix || _matrix, options);
		const rect = res.rect || res;
		const nonscaling = res.nonscaling;
		if (cacheKey) {
			if (!this._bounds) {
				this._bounds = this._bounds = {};
			}
			cached = this._bounds[cacheKey] = {
				rect: rect.clone(),
				nonscaling: nonscaling,
				internal: isInternal
			};
		}
		return {
			rect: rect,
			nonscaling: nonscaling
		};
	};
	Item.prototype._decompose = function () {
		return this._applyMatrix
			? null
			: this._decomposed || (this._decomposed = this._matrix.decompose());
	};
	Item.prototype.getRotation = function () {
		const decomposed = this._decompose();
		return decomposed ? decomposed.rotation : 0;
	};
	Item.prototype.setApplyMatrix = function (apply) {
		if (this._applyMatrix = this._canApplyMatrix && !!apply)
			this.transform(null, true);
	};
	Item.prototype._getOwner = Item.prototype.getParent = function () {
		return this._parent;
	};
	Item.prototype.getChildren = function () {
		return this._children;
	};
	Item.prototype.setChildren = function (items) {
		this.removeChildren();
		this.addChildren(items);
	};
	Item.prototype.getFirstChild = function () {
		return this._children && this._children[0] || null;
	};
	Item.prototype.getLastChild = function () {
		return this._children && this._children[this._children.length - 1]
			|| null;
	};
	Item.prototype.getNextSibling = function () {
		const owner = this._getOwner();
		return owner && owner._children[this._index + 1] || null;
	};
	Item.prototype.getPreviousSibling = function () {
		const owner = this._getOwner();
		return owner && owner._children[this._index - 1] || null;
	};
	Item.prototype.getIndex = function () {
		return this._index;
	};
	Item.prototype.equals = function (item) {
		return item === this || item
			&& this._matrix.equals(item._matrix)
			&& this._equals(item);
	};
	Item.prototype._equals = function (item) {
		return Base.equals(this._children, item._children);
	};
	Item.prototype.clone = function (options) {
		const copy = new this.constructor({ insert: false });
		if (this._children)
			copy.copyAttributes(this);

		const deep = Base.pick(options ? options.deep : undefined, true);
		if (!this._children || deep)
			copy.copyContent(this);

		if (!this._children)
			copy.copyAttributes(this);

		const shouldInsert = Base.pick(options ? options.insert : undefined, options === undefined || options === true);
		if (shouldInsert)
			copy.insertAbove(this);
		return copy;
	};
	Item.prototype.copyContent = function (source) {
		for (let i = 0, l = source._children && source._children.length; i < l; i++) {
			this.addChild(source._children[i].clone(false), true);
		}
	};
	Item.prototype.copyAttributes = function (source, excludeMatrix) {
		if (!excludeMatrix)
			this._matrix.set(source._matrix, true);
		this.setApplyMatrix(source._applyMatrix);
		this.setPivot(source._pivot);
		const data = source._data;
		this._data = data ? Object.assign(new data.constructor(), data) : null;
	};
	Item.prototype.contains = function () {
		return (
			this._matrix.isInvertible() &&
			!!this._contains(this._matrix._inverseTransform(Point.read(arguments)))
		);
	};
	Item.prototype._contains = function (point) {
		if (this._children) {
			for (let i = this._children.length - 1; i >= 0; i--) {
				if (this._children[i].contains(point))
					return true;
			}
			return false;
		}
		return point.isInside(this.getInternalBounds());
	};
	Item.prototype.isInside = function () {
		return Rectangle.read(arguments).contains(this.getBounds());
	};
	Item.prototype.addChild = function (item) {
		return this.insertChild(undefined, item);
	};
	Item.prototype.insertChild = function (index, item) {
		const res = item ? this.insertChildren(index, [item]) : null;
		return res && res[0];
	};
	Item.prototype.addChildren = function (items) {
		return this.insertChildren(this._children.length, items);
	};
	Item.prototype.insertChildren = function (index, items) {
		if (this._children && items && items.length > 0) {
			items = Base.slice(items);
			const inserted = {};
			for (let i = items.length - 1; i >= 0; i--) {
				const item = items[i],
					id = item && item._id;
				if (!item || inserted[id]) {
					items.splice(i, 1);
				} else {
					item._remove(false, true);
					inserted[id] = true;
				}
			}
			Base.splice(this._children, items, index, 0);
			for (let i = 0, l = items.length; i < l; i++) {
				const item = items[i];
				item._parent = this;
			}
			this._changed(11);
		} else {
			items = null;
		}
		return items;
	};
	Item.prototype._insertAt = function (item, offset) {
		const owner = item && item._getOwner(),
			res = item !== this && owner ? this : null;
		if (res) {
			res._remove(false, true);
			owner._insertChild(item._index + offset, res);
		}
		return res;
	};
	Item.prototype._insertChild = function(index, item) {
		var res = item ? this.insertChildren(index, [item]) : null;
		return res && res[0];
	};
	Item.prototype.insertAbove = function (item) {
		return this._insertAt(item, 1);
	};
	Item.prototype.insertBelow = function (item) {
		return this._insertAt(item, 0);
	};
	Item.prototype.reduce = function (options) {
		if (this._children && this._children.length === 1) {
			const child = this._children[0].reduce(options);
			if (this._parent) {
				child.insertAbove(this);
				this.remove();
			} else {
				child.remove();
			}
			return child;
		}
		return this;
	};
	Item.prototype._remove = function (notifySelf, notifyParent) {
		const owner = this._getOwner();
		if (owner) {
			if (this._index != null) {
				Base.splice(owner._children, null, this._index, 1);
			}
			if (notifyParent)
				owner._changed(11, this);
			this._parent = null;
			return true;
		}
		return false;
	};
	Item.prototype.remove = function () {
		return this._remove(true, true);
	};
	Item.prototype.replaceWith = function (item) {
		const ok = item && item.insertBelow(this);
		if (ok) { this.remove(); }
		return ok;
	};
	Item.prototype.clear = Item.prototype.removeChildren = function (start, end) {
		if (!this._children) { return null; }
		start = start || 0;
		end = Base.pick(end, this._children.length);
		const removed = Base.splice(this._children, null, start, end - start);
		for (let i = removed.length - 1; i >= 0; i--) {
			removed[i]._remove(true, false);
		}
		if (removed.length > 0)
			this._changed(11);
		return removed;
	};
	Item.prototype.isEmpty = function (recursively) {
		const numChildren = this._children ? this._children.length : 0;
		if (recursively) {
			for (let i = 0; i < numChildren; i++) {
				if (!this._children[i].isEmpty(recursively)) {
					return false;
				}
			}
			return true;
		}
		return !numChildren;
	};
	Item.prototype._getOrder = function (item) {
		function getList(item) {
			const list = [];
			do {
				list.unshift(item);
			} while (item = item._parent);
			return list;
		}
		const list1 = getList(this);
		const list2 = getList(item);
		for (let i = 0, l = Math.min(list1.length, list2.length); i < l; i++) {
			if (list1[i] != list2[i]) {
				return list1[i]._index < list2[i]._index ? 1 : -1;
			}
		}
		return 0;
	};
	Item.prototype.translate = function () {
		const mx = new Matrix();
		return this.transform(mx.translate.apply(mx, arguments));
	};
	Item.prototype.transform = function (matrix, _applyRecursively, _setApplyMatrix) {
		const _matrix = this._matrix;
		const transformMatrix = matrix && !matrix.isIdentity();
		let applyMatrix = (
			_setApplyMatrix && this._canApplyMatrix ||
			this._applyMatrix && (transformMatrix || !_matrix.isIdentity() || _applyRecursively && this._children)
		);
		if (!transformMatrix && !applyMatrix) { return this; }
		if (transformMatrix) {
			if (!matrix.isInvertible() && _matrix.isInvertible())
				_matrix._backup = _matrix.getValues();
			_matrix.prepend(matrix, true);
		}

		if (applyMatrix && (applyMatrix = this._transformContent(_matrix, _applyRecursively, _setApplyMatrix))) {
			if (this._pivot)
				_matrix._transformPoint(this._pivot, this._pivot, true);
			_matrix.reset(true);
			if (_setApplyMatrix && this._canApplyMatrix)
				this._applyMatrix = true;
		}
		const bounds = this._bounds;
		const position = this._position;
		if (transformMatrix || applyMatrix) {
			this._changed(25);
		}
		const decomp = transformMatrix && bounds && matrix.decompose();
		if (decomp && decomp.skewing.isZero() && decomp.rotation % 90 === 0) {
			for (let key in bounds) {
				const cache = bounds[key];
				if (cache.nonscaling) {
					delete bounds[key];
				} else if (applyMatrix || !cache.internal) {
					matrix._transformBounds(cache.rect, cache.rect);
				}
			}
			this._bounds = bounds;
			const cached = bounds['000'];
			if (cached) {
				this._position = this._getPositionFromBounds(cached.rect);
			}
		} else if (transformMatrix && position && this._pivot) {
			this._position = matrix._transformPoint(position, position);
		}
		return this;
	};
	Item.prototype._transformContent = function (matrix, applyRecursively, setApplyMatrix) {
		const children = this._children;
		if (children) {
			for (let i = 0, l = children.length; i < l; i++) {
				children[i].transform(matrix, applyRecursively, setApplyMatrix);
			}
			return true;
		}
	};

	Item._updateBoundsCache = function (parent, item) {
		if (parent && item) {
			const id = item._id;
			const ref = parent._boundsCache = parent._boundsCache || { ids: {}, list: [] };
			if (!ref.ids[id]) {
				ref.list.push(item);
				ref.ids[id] = item;
			}
		}
	};
	Item._clearBoundsCache = function (item) {
		const cache = item._boundsCache;
		if (cache) {
			item._bounds = item._position = item._boundsCache = undefined;
			for (let i = 0, list = cache.list, l = list.length; i < l; i++) {
				const other = list[i];
				if (other !== item) {
					other._bounds = other._position = undefined;
					if (other._boundsCache)
						Item._clearBoundsCache(other);
				}
			}
		}
	};
	Item._getBounds = function (items, matrix, options) {
		let x1 = Infinity;
		let x2 = -x1;
		let y1 = x1;
		let y2 = x2;
		let nonscaling = false;
		options = options || {};
		for (let i = 0, l = items.length; i < l; i++) {
			const item = items[i];
			if (!item.isEmpty(true)) {
				const bounds = item._getCachedBounds(matrix && matrix.appended(item._matrix), options, true);
				x1 = Math.min(bounds.rect.x, x1);
				y1 = Math.min(bounds.rect.y, y1);
				x2 = Math.max(bounds.rect.x + bounds.rect.width, x2);
				y2 = Math.max(bounds.rect.y + bounds.rect.height, y2);
				if (bounds.nonscaling)
					nonscaling = true;
			}
		}
		return {
			rect: isFinite(x1)
				? new Rectangle(x1, y1, x2 - x1, y2 - y1)
				: new Rectangle(),
			nonscaling: nonscaling
		};
	};

	const Segment = function (arg0, arg1, arg2, arg3, arg4, arg5) {
		for (let i = 0, l = arguments.length; i < l; i++) {
			const src = arguments[i];
			if (src)
				Object.assign(this, src);
		}
		const count = arguments.length;
		let point, handleIn, handleOut;
		if (count > 0) {
			if (arg0 == null || typeof arg0 === 'object') {
				if (count === 1 && arg0 && arg0.point) {
					point = arg0.point;
					handleIn = arg0.handleIn;
					handleOut = arg0.handleOut;
				} else {
					point = arg0;
					handleIn = arg1;
					handleOut = arg2;
				}
			} else {
				point = [arg0, arg1];
				handleIn = arg2 !== undefined ? [arg2, arg3] : null;
				handleOut = arg4 !== undefined ? [arg4, arg5] : null;
			}
		}
		this._point = new Point(point, this);
		this._handleIn = new Point(handleIn, this);
		this._handleOut = new Point(handleOut, this);
	};
	InitClassWithStatics(Segment, Base);

	Segment.prototype._changed = function (point) {
		if (!this._path) { return; }

		const curves = this._path._curves;
		const index = this._index;
		let curve;
		if (curves) {
			if ((!point || point === this._point || point === this._handleIn) && (curve = index > 0 ? curves[index - 1] : this._path._closed ? curves[curves.length - 1] : null))
				curve._changed();
			if ((!point || point === this._point || point === this._handleOut) && (curve = curves[index]))
				curve._changed();
		}
		this._path._changed(41);
	};
	Segment.prototype.getPoint = function () {
		return this._point;
	};
	Segment.prototype.getHandleIn = function () {
		return this._handleIn;
	};
	Segment.prototype.setHandleIn = function () {
		this._handleIn.set(Point.read(arguments));
	};
	Segment.prototype.getHandleOut = function () {
		return this._handleOut;
	};
	Segment.prototype.setHandleOut = function () {
		const newPoint = Point.read(arguments)
		this._handleOut.set(newPoint);
	};
	Segment.prototype.hasHandles = function () {
		return !this._handleIn.isZero() || !this._handleOut.isZero();
	};
	Segment.prototype.isSmooth = function () {
		return !this._handleIn.isZero() && !this._handleOut.isZero() && this._handleIn.isCollinear(this._handleOut);
	};
	Segment.prototype.clearHandles = function () {
		this._handleIn._set(0, 0);
		this._handleOut._set(0, 0);
	};
	Segment.prototype.getIndex = function () {
		return this._index !== undefined ? this._index : null;
	};
	Segment.prototype.getPath = function () {
		return this._path || null;
	};
	Segment.prototype.getCurve = function () {
		const path = this._path;
		let index = this._index;
		if (path) {
			if (index > 0 && !path._closed
				&& index === path._segments.length - 1)
				index--;
			return path.getCurves()[index] || null;
		}
		return null;
	};
	Segment.prototype.getLocation = function () {
		const curve = this.getCurve();
		return curve ? new CurveLocation(curve, this === curve._segment1 ? 0 : 1) : null;
	};
	Segment.prototype.getNext = function () {
		const segments = this._path && this._path._segments;
		return segments && (segments[this._index + 1] || this._path._closed && segments[0]) || null;
	};
	Segment.prototype.getPrevious = function () {
		const segments = this._path && this._path._segments;
		return segments && (segments[this._index - 1] || this._path._closed && segments[segments.length - 1]) || null;
	};
	Segment.prototype.isFirst = function () {
		return !this._index;
	};
	Segment.prototype.isLast = function () {
		return this._path && this._index === this._path._segments.length - 1 || false;
	};
	Segment.prototype.reverse = function () {
		const handleIn = this._handleIn;
		const handleOut = this._handleOut;
		const tmp = handleIn.clone();
		handleIn.set(handleOut);
		handleOut.set(tmp);
	};
	Segment.prototype.reversed = function () {
		return new Segment(this._point, this._handleOut, this._handleIn);
	};
	Segment.prototype.remove = function () {
		return this._path ? !!this._path.removeSegment(this._index) : false;
	};
	Segment.prototype.clone = function () {
		return new Segment(this._point, this._handleIn, this._handleOut);
	};
	Segment.prototype.equals = function (segment) {
		return segment === this || segment
			&& this._point.equals(segment._point)
			&& this._handleIn.equals(segment._handleIn)
			&& this._handleOut.equals(segment._handleOut)
			|| false;
	};
	Segment.prototype.transform = function (matrix) {
		this._transformCoordinates(matrix, new Array(6), true);
		this._changed();
	};
	Segment.prototype._transformCoordinates = function (matrix, coords, change) {
		const handleIn = !change || !this._handleIn.isZero() ? this._handleIn : null;
		const handleOut = !change || !this._handleOut.isZero() ? this._handleOut : null;
		let x = this._point.getX();
		let y = this._point.getY();
		let i = 2;
		coords[0] = x;
		coords[1] = y;
		if (handleIn) {
			coords[i++] = handleIn.getX() + x;
			coords[i++] = handleIn.getY() + y;
		}
		if (handleOut) {
			coords[i++] = handleOut.getX() + x;
			coords[i++] = handleOut.getY() + y;
		}
		if (matrix) {
			matrix._transformCoordinates(coords, coords, i / 2);
			x = coords[0];
			y = coords[1];
			if (change) {
				this._point.x = x;
				this._point.y = y;
				i = 2;
				if (handleIn) {
					handleIn.x = coords[i++] - x;
					handleIn.y = coords[i++] - y;
				}
				if (handleOut) {
					handleOut.x = coords[i++] - x;
					handleOut.y = coords[i++] - y;
				}
			} else {
				if (!handleIn) {
					coords[i++] = x;
					coords[i++] = y;
				}
				if (!handleOut) {
					coords[i++] = x;
					coords[i++] = y;
				}
			}
		}
		return coords;
	};

	const Curve = function (arg0, arg1, arg2, arg3, arg4, arg5, arg6, arg7) {
		const count = arguments.length;
		let seg1, seg2;
		let point1, point2;
		let handle1, handle2;
		if (count === 3) {
			this._path = arg0;
			seg1 = arg1;
			seg2 = arg2;
		} else if (!count) {
			seg1 = new Segment();
			seg2 = new Segment();
		} else if (count === 1) {
			if (arg0.segment1) {
				seg1 = new Segment(arg0.segment1);
				seg2 = new Segment(arg0.segment2);
			} else if (arg0.point1) {
				point1 = arg0.point1;
				handle1 = arg0.handle1;
				handle2 = arg0.handle2;
				point2 = arg0.point2;
			} else if (Array.isArray(arg0)) {
				point1 = [arg0[0], arg0[1]];
				point2 = [arg0[6], arg0[7]];
				handle1 = [arg0[2] - arg0[0], arg0[3] - arg0[1]];
				handle2 = [arg0[4] - arg0[6], arg0[5] - arg0[7]];
			}
		} else if (count === 2) {
			seg1 = new Segment(arg0);
			seg2 = new Segment(arg1);
		} else if (count === 4) {
			point1 = arg0;
			handle1 = arg1;
			handle2 = arg2;
			point2 = arg3;
		} else if (count === 8) {
			point1 = [arg0, arg1];
			point2 = [arg6, arg7];
			handle1 = [arg2 - arg0, arg3 - arg1];
			handle2 = [arg4 - arg6, arg5 - arg7];
		}
		this._segment1 = seg1 || new Segment(point1, null, handle1);
		this._segment2 = seg2 || new Segment(point2, handle2, null);
	};
	InitClassWithStatics(Curve, Base);

	Curve.prototype._changed = function () {
		this._length = this._bounds = undefined;
	};
	Curve.prototype.clone = function () {
		return new Curve(this._segment1, this._segment2);
	};
	Curve.prototype.classify = function () {
		return Curve.classify(this.getValues());
	};
	Curve.prototype.remove = function () {
		let removed = false;
		if (this._path) {
			const segment2 = this._segment2;
			const handleOut = segment2._handleOut;
			removed = segment2.remove();
			if (removed)
				this._segment1._handleOut.set(handleOut);
		}
		return removed;
	};
	Curve.prototype.getPoint1 = function () {
		return this._segment1._point;
	};
	Curve.prototype.getPoint2 = function () {
		return this._segment2._point;
	};
	Curve.prototype.getSegment1 = function () {
		return this._segment1;
	};
	Curve.prototype.getSegment2 = function () {
		return this._segment2;
	};
	Curve.prototype.getPath = function () {
		return this._path;
	};
	Curve.prototype.getIndex = function () {
		return this._segment1._index;
	};
	Curve.prototype.getNext = function () {
		const curves = this._path && this._path._curves;
		return curves && (curves[this._segment1._index + 1]
			|| this._path._closed && curves[0]) || null;
	};
	Curve.prototype.getPrevious = function () {
		const curves = this._path && this._path._curves;
		return curves && (curves[this._segment1._index - 1]
			|| this._path._closed && curves[curves.length - 1]) || null;
	};
	Curve.prototype.isFirst = function () {
		return !this._segment1._index;
	};
	Curve.prototype.isLast = function () {
		const path = this._path;
		return path && this._segment1._index === path._curves.length - 1
			|| false;
	};
	Curve.prototype.getValues = function (matrix) {
		return Curve.getValues(this._segment1, this._segment2, matrix);
	};
	Curve.prototype.getLength = function () {
		if (this._length == null)
			this._length = Curve.getLength(this.getValues(), 0, 1);
		return this._length;
	};
	Curve.prototype.getArea = function () {
		return Curve.getArea(this.getValues());
	};
	Curve.prototype.getLine = function () {
		return new Line(this._segment1._point, this._segment2._point);
	};
	Curve.prototype.getPart = function (from, to) {
		return new Curve(Curve.getPart(this.getValues(), from, to));
	};
	Curve.prototype.getPartLength = function (from, to) {
		return Curve.getLength(this.getValues(), from, to);
	};
	Curve.prototype.divideAtTime = function (time, _setHandles) {
		const tMin = 1e-8, tMax = 1 - tMin;
		let res = null;
		if (time >= tMin && time <= tMax) {
			const parts = Curve.subdivide(this.getValues(), time);
			const left = parts[0];
			const right = parts[1];
			const setHandles = _setHandles || this.hasHandles();
			const seg1 = this._segment1;
			const seg2 = this._segment2;
			const path = this._path;
			if (setHandles) {
				seg1._handleOut._set(left[2] - left[0], left[3] - left[1]);
				seg2._handleIn._set(right[4] - right[6], right[5] - right[7]);
			}
			const x = left[6];
			const y = left[7];
			const segment = new Segment(new Point(x, y),
				setHandles && new Point(left[4] - x, left[5] - y),
				setHandles && new Point(right[2] - x, right[3] - y)
			);
			if (path) {
				path.insert(seg1._index + 1, segment);
				res = this.getNext();
			} else {
				this._segment2 = segment;
				this._changed();
				res = new Curve(segment, seg2);
			}
		}
		return res;
	};
	Curve.prototype.divide = function (offset, isTime) {
		return this.divideAtTime(offset === undefined ? 0.5 : isTime ? offset
			: this.getTimeAt(offset));
	};
	Curve.prototype.reversed = function () {
		return new Curve(this._segment2.reversed(), this._segment1.reversed());
	};
	Curve.prototype.clearHandles = function () {
		this._segment1._handleOut._set(0, 0);
		this._segment2._handleIn._set(0, 0);
	};
	Curve.prototype.hasHandles = function () {
		return !this._segment1._handleOut.isZero()
			|| !this._segment2._handleIn.isZero();
	};
	Curve.prototype.hasLength = function (epsilon) {
		return (!this.getPoint1().equals(this.getPoint2()) || this.hasHandles())
			&& this.getLength() > (epsilon || 0);
	};
	Curve.prototype.isCollinear = function (curve) {
		return curve && this.isStraight() && curve.isStraight()
			&& this.getLine().isCollinear(curve.getLine());
	};
	Curve.prototype.isStraight = function (epsilon) {
		const test = function (p1, h1, h2, p2) {
			if (h1.isZero() && h2.isZero()) {
				return true;
			} else {
				const v = p2.subtract(p1);
				if (v.isZero()) { return false; }
				if (v.isCollinear(h1) && v.isCollinear(h2)) {
					const l = new Line(p1, p2);
					const epsilon = 1e-7;
					if (l.getDistance(p1.add(h1)) < epsilon && l.getDistance(p2.add(h2)) < epsilon) {
						const div = v.dot(v);
						const s1 = v.dot(h1) / div;
						const s2 = v.dot(h2) / div;
						return s1 >= 0 && s1 <= 1 && s2 <= 0 && s2 >= -1;
					}
				}
			}
			return false;
		}
		return test(this._segment1._point, this._segment1._handleOut, this._segment2._handleIn, this._segment2._point, epsilon);
	};
	Curve.prototype.getLocationAt = function (offset, _isTime) {
		return this.getLocationAtTime(_isTime ? offset : this.getTimeAt(offset));
	};
	Curve.prototype.getLocationAtTime = function (t) {
		return t != null && t >= 0 && t <= 1
			? new CurveLocation(this, t)
			: null;
	};
	Curve.prototype.getTimeAt = function (offset, start) {
		return Curve.getTimeAt(this.getValues(), offset, start);
	};
	Curve.prototype.getTimesWithTangent = function () {
		const tangent = Point.read(arguments);
		return !tangent.isZero() ? Curve.getTimesWithTangent(this.getValues(), tangent) : [];
	};
	Curve.prototype.getTimeOf = function () {
		return Curve.getTimeOf(this.getValues(), Point.read(arguments));
	};
	Curve.prototype.getNearestLocation = function () {
		const point = Point.read(arguments);
		const values = this.getValues();
		const t = Curve.getNearestTime(values, point);
		const pt = Curve.getPoint(values, t);
		return new CurveLocation(this, t, pt, null, point.getDistance(pt));
	};
	Curve.prototype.getNearestPoint = function () {
		const loc = this.getNearestLocation.apply(this, arguments);
		return loc ? loc.getPoint() : loc;
	};
	Curve.prototype.getPointAt = function (location, _isTime) {
		const values = this.getValues();
		return Curve.getPoint(values, _isTime ? location : Curve.getTimeAt(values, location));
	};
	Curve.prototype.getPointAtTime = function (time) {
		return Curve.getPoint(this.getValues(), time);
	};
	Curve.prototype.getTangentAt = function (location, _isTime) {
		const values = this.getValues();
		return Curve.getTangent(values, _isTime ? location : Curve.getTimeAt(values, location));
	};
	Curve.prototype.getTangentAtTime = function (time) {
		return Curve.getTangent(this.getValues(), time);
	};
	Curve.prototype.getNormalAt = function (location, _isTime) {
		const values = this.getValues();
		return Curve.getNormal(values, _isTime ? location : Curve.getTimeAt(values, location));
	};
	Curve.prototype.getNormalAtTime = function (time) {
		return Curve.getNormal(this.getValues(), time);
	};
	Curve.prototype.getWeightedTangentAt = function (location) {
		const values = this.getValues();
		return Curve.getWeightedTangent(values, location);
	};
	Curve.prototype.getWeightedNormalAt = function (location) {
		const values = this.getValues();
		return Curve.getWeightedNormal(values, location);
	};
	Curve.prototype.getCurvatureAt = function (location) {
		const values = this.getValues();
		return Curve.getCurvature(values, location);
	};
	Curve.prototype.getIntersections = function (curve) {
		const v1 = this.getValues();
		const v2 = curve && curve !== this && curve.getValues();
		return v2
			? Curve.getCurveIntersections(v1, v2, this, curve, [])
			: Curve.getSelfIntersection(v1, this, []);
	};

	Curve.getValues = function (segment1, segment2, matrix, straight) {
		const p1 = segment1._point;
		const h1 = segment1._handleOut;
		const h2 = segment2._handleIn;
		const p2 = segment2._point;
		const x1 = p1.x;
		const y1 = p1.y;
		const x2 = p2.x;
		const y2 = p2.y;
		const values = straight
			? [x1, y1, x1, y1, x2, y2, x2, y2]
			: [x1, y1, x1 + h1.getX(), y1 + h1.getY(), x2 + h2.getX(), y2 + h2.getY(), x2, y2];
		if (matrix) { matrix._transformCoordinates(values, values, 4); }
		return values;
	};
	Curve.subdivide = function (v, t) {
		if (t === undefined) { t = 0.5; }
		const x0 = v[0], y0 = v[1];
		const x1 = v[2], y1 = v[3];
		const x2 = v[4], y2 = v[5];
		const x3 = v[6], y3 = v[7];
		const u = 1 - t;
		const x4 = u * x0 + t * x1;
		const y4 = u * y0 + t * y1;
		const x5 = u * x1 + t * x2;
		const y5 = u * y1 + t * y2;
		const x6 = u * x2 + t * x3;
		const y6 = u * y2 + t * y3;
		const x7 = u * x4 + t * x5;
		const y7 = u * y4 + t * y5;
		const x8 = u * x5 + t * x6;
		const y8 = u * y5 + t * y6;
		const x9 = u * x7 + t * x8;
		const y9 = u * y7 + t * y8;
		return [
			[x0, y0, x4, y4, x7, y7, x9, y9],
			[x9, y9, x8, y8, x6, y6, x3, y3]
		];
	};
	Curve.getMonoCurves = function (v, dir) {
		const curves = [];
		const io = dir ? 0 : 1;
		const o0 = v[io + 0];
		const o1 = v[io + 2];
		const o2 = v[io + 4];
		const o3 = v[io + 6];
		if ((o0 >= o1) === (o1 >= o2) && (o1 >= o2) === (o2 >= o3) || Curve.isStraight(v)) {
			curves.push(v);
		} else {
			const a = 3 * (o1 - o2) - o0 + o3;
			const b = 2 * (o0 + o2) - 4 * o1;
			const c = o1 - o0;
			const tMin = 1e-8;
			const tMax = 1 - tMin;
			const roots = [];
			const n = Numerical.solveQuadratic(a, b, c, roots, tMin, tMax);
			if (!n) {
				curves.push(v);
			} else {
				roots.sort();
				let t = roots[0];
				let parts = Curve.subdivide(v, t);
				curves.push(parts[0]);
				if (n > 1) {
					t = (roots[1] - t) / (1 - t);
					parts = Curve.subdivide(parts[1], t);
					curves.push(parts[0]);
				}
				curves.push(parts[1]);
			}
		}
		return curves;
	};
	Curve.solveCubic = function (v, coord, val, roots, min, max) {
		const v0 = v[coord];
		const v1 = v[coord + 2];
		const v2 = v[coord + 4];
		const v3 = v[coord + 6];
		let res = 0;
		if (!(v0 < val && v3 < val && v1 < val && v2 < val || v0 > val && v3 > val && v1 > val && v2 > val)) {
			const c = 3 * (v1 - v0);
			const b = 3 * (v2 - v1) - c;
			const a = v3 - v0 - c - b;
			res = Numerical.solveCubic(a, b, c, v0 - val, roots, min, max);
		}
		return res;
	};
	Curve.getTimeOf = function (v, point) {
		const p0 = new Point(v[0], v[1]);
		const p3 = new Point(v[6], v[7]);
		const t = point.isClose(p0, Numerical.EPSILON) ? 0
			: point.isClose(p3, Numerical.EPSILON) ? 1
				: null;
		if (t === null) {
			const coords = [point.x, point.y];
			const roots = [];
			for (let c = 0; c < 2; c++) {
				const count = Curve.solveCubic(v, c, coords[c], roots, 0, 1);
				for (let i = 0; i < count; i++) {
					const u = roots[i];
					if (point.isClose(Curve.getPoint(v, u), Numerical.GEOMETRIC_EPSILON))
						return u;
				}
			}
		}
		return point.isClose(p0, Numerical.GEOMETRIC_EPSILON) ? 0
			: point.isClose(p3, Numerical.GEOMETRIC_EPSILON) ? 1
				: null;
	};
	Curve.getNearestTime = function (v, point) {
		if (Curve.isStraight(v)) {
			const x0 = v[0], y0 = v[1];
			const x3 = v[6], y3 = v[7];
			const vx = x3 - x0, vy = y3 - y0;
			const det = vx * vx + vy * vy;
			if (det === 0) { return 0; }

			const u = ((point.x - x0) * vx + (point.y - y0) * vy) / det;
			return u < Numerical.EPSILON ? 0
				: u > 0.999999999999 ? 1
					: Curve.getTimeOf(v, new Point(x0 + u * vx, y0 + u * vy));
		}

		const count = 100;
		let minDist = Infinity;
		let minT = 0;

		function refine(t) {
			if (t >= 0 && t <= 1) {
				const dist = point.getDistance(Curve.getPoint(v, t), true);
				if (dist < minDist) {
					minDist = dist;
					minT = t;
					return true;
				}
			}
		}

		for (let i = 0; i <= count; i++)
			refine(i / count);

		let step = 1 / (count * 2);
		while (step > 1e-8) {
			if (!refine(minT - step) && !refine(minT + step))
				step /= 2;
		}
		return minT;
	};
	Curve.getPart = function (v, from, to) {
		const flip = from > to;
		if (flip) {
			const tmp = from;
			from = to;
			to = tmp;
		}
		if (from > 0)
			v = Curve.subdivide(v, from)[1];
		if (to < 1)
			v = Curve.subdivide(v, (to - from) / (1 - from))[0];
		return flip ? [v[6], v[7], v[4], v[5], v[2], v[3], v[0], v[1]] : v;
	};
	Curve.getArea = function (v) {
		const x0 = v[0], y0 = v[1];
		const x1 = v[2], y1 = v[3];
		const x2 = v[4], y2 = v[5];
		const x3 = v[6], y3 = v[7];
		return 3 * ((y3 - y0) * (x1 + x2) - (x3 - x0) * (y1 + y2)
			+ y1 * (x0 - x2) - x1 * (y0 - y2)
			+ y3 * (x2 + x0 / 3) - x3 * (y2 + y0 / 3)) / 20;
	};
	Curve.getBounds = function (v) {
		const min = v.slice(0, 2);
		const max = min.slice();
		const roots = [0, 0];
		for (let i = 0; i < 2; i++)
			Curve._addBounds(v[i], v[i + 2], v[i + 4], v[i + 6], i, 0, min, max, roots);
		return new Rectangle(min[0], min[1], max[0] - min[0], max[1] - min[1]);
	};
	Curve._addBounds = function (v0, v1, v2, v3, coord, padding, min, max, roots) {
		function add(value, padding) {
			const left = value - padding;
			const right = value + padding;
			if (left < min[coord])
				min[coord] = left;
			if (right > max[coord])
				max[coord] = right;
		}

		padding /= 2;
		const minPad = min[coord] + padding;
		const maxPad = max[coord] - padding;
		if (v0 < minPad || v1 < minPad || v2 < minPad || v3 < minPad ||
			v0 > maxPad || v1 > maxPad || v2 > maxPad || v3 > maxPad) {
			if (v1 < v0 != v1 < v3 && v2 < v0 != v2 < v3) {
				add(v0, 0);
				add(v3, 0);
			} else {
				const a = 3 * (v1 - v2) - v0 + v3;
				const b = 2 * (v0 + v2) - 4 * v1;
				const c = v1 - v0;
				const count = Numerical.solveQuadratic(a, b, c, roots);
				const tMin = 1e-8;
				const tMax = 1 - tMin;
				add(v3, 0);
				for (let i = 0; i < count; i++) {
					const t = roots[i];
					const u = 1 - t;
					if (tMin <= t && t <= tMax)
						add(u * u * u * v0
							+ 3 * u * u * t * v1
							+ 3 * u * t * t * v2
							+ t * t * t * v3,
							padding
						);
				}
			}
		}
	};
	Curve.isStraight = function (v, epsilon) {
		const x0 = v[0];
		const y0 = v[1];
		const x3 = v[6];
		const y3 = v[7];

		function test(p1, h1, h2, p2) {
			if (h1.isZero() && h2.isZero()) {
				return true;
			} else {
				const v = p2.subtract(p1);
				if (v.isZero()) { return false; }

				if (v.isCollinear(h1) && v.isCollinear(h2)) {
					const l = new Line(p1, p2);
					if (l.getDistance(p1.add(h1)) < Numerical.GEOMETRIC_EPSILON &&
						l.getDistance(p2.add(h2)) < Numerical.GEOMETRIC_EPSILON) {
						const div = v.dot(v);
						const s1 = v.dot(h1) / div;
						const s2 = v.dot(h2) / div;
						return s1 >= 0 && s1 <= 1 && s2 <= 0 && s2 >= -1;
					}
				}
			}
			return false;
		};

		return test(
			new Point(x0, y0),
			new Point(v[2] - x0, v[3] - y0),
			new Point(v[4] - x3, v[5] - y3),
			new Point(x3, y3), epsilon);
	};
	Curve.getLengthIntegrand = function (v) {
		const x0 = v[0], y0 = v[1];
		const x1 = v[2], y1 = v[3];
		const x2 = v[4], y2 = v[5];
		const x3 = v[6], y3 = v[7];

		const ax = 9 * (x1 - x2) + 3 * (x3 - x0);
		const bx = 6 * (x0 + x2) - 12 * x1;
		const cx = 3 * (x1 - x0);

		const ay = 9 * (y1 - y2) + 3 * (y3 - y0);
		const by = 6 * (y0 + y2) - 12 * y1;
		const cy = 3 * (y1 - y0);

		return function (t) {
			const dx = (ax * t + bx) * t + cx;
			const dy = (ay * t + by) * t + cy;
			return Math.sqrt(dx * dx + dy * dy);
		};
	};
	Curve.getIterations = function (a, b) {
		return Math.max(2, Math.min(16, Math.ceil(Math.abs(b - a) * 32)));
	};
	Curve.evaluate = function (v, t, type, normalized) {
		if (t == null || t < 0 || t > 1) { return null; }

		let x0 = v[0];
		let y0 = v[1];
		let x1 = v[2];
		let y1 = v[3];
		let x2 = v[4];
		let y2 = v[5];
		let x3 = v[6];
		let y3 = v[7];
		if (Numerical.isZero(x1 - x0) && Numerical.isZero(y1 - y0)) {
			x1 = x0;
			y1 = y0;
		}
		if (Numerical.isZero(x2 - x3) && Numerical.isZero(y2 - y3)) {
			x2 = x3;
			y2 = y3;
		}
		const cx = 3 * (x1 - x0);
		const bx = 3 * (x2 - x1) - cx;
		const ax = x3 - x0 - cx - bx;
		const cy = 3 * (y1 - y0);
		const by = 3 * (y2 - y1) - cy;
		const ay = y3 - y0 - cy - by;
		let x, y;
		if (type === 0) {
			x = t === 0
				? x0
				: t === 1
					? x3
					: ((ax * t + bx) * t + cx) * t + x0;
			y = t === 0
				? y0
				: t === 1
					? y3
					: ((ay * t + by) * t + cy) * t + y0;
		} else {
			const tMin = 1e-8;
			const tMax = 1 - tMin;
			if (t < tMin) {
				x = cx;
				y = cy;
			} else if (t > tMax) {
				x = 3 * (x3 - x2);
				y = 3 * (y3 - y2);
			} else {
				x = (3 * ax * t + 2 * bx) * t + cx;
				y = (3 * ay * t + 2 * by) * t + cy;
			}
			if (normalized) {
				if (x === 0 && y === 0 && (t < tMin || t > tMax)) {
					x = x2 - x1;
					y = y2 - y1;
				}
				const len = Math.sqrt(x * x + y * y);
				if (len) {
					x /= len;
					y /= len;
				}
			}
			if (type === 3) {
				const _x2 = 6 * ax * t + 2 * bx;
				const _y2 = 6 * ay * t + 2 * by;
				const d = Math.pow(x * x + y * y, 3 / 2);

				x = d !== 0 ? (x * _y2 - y * _x2) / d : 0;
				y = 0;
			}
		}
		return type === 2 ? new Point(y, -x) : new Point(x, y);
	};
	Curve.classify = function (v) {
		const x0 = v[0], y0 = v[1];
		const x1 = v[2], y1 = v[3];
		const x2 = v[4], y2 = v[5];
		const x3 = v[6], y3 = v[7];
		const a1 = x0 * (y3 - y2) + y0 * (x2 - x3) + x3 * y2 - y3 * x2;
		const a2 = x1 * (y0 - y3) + y1 * (x3 - x0) + x0 * y3 - y0 * x3;
		const a3 = x2 * (y1 - y0) + y2 * (x0 - x1) + x1 * y0 - y1 * x0;
		let d3 = 3 * a3;
		let d2 = d3 - a2;
		let d1 = d2 - a2 + a1;
		const l = Math.sqrt(d1 * d1 + d2 * d2 + d3 * d3);
		const s = l !== 0 ? 1 / l : 0;
		d1 *= s;
		d2 *= s;
		d3 *= s;

		function type(type, t1, t2) {
			const hasRoots = t1 !== undefined;
			let t1Ok = hasRoots && t1 > 0 && t1 < 1;
			let t2Ok = hasRoots && t2 > 0 && t2 < 1;
			if (hasRoots && (!(t1Ok || t2Ok) || type === CURVE_TYPES.loop && !(t1Ok && t2Ok))) {
				type = CURVE_TYPES.arc;
				t1Ok = t2Ok = false;
			}
			return {
				type: type,
				roots: t1Ok || t2Ok
					? t1Ok && t2Ok
						? t1 < t2
							? [t1, t2]
							: [t2, t1]
						: [t1Ok ? t1 : t2]
					: null
			};
		}

		if (Numerical.isZero(d1)) {
			return Numerical.isZero(d2)
				? type(Numerical.isZero(d3) ? CURVE_TYPES.line : CURVE_TYPES.quadratic)
				: type(CURVE_TYPES.serpentine, d3 / (3 * d2));
		}

		const d = 3 * d2 * d2 - 4 * d1 * d3;
		if (Numerical.isZero(d)) {
			return type(CURVE_TYPES.cusp, d2 / (2 * d1));
		}

		const f1 = d > 0 ? Math.sqrt(d / 3) : Math.sqrt(-d);
		const f2 = 2 * d1;
		return type(d > 0 ? CURVE_TYPES.serpentine : CURVE_TYPES.loop, (d2 + f1) / f2, (d2 - f1) / f2);
	};
	Curve.getLength = function (v, a, b, ds) {
		if (a === undefined) { a = 0; }
		if (b === undefined) { b = 1; }

		if (Curve.isStraight(v)) {
			let c = v;
			if (b < 1) {
				c = Curve.subdivide(c, b)[0];
				a /= b;
			}
			if (a > 0) {
				c = Curve.subdivide(c, a)[1];
			}
			const dx = c[6] - c[0];
			const dy = c[7] - c[1];
			return Math.sqrt(dx * dx + dy * dy);
		}
		return Numerical.integrate(ds || Curve.getLengthIntegrand(v), a, b, Curve.getIterations(a, b));
	};
	Curve.getTimeAt = function (v, offset, start) {
		if (start === undefined) { start = offset < 0 ? 1 : 0; }
		if (offset === 0) { return start; }

		const forward = offset > 0;
		const a = forward ? start : 0;
		const b = forward ? 1 : start;
		const ds = Curve.getLengthIntegrand(v);
		const rangeLength = Curve.getLength(v, a, b, ds);
		const diff = Math.abs(offset) - rangeLength;

		if (Math.abs(diff) < Numerical.EPSILON) {
			return forward ? b : a;
		} else if (diff > Numerical.EPSILON) {
			return null;
		}

		let length = 0;
		function f(t) {
			length += Numerical.integrate(ds, start, t, Curve.getIterations(start, t));
			start = t;
			return length - offset;
		}

		const guess = offset / rangeLength;
		return Numerical.findRoot(f, ds, start + guess, a, b, 32, Numerical.EPSILON);
	};
	Curve.getPoint = function (v, t) {
		return Curve.evaluate(v, t, 0, false);
	};
	Curve.getTangent = function (v, t) {
		return Curve.evaluate(v, t, 1, true);
	};
	Curve.getNormal = function (v, t) {
		return Curve.evaluate(v, t, 2, true);
	};
	Curve.getWeightedNormal = function (v, t) {
		return Curve.evaluate(v, t, 2, false);
	};
	Curve.getCurvature = function (v, t) {
		return Curve.evaluate(v, t, 3, false).x;
	};
	Curve.getPeaks = function (v) {
		const x0 = v[0], y0 = v[1], x1 = v[2], y1 = v[3], x2 = v[4], y2 = v[5], x3 = v[6], y3 = v[7];
		const ax = -x0 + 3 * x1 - 3 * x2 + x3;
		const bx = 3 * x0 - 6 * x1 + 3 * x2;
		const cx = -3 * x0 + 3 * x1;
		const ay = -y0 + 3 * y1 - 3 * y2 + y3;
		const by = 3 * y0 - 6 * y1 + 3 * y2;
		const cy = -3 * y0 + 3 * y1;

		const tMin = 1e-8;
		const tMax = 1 - tMin;

		const roots = [];
		Numerical.solveCubic(
			9 * (ax * ax + ay * ay),
			9 * (ax * bx + by * ay),
			2 * (bx * bx + by * by) + 3 * (cx * ax + cy * ay),
			(cx * bx + by * cy),
			roots, tMin, tMax
		);
		return roots.sort();
	};
	Curve.addLocation = function (locations, include, c1, t1, c2, t2, overlap) {
		const excludeStart = !overlap && c1.getPrevious() === c2;
		const excludeEnd = !overlap && c1 !== c2 && c1.getNext() === c2;
		const tMin = 1e-8;
		const tMax = 1 - tMin;
		if (t1 !== null && t1 >= (excludeStart ? tMin : 0) && t1 <= (excludeEnd ? tMax : 1)) {
			if (t2 !== null && t2 >= (excludeEnd ? tMin : 0) && t2 <= (excludeStart ? tMax : 1)) {
				const loc1 = new CurveLocation(c1, t1, null, overlap);
				const loc2 = new CurveLocation(c2, t2, null, overlap);
				loc1._intersection = loc2;
				loc2._intersection = loc1;
				if (!include || include(loc1)) {
					CurveLocation.insert(locations, loc1, true);
				}
			}
		}
	};
	Curve.addCurveIntersections = function (v1, v2, c1, c2, locations, include, flip, recursion, calls, tMin, tMax, uMin, uMax) {
		if (++calls >= 4096 || ++recursion >= 40)
			return calls;

		const fatLineEpsilon = 1e-9;
		const q0x = v2[0], q0y = v2[1], q3x = v2[6], q3y = v2[7];

		const d1 = Line.getSignedDistance(q0x, q0y, q3x, q3y, v2[2], v2[3]);
		const d2 = Line.getSignedDistance(q0x, q0y, q3x, q3y, v2[4], v2[5]);

		const factor = d1 * d2 > 0 ? 3 / 4 : 4 / 9;
		const dMin = factor * Math.min(0, d1, d2);
		const dMax = factor * Math.max(0, d1, d2);

		const dp0 = Line.getSignedDistance(q0x, q0y, q3x, q3y, v1[0], v1[1]);
		const dp1 = Line.getSignedDistance(q0x, q0y, q3x, q3y, v1[2], v1[3]);
		const dp2 = Line.getSignedDistance(q0x, q0y, q3x, q3y, v1[4], v1[5]);
		const dp3 = Line.getSignedDistance(q0x, q0y, q3x, q3y, v1[6], v1[7]);

		const hull = Curve.getConvexHull(dp0, dp1, dp2, dp3);
		const top = hull[0];
		const bottom = hull[1];

		let tMinClip;
		let tMaxClip;
		if (d1 === 0 && d2 === 0 && dp0 === 0 && dp1 === 0 && dp2 === 0 && dp3 === 0
			|| (tMinClip = Curve.clipConvexHull(top, bottom, dMin, dMax)) == null
			|| (tMaxClip = Curve.clipConvexHull(top.reverse(), bottom.reverse(), dMin, dMax)) == null) {
			return calls;
		}

		const tMinNew = tMin + (tMax - tMin) * tMinClip;
		const tMaxNew = tMin + (tMax - tMin) * tMaxClip;
		if (Math.max(uMax - uMin, tMaxNew - tMinNew) < fatLineEpsilon) {
			const t = (tMinNew + tMaxNew) / 2;
			const u = (uMin + uMax) / 2;
			Curve.addLocation(locations, include, flip ? c2 : c1, flip ? u : t, flip ? c1 : c2, flip ? t : u);
		} else {
			v1 = Curve.getPart(v1, tMinClip, tMaxClip);
			const uDiff = uMax - uMin;
			if (tMaxClip - tMinClip > 0.8) {
				if (tMaxNew - tMinNew > uDiff) {
					const parts = Curve.subdivide(v1, 0.5);
					const t = (tMinNew + tMaxNew) / 2;
					calls = Curve.addCurveIntersections(
						v2, parts[0], c2, c1, locations, include, !flip,
						recursion, calls, uMin, uMax, tMinNew, t);
					calls = Curve.addCurveIntersections(
						v2, parts[1], c2, c1, locations, include, !flip,
						recursion, calls, uMin, uMax, t, tMaxNew);
				} else {
					const parts = Curve.subdivide(v2, 0.5);
					const u = (uMin + uMax) / 2;
					calls = Curve.addCurveIntersections(
						parts[0], v1, c2, c1, locations, include, !flip,
						recursion, calls, uMin, u, tMinNew, tMaxNew);
					calls = Curve.addCurveIntersections(
						parts[1], v1, c2, c1, locations, include, !flip,
						recursion, calls, u, uMax, tMinNew, tMaxNew);
				}
			} else {
				if (uDiff === 0 || uDiff >= fatLineEpsilon) {
					calls = Curve.addCurveIntersections(
						v2, v1, c2, c1, locations, include, !flip,
						recursion, calls, uMin, uMax, tMinNew, tMaxNew);
				} else {
					calls = Curve.addCurveIntersections(
						v1, v2, c1, c2, locations, include, flip,
						recursion, calls, tMinNew, tMaxNew, uMin, uMax);
				}
			}
		}
		return calls;
	};
	Curve.getConvexHull = function (dq0, dq1, dq2, dq3) {
		const p0 = [0, dq0];
		const p1 = [1 / 3, dq1];
		const p2 = [2 / 3, dq2];
		const p3 = [1, dq3];
		const dist1 = dq1 - (2 * dq0 + dq3) / 3;
		const dist2 = dq2 - (dq0 + 2 * dq3) / 3;

		let hull;
		if (dist1 * dist2 < 0) {
			hull = [[p0, p1, p3], [p0, p2, p3]];
		} else {
			const distRatio = dist1 / dist2;
			hull = [
				distRatio >= 2
					? [p0, p1, p3]
					: distRatio <= 0.5
						? [p0, p2, p3]
						: [p0, p1, p2, p3],
				[p0, p3]
			];
		}
		return (dist1 || dist2) < 0 ? hull.reverse() : hull;
	};
	Curve.clipConvexHull = function (hullTop, hullBottom, dMin, dMax) {
		if (hullTop[0][1] < dMin) {
			return Curve.clipConvexHullPart(hullTop, true, dMin);
		} else if (hullBottom[0][1] > dMax) {
			return Curve.clipConvexHullPart(hullBottom, false, dMax);
		} else {
			return hullTop[0][0];
		}
	};
	Curve.clipConvexHullPart = function (part, top, threshold) {
		let px = part[0][0];
		let py = part[0][1];
		for (let i = 1, l = part.length; i < l; i++) {
			const qx = part[i][0];
			const qy = part[i][1];
			if (top ? qy >= threshold : qy <= threshold) {
				return qy === threshold ? qx : px + (threshold - py) * (qx - px) / (qy - py);
			}
			px = qx;
			py = qy;
		}
		return null;
	};
	Curve.getCurveLineIntersections = function (v, px, py, vx, vy) {
		if (Numerical.isZero(vx) && Numerical.isZero(vy)) {
			const t = Curve.getTimeOf(v, new Point(px, py));
			return t === null ? [] : [t];
		}

		const angle = Math.atan2(-vy, vx);
		const sin = Math.sin(angle);
		const cos = Math.cos(angle);
		const rv = [];
		const roots = [];
		for (let i = 0; i < 8; i += 2) {
			const x = v[i] - px;
			const y = v[i + 1] - py;
			rv.push(x * cos - y * sin, x * sin + y * cos);
		}

		Curve.solveCubic(rv, 1, 0, roots, 0, 1);
		return roots;
	};
	Curve.addCurveLineIntersections = function (v1, v2, c1, c2, locations, include, flip) {
		const x1 = v2[0];
		const y1 = v2[1];
		const x2 = v2[6];
		const y2 = v2[7];
		const roots = Curve.getCurveLineIntersections(v1, x1, y1, x2 - x1, y2 - y1);
		for (let i = 0, l = roots.length; i < l; i++) {
			const t1 = roots[i];
			const p1 = Curve.getPoint(v1, t1);
			const t2 = Curve.getTimeOf(v2, p1);
			if (t2 !== null) {
				flip
					? Curve.addLocation(locations, include, c2, t2, c1, t1)
					: Curve.addLocation(locations, include, c1, t1, c2, t2);
			}
		}
	};
	Curve.addLineIntersection = function (v1, v2, c1, c2, locations, include) {
		const pt = Line.intersect(v1[0], v1[1], v1[6], v1[7], v2[0], v2[1], v2[6], v2[7]);
		if (pt) {
			Curve.addLocation(locations, include,
				c1, Curve.getTimeOf(v1, pt),
				c2, Curve.getTimeOf(v2, pt)
			);
		}
	};
	Curve.getCurveIntersections = function (v1, v2, c1, c2, locations, include) {
		if (Math.max(v1[0], v1[2], v1[4], v1[6]) + Numerical.EPSILON > Math.min(v2[0], v2[2], v2[4], v2[6])
			&& Math.min(v1[0], v1[2], v1[4], v1[6]) - Numerical.EPSILON < Math.max(v2[0], v2[2], v2[4], v2[6])
			&& Math.max(v1[1], v1[3], v1[5], v1[7]) + Numerical.EPSILON > Math.min(v2[1], v2[3], v2[5], v2[7])
			&& Math.min(v1[1], v1[3], v1[5], v1[7]) - Numerical.EPSILON < Math.max(v2[1], v2[3], v2[5], v2[7])) {
			const overlaps = Curve.getOverlaps(v1, v2);
			if (overlaps) {
				for (let i = 0; i < 2; i++) {
					const overlap = overlaps[i];
					Curve.addLocation(locations, include, c1, overlap[0], c2, overlap[1], true);
				}
			} else {
				const straight1 = Curve.isStraight(v1);
				const straight2 = Curve.isStraight(v2);
				const straight = straight1 && straight2;
				const flip = straight1 && !straight2;
				const before = locations.length;

				const addIntersectionsFunction = straight
					? Curve.addLineIntersection
					: straight1 || straight2
						? Curve.addCurveLineIntersections
						: Curve.addCurveIntersections

				flip
					? addIntersectionsFunction(v2, v1, c2, c1, locations, include, flip, 0, 0, 0, 1, 0, 1)
					: addIntersectionsFunction(v1, v2, c1, c2, locations, include, flip, 0, 0, 0, 1, 0, 1);

				if (!straight || locations.length === before) {
					for (let i = 0; i < 4; i++) {
						const t1 = i >> 1;
						const t2 = i & 1;
						const i1 = t1 * 6;
						const i2 = t2 * 6;
						const p1 = new Point(v1[i1], v1[i1 + 1]);
						const p2 = new Point(v2[i2], v2[i2 + 1]);
						if (p1.isClose(p2, Numerical.EPSILON)) {
							Curve.addLocation(locations, include, c1, t1, c2, t2);
						}
					}
				}
			}
		}
		return locations;
	};
	Curve.getSelfIntersection = function (v1, c1, locations, include) {
		const info = Curve.classify(v1);
		if (info.type === CURVE_TYPES.loop) {
			const roots = info.roots;
			Curve.addLocation(locations, include, c1, roots[0], c1, roots[1]);
		}
		return locations;
	};
	Curve.getIntersections = function (curves1, curves2, include, matrix1, matrix2, _returnFirst) {
		const epsilon = 1e-7;
		const self = !curves2;
		if (self)
			curves2 = curves1;

		const values1 = new Array(curves1.length);
		const values2 = self ? values1 : new Array(curves2.length);
		const locations = [];
		for (let i = 0; i < curves1.length; i++) {
			values1[i] = curves1[i].getValues(matrix1);
		}
		if (!self) {
			for (let i = 0; i < curves2.length; i++) {
				values2[i] = curves2[i].getValues(matrix2);
			}
		}

		const boundsCollisions = CollisionDetection.findCurveBoundsCollisions(values1, values2, epsilon);
		for (let index1 = 0; index1 < curves1.length; index1++) {
			const curve1 = curves1[index1];
			const v1 = values1[index1];
			if (self) { Curve.getSelfIntersection(v1, curve1, locations, include); }

			const collisions1 = boundsCollisions[index1];
			if (collisions1) {
				for (let j = 0; j < collisions1.length; j++) {
					if (_returnFirst && locations.length)
						return locations;

					const index2 = collisions1[j];
					if (!self || index2 > index1) {
						const curve2 = curves2[index2];
						const v2 = values2[index2];
						Curve.getCurveIntersections(v1, v2, curve1, curve2, locations, include);
					}
				}
			}
		}
		return locations;
	};
	Curve.getOverlaps = function (v1, v2) {

		function getSquaredLineLength(v) {
			const x = v[6] - v[0];
			const y = v[7] - v[1];
			return x * x + y * y;
		}

		const timeEpsilon = 1e-8;
		const geomEpsilon = 1e-7;
		let straight1 = Curve.isStraight(v1);
		let straight2 = Curve.isStraight(v2);
		let straightBoth = straight1 && straight2;
		const flip = getSquaredLineLength(v1) < getSquaredLineLength(v2);
		const l1 = flip ? v2 : v1;
		const l2 = flip ? v1 : v2;
		const px = l1[0], py = l1[1];
		const vx = l1[6] - px, vy = l1[7] - py;
		if (Line.getDistance(px, py, vx, vy, l2[0], l2[1], true) < geomEpsilon
			&& Line.getDistance(px, py, vx, vy, l2[6], l2[7], true) < geomEpsilon) {
			if (!straightBoth
				&& Line.getDistance(px, py, vx, vy, l1[2], l1[3], true) < geomEpsilon
				&& Line.getDistance(px, py, vx, vy, l1[4], l1[5], true) < geomEpsilon
				&& Line.getDistance(px, py, vx, vy, l2[2], l2[3], true) < geomEpsilon
				&& Line.getDistance(px, py, vx, vy, l2[4], l2[5], true) < geomEpsilon) {
				straight1 = straight2 = straightBoth = true;
			}
		} else if (straightBoth) {
			return null;
		}
		if (straight1 ^ straight2) {
			return null;
		}

		const v = [v1, v2];
		let pairs = [];
		for (let i = 0; i < 4 && pairs.length < 2; i++) {
			const i1 = i & 1;
			const i2 = i1 ^ 1;
			const t1 = i >> 1;
			const t2 = Curve.getTimeOf(v[i1], new Point(v[i2][t1 ? 6 : 0], v[i2][t1 ? 7 : 1]));

			if (t2 != null) {
				const pair = i1 ? [t1, t2] : [t2, t1];
				if (!pairs.length || Math.abs(pair[0] - pairs[0][0]) > timeEpsilon && Math.abs(pair[1] - pairs[0][1]) > timeEpsilon) {
					pairs.push(pair);
				}
			}
			if (i > 2 && !pairs.length)
				break;
		}
		if (pairs.length !== 2) {
			pairs = null;
		} else if (!straightBoth) {
			const o1 = Curve.getPart(v1, pairs[0][0], pairs[1][0]);
			const o2 = Curve.getPart(v2, pairs[0][1], pairs[1][1]);
			if (Math.abs(o2[2] - o1[2]) > geomEpsilon
				|| Math.abs(o2[3] - o1[3]) > geomEpsilon
				|| Math.abs(o2[4] - o1[4]) > geomEpsilon
				|| Math.abs(o2[5] - o1[5]) > geomEpsilon)
				pairs = null;
		}
		return pairs;
	};
	Curve.getTimesWithTangent = function (v, tangent) {
		const x0 = v[0], y0 = v[1], x1 = v[2], y1 = v[3], x2 = v[4], y2 = v[5], x3 = v[6], y3 = v[7];

		const normalized = tangent.normalize();
		const tx = normalized.x;
		const ty = normalized.y;

		const ax = 3 * x3 - 9 * x2 + 9 * x1 - 3 * x0;
		const ay = 3 * y3 - 9 * y2 + 9 * y1 - 3 * y0;
		const bx = 6 * x2 - 12 * x1 + 6 * x0;
		const by = 6 * y2 - 12 * y1 + 6 * y0;
		const cx = 3 * x1 - 3 * x0;
		const cy = 3 * y1 - 3 * y0;

		let den = 2 * ax * ty - 2 * ay * tx;
		const times = [];
		if (Math.abs(den) < Numerical.CURVETIME_EPSILON) {
			const num = ax * cy - ay * cx;
			den = ax * by - ay * bx;
			if (den != 0) {
				const t = -num / den;
				if (t >= 0 && t <= 1) times.push(t);
			}
		} else {
			const delta = (bx * bx - 4 * ax * cx) * ty * ty +
				(-2 * bx * by + 4 * ay * cx + 4 * ax * cy) * tx * ty +
				(by * by - 4 * ay * cy) * tx * tx;
			const k = bx * ty - by * tx;

			if (delta >= 0 && den != 0) {
				const d = Math.sqrt(delta);
				const t0 = -(k + d) / den;
				const t1 = (-k + d) / den;
				if (t0 >= 0 && t0 <= 1) times.push(t0);
				if (t1 >= 0 && t1 <= 1) times.push(t1);
			}
		}
		return times;
	};

	const CurveLocation = function (curve, time, point, _overlap, _distance) {
		if (time >= 0.99999999) {
			const next = curve.getNext();
			if (next) {
				time = 0;
				curve = next;
			}
		}
		this._setCurve(curve);
		this._time = time;
		this._point = point || curve.getPointAtTime(time);
		this._overlap = _overlap;
		this._distance = _distance;
		this._intersection = this._next = this._previous = null;
	};
	InitClassWithStatics(CurveLocation, Base);

	CurveLocation.prototype._setPath = function (path) {
		this._path = path;
		this._version = path ? path._version : 0;
	};
	CurveLocation.prototype._setCurve = function (curve) {
		this._setPath(curve._path);
		this._curve = curve;
		this._segment = null;
		this._segment1 = curve._segment1;
		this._segment2 = curve._segment2;
	};
	CurveLocation.prototype._setSegment = function (segment) {
		const curve = segment.getCurve();
		if (curve) {
			this._setCurve(curve);
		} else {
			this._setPath(segment._path);
			this._segment1 = segment;
			this._segment2 = null;
		}
		this._segment = segment;
		this._time = segment === this._segment1 ? 0 : 1;
		this._point = segment._point.clone();
	};
	CurveLocation.prototype.getSegment = function () {
		let segment = this._segment;
		if (!segment) {
			const curve = this.getCurve();
			const time = this.getTime();
			if (time === 0) {
				segment = curve._segment1;
			} else if (time === 1) {
				segment = curve._segment2;
			} else if (time != null) {
				segment = curve.getPartLength(0, time) < curve.getPartLength(time, 1)
					? curve._segment1
					: curve._segment2;
			}
			this._segment = segment;
		}
		return segment;
	};
	CurveLocation.prototype.getCurve = function () {
		if (this._path && this._path._version !== this._version) {
			this._time = this._offset = this._curveOffset = this._curve = null;
		}

		const that = this;
		function trySegment(segment) {
			const curve = segment && segment.getCurve();
			if (curve && (that._time = curve.getTimeOf(that._point)) != null) {
				that._setCurve(curve);
				return curve;
			}
		}

		return this._curve
			|| trySegment(this._segment)
			|| trySegment(this._segment1)
			|| trySegment(this._segment2.getPrevious());
	};
	CurveLocation.prototype.getPath = function () {
		const curve = this.getCurve();
		return curve && curve._path;
	};
	CurveLocation.prototype.getIndex = function () {
		const curve = this.getCurve();
		return curve && curve.getIndex();
	};
	CurveLocation.prototype.getTime = function () {
		const curve = this.getCurve();
		return curve && this._time == null
			? this._time = curve.getTimeOf(this._point)
			: this._time;
	};
	CurveLocation.prototype.getPoint = function () {
		return this._point;
	};
	CurveLocation.prototype.getOffset = function () {
		let offset = this._offset;
		if (offset == null) {
			offset = 0;
			const path = this.getPath();
			const index = this.getIndex();
			if (path && index != null) {
				const curves = path.getCurves();
				for (let i = 0; i < index; i++)
					offset += curves[i].getLength();
			}
			this._offset = offset += this.getCurveOffset();
		}
		return offset;
	};
	CurveLocation.prototype.getCurveOffset = function () {
		let offset = this._curveOffset;
		if (offset == null) {
			const curve = this.getCurve();
			const time = this.getTime();
			this._curveOffset = offset = time != null && curve && curve.getPartLength(0, time);
		}
		return offset;
	};
	CurveLocation.prototype.getIntersection = function () {
		return this._intersection;
	};
	CurveLocation.prototype.getDistance = function () {
		return this._distance;
	};
	CurveLocation.prototype.divide = function () {
		const curve = this.getCurve();
		const res = curve && curve.divideAtTime(this.getTime());
		if (res) { this._setSegment(res._segment1); }
		return res;
	};
	CurveLocation.prototype.equals = function (loc, _ignoreOther) {
		if (this === loc) return true;
		if (!(loc instanceof CurveLocation)) return false;

		const curve1 = this.getCurve();
		const curve2 = loc.getCurve();
		const samePath = curve1._path === curve2._path;
		if (!samePath) return false;

		const offsetDifference = Math.abs(this.getOffset() - loc.getOffset());
		const closeOffsets = (
			offsetDifference < Numerical.GEOMETRIC_EPSILON ||
			(curve1._path && Math.abs(curve1._path.getLength() - offsetDifference) < Numerical.GEOMETRIC_EPSILON)
		);

		const intersection1 = !_ignoreOther && this._intersection;
		const intersection2 = !_ignoreOther && loc._intersection;

		const matchingIntersections = !intersection1 && !intersection2 || (intersection1 && intersection2 && intersection1.equals(intersection2, true));
		return closeOffsets && matchingIntersections;
	};
	CurveLocation.prototype.isTouching = function () {
		if (this._intersection && this.getTangent().isCollinear(this._intersection.getTangent())) {
			const curve1 = this.getCurve();
			const curve2 = this._intersection.getCurve();
			return !(curve1.isStraight() && curve2.isStraight() && curve1.getLine().intersect(curve2.getLine()));
		}
		return false;
	};
	CurveLocation.prototype.isCrossing = function () {
		if (!this._intersection) { return false; }

		const t1 = this.getTime();
		const t2 = this._intersection.getTime();
		const tMin = 1e-8;
		const tMax = 1 - tMin;
		const t1Inside = t1 >= tMin && t1 <= tMax;
		const t2Inside = t2 >= tMin && t2 <= tMax;
		if (t1Inside && t2Inside)
			return !this.isTouching();

		let c2 = this.getCurve();
		let c1 = c2 && t1 < tMin ? c2.getPrevious() : c2;
		let c4 = this._intersection.getCurve();
		let c3 = c4 && t2 < tMin ? c4.getPrevious() : c4;
		if (t1 > tMax)
			c2 = c2.getNext();
		if (t2 > tMax)
			c4 = c4.getNext();
		if (!c1 || !c2 || !c3 || !c4)
			return false;

		const offsets = [];

		function addOffsets(curve, end) {
			const v = curve.getValues();
			const roots = Curve.classify(v).roots || Curve.getPeaks(v);
			const count = roots.length;
			const offset = Curve.getLength(v, end && count ? roots[count - 1] : 0, !end && count ? roots[0] : 1);
			offsets.push(count ? offset : offset / 32);
		}

		function isInRange(angle, min, max) {
			return min < max
				? angle > min && angle < max
				: angle > min || angle < max;
		}

		if (!t1Inside) {
			addOffsets(c1, true);
			addOffsets(c2, false);
		}
		if (!t2Inside) {
			addOffsets(c3, true);
			addOffsets(c4, false);
		}
		const pt = this.getPoint();
		const offset = Math.min.apply(Math, offsets);

		const v2 = t1Inside ? c2.getTangentAtTime(t1) : c2.getPointAt(offset).subtract(pt);
		const v1 = t1Inside ? v2.negate() : c1.getPointAt(-offset).subtract(pt);
		const v4 = t2Inside ? c4.getTangentAtTime(t2) : c4.getPointAt(offset).subtract(pt);
		const v3 = t2Inside ? v4.negate() : c3.getPointAt(-offset).subtract(pt);

		const a1 = v1.getAngle();
		const a2 = v2.getAngle();
		const a3 = v3.getAngle();
		const a4 = v4.getAngle();

		return !!(t1Inside
			? (isInRange(a1, a3, a4) ^ isInRange(a2, a3, a4)) && (isInRange(a1, a4, a3) ^ isInRange(a2, a4, a3))
			: (isInRange(a3, a1, a2) ^ isInRange(a4, a1, a2)) && (isInRange(a3, a2, a1) ^ isInRange(a4, a2, a1)));
	};
	CurveLocation.prototype.hasOverlap = function () {
		return !!this._overlap;
	};
	CurveLocation.prototype.getTangent = function () {
		const curve = this.getCurve();
		const time = this.getTime();
		return time != null && curve && curve.getTangentAt(time, true);
	};
	CurveLocation.prototype.getNormal = function () {
		const curve = this.getCurve();
		const time = this.getTime();
		return time != null && curve && curve.getNormalAt(time, true);
	};
	CurveLocation.prototype.getWeightedTangent = function () {
		const curve = this.getCurve();
		const time = this.getTime();
		return time != null && curve && curve.getWeightedTangentAt(time);
	};
	CurveLocation.prototype.getWeightedNormal = function () {
		const curve = this.getCurve();
		const time = this.getTime();
		return time != null && curve && curve.getWeightedNormalAt(time);
	};
	CurveLocation.prototype.getCurvature = function () {
		const curve = this.getCurve();
		const time = this.getTime();
		return time != null && curve && curve.getCurvatureAt(time);
	};

	CurveLocation.insert = function (locations, loc, merge) {
		const length = locations.length;

		function search(index, dir) {
			for (let i = index + dir; i >= -1 && i <= length; i += dir) {
				const loc2 = locations[((i % length) + length) % length];
				if (!loc.getPoint().isClose(loc2.getPoint(), 1e-7))
					break;
				if (loc.equals(loc2))
					return loc2;
			}
			return null;
		}

		let l = 0;
		let r = length - 1;

		while (l <= r) {
			const m = (l + r) >>> 1;
			const loc2 = locations[m];
			let found;
			if (merge && (found = loc.equals(loc2) ? loc2 : (search(m, -1) || search(m, 1)))) {
				if (loc._overlap) {
					found._overlap = found._intersection._overlap = true;
				}
				return found;
			}
			const path1 = loc.getPath();
			const path2 = loc2.getPath();
			const diff = path1 !== path2
				? path1._id - path2._id
				: (loc.getIndex() + loc.getTime()) - (loc2.getIndex() + loc2.getTime());
			diff < 0 ? (r = m - 1) : (l = m + 1);
		}
		locations.splice(l, 0, loc);
		return loc;
	};
	CurveLocation.expand = function (locations) {
		const expanded = locations.slice();
		for (let i = locations.length - 1; i >= 0; i--) {
			CurveLocation.insert(expanded, locations[i]._intersection, false);
		}
		return expanded;
	};

	const PathItem = function PathItem() { };
	InitClassWithStatics(PathItem, Item);

	PathItem.prototype.isClockwise = function () {
		return this.getArea() >= 0;
	};
	PathItem.prototype.setClockwise = function (clockwise) {
		if (this.isClockwise() != (clockwise = !!clockwise))
			this.reverse();
	};
	PathItem.prototype._contains = function (point) {
		const winding = point.isInside(this.getBounds({ internal: true, handle: true }))
			? this._getWinding(point)
			: {};
		return winding.onPath || !!(winding.winding);
	};
	PathItem.prototype.getIntersections = function (path, include, _matrix, _returnFirst) {
		const self = this === path || !path;
		const matrix1 = this._matrix._orNullIfIdentity();
		const matrix2 = self ? matrix1 : (_matrix || path._matrix)._orNullIfIdentity();
		return self || this.getBounds(matrix1).intersects(path.getBounds(matrix2), 1e-12)
			? Curve.getIntersections(this.getCurves(), !self && path.getCurves(), include, matrix1, matrix2, _returnFirst)
			: [];
	};
	PathItem.prototype.getNearestLocation = function () {
		const point = Point.read(arguments);
		const curves = this.getCurves();
		let minDist = Infinity;
		let minLoc = null;
		for (let i = 0, l = curves.length; i < l; i++) {
			const loc = curves[i].getNearestLocation(point);
			if (loc._distance < minDist) {
				minDist = loc._distance;
				minLoc = loc;
			}
		}
		return minLoc;
	};
	PathItem.prototype.getNearestPoint = function () {
		const loc = this.getNearestLocation.apply(this, arguments);
		return loc ? loc.getPoint() : loc;
	};
	PathItem.prototype.compare = function (path) {
		if (!path) return false;

		const paths1 = this._children || [this];
		const paths2 = path._children ? path._children.slice() : [path];
		const length1 = paths1.length;
		const length2 = paths2.length;

		const boundsOverlaps = CollisionDetection.findItemBoundsCollisions(paths1, paths2, Numerical.GEOMETRIC_EPSILON);

		let matched = Array(length2).fill(false);
		let matchCount = 0;
		let allMatched = true;

		for (let i1 = length1 - 1; i1 >= 0 && allMatched; i1--) {
			const path1 = paths1[i1];
			const pathBoundsOverlaps = boundsOverlaps[i1];
			let pathMatched = false;

			if (pathBoundsOverlaps) {
				for (let i2 = pathBoundsOverlaps.length - 1; i2 >= 0 && !pathMatched; i2--) {
					const pathIndex = pathBoundsOverlaps[i2];
					if (path1.compare(paths2[pathIndex])) {
						if (!matched[pathIndex]) {
							matched[pathIndex] = true;
							matchCount++;
						}
						pathMatched = true;
					}
				}
			}
			if (!pathMatched) allMatched = false;
		}
		return allMatched && matchCount === length2;
	};


	PathItem.prototype._getWinding = function (point, dir, closed) {
		return PathItem.getWinding(point, this.getCurves(), dir, closed);
	};
	PathItem.prototype.unite = function (path) {
		return PathItem.traceBoolean(this, path, OPERATIONS.unite);
	};
	PathItem.prototype.intersect = function (path) {
		return PathItem.traceBoolean(this, path, OPERATIONS.intersect);
	};
	PathItem.prototype.subtract = function (path) {
		return PathItem.traceBoolean(this, path, OPERATIONS.subtract);
	};
	PathItem.prototype.exclude = function (path) {
		return PathItem.traceBoolean(this, path, OPERATIONS.exclude);
	};
	PathItem.prototype.divide = function (argument) {
		// Original version with only two paths
		if (!Array.isArray(argument)) {
			const path = argument;
			return PathItem.createResult([
				this.exclude(path),
				this.intersect(path)
			], true, this, path);
		}

		// Version for multiple paths
		const paths = argument;

		function calculateUniversumBounds(paths) {
			const pathBounds = paths.map(function (path) { return path.getBounds(); });
			const left = Math.min.apply(null, pathBounds.map(function (bounds) { return bounds.getLeft(); }));
			const top = Math.min.apply(null, pathBounds.map(function (bounds) { return bounds.getTop(); }));
			const right = Math.max.apply(null, pathBounds.map(function (bounds) { return bounds.getLeft() + bounds.getWidth(); }));
			const bottom = Math.max.apply(null, pathBounds.map(function (bounds) { return bounds.getTop() + bounds.getHeight(); }));
			return [left, top, right, bottom];
		}
		const bounds = calculateUniversumBounds(paths);
		const delta = 1; // Expand universum so that it accurately includes all paths
		const universum = new Path.Rectangle(bounds[0] - delta, bounds[1] - delta, bounds[2] + delta, bounds[3] + delta);

		const areas = [];
		for (let option = 1, totalCombinations = Math.pow(2, paths.length); option < totalCombinations; option++) {
			let result = universum;
			for (let i = 0; i < paths.length; i++) {
				const path = paths[i];
				const isBitSet = (option & (1 << i)) !== 0;
				if (isBitSet) {
					result = result.intersect(path);
				} else {
					result = result.intersect(universum.subtract(path));
				}
			}
			if (!result.isEmpty()) {
				result._option = option;
				areas.push(result);
			}
		}

		function splitCompoundPath(compoundPath) {
			const split = [];
			const paths = compoundPath.getChildren().slice();
			const hasIntersection = function (p1, p2) { return !p1.intersect(p2).isEmpty(); };
			paths.forEach(function (path) {
				const intersects = paths.some(function (p1) {
					return p1 !== path && hasIntersection(p1, path);
				});
				if (!intersects) {
					split.push(path);
					path.remove();
				}
			});

			if (compoundPath.getChildren().length) {
				split.push(compoundPath);
			}
			return split;
		}

		const fragments = areas.flatMap(function (area) {
			if (area instanceof Path) { return [area]; }
			if (area instanceof CompoundPath) { return splitCompoundPath(area); }
		});

		return fragments;
	};
	PathItem.prototype.resolveCrossings = function () {
		let paths = this._children || [this];
		let hasOverlaps = false;
		let hasCrossings = false;

		function hasOverlap(seg, path) {
			const inter = seg && seg._intersection;
			return inter && inter._overlap && inter._path === path;
		}

		let intersections = this.getIntersections(null, function (inter) {
			return (inter.hasOverlap() && (hasOverlaps = true)) ||
				(inter.isCrossing() && (hasCrossings = true));
		});

		const clearCurves = hasOverlaps && hasCrossings ? [] : null;
		intersections = CurveLocation.expand(intersections);

		if (hasOverlaps) {
			const overlaps = PathItem.divideLocations(intersections, function (inter) {
				return inter.hasOverlap();
			}, clearCurves);

			for (let i = overlaps.length - 1; i >= 0; i--) {
				const overlap = overlaps[i];
				const path = overlap._path;
				const seg = overlap._segment;
				const prev = seg.getPrevious();
				const next = seg.getNext();

				if (hasOverlap(prev, path) && hasOverlap(next, path)) {
					seg.remove();
					prev._handleOut._set(0, 0);
					next._handleIn._set(0, 0);
					if (prev !== seg && !prev.getCurve().hasLength()) {
						next._handleIn.set(prev._handleIn);
						prev.remove();
					}
				}
			}
		}
		if (hasCrossings) {
			PathItem.divideLocations(intersections, hasOverlaps && function (inter) {
				const curve1 = inter.getCurve();
				const seg1 = inter.getSegment();
				const other = inter._intersection;
				const curve2 = other._curve;
				const seg2 = other._segment;
				if (curve1 && curve2 && curve1._path && curve2._path) { return true; }
				if (seg1) { seg1._intersection = null; }
				if (seg2) { seg2._intersection = null; }
			}, clearCurves);

			if (clearCurves) PathItem.clearCurveHandles(clearCurves);

			paths = PathItem.tracePaths(Base.each(paths, function (path) {
				this.push.apply(this, path._segments);
			}, []));
		}

		let item;
		const length = paths.length;
		if (length > 1 && this._children) {
			if (paths !== this._children) { this.setChildren(paths); }
			item = this;
		} else if (length === 1 && !this._children) {
			if (paths[0] !== this) { this.setSegments(paths[0].removeSegments()); }
			item = this;
		}

		if (!item) {
			item = new CompoundPath({ insert: false });
			item.addChildren(paths);
			item = item.reduce();
			item.copyAttributes(this);
			this.replaceWith(item);
		}
		return item;
	};
	PathItem.prototype.reorient = function (nonZero, clockwise) {
		if (this._children && this._children.length) {
			const reorientedChildren = PathItem.reorientPaths(
				this.removeChildren(),
				function (w) {
					return !!(nonZero ? w : w & 1);
				},
				clockwise
			);
			this.setChildren(reorientedChildren);
		} else if (clockwise !== undefined) {
			this.setClockwise(clockwise);
		}
		return this;
	};
	PathItem.prototype.getInteriorPoint = function () {
		const bounds = this.getBounds();
		const point = bounds.getCenter(true);
		if (this.contains(point)) return point;

		const curves = this.getCurves();
		const y = point.y;
		const intercepts = [];
		const roots = [];

		curves.forEach(function (curve) {
			const v = curve.getValues();
			const [o0, o1, o2, o3] = [v[1], v[3], v[5], v[7]];

			if (y >= Math.min(o0, o1, o2, o3) && y <= Math.max(o0, o1, o2, o3)) {
				const monoCurves = Curve.getMonoCurves(v);

				monoCurves.forEach(function (mv) {
					const mo0 = mv[1];
					const mo3 = mv[7];

					if (mo0 !== mo3 && (y >= Math.min(mo0, mo3) && y <= Math.max(mo0, mo3))) {
						const x = y === mo0
							? mv[0]
							: y === mo3
								? mv[6]
								: Curve.solveCubic(mv, 1, y, roots, 0, 1) === 1
									? Curve.getPoint(mv, roots[0]).x
									: (mv[0] + mv[6]) / 2;

						intercepts.push(x);
					}
				});
			}
		});

		if (intercepts.length > 1) {
			intercepts.sort(function (a, b) { return a - b; });
			point.x = (intercepts[0] + intercepts[1]) / 2;
		}
		return point;
	};

	PathItem.getPaths = function (path) {
		return path._children || [path];
	};
	PathItem.preparePath = function (path) {
		let res = path
			.clone(false)
			.reduce({ simplify: true })
			.transform(null, true, true);

		const paths = PathItem.getPaths(res);
		for (let i = 0, l = paths.length; i < l; i++) {
			const path = paths[i];
			if (!path._closed && !path.isEmpty()) {
				path.closePath(Numerical.EPSILON);
				path.getFirstSegment().setHandleIn(0, 0);
				path.getLastSegment().setHandleOut(0, 0);
			}
		}
		return res.resolveCrossings().reorient(true, true);
	};
	PathItem.createResult = function (paths, simplify, path1, path2) {
		let result = new CompoundPath({ insert: false });
		result.addChildren(paths, true);
		result = result.reduce({ simplify: simplify });
		result.copyAttributes(path1, true);
		return result;
	};
	PathItem.filterIntersection = function (inter) {
		return inter.hasOverlap() || inter.isCrossing();
	};
	PathItem.traceBoolean = function (path1, path2, operation) {
		const operators = {
			'1': { '1': true, '2': true, 'unite': true },
			'2': { '2': true, 'intersect': true },
			'3': { '1': true, 'subtract': true },
			'4': { '1': true, '-1': true, 'exclude': true },
		};
		const operator = operators[operation];

		const _path1 = PathItem.preparePath(path1);
		const _path2 = path2 && path1 !== path2
			? PathItem.preparePath(path2)
			: null;

		if (_path2 && (operator['subtract'] || operator['exclude']) ^ (_path2.isClockwise() ^ _path1.isClockwise())) {
			_path2.reverse();
		}

		const crossings = PathItem.divideLocations(
			CurveLocation.expand(_path1.getIntersections(_path2, PathItem.filterIntersection))
		);
		const paths1 = PathItem.getPaths(_path1);
		const paths2 = _path2 ? PathItem.getPaths(_path2) : null;
		const segments = [];
		const curves = [];

		function collectPaths(paths) {
			for (let i = 0, l = paths.length; i < l; i++) {
				const path = paths[i];
				segments.push.apply(segments, path._segments);
				curves.push.apply(curves, path.getCurves());
				path._overlapsOnly = true;
			}
		}

		function getCurves(indices) {
			const list = [];
			for (let i = 0, l = indices && indices.length; i < l; i++) {
				list.push(curves[indices[i]]);
			}
			return list;
		}

		let paths;
		if (crossings.length) {
			collectPaths(paths1);
			if (paths2) collectPaths(paths2);

			const curvesValues = new Array(curves.length);
			for (let i = 0, l = curves.length; i < l; i++) {
				curvesValues[i] = curves[i].getValues();
			}

			const curveCollisions = CollisionDetection.findCurveBoundsCollisions(curvesValues, curvesValues, 0, true);
			const curveCollisionsMap = {};

			for (let i = 0; i < curves.length; i++) {
				const curve = curves[i];
				const id = curve._path._id;
				const map = curveCollisionsMap[id] = curveCollisionsMap[id] || {};
				map[curve.getIndex()] = {
					hor: getCurves(curveCollisions[i].hor),
					ver: getCurves(curveCollisions[i].ver)
				};
			}
			for (let i = 0, l = crossings.length; i < l; i++) {
				PathItem.propagateWinding(crossings[i]._segment, _path1, _path2, curveCollisionsMap, operator);
			}
			for (let i = 0, l = segments.length; i < l; i++) {
				const segment = segments[i];
				const inter = segment._intersection;

				if (!segment._winding) {
					PathItem.propagateWinding(segment, _path1, _path2, curveCollisionsMap, operator);
				}
				if (!(inter && inter._overlap)) {
					segment._path._overlapsOnly = false;
				}
			}

			paths = PathItem.tracePaths(segments, operator);
		} else {
			paths = PathItem.reorientPaths(
				paths2 ? paths1.concat(paths2) : paths1.slice(),
				function (w) {
					return !!operator[w];
				}
			);
		}

		return PathItem.createResult(paths, true, path1, path2);
	};
	PathItem.linkIntersections = function (from, to) {
		let prev = from;
		while (prev) {
			if (prev === to)
				return;
			prev = prev._previous;
		}
		while (from._next && from._next !== to)
			from = from._next;
		if (!from._next) {
			while (to._previous)
				to = to._previous;
			from._next = to;
			to._previous = from;
		}
	};
	PathItem.clearCurveHandles = function (curves) {
		for (let i = curves.length - 1; i >= 0; i--)
			curves[i].clearHandles();
	};
	PathItem.reorientPaths = function (paths, isInside, clockwise) {
		const length = paths ? paths.length : 0;
		if (!length) { return paths; }

		const lookup = Base.each(paths, function (path, i) {
			this[path._id] = {
				container: null,
				winding: path.isClockwise() ? 1 : -1,
				index: i
			};
		}, {});

		// Сортировка путей по площади
		const sorted = paths.slice().sort(function (a, b) {
			return Math.abs(b.getArea()) - Math.abs(a.getArea());
		});

		const first = sorted[0];
		if (clockwise == null) { clockwise = first.isClockwise(); }

		const collisions = CollisionDetection.findItemBoundsCollisions(sorted, null, Numerical.GEOMETRIC_EPSILON);
		for (let i = 0; i < length; i++) {
			const path1 = sorted[i];
			const entry1 = lookup[path1._id];
			let containerWinding = 0;
			const indices = collisions[i];

			if (indices) {
				let point = null;
				for (let j = indices.length - 1; j >= 0; j--) {
					if (indices[j] < i) {
						point = point || path1.getInteriorPoint();
						const path2 = sorted[indices[j]];
						if (path2.contains(point)) {
							const entry2 = lookup[path2._id];
							containerWinding = entry2.winding;
							entry1.winding += containerWinding;
							entry1.container = entry2['exclude'] ? entry2.container : path2;
							break;
						}
					}
				}
			}
			if (isInside(entry1.winding) === isInside(containerWinding)) {
				entry1['exclude'] = true;
				paths[entry1.index] = null;
			} else {
				path1.setClockwise(entry1.container ? !entry1.container.isClockwise() : clockwise);
			}
		}
		return paths;
	};
	PathItem.divideLocations = function (locations, include, clearLater) {
		const results = include && [];
		const tMin = 1e-8;
		const tMax = 1 - tMin;
		const clearCurves = clearLater || [];
		const clearLookup = clearLater && {};

		let clearHandles = false;
		let renormalizeLocs;
		let prevCurve;
		let prevTime;

		function getId(curve) {
			return curve._path._id + '.' + curve._segment1._index;
		}

		for (let i = (clearLater && clearLater.length) - 1; i >= 0; i--) {
			const curve = clearLater[i];
			if (curve._path) { clearLookup[getId(curve)] = true; }
		}

		for (let i = locations.length - 1; i >= 0; i--) {
			const loc = locations[i];
			const origTime = loc._time;
			let time = loc._time;
			if (loc._curve) {
				if (loc._curve !== prevCurve) {
					clearHandles = !loc._curve.hasHandles() || clearLookup && clearLookup[getId(loc._curve)];
					renormalizeLocs = [];
					prevTime = null;
					prevCurve = loc._curve;
				} else if (prevTime >= tMin) {
					time /= prevTime;
				}
			}
			const exclude = include && !include(loc);
			if (exclude) {
				if (renormalizeLocs) { renormalizeLocs.push(loc); }
				continue;
			} else if (include) {
				results.unshift(loc);
			}
			prevTime = origTime;

			let segment;
			if (time < tMin) {
				segment = loc._curve._segment1;
			} else if (time > tMax) {
				segment = loc._curve._segment2;
			} else {
				const newCurve = loc._curve.divideAtTime(time, true);
				if (clearHandles)
					clearCurves.push(loc._curve, newCurve);
				segment = newCurve._segment1;
				for (let j = renormalizeLocs.length - 1; j >= 0; j--) {
					const l = renormalizeLocs[j];
					l._time = (l._time - time) / (1 - time);
				}
			}
			loc._setSegment(segment);

			const inter = segment._intersection;
			const dest = loc._intersection;
			if (inter) {
				PathItem.linkIntersections(inter, dest);
				let other = inter;
				while (other) {
					PathItem.linkIntersections(other._intersection, inter);
					other = other._next;
				}
			} else {
				segment._intersection = dest;
			}
		}
		if (!clearLater) { PathItem.clearCurveHandles(clearCurves); }

		return results || locations;
	};
	PathItem.getWinding = function (point, curves, dir, closed, dontFlip) {
		const curvesList = Array.isArray(curves) ? curves : (dir ? curves.hor : curves.ver);
		const ia = dir ? 1 : 0;
		const io = ia ^ 1;
		const pv = [point.x, point.y];
		const pa = pv[ia];
		const po = pv[io];
		const windingEpsilon = 1e-9;
		const qualityEpsilon = 1e-6;
		const paL = pa - windingEpsilon;
		const paR = pa + windingEpsilon;

		let windingL = 0, windingR = 0, pathWindingL = 0, pathWindingR = 0;
		let onPath = false, onAnyPath = false, quality = 1;
		let roots = [], vPrev, vClose;

		function addWinding(v) {
			const o0 = v[io + 0];
			const o3 = v[io + 6];
			if (po < Math.min(o0, o3) || po > Math.max(o0, o3)) { return; }

			const a0 = v[ia + 0];
			const a1 = v[ia + 2];
			const a2 = v[ia + 4];
			const a3 = v[ia + 6];

			if (o0 === o3) {
				if ((a0 < paR && a3 > paL) || (a3 < paR && a0 > paL)) { onPath = true; }
				return;
			}

			const t = po === o0 ? 0 : (po === o3 ? 1 :
				(paL > Math.max(a0, a1, a2, a3) || paR < Math.min(a0, a1, a2, a3) ? 1 :
					Curve.solveCubic(v, io, po, roots, 0, 1) > 0 ? roots[0] : 1));

			const a = t === 0 ? a0 : (t === 1 ? a3 : (dir ? Curve.getPoint(v, t).y : Curve.getPoint(v, t).x));
			const winding = o0 > o3 ? 1 : -1;
			const windingPrev = vPrev[io] > vPrev[io + 6] ? 1 : -1;
			const a3Prev = vPrev[ia + 6];

			if (po !== o0) {
				if (a < paL) {
					pathWindingL += winding;
				} else if (a > paR) {
					pathWindingR += winding;
				} else {
					onPath = true;
				}
				if (a > pa - qualityEpsilon && a < pa + qualityEpsilon) {
					quality /= 2;
				}
			} else {
				if (winding !== windingPrev) {
					if (a0 < paL) {
						pathWindingL += winding;
					} else if (a0 > paR) {
						pathWindingR += winding;
					}
				} else if (a0 != a3Prev) {
					if (a3Prev < paR && a > paR) {
						pathWindingR += winding;
						onPath = true;
					} else if (a3Prev > paL && a < paL) {
						pathWindingL += winding;
						onPath = true;
					}
				}
				quality /= 4;
			}

			vPrev = v;
			return !dontFlip && a > paL && a < paR
				&& (dir ? (Curve.getTangent(v, t).x === 0) : (Curve.getTangent(v, t).y === 0))
				&& PathItem.getWinding(point, curves, !dir, closed, true);
		}

		function handleCurve(v) {
			const o0 = v[io + 0];
			const o1 = v[io + 2];
			const o2 = v[io + 4];
			const o3 = v[io + 6];

			if (po <= Math.max(o0, o1, o2, o3) && po >= Math.min(o0, o1, o2, o3)) {
				const a0 = v[ia + 0];
				const a1 = v[ia + 2];
				const a2 = v[ia + 4];
				const a3 = v[ia + 6];
				const monoCurves = (paL > Math.max(a0, a1, a2, a3) || paR < Math.min(a0, a1, a2, a3))
					? [v]
					: Curve.getMonoCurves(v, dir);

				let res;
				for (let i = 0, l = monoCurves.length; i < l; i++) {
					if (res = addWinding(monoCurves[i])) {
						return res;
					}
				}
			}
		}

		for (let i = 0, l = curvesList.length; i < l; i++) {
			const curve = curvesList[i];
			const path = curve._path;
			const v = curve.getValues();
			if (!i || curvesList[i - 1]._path !== path) {
				vPrev = null;
				if (!path._closed) {
					vClose = Curve.getValues(
						path.getLastCurve().getSegment2(),
						curve.getSegment1(),
						null, !closed
					);
					if (vClose[io] !== vClose[io + 6]) {
						vPrev = vClose;
					}
				}
				if (!vPrev) {
					vPrev = v;
					let prev = path.getLastCurve();
					while (prev && prev !== curve) {
						const v2 = prev.getValues();
						if (v2[io] !== v2[io + 6]) {
							vPrev = v2;
							break;
						}
						prev = prev.getPrevious();
					}
				}
			}

			let res;
			if (res = handleCurve(v)) { return res; }
			if (i + 1 === l || curvesList[i + 1]._path !== path) {
				if (vClose && (res = handleCurve(vClose))) {
					return res;
				}
				if (onPath && !pathWindingL && !pathWindingR) {
					pathWindingL = pathWindingR = path.isClockwise(closed) ^ dir ? 1 : -1;
				}
				windingL += pathWindingL;
				windingR += pathWindingR;
				pathWindingL = pathWindingR = 0;
				if (onPath) {
					onAnyPath = true;
					onPath = false;
				}
				vClose = null;
			}
		}
		windingL = Math.abs(windingL);
		windingR = Math.abs(windingR);

		return {
			winding: Math.max(windingL, windingR),
			windingL: windingL,
			windingR: windingR,
			quality: quality,
			onPath: onAnyPath
		};
	};
	PathItem.propagateWinding = function (segment, path1, path2, curveCollisionsMap, operator) {
		const chain = [];
		const start = segment;
		let totalLength = 0;
		let winding = { winding: 0, quality: -1 };

		do {
			const curve = segment.getCurve();
			if (curve) {
				const length = curve.getLength();
				chain.push({ segment, curve, length });
				totalLength += length
			}
			segment = segment.getNext();
		} while (segment && !segment._intersection && segment !== start);

		const offsets = [0.5, 0.25, 0.75];
		const tMin = 1e-3;
		const tMax = 1 - tMin;

		for (let i = 0; i < offsets.length && winding.quality < 0.5; i++) {
			let lengthAtOffset = totalLength * offsets[i];

			for (let j = 0, l = chain.length; j < l; j++) {
				const entry = chain[j];
				const curveLength = entry.length;
				if (lengthAtOffset <= curveLength) {
					const curve = entry.curve;
					const path = curve._path;
					const parent = path._parent;
					const operand = parent instanceof CompoundPath ? parent : path;
					const t = Numerical.clamp(curve.getTimeAt(lengthAtOffset), tMin, tMax);
					const pt = curve.getPointAtTime(t);
					const dir = Math.abs(curve.getTangentAtTime(t).y) < Math.SQRT1_2;

					let wind = null;
					if (operator['subtract'] && path2) {
						const otherPath = operand === path1 ? path2 : path1;
						const pathWinding = otherPath._getWinding(pt, dir, true);
						if (operand === path1 && pathWinding.winding ||
							operand === path2 && !pathWinding.winding) {
							if (pathWinding.quality < 1) {
								continue;
							} else {
								wind = { winding: 0, quality: 1 };
							}
						}
					}
					wind = wind || PathItem.getWinding(pt, curveCollisionsMap[path._id][curve.getIndex()], dir, true);
					if (wind.quality > winding.quality) {
						winding = wind;
					}
					break;
				}
				lengthAtOffset -= curveLength;
			}
		}
		for (let j = chain.length - 1; j >= 0; j--) {
			chain[j].segment._winding = winding;
		}
	};
	PathItem.tracePaths = function (segments, operator) {
		const paths = [];
		let starts;

		function isValid(seg) {
			let winding;
			return !!(seg && !seg._visited && (!operator
				|| operator[(winding = seg._winding || {}).winding]
				&& !(operator['unite'] && winding.winding === 2
					&& winding.windingL && winding.windingR)));
		}

		function isStart(seg) {
			if (seg) {
				for (let i = 0, l = starts.length; i < l; i++) {
					if (seg === starts[i]) { return true; }
				}
			}
			return false;
		}

		function visitPath(path) {
			for (let i = 0, l = path._segments.length; i < l; i++) {
				path._segments[i]._visited = true;
			}
		}

		function getCrossingSegments(segment, collectStarts) {
			let inter = segment._intersection;
			const start = inter;
			const crossings = [];
			if (collectStarts)
				starts = [segment];

			function collect(inter, end) {
				while (inter && inter !== end) {
					const other = inter._segment;
					const path = other && other._path;
					if (path) {
						const next = other.getNext() || path.getFirstSegment();
						const nextInter = next._intersection;
						const isCrossingValid = isStart(other) || isStart(next) || next && (isValid(other) && (isValid(next) || nextInter && isValid(nextInter._segment)));
						if (other !== segment && isCrossingValid) { crossings.push(other); }
						if (collectStarts) { starts.push(other); }
					}
					inter = inter._next;
				}
			}

			if (inter) {
				collect(inter);
				while (inter && inter._previous)
					inter = inter._previous;
				collect(inter, start);
			}
			return crossings;
		}

		segments.sort(function (seg1, seg2) {
			const inter1 = seg1._intersection, inter2 = seg2._intersection;
			const over1 = inter1 ? inter1._overlap : false;
			const over2 = inter2 ? inter2._overlap : false;

			if (over1 ^ over2) { return over1 ? 1 : -1; }
			if (!inter1 ^ !inter2) { return inter1 ? 1 : -1; }
			if (seg1._path !== seg2._path) { return seg1._path._id - seg2._path._id; }
			return seg1._index - seg2._index;
		});

		for (let i = 0; i < segments.length; i++) {
			let segment = segments[i];
			let isValidSegment = isValid(segment);
			if (isValidSegment && segment._path._overlapsOnly) {
				const path1 = segment._path;
				const path2 = segment._intersection._segment._path;

				if (path1.compare(path2)) {
					if (path1.getArea()) {
						paths.push(path1.clone(false));
					}
					visitPath(path1);
					visitPath(path2);
					isValidSegment = false;
				}
			}

			let visitedSegments;
			let branches = [];
			let currentPath = null;
			let isFinished = false;
			let isClosed = true;
			let branch, handleIn;
			while (isValidSegment) {
				const isFirstSegment = !currentPath;
				const crossings = getCrossingSegments(segment, isFirstSegment);
				const otherSegment = crossings.shift();
				isFinished = !isFirstSegment && (isStart(segment) || isStart(otherSegment));
				const isCrossing = !isFinished && otherSegment;
				if (isFirstSegment) {
					currentPath = new Path({ insert: false });
					branch = null;
				}
				if (isFinished) {
					if (segment.isFirst() || segment.isLast()) { isClosed = segment._path._closed; }
					segment._visited = true;
					break;
				}
				if (isCrossing && branch) {
					branches.push(branch);
					branch = null;
				}
				if (!branch) {
					if (isCrossing) { crossings.push(segment); }
					branch = {
						start: currentPath._segments.length,
						crossings: crossings,
						visited: visitedSegments = [],
						handleIn: handleIn
					};
				}
				if (isCrossing) { segment = otherSegment; }
				if (!isValid(segment)) {
					currentPath.removeSegments(branch.start);
					visitedSegments.forEach(function (segment) { segment._visited = false; })
					visitedSegments.length = 0;
					do {
						segment = branch && branch.crossings.shift();
						if (!segment || !segment._path) {
							segment = null;
							branch = branches.pop();
							if (branch) {
								visitedSegments = branch.visited;
								handleIn = branch.handleIn;
							}
						}
					} while (branch && !isValid(segment));
					if (!segment) {
						break;
					}
				}
				const nextSegment = segment.getNext();
				currentPath.add(new Segment(segment._point, handleIn, nextSegment && segment._handleOut));
				segment._visited = true;
				visitedSegments.push(segment);
				segment = nextSegment || segment._path.getFirstSegment();
				handleIn = nextSegment && nextSegment._handleIn;
			}

			if (isFinished && isClosed) {
				currentPath.getFirstSegment().setHandleIn(handleIn);
				currentPath.setClosed(isClosed);
			}
			if (isFinished && currentPath.getArea() !== 0) {
				paths.push(currentPath);
			}
		}
		return paths;
	};

	const Path = function (arg) {
		this._closed = false;
		this._segments = [];
		this._version = 0;

		const isArrayArg = Array.isArray(arg);
		const isObjectElement = isArrayArg && typeof arg[0] === 'object';
		const isValidObject = arg && (arg.size === undefined && (arg.x !== undefined || arg.point !== undefined));
		const segments = isArrayArg
			? isObjectElement ? arg : arguments
			: isValidObject ? arguments : null;
		segments && segments.length > 0
			? this.setSegments(segments)
			: this._curves = undefined;
		this._initialize(!segments && arg);
	};
	InitClassWithStatics(Path, PathItem);

	Path.prototype._equals = function (item) {
		return this._closed === item._closed
			&& Base.equals(this._segments, item._segments);
	};
	Path.prototype.copyContent = function (source) {
		this.setSegments(source._segments);
		this._closed = source._closed;
	};
	Path.prototype._changed = function _changed(flags) {
		Item.prototype._changed.call(this, flags);
		if (flags & 8) {
			this._length = this._area = undefined;
			if (flags & 32) {
				this._version++;
			} else if (this._curves) {
				for (let i = 0, l = this._curves.length; i < l; i++)
					this._curves[i]._changed();
			}
		} else if (flags & 64) {
			this._bounds = undefined;
		}
	};
	Path.prototype.getSegments = function () {
		return this._segments;
	};
	Path.prototype.setSegments = function (segments) {
		this._segments.length = 0;
		this._curves = undefined;
		let length = segments && segments.length;
		if (!length) { return; }

		const last = segments[length - 1];
		if (typeof last === 'boolean') {
			this.setClosed(last);
			length--;
		}
		this._add(Segment.readList(segments, 0, {}, length));

	};
	Path.prototype.getFirstSegment = function () {
		return this._segments[0];
	};
	Path.prototype.getLastSegment = function () {
		return this._segments[this._segments.length - 1];
	};
	Path.prototype.getCurves = function () {
		let curves = this._curves;
		let segments = this._segments;
		if (!curves) {
			const length = this._countCurves();
			curves = this._curves = new Array(length);
			for (let i = 0; i < length; i++)
				curves[i] = new Curve(this, segments[i],
					segments[i + 1] || segments[0]);
		}
		return curves;
	};
	Path.prototype.getFirstCurve = function () {
		return this.getCurves()[0];
	};
	Path.prototype.getLastCurve = function () {
		const curves = this.getCurves();
		return curves[curves.length - 1];
	};
	Path.prototype.isClosed = function () {
		return this._closed;
	};
	Path.prototype.setClosed = function (closed) {
		if (this._closed != (closed = !!closed)) {
			this._closed = closed;
			if (this._curves) {
				const length = this._curves.length = this._countCurves();
				if (closed)
					this._curves[length - 1] = new Curve(this,
						this._segments[length - 1], this._segments[0]);
			}
			this._changed(41);
		}
	};
	Path.prototype.isEmpty = function () {
		return !this._segments.length;
	};
	Path.prototype._transformContent = function (matrix) {
		const segments = this._segments;
		const coords = new Array(6);
		for (let i = 0, l = segments.length; i < l; i++)
			segments[i]._transformCoordinates(matrix, coords, true);
		return true;
	};
	Path.prototype._add = function (segs, index) {
		const segments = this._segments;
		const curves = this._curves;
		const amount = segs.length;
		const append = index == null;
		index = append ? segments.length : index;

		for (let i = 0; i < amount; i++) {
			let segment = segs[i];
			if (segment._path) {
				segment = segs[i] = segment.clone();
			}
			segment._path = this;
			segment._index = index + i;
		}
		if (append) {
			segments.push.apply(segments, segs);
		} else {
			segments.splice.apply(segments, [index, 0].concat(segs));
			for (let i = index + amount, l = segments.length; i < l; i++) {
				segments[i]._index = i;
			}
		}
		if (curves) {
			const total = this._countCurves();
			const start = index > 0 && index + amount - 1 === total ? index - 1 : index;
			let insert = start;
			const end = Math.min(start + amount, total);
			if (segs._curves) {
				curves.splice.apply(curves, [start, 0].concat(segs._curves));
				insert += segs._curves.length;
			}
			for (let i = insert; i < end; i++) {
				curves.splice(i, 0, new Curve(this, null, null));
			}
			this._adjustCurves(start, end);
		}
		this._changed(41);
		return segs;
	};
	Path.prototype._adjustCurves = function (start, end) {
		const segments = this._segments;
		const curves = this._curves;
		let curve;
		for (let i = start; i < end; i++) {
			curve = curves[i];
			curve._path = this;
			curve._segment1 = segments[i];
			curve._segment2 = segments[i + 1] || segments[0];
			curve._changed();
		}
		if (curve = curves[this._closed && !start ? segments.length - 1 : start - 1]) {
			curve._segment2 = segments[start] || segments[0];
			curve._changed();
		}
		if (curve = curves[end]) {
			curve._segment1 = segments[end];
			curve._changed();
		}
	};
	Path.prototype._countCurves = function () {
		const length = this._segments.length;
		return !this._closed && length > 0 ? length - 1 : length;
	};
	Path.prototype.add = function (segment1) {
		return arguments.length > 1 && typeof segment1 !== 'number'
			? this._add(Segment.readList(arguments))
			: this._add([Segment.read(arguments)])[0];
	};
	Path.prototype.insert = function (index, segment1) {
		return arguments.length > 2 && typeof segment1 !== 'number'
			? this._add(Segment.readList(arguments, 1), index)
			: this._add([Segment.read(arguments, 1)], index)[0];
	};
	Path.prototype.addSegment = function () {
		return this._add([Segment.read(arguments)])[0];
	};
	Path.prototype.removeSegment = function (index) {
		return this.removeSegments(index, index + 1)[0] || null;
	};
	Path.prototype.removeSegments = function (start, end, _includeCurves) {
		if (start == null) start = 0;
		if (end == null) end = this._segments.length;
		const segments = this._segments;
		const curves = this._curves;
		const count = segments.length;
		const removed = segments.splice(start, end - start);
		const amount = removed.length;
		if (!amount) return removed;

		for (let i = 0; i < amount; i++) {
			removed[i]._index = removed[i]._path = null;
		}
		for (let i = start, l = segments.length; i < l; i++) {
			segments[i]._index = i;
		}
		if (curves) {
			const index = (start > 0 && end === count + (this._closed ? 1 : 0)) ? start - 1 : start;
			const removedCurves = curves.splice(index, amount);
			for (let i = removedCurves.length - 1; i >= 0; i--) {
				// Есть баг с файлом "shapesMerge - remove curves _path bug" (загрузил к себе в личные документы на nct)
				// removedCurves[i]._path = null;
			}
			if (_includeCurves) {
				removed._curves = removedCurves.slice(1);
			}
			this._adjustCurves(index, index);
		}
		this._changed(41);
		return removed;
	};
	Path.prototype.hasHandles = function () {
		for (let i = 0, l = this._segments.length; i < l; i++) {
			if (this._segments[i].hasHandles()) { return true; }
		}
		return false;
	};
	Path.prototype.clearHandles = function () {
		for (let i = 0, l = this._segments.length; i < l; i++)
			this._segments[i].clearHandles();
	};
	Path.prototype.getLength = function () {
		if (this._length == null) {
			const curves = this.getCurves();
			let length = 0;
			for (let i = 0, l = curves.length; i < l; i++)
				length += curves[i].getLength();
			this._length = length;
		}
		return this._length;
	};
	Path.prototype.getArea = function () {
		if (this._area != null) { return this._area; }

		let area = 0;
		const length = this._segments.length;
		for (let i = 0; i < length; i++) {
			const nextIndex = (i + 1) % length;
			const isLastSegment = (i === length - 1);
			area += Curve.getArea(Curve.getValues(
				this._segments[i],
				this._segments[isLastSegment ? 0 : nextIndex],
				null, isLastSegment && !this._closed
			));
		}
		this._area = area;
		return area;
	};
	Path.prototype.join = function (path, tolerance) {
		const epsilon = tolerance || 0;
		if (path && path !== this) {
			let last1 = this.getLastSegment();
			let last2 = path.getLastSegment();
			if (!last2) { return this; }

			if (last1 && last1._point.isClose(last2._point, epsilon)) { path.reverse(); }
			const first2 = path.getFirstSegment();
			if (last1 && last1._point.isClose(first2._point, epsilon)) {
				last1.setHandleOut(first2._handleOut);
				this._add(path._segments.slice(1));
			} else {
				const first1 = this.getFirstSegment();
				if (first1 && first1._point.isClose(first2._point, epsilon)) { path.reverse(); }
				last2 = path.getLastSegment();
				if (first1 && first1._point.isClose(last2._point, epsilon)) {
					first1.setHandleIn(last2._handleIn);
					this._add(path._segments.slice(0, path._segments.length - 1), 0);
				} else {
					this._add(path._segments.slice());
				}
			}
			if (path._closed) { this._add([path._segments[0]]); }
			path.remove();
		}
		const first = this.getFirstSegment();
		const last = this.getLastSegment();
		if (first !== last && first._point.isClose(last._point, epsilon)) {
			first.setHandleIn(last._handleIn);
			last.remove();
			this.setClosed(true);
		}
		return this;
	};
	Path.prototype.reduce = function (options) {
		const curves = this.getCurves();
		const simplify = options && options.simplify;
		const tolerance = simplify ? 1e-7 : 0;
		for (let i = curves.length - 1; i >= 0; i--) {
			const curve = curves[i];
			if (!curve.hasHandles() && (!curve.hasLength(tolerance) || simplify && curve.isCollinear(curve.getNext())))
				curve.remove();
		}
		return this;
	};
	Path.prototype.reverse = function () {
		this._segments.reverse();
		for (let i = 0, l = this._segments.length; i < l; i++) {
			const segment = this._segments[i];
			const handleIn = segment._handleIn;
			segment._handleIn = segment._handleOut;
			segment._handleOut = handleIn;
			segment._index = i;
		}
		this._curves = null;
		this._changed(9);
	};
	Path.prototype.compare = function (path) {
		if (!path || path instanceof CompoundPath)
			return PathItem.prototype.compare.call(this, path);
		const curves1 = this.getCurves();
		const curves2 = path.getCurves();
		if (!curves1.length || !curves2.length) {
			return curves1.length == curves2.length;
		}
		let v1 = curves1[0].getValues();
		const values2 = [];
		let pos1 = 0, pos2;
		let end1 = 0, end2;
		for (let i = 0; i < curves2.length; i++) {
			const v2 = curves2[i].getValues();
			values2.push(v2);
			const overlaps = Curve.getOverlaps(v1, v2);
			if (overlaps) {
				pos2 = !i && overlaps[0][0] > 0 ? curves2.length - 1 : i;
				end2 = overlaps[0][1];
				break;
			}
		}
		let v2 = values2[pos2];
		let start2;
		while (v1 && v2) {
			const overlaps = Curve.getOverlaps(v1, v2);
			if (overlaps) {
				const t1 = overlaps[0][0];
				if (Math.abs(t1 - end1) < 1e-8) {
					end1 = overlaps[1][0];
					if (end1 === 1) {
						v1 = ++pos1 < curves1.length ? curves1[pos1].getValues() : null;
						end1 = 0;
					}
					const t2 = overlaps[0][1];
					if (Math.abs(t2 - end2) < 1e-8) {
						if (!start2)
							start2 = [pos2, t2];
						end2 = overlaps[1][1];
						if (end2 === 1) {
							if (++pos2 >= curves2.length)
								pos2 = 0;
							v2 = values2[pos2] || curves2[pos2].getValues();
							end2 = 0;
						}
						if (!v1) {
							return start2[0] === pos2 && start2[1] === end2;
						}
						continue;
					}
				}
			}
			break;
		}
		return false;
	};
	Path.prototype.getLocationAt = function (offset) {
		if (typeof offset === 'number') {
			const curves = this.getCurves();
			let length = 0;
			for (let i = 0, l = curves.length; i < l; i++) {
				const start = length;
				const curve = curves[i];
				length += curve.getLength();
				if (length > offset) {
					return curve.getLocationAt(offset - start);
				}
			}
			if (curves.length > 0 && offset <= this.getLength()) {
				return new CurveLocation(curves[curves.length - 1], 1);
			}
		} else if (offset && offset.getPath && offset.getPath() === this) {
			return offset;
		}
		return null;
	};
	Path.prototype.getPointAt = function (offset) {
		const loc = this.getLocationAt(offset);
		return loc && loc.getPoint();
	};
	Path.prototype.getTangentAt = function (offset) {
		const loc = this.getLocationAt(offset);
		return loc && loc.getTangent();
	};
	Path.prototype.getNormalAt = function (offset) {
		const loc = this.getLocationAt(offset);
		return loc && loc.getNormal();
	};
	Path.prototype.getCurvatureAt = function (offset) {
		const loc = this.getLocationAt(offset);
		return loc && loc.getCurvature();
	};
	Path.prototype.moveTo = function () {
		if (this._segments.length === 1)
			this.removeSegment(0);
		if (!this._segments.length)
			this._add([new Segment(Point.read(arguments))]);
	};
	Path.prototype.lineTo = function () {
		this._add([new Segment(Point.read(arguments))]);
	};
	Path.prototype.cubicCurveTo = function () {
		const handle1 = Point.read(arguments);
		const handle2 = Point.read(arguments);
		const to = Point.read(arguments);
		const current = Path.getCurrentSegment(this);
		current.setHandleOut(handle1.subtract(current._point));
		this._add([new Segment(to, handle2.subtract(to))]);
	};
	Path.prototype.quadraticCurveTo = function () {
		const handle = Point.read(arguments);
		const to = Point.read(arguments);
		const current = Path.getCurrentSegment(this)._point;
		this.cubicCurveTo(
			handle.add(current.subtract(handle).multiply(1 / 3)),
			handle.add(to.subtract(handle).multiply(1 / 3)),
			to
		);
	};
	Path.prototype.closePath = function (tolerance) {
		this.setClosed(true);
		this.join(this, tolerance);
	};
	Path.prototype._getBounds = function (matrix, options) {
		return options.handle
			? Path.getHandleBounds(this._segments, this._closed, this, matrix, options)
			: Path.getBounds(this._segments, this._closed, this, matrix, options);
	};

	Path.getCurrentSegment = function (that) {
		const segments = that._segments;
		if (!segments.length) { throw new Error('Use a moveTo() command first'); }
		return segments[segments.length - 1];
	};
	Path.getBounds = function (segments, closed, _path, matrix) {
		const first = segments[0];
		if (!first) { return new Rectangle(); }

		let coords = new Array(6);
		let prevCoords = first._transformCoordinates(matrix, new Array(6));
		const min = prevCoords.slice(0, 2);
		const max = min.slice();
		const roots = new Array(2);

		function processSegment(segment) {
			segment._transformCoordinates(matrix, coords);
			for (let i = 0; i < 2; i++) {
				Curve._addBounds(
					prevCoords[i],
					prevCoords[i + 4],
					coords[i + 2],
					coords[i],
					i, 0, min, max, roots);
			}
			const tmp = prevCoords;
			prevCoords = coords;
			coords = tmp;
		}

		for (let i = 1, l = segments.length; i < l; i++)
			processSegment(segments[i]);
		if (closed)
			processSegment(first);
		return new Rectangle(min[0], min[1], max[0] - min[0], max[1] - min[1]);
	};
	Path.getHandleBounds = function (segments, _closed, _path, matrix) {
		const coords = new Array(6);
		let x1 = Infinity;
		let x2 = -x1;
		let y1 = x1;
		let y2 = x2;
		for (let i = 0, l = segments.length; i < l; i++) {
			segments[i]._transformCoordinates(matrix, coords);
			for (let j = 0; j < 6; j += 2) {
				const x = coords[j];
				const y = coords[j + 1];
				if (x < x1) x1 = x;
				if (x > x2) x2 = x;
				if (y < y1) y1 = y;
				if (y > y2) y2 = y;
			}
		}
		return new Rectangle(x1, y1, x2 - x1, y2 - y1);
	};
	Path.Rectangle = function (left, top, right, bottom) {
		const path = new Path();
		path.moveTo(left, top);
		path.lineTo(right, top);
		path.lineTo(right, bottom);
		path.lineTo(left, bottom);
		path.closePath();
		return path;
	};

	const CompoundPath = function (arg) {
		this._children = [];
		this._namedChildren = {};
		if (!this._initialize(arg)) {
			this.addChildren(Array.isArray(arg) ? arg : arguments);
		}
	};
	InitClassWithStatics(CompoundPath, PathItem);

	CompoundPath.prototype.insertChildren = function insertChildren(index, items) {
		let list = items;
		const first = list[0];
		if (first && typeof first[0] === 'number')
			list = [list];
		for (let i = items.length - 1; i >= 0; i--) {
			const item = list[i];
			if (list === items && !(item instanceof Path))
				list = Base.slice(list);
			if (Array.isArray(item)) {
				list[i] = new Path({ segments: item, insert: false });
			} else if (item instanceof CompoundPath) {
				list.splice.apply(list, [i, 1].concat(item.removeChildren()));
				item.remove();
			}
		}
		return Item.prototype.insertChildren.call(this, index, list);
	};
	CompoundPath.prototype.reduce = function reduce(options) {
		for (let i = this._children.length - 1; i >= 0; i--) {
			const path = this._children[i].reduce(options);
			if (path.isEmpty())
				path.remove();
		}
		if (!this._children.length) {
			const path = new Path({ insert: false });
			path.copyAttributes(this);
			path.insertAbove(this);
			this.remove();
			return path;
		}
		return Item.prototype.reduce.call(this);
	};
	CompoundPath.prototype.isClosed = function () {
		for (let i = 0, l = this._children.length; i < l; i++) {
			if (!this._children[i]._closed)
				return false;
		}
		return true;
	};
	CompoundPath.prototype.setClosed = function (closed) {
		for (let i = 0, l = this._children.length; i < l; i++) {
			this._children[i].setClosed(closed);
		}
	};
	CompoundPath.prototype.getFirstSegment = function () {
		const first = this.getFirstChild();
		return first && first.getFirstSegment();
	};
	CompoundPath.prototype.getLastSegment = function () {
		const last = this.getLastChild();
		return last && last.getLastSegment();
	};
	CompoundPath.prototype.getCurves = function () {
		const curves = [];
		for (let i = 0, l = this._children.length; i < l; i++) {
			curves.push.apply(curves, this._children[i].getCurves());
		}
		return curves;
	};
	CompoundPath.prototype.getFirstCurve = function () {
		const first = this.getFirstChild();
		return first && first.getFirstCurve();
	};
	CompoundPath.prototype.getLastCurve = function () {
		const last = this.getLastChild();
		return last && last.getLastCurve();
	};
	CompoundPath.prototype.getArea = function () {
		let area = 0;
		for (let i = 0, l = this._children.length; i < l; i++)
			area += this._children[i].getArea();
		return area;
	};
	CompoundPath.prototype.isEmpty = function () {
		let empty = true;
		for (let i = 0, l = this._children.length; i < l; i++)
			empty *= this._children[i].isEmpty();
		return empty;
	};
	CompoundPath.prototype.getLength = function () {
		let length = 0;
		for (let i = 0, l = this._children.length; i < l; i++)
			length += this._children[i].getLength();
		return length;
	};
	CompoundPath.prototype.moveTo = function () {
		const current = CompoundPath.getCurrentPath(this);
		const path = current && current.isEmpty() ? current : new Path({ insert: false });
		if (path !== current) { this.addChild(path); }
		path.moveTo.apply(path, arguments);
	};
	CompoundPath.prototype.closePath = function (tolerance) {
		CompoundPath.getCurrentPath(this, true).closePath(tolerance);
	};
	CompoundPath.prototype.lineTo = function () {
		const path = CompoundPath.getCurrentPath(this, true);
		path.lineTo.apply(path, arguments);
	};
	CompoundPath.prototype.cubicCurveTo = function () {
		const path = CompoundPath.getCurrentPath(this, true);
		path.cubicCurveTo.apply(path, arguments);
	};
	CompoundPath.prototype.quadraticCurveTo = function () {
		const path = CompoundPath.getCurrentPath(this, true);
		path.quadraticCurveTo.apply(path, arguments);
	};
	CompoundPath.prototype.reverse = function (param) {
		let res;
		for (let i = 0, l = this._children.length; i < l; i++) {
			res = this._children[i].reverse(param) || res;
		}
		return res;
	};

	CompoundPath.getCurrentPath = function (that, check) {
		if (check && !that._children.length)
			throw new Error('Use a moveTo() command first');
		return that._children[that._children.length - 1];
	};

	// EXPORTS

	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon']['PathBoolean'] = {}
	window['AscCommon']['PathBoolean']['CompoundPath'] = CompoundPath;

	CompoundPath.prototype['divide'] = PathItem.prototype.divide;
	CompoundPath.prototype['unite'] = PathItem.prototype.unite;
	CompoundPath.prototype['intersect'] = PathItem.prototype.intersect;
	CompoundPath.prototype['subtract'] = PathItem.prototype.subtract;
	CompoundPath.prototype['exclude'] = PathItem.prototype.exclude;

	Path.prototype['divide'] = PathItem.prototype.divide;
	Path.prototype['unite'] = PathItem.prototype.unite;
	Path.prototype['intersect'] = PathItem.prototype.intersect;
	Path.prototype['subtract'] = PathItem.prototype.subtract;
	Path.prototype['exclude'] = PathItem.prototype.exclude;

	CompoundPath.prototype['moveTo'] = CompoundPath.prototype.moveTo;
	CompoundPath.prototype['lineTo'] = CompoundPath.prototype.lineTo;
	CompoundPath.prototype['cubicCurveTo'] = CompoundPath.prototype.cubicCurveTo;
	CompoundPath.prototype['closePath'] = CompoundPath.prototype.closePath;
	CompoundPath.prototype['getChildren'] = Item.prototype.getChildren;
	CompoundPath.prototype['getBounds'] = Item.prototype.getBounds;
	CompoundPath.prototype['getPosition'] = Item.prototype.getPosition;
	CompoundPath.prototype['setPosition'] = Item.prototype.setPosition;

	Path.prototype['getSegments'] = Path.prototype.getSegments;
	Path.prototype['isClosed'] = Path.prototype.isClosed;
	Path.prototype['getBounds'] = Item.prototype.getBounds;
	Path.prototype['getPosition'] = Item.prototype.getPosition;
	Path.prototype['setPosition'] = Item.prototype.setPosition;

	Segment.prototype['isFirst'] = Segment.prototype.isFirst;
	Segment.prototype['isLast'] = Segment.prototype.isLast;
	Segment.prototype['getPrevious'] = Segment.prototype.getPrevious;
	Segment.prototype['getNext'] = Segment.prototype.getNext;
	Segment.prototype['getPoint'] = Segment.prototype.getPoint;
	Segment.prototype['getHandleOut'] = Segment.prototype.getHandleOut;
	Segment.prototype['getHandleIn'] = Segment.prototype.getHandleIn;

	Rectangle.prototype['getTopLeft'] = Rectangle.prototype.getTopLeft;
	Rectangle.prototype['getWidth'] = Rectangle.prototype.getWidth;
	Rectangle.prototype['getHeight'] = Rectangle.prototype.getHeight;
	Rectangle.prototype['getLeft'] = Rectangle.prototype.getLeft;
	Rectangle.prototype['getTop'] = Rectangle.prototype.getTop;

	Point.prototype['subtract'] = Point.prototype.subtract;
	Point.prototype['getX'] = Point.prototype.getX;
	Point.prototype['getY'] = Point.prototype.getY;

})(window);
