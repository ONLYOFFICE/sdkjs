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

module.exports = function(grunt) {
	function loadConfig(name) {
		var config = require(pathConfigs +'/' + name + '.json');
		if (config) {
			grunt.log.ok((name + ' config loaded successfully').green);
			return config;
		}
		grunt.log.error().writeln(('could not load' + name + 'config file').red);
		return null;
	}
	function getExterns(config) {
		var externs = config['externs'];
		var result = [];
		for (var i = 0; i < externs.length; ++i) {
			result.push('--externs=' + externs[i]);
		}
		return result;
	}
	function getFilesMin(config) {
		var result = config['min'];
		if (grunt.option('mobile')) {
			result = config['mobile_banners']['min'].concat(result);
		}
		if (grunt.option('desktop')) {
			result = result.concat(config['desktop']['min']);
		}
		return result;
	}
	function getFilesAll(config) {
		var result = config['common'];
		if (grunt.option('mobile')) {
			result = config['mobile_banners']['common'].concat(result);

			var excludeFiles = config['exclude_mobile'];
			result = result.filter(function(item) {
				return -1 === excludeFiles.indexOf(item);
			});
			result = result.concat(config['mobile']);
		}
		if (!grunt.option('noprivate')) {
			result = result.concat(config['private']);
		}
		if (grunt.option('desktop')) {
			result = result.concat(config['desktop']['common']);
		}
		return result;
	}
	function getSdkPath(min, name) {
		return path.join(name, min ? 'sdk-all-min.js' : 'sdk-all.js');
	}

	var path = require('path');
	var pathConfigs = grunt.option('src') || './configs';
	var level = grunt.option('level') || 'ADVANCED';
	var formatting = grunt.option('formatting') || '';

	require('google-closure-compiler').grunt(grunt, {
		platform: 'java',
		extraArguments: ['-Xms2048m']
	});

	grunt.loadNpmTasks('grunt-contrib-clean');
	grunt.loadNpmTasks('grunt-contrib-concat');
	grunt.loadNpmTasks('grunt-contrib-copy');
	grunt.loadNpmTasks('grunt-replace');
	grunt.loadNpmTasks('grunt-split-file');

	grunt.registerTask('build-sdk', 'Build SDK', function () {
		var configFonts = loadConfig('fonts');
		var configExterns = loadConfig('externs');
		var configWord = loadConfig('word');
		var configCell = loadConfig('cell');
		var configSlide = loadConfig('slide');
		if (!configFonts || !configExterns || !configWord || !configCell || !configSlide) {
			return;
		}
		configWord = configWord['sdk'];
		configCell = configCell['sdk'];
		configSlide = configSlide['sdk'];

		// crete empty.js for polyfills
		var emptyJs = 'empty.js';
		grunt.file.write(emptyJs, '');

		var optionsSdkMin ={
			banner: '',
			footer: 'window["split"]="split";'
		};
		var optionsSdkAll = {};
		if (!grunt.option('noclosure')) {
			optionsSdkAll = {
				banner: '(function(window, undefined) {',
				footer: '})(window);'
			};
		}
		var fontsWasmTmp = 'fonts-wasm-tmp.js';
		var fontsJsTmp = 'fonts-js-tmp.js';
		var sdkMinTmp = 'sdk-min-tmp.js';
		var sdkAllTmp = 'sdk-all-tmp.js';
		var sdkWordTmp = 'sdk-word-tmp.js';
		var sdkCellTmp = 'sdk-cell-tmp.js';
		var sdkSlideTmp = 'sdk-slide-tmp.js';

		var compilerArgs = getExterns(configExterns);
		if (grunt.option('map')) {
			compilerArgs.push('--property_renaming_report=sdk-all.props.js.map');
			compilerArgs.push('--variable_renaming_report=sdk-all.vars.js.map');
		}
		if (formatting) {
			compilerArgs.push('--formatting=' + formatting);
		}

		grunt.initConfig({
			concat: {
				wasm: {
					src: configFonts['wasm'],
					dest: fontsWasmTmp
				},
				js: {
					src: configFonts['js'],
					dest: fontsJsTmp
				},
				wordsdkmin: {
					options: optionsSdkMin,
					src: getFilesMin(configWord),
					dest: sdkMinTmp
				},
				wordsdkall: {
					options: optionsSdkAll,
					src: getFilesAll(configWord),
					dest: sdkAllTmp
				},
				wordall: {
					src: [sdkMinTmp, sdkAllTmp],
					dest: sdkWordTmp
				},
				cellsdkmin: {
					options: optionsSdkMin,
					src: getFilesMin(configCell),
					dest: sdkMinTmp
				},
				cellsdkall: {
					options: optionsSdkAll,
					src: getFilesAll(configCell),
					dest: sdkAllTmp
				},
				cellall: {
					src: [sdkMinTmp, sdkAllTmp],
					dest: sdkCellTmp
				},
				slidesdkmin: {
					options: optionsSdkMin,
					src: getFilesMin(configSlide),
					dest: sdkMinTmp
				},
				slidesdkall: {
					options: optionsSdkAll,
					src: getFilesAll(configSlide),
					dest: sdkAllTmp
				},
				slideall: {
					src: [sdkMinTmp, sdkAllTmp],
					dest: sdkSlideTmp
				}
			},
			'closure-compiler': {
				js: {
					options: {
						args: compilerArgs.concat(
							'--rewrite_polyfills=true', '--jscomp_off=checkVars',
							'--warning_level=QUIET', '--compilation_level=' + level,
							'--module=polyfill:1:', '--js=' + emptyJs,
							'--module=fontswasm:1:polyfill', '--js=' + fontsWasmTmp,
							'--module=fontsjs:1:fontswasm', '--js=' + fontsJsTmp,
							'--module=word:1:fontswasm', '--js=' + sdkWordTmp,
							'--module=cell:1:fontswasm', '--js=' + sdkCellTmp,
							'--module=slide:1:fontswasm', '--js=' + sdkSlideTmp)
					}
				}
			},
			clean: {
				tmp: {
					options: {
						force: true
					},
					src: [
						emptyJs,
						fontsWasmTmp,
						fontsJsTmp,
						sdkMinTmp,
						sdkAllTmp,
						sdkWordTmp,
						sdkCellTmp,
						sdkSlideTmp
					]
				}
			}
		});
	});
	grunt.registerTask('license', 'Add license', function () {
		const appCopyright = "Copyright (C) Ascensio System SIA 2012-" + grunt.template.today('yyyy') +". All rights reserved";
		const publisherUrl = "https://www.onlyoffice.com/";
		var cache = '*.cache';
		var fonts = '../common/libfont/';
		var word = '../word/';
		var cell = '../cell/';
		var slide = '../slide/';
		var polyfill = 'polyfill.js';
		var fontsWasm = 'fontswasm.js';
		var fontsJs = 'fontsjs.js';
		var fontFile = 'fonts.js';
		var wordJs = 'word.js';
		var cellJs = 'cell.js';
		var slideJs = 'slide.js';
		var license = 'license.header';
		var deploy = '../deploy/sdkjs/';
		var sdkjspattern = 'sdk-*.js';
		var splitLine;
		if ('ADVANCED' === level) {
			splitLine = ('PRETTY_PRINT' === formatting) ? 'window.split = "split";' : 'window.split="split";';
		} else {
			splitLine = ('PRETTY_PRINT' === formatting) ? 'window["split"] = "split";' : 'window["split"]="split";';
		}
		var splitOptions = {
			separator: splitLine,
			prefix: ["sdk-all-min", "sdk-all"]
		};

		var concatSdk = {files:{}};
		var concatSdkFiles = concatSdk['files'];
		concatSdkFiles[fontsWasm] = [license, fontsWasm];
		concatSdkFiles[fontsJs] = [license, fontsJs];
		concatSdkFiles[getSdkPath(true, word)] = [license, polyfill, getSdkPath(true, word)];
		concatSdkFiles[getSdkPath(false, word)] = [license, getSdkPath(false, word)];
		concatSdkFiles[getSdkPath(true, cell)] = [license, polyfill, getSdkPath(true, cell)];
		concatSdkFiles[getSdkPath(false, cell)] = [license, getSdkPath(false, cell)];
		concatSdkFiles[getSdkPath(true, slide)] = [license, polyfill, getSdkPath(true, slide)];
		concatSdkFiles[getSdkPath(false, slide)] = [license, getSdkPath(false, slide)];

		grunt.initConfig({
			splitfile: {
				word: {
					options: splitOptions,
					dest: word,
					src: wordJs
				},
				cell: {
					options: splitOptions,
					dest: cell,
					src: cellJs
				},
				slide: {
					options: splitOptions,
					dest: slide,
					src: slideJs
				}
			},
			concat: {
				sdk: concatSdk
			},
			replace: {
				version: {
					options: {
						patterns: [
							{
								json: {
									AppCopyright: process.env['APP_COPYRIGHT'] || appCopyright,
									PublisherUrl: process.env['PUBLISHER_URL'] || publisherUrl,
									Version: process.env['PRODUCT_VERSION'] || '0.0.0',
									Build: process.env['BUILD_NUMBER'] || '0'
								}
							}
						]
					},
					files: [
						{src: [fontsWasm], dest: path.join(fonts, 'wasm', fontFile)},
						{src: [fontsJs], dest: path.join(fonts, 'js', fontFile)},
						{src: [getSdkPath(true, word), getSdkPath(false, word)], dest: word},
						{src: [getSdkPath(true, cell), getSdkPath(false, cell)], dest: cell},
						{src: [getSdkPath(true, slide), getSdkPath(false, slide)], dest: slide}
					]
				}
			},
			clean: {
				tmp: {
					options: {
						force: true
					},
					src: [
						polyfill,
						fontsWasm,
						fontsJs,
						wordJs,
						cellJs,
						slideJs,
						word + cache,
						cell + cache,
						slide + cache
					]
				}
			},
			copy: {
				sdkjs: {
					files: [
						{
							expand: true,
							cwd: '../common/',
							src: [
								'Images/*',
								'Native/*.js',
								'libfont/js/fonts.*',
								'libfont/wasm/fonts.*'
							],
							dest: path.join(deploy, 'common')
						},
						{
							expand: true,
							src: path.join(word, sdkjspattern),
							dest: path.join(deploy, 'word')
						},
						{
							expand: true,
							cwd: path.join(cell, 'css'),
							src: '*.css',
							dest: path.join(deploy, 'cell', 'css')
						},
						{
							expand: true,
							src: path.join(cell, sdkjspattern),
							dest: path.join(deploy, 'cell')
						},
						{
							expand: true,
							cwd: path.join(slide, 'themes'),
							src: '**/**',
							dest: path.join(deploy, 'slide', 'themes')
						},
						{
							expand: true,
							src: path.join(slide, sdkjspattern),
							dest: path.join(deploy, 'slide')
						}
					]
				}
			}
		})
	});
	grunt.registerTask('default', ['build-sdk', 'concat', 'closure-compiler', 'clean', 'license', 'splitfile', 'concat', 'replace', 'clean', 'copy']);
};
