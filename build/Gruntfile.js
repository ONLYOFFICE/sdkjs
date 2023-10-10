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

module.exports = function(grunt) {
	function loadConfig(pathConfigs, name) {
		let config;
		try {
			const file = path.join(pathConfigs, name + '.json');
			if (grunt.file.exists(file)) {
				config = grunt.file.readJSON(file);
				grunt.log.ok((name + ' config loaded successfully').green);
			}
		} catch (e) {
			grunt.log.error().writeln(('could not load' + name + 'config file').red);
		}
		return config;
	}
	function fixPath(obj, basePath = '') {
		function fixPathArray(arrPaths, basePath = '') {
			arrPaths.forEach((element, index) => {
				arrPaths[index] = path.join(basePath, element);
			});
		}
		if (Array.isArray(obj))
			return fixPathArray(obj, basePath);
		for (let prop in obj) {
			fixPath(obj[prop], basePath);
		}
	}
	function fixUrl(arrPaths, basePath = '') {
		const url = require('url');
		arrPaths.forEach((element, index) => {
			arrPaths[index] = url.resolve(basePath, element);
		});
	}
	function getConfigs() {
		const configs = new CConfig(grunt.option('src') || '../');

		let addons = grunt.option('addon') || [];
		if (!Array.isArray(addons)) {
			addons = [addons];
		}
		addons.forEach(element => configs.append(grunt.file.isDir(element) ? element : path.join('../../', element)));

		return configs;
	}
	function writeScripts(config, name) {
		const develop = '../develop/sdkjs/';
		const fileName = 'scripts.js';
		const files = ['../vendor/polyfill.js', '../common/applyDocumentChanges.js', '../common/AllFonts.js'].concat(getFilesMin(config), getFilesAll(config));
		fixUrl(files, '../../../../sdkjs/build/');

		grunt.file.write(path.join(develop, name, fileName), 'var sdk_scripts = [\n\t"' + files.join('",\n\t"') + '"\n];');
	}

	function CConfig(pathConfigs) {
		this.externs = null;
		this.word = null;
		this.cell = null;
		this.slide = null;

		this.append(pathConfigs);
	}

	CConfig.prototype.append = function (basePath = '') {
		const pathConfigs = path.join(basePath, 'configs');
		
		function appendOption(name) {
			const option = loadConfig(pathConfigs, name);
			if (!option)
				return;
			
			fixPath(option, basePath);
			
			if (!this[name]) {
				this[name] = option;
				return;
			}
			
			function mergeProps(base, addon) {
				for (let prop in addon)
				{
					if (Array.isArray(addon[prop])) {
						base[prop] = Array.isArray(base[prop]) ? base[prop].concat(addon[prop]) : addon[prop];
					} else {
						if (!base[prop]) 
							base[prop] = {};
						mergeProps(base[prop], addon[prop]);						
					}
				}
			}
			
			mergeProps(this[name], option);			
		}
		
		appendOption.call(this, 'externs');
		appendOption.call(this, 'word');
		appendOption.call(this, 'cell');
		appendOption.call(this, 'slide');
	};
	CConfig.prototype.valid = function () {
		return this.externs && this.word && this.cell && this.slide;
	};

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
		if (grunt.option('desktop')) {
			result = result.concat(config['desktop']['common']);
		}
		return result;
	}

	const path = require('path');
	const level = grunt.option('level') || 'ADVANCED';
	const formatting = grunt.option('formatting') || '';

	require('google-closure-compiler').grunt(grunt, {
		platform: ['native', 'java'],
		extraArguments: ['-Xms2048m']
	});

	grunt.loadNpmTasks('grunt-contrib-clean');
	grunt.loadNpmTasks('grunt-contrib-copy');

	const configs = getConfigs();
	if (!configs.valid()) {
		return;
	}

	const configWord = configs.word['sdk'];
	const configCell = configs.cell['sdk'];
	const configSlide = configs.slide['sdk'];
	const deploy = '../deploy/sdkjs/';

	const compilerArgs = getExterns(configs.externs);
	if (formatting) {
		compilerArgs.push('--formatting=' + formatting);
	}
	if (grunt.option('map')) {
		grunt.file.mkdir(path.join('./maps'));
		compilerArgs.push('--create_source_map=' + path.join('maps/%outname%.map'));
		compilerArgs.push('--source_map_format=V3');
		compilerArgs.push('--source_map_include_content=true');
	}

	const appCopyright = process.env['APP_COPYRIGHT'] || "Copyright (C) Ascensio System SIA 2012-" + grunt.template.today('yyyy') +". All rights reserved";
	const publisherUrl = process.env['PUBLISHER_URL'] || "https://www.onlyoffice.com/";
	const companyName = process.env['COMPANY_NAME'] || 'onlyoffice';
	const version = process.env['PRODUCT_VERSION'] || '0.0.0';
	const buildNumber = process.env['BUILD_NUMBER'] || '0';
	const beta = grunt.option('beta') || 'false';

	let license = grunt.file.read(path.join('./license.header'));
	license = license.replace('@@AppCopyright', appCopyright);
	license = license.replace('@@PublisherUrl', publisherUrl);
	license = license.replace('@@Version', version);
	license = license.replace('@@Build', buildNumber);

	function getCompileConfig(sdkmin, sdkall, outmin, outall, name) {
		const args = compilerArgs.concat (
		`--define=window.AscCommon.g_cCompanyName='${companyName}'`,
		`--define=window.AscCommon.g_cProductVersion='${version}'`,
		`--define=window.AscCommon.g_cBuildNumber='${buildNumber}'`,
		`--define=window.AscCommon.g_cIsBeta='${beta}'`,
		'--rewrite_polyfills=true',
		'--warning_level=QUIET',
		'--language_out=ECMASCRIPT5',
		'--compilation_level=' + level,
		...sdkmin.map((file) => ('--js=' + file)),
		`--chunk=${outmin}:${sdkmin.length}`,
		`--chunk_wrapper=${outmin}:${license}\n%s`,
		...sdkall.map((file) => ('--js=' + file)),
		`--chunk=${outall}:${sdkall.length}:${outmin}`,
		`--chunk_wrapper=${outall}:${license}\n(function(window, undefined) {%s})(window);`);
		if (grunt.option('map')) {
			args.push('--property_renaming_report=' + path.join(`maps/${name}.props.js.map`));
			args.push('--variable_renaming_report=' + path.join(`maps/${name}.vars.js.map`));
		}
		return {
			'closure-compiler': {
				js: {
					options: {
						args: args,
					}
				}
			}
		}
	}
	grunt.registerTask('compile-word', 'Compile Word SDK', function () {
		grunt.initConfig(getCompileConfig(getFilesMin(configWord), getFilesAll(configWord), 'word-all-min', 'word-all', 'word'));
		grunt.task.run('closure-compiler');
	});
	grunt.registerTask('compile-cell', 'Compile Cell SDK', function () {
		grunt.initConfig(getCompileConfig(getFilesMin(configCell), getFilesAll(configCell), 'cell-all-min', 'cell-all', 'cell'));
		grunt.task.run('closure-compiler');
	});
	grunt.registerTask('compile-slide', 'Compile Slide SDK', function () {
		grunt.initConfig(getCompileConfig(getFilesMin(configSlide), getFilesAll(configSlide), 'slide-all-min', 'slide-all', 'slide'));
		grunt.task.run('closure-compiler');
	});
	grunt.registerTask('compile-sdk', ['compile-word', 'compile-cell', 'compile-slide']);

	grunt.registerTask('clean-deploy', 'Clean files after deploy', function() {
		grunt.initConfig({
			clean: {
				tmp: {
					options: {
						force: true
					},
					src: [
						wordJsAll,
						wordJsMin,
						cellJsAll,
						cellJsMin,
						slideJsAll,
						slideJsMin,
					]
				}
			}
		});
		grunt.task.run('clean');
	});
	const word = path.join(deploy, 'word');
	const cell = path.join(deploy, 'cell');
	const slide = path.join(deploy, 'slide');
	
	const wordJsAll = 'word-all.js';
	const wordJsMin = 'word-all-min.js';
	const cellJsAll = 'cell-all.js';
	const cellJsMin = 'cell-all-min.js';
	const slideJsAll = 'slide-all.js';
	const slideJsMin = 'slide-all-min.js';
	grunt.registerTask('deploy-sdk', 'Deploy SDK files', function () {
		grunt.initConfig({
			clean: {
				deploy: {
					options: {
						force: true
					},
					src: [
						deploy
					]
				}
			},
			copy: {
				sdkjs: {
					files: [
						{
							expand: true,
							src: [
								slideJsAll,
								slideJsMin
							],
							dest: slide,
							rename: function (dest, src) {
								return path.join(dest , src.replace('slide', 'sdk'));
							}
						},
						{
							expand: true,
							src: [
								wordJsAll,
								wordJsMin
							],
							dest: word,
							rename: function (dest, src) {
								return path.join(dest , src.replace('word', 'sdk'));
							}
						},
						{
							expand: true,
							src: [
								cellJsAll,
								cellJsMin
							],
							dest: cell,
							rename: function (dest, src) {
								return path.join(dest , src.replace('cell', 'sdk'));
							}
						},
						{
							expand: true,
							cwd: '../common/',
							src: [
								'Charts/ChartStyles.js',
								'SmartArts/SmartArtData/*',
								'SmartArts/SmartArtDrawing/*',
								'Images/*',
								'Images/placeholders/*',
								'Images/content_controls/*',
								'Images/cursors/*',
								'Images/reporter/*',
								'Images/icons/*',
								'Native/*.js',
								'libfont/engine/*',
								'spell/spell/*',
								'hash/hash/*',
								'zlib/engine/*'
							],
							dest: path.join(deploy, 'common')
						},
						{
							expand: true,
							cwd: '../cell/css',
							src: '*.css',
							dest: path.join(cell, 'css')
						},
						{
							expand: true,
							cwd: '../slide/themes',
							src: '**/**',
							dest: path.join(slide, 'themes')
						},
						{
							expand: true,
							cwd: '../pdf/',
							src: 'src/engine/*',
							dest: path.join(deploy, 'pdf')
						}
					]
				}
			},
		});
	});
	grunt.registerTask('clean-develop', 'Clean develop scripts', function () {
		const develop = '../develop/sdkjs/';
		grunt.initConfig({
			clean: {
				tmp: {
					options: {
						force: true
					}, src: [develop]
				}
			}
		});
	});
	grunt.registerTask('build-develop', 'Build develop scripts', function () {
		const configs = getConfigs();
		if (!configs.valid()) {
			return;
		}

		writeScripts(configs.word['sdk'], 'word');
		writeScripts(configs.cell['sdk'], 'cell');
		writeScripts(configs.slide['sdk'], 'slide');
	});
	grunt.registerTask('deploy', ['deploy-sdk', 'clean', 'copy']);
	grunt.registerTask('default', ['compile-sdk', 'deploy', 'clean-deploy']);
	grunt.registerTask('develop', ['clean-develop', 'clean', 'build-develop']);
};
