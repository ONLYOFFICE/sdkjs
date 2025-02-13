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
		let files = ['../vendor/polyfill.js', '../common/AllFonts.js'];
		if (grunt.option('compiled')) {
			//todo set window['AscNotLoadAllScript'] = false; (in applyDocumentChanges.js)
			files.push(deploy + name + '/sdk-all-min.js');
		} else {
			files = files.concat(['../common/applyDocumentChanges.js'], getFilesMin(config), getFilesAll(config));
		}
		fixUrl(files, '../../../../sdkjs/build/');

		grunt.file.write(path.join(develop, name, fileName), 'var sdk_scripts = [\n\t"' + files.join('",\n\t"') + '"\n];');
	}

	function CConfig(pathConfigs) {
		this.externs = null;
		this.word = null;
		this.cell = null;
		this.slide = null;
		this.visio = null;

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
		appendOption.call(this, 'visio');
	};
	CConfig.prototype.valid = function () {
		return this.externs && this.word && this.cell && this.slide && this.visio;
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
	const deploy = '../deploy/sdkjs/';
	const word = path.join(deploy, 'word');
	const cell = path.join(deploy, 'cell');
	const slide = path.join(deploy, 'slide');
	const visio = path.join(deploy, 'visio');

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
	const otherFiles = [
		{
			cwd: '../common/',
			src: [
				'Drawings/Format/path-boolean-min.js',
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
				'zlib/engine/*',
				'serviceworker/*'
			],
			dest: path.join(deploy, 'common'),
			name: 'common'
		},
		{
			cwd: '../cell/css',
			src: ['*.css'],
			dest: path.join(cell, 'css'),
			name: 'cell-css'
		},
		{
			cwd: '../slide/themes',
			src: ['**/**'],
			dest: path.join(slide, 'themes'),
			name: 'slide-themes'
		},
		{
			cwd: '../pdf/',
			src: [
				'src/engine/*',
				'src/annotations/stamps.json'
			],
			dest: path.join(deploy, 'pdf'),
			name: 'pdf'
		}
	];
	const configWord = configs.word['sdk'];
	const configCell = configs.cell['sdk'];
	const configSlide = configs.slide['sdk'];
	const configVisio = configs.visio['sdk'];

	const compilerArgs = getExterns(configs.externs);
	if (formatting) {
		compilerArgs.push('--formatting=' + formatting);
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

	function getCompileConfig(sdkmin, sdkall, outmin, outall, name, pathPrefix) {
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
		`--chunk_wrapper=${outall}:${license}\n(function(window, undefined) {%s})(window);`,
		`--chunk_output_path_prefix=${pathPrefix}`);
		if (grunt.option('map')) {
			grunt.file.mkdir(path.join('./maps'));
			args.push('--property_renaming_report=' + path.join(`maps/${name}.props.js.map`));
			args.push('--variable_renaming_report=' + path.join(`maps/${name}.vars.js.map`));
			args.push('--create_source_map=' + path.join(`%outname%.map`));
			args.push('--source_map_format=V3');
			args.push('--source_map_include_content=true');
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
		grunt.initConfig(getCompileConfig(getFilesMin(configWord), getFilesAll(configWord), 'sdk-all-min', 'sdk-all', 'word', path.join(word , '/')));
		grunt.task.run('closure-compiler');
	});
	grunt.registerTask('compile-cell', 'Compile Cell SDK', function () {
		grunt.initConfig(getCompileConfig(getFilesMin(configCell), getFilesAll(configCell), 'sdk-all-min', 'sdk-all', 'cell', path.join(cell , '/')));
		grunt.task.run('closure-compiler');
	});
	grunt.registerTask('compile-slide', 'Compile Slide SDK', function () {
		grunt.initConfig(getCompileConfig(getFilesMin(configSlide), getFilesAll(configSlide), 'sdk-all-min', 'sdk-all', 'slide', path.join(slide , '/')));
		grunt.task.run('closure-compiler');
	});
	grunt.registerTask('compile-visio', 'Compile Visio SDK', function () {
		grunt.initConfig(getCompileConfig(getFilesMin(configVisio), getFilesAll(configVisio), 'sdk-all-min', 'sdk-all', 'visio', path.join(visio , '/')));
		grunt.task.run('closure-compiler');
	});
	grunt.registerTask('copy-maps', 'Copy maps from deploy to build', function() {
		grunt.initConfig({
			copy: {
				word: {
					files: [
						{
							expand: true,
							cwd: word,
							src: [
								'sdk-all-min.js.map',
								'sdk-all.js.map',
							],
							dest: 'maps',
							rename: function (dest, src) {
								return path.join(dest , src.replace('sdk', 'word'));
							}
						}
					]
				},
				cell: {
					files: [
						{
							expand: true,
							cwd: cell,
							src: [
								'sdk-all-min.js.map',
								'sdk-all.js.map',
							],
							dest: 'maps',
							rename: function (dest, src) {
								return path.join(dest , src.replace('sdk', 'cell'));
							}
						}
					]
				},
				slide: {
					files: [
						{
							expand: true,
							cwd: slide,
							src: [
								'sdk-all-min.js.map',
								'sdk-all.js.map',
							],
							dest: 'maps',
							rename: function (dest, src) {
								return path.join(dest , src.replace('sdk', 'slide'));
							}
						}
					]
				},
				visio: {
					files: [
						{
							expand: true,
							cwd: visio,
							src: [
								'sdk-all-min.js.map',
								'sdk-all.js.map',
							],
							dest: 'maps',
							rename: function (dest, src) {
								return path.join(dest , src.replace('sdk', 'visio'));
							}
						}
					]
				}
			},
			clean: {
				deploy: {
					options: {
						force: true
					},
					src: [
						path.join(word, 'sdk-all-min.js.map'),
						path.join(word, 'sdk-all.js.map'),
						path.join(cell, 'sdk-all-min.js.map'),
						path.join(cell, 'sdk-all.js.map'),
						path.join(slide, 'sdk-all-min.js.map'),
						path.join(slide, 'sdk-all.js.map'),
						path.join(visio, 'sdk-all-min.js.map'),
						path.join(visio, 'sdk-all.js.map'),
					]
				}
			}
		});
		grunt.task.run('copy', 'clean');
	});
	grunt.registerTask('compile-sdk', ['compile-word', 'compile-cell', 'compile-slide', 'compile-visio']);
	grunt.registerTask('clean-deploy', 'Clean deploy folder before deploying', function () {
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
			}
		});
		grunt.task.run('clean');
	});
	const glob = require('glob');
	const ignoreFiles = ['jquery_native', 'fonts_ie', 'spell_ie', 'engine_ie', 'zlib_ie', 'drawingfile_ie', 'themes'];
	/**
	 * @param {string[]} paths
	 * @param {string} cwd
	 * @return {[string[], string[]]}
	 */
	function splitJSFiles(paths, cwd) {
		const jsFiles = [];
		const noJSFiles = [];
		paths.forEach((p) => {
			glob.sync(p, {
				cwd: cwd,
			}).forEach((f) => {
				if (path.extname(f) === '.js' && !ignoreFiles.includes(path.parse(f).name)) {
					jsFiles.push(path.join(f));
				} else {
					noJSFiles.push(path.join(f));
				}
			})
		});
		return [jsFiles, noJSFiles];
	}
	function getOtherCompileConfig(o, jsFile) {
		return {
			'closure-compiler': {
				js: {
					options: {
						args: [
							'--language_out=ECMASCRIPT5',
							'--compilation_level=WHITESPACE_ONLY',
							'--rewrite_polyfills=true',
							'--warning_level=QUIET',
							`--js=${path.join(o.cwd, jsFile)}`,
							`--js_output_file=${path.join(o.dest, jsFile)}`,
							`--output_wrapper=${license}\n%output%`
						]
					}
				}
			}
		}
	}
	function getOtherCopyConfig(o, noJSFiles) {
		return {
			copy: {
				sdkjs: {
					files: noJSFiles.map(f => ({
						expand: true,
						cwd: o.cwd,
						src: f,
						dest: o.dest
					}))
				}
			}
		}
	}
	grunt.registerTask('copy-other', 'Copy other SDK files', function () {
		const compilerTasks = [];
		const copyTasks = [];
		otherFiles.forEach((o) => {
			const [jsFiles, noJSFiles] = splitJSFiles(o.src, o.cwd);
			if (jsFiles.length !== 0) {
				jsFiles.forEach((f) => {
					grunt.registerTask(`compile-${path.join(o.dest, f)}`, `Compiling ${path.join(o.dest, f)}`, function() {
						grunt.initConfig(getOtherCompileConfig(o, f));
						grunt.task.run('closure-compiler');
					});
					compilerTasks.push(`compile-${path.join(o.dest, f)}`);
				});
			}
			if (noJSFiles.length !== 0) {
				grunt.registerTask(`copy-${path.normalize(o.name)}`, `Copying files ${path.normalize(o.name)}`, function() {
					grunt.initConfig(getOtherCopyConfig(o, noJSFiles));
					grunt.task.run('copy');
				});
				copyTasks.push(`copy-${path.normalize(o.name)}`);
			}
		});
		grunt.task.run(compilerTasks);
		grunt.task.run(copyTasks);
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
		grunt.task.run('clean');
	});
	grunt.registerTask('build-develop', 'Build develop scripts', function () {
		const configs = getConfigs();
		if (!configs.valid()) {
			return;
		}

		writeScripts(configs.word['sdk'], 'word');
		writeScripts(configs.cell['sdk'], 'cell');
		writeScripts(configs.slide['sdk'], 'slide');
		writeScripts(configs.visio['sdk'], 'visio');
	});
	const defaultTasks = ['clean-deploy', 'compile-sdk', 'copy-other'];
	if (grunt.option('map')) {
		defaultTasks.push('copy-maps');
	}
	grunt.registerTask('default', defaultTasks);
	grunt.registerTask('develop', ['clean-develop', 'build-develop']);
};
