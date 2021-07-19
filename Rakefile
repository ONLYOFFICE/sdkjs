# frozen_string_literal: true

excluded_source_paths = %w[/vendor/ /externs/ /jquery_native.js]

desc 'Find if any files do not containing license'
task :check_file_without_license do
  license_header = 'Copyright Ascensio System'
  all_js_files = Dir['./**/*.js']
  files_without_license = []
  all_js_files.each do |file|
    next if excluded_source_paths.any? { |exclude| file.include?(exclude) }

    unless File.read(file).include?(license_header)
      files_without_license << file
    end
  end
  unless files_without_license.empty?
    raise("Files without license: #{files_without_license}")
  end
end

desc 'Find if any files do not containing correct latvian adress'
task :check_file_without_latvian_adress do
  latvian_address = 'LV-1050'
  all_js_files = Dir['./**/*.js']
  files_without_address = []
  all_js_files.each do |file|
    next if excluded_source_paths.any? { |exclude| file.include?(exclude) }

    unless File.read(file).include?(latvian_address)
      files_without_address << file
    end
  end
  unless files_without_address.empty?
    raise("Files without latvian address `#{latvian_address}`: #{files_without_address}")
  end
end

desc 'Find files without LF ending'
task :check_file_without_lf_ending do
  all_files = Dir['./**/*.js']
  files_without_ending = []
  all_files.each do |file|
    files_without_ending << file if File.read(file).include?("\r")
  end
  unless files_without_ending.empty?
    raise("Files without LF ending: #{files_without_ending}")
  end
end

desc 'Check files for trailing newline'
task :check_file_trailing_newline do
  all_files = Dir['./**/*.js']
  files_without_trailing_line = []
  all_files.each do |file|
    unless IO.readlines(file)[-1].end_with?("\n")
      files_without_trailing_line << file
    end
  end
  unless files_without_trailing_line.empty?
    raise("Files without newline at end: #{files_without_trailing_line}")
  end
end

desc 'Perform check on source files'
task :check_source_files do
  Rake::Task['check_file_without_license'].invoke
  Rake::Task['check_file_without_lf_ending'].invoke
  Rake::Task['check_file_trailing_newline'].invoke
  Rake::Task['check_file_without_latvian_adress'].invoke
end
