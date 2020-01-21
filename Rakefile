# frozen_string_literal: true

desc 'Find if any files do not containing license'
task :check_file_without_license do
  license_header = 'Copyright Ascensio System'
  all_js_files = Dir['./**/*.js']
  excluded_pattern = %w[/vendor/ /externs/ /jquery_native.js]
  files_without_license = []
  all_js_files.each do |file|
    next if excluded_pattern.any? { |exclude| file.include?(exclude) }

    unless File.read(file).include?(license_header)
      files_without_license << file
    end
  end
  unless files_without_license.empty?
    raise("Files without license: #{files_without_license}")
  end
end
