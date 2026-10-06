# frozen_string_literal: true

require_relative "lib/fast_xlsx/version"

Gem::Specification.new do |spec|
  spec.name = "fast_xlsx"
  spec.version = FastXlsx::VERSION
  spec.authors = ["Zac"]
  spec.email = ["579103+7a6163@users.noreply.github.com"]

  spec.summary = "Excel .xlsx writer for Ruby with a Rust engine"
  spec.description = "Writes Excel .xlsx files from Ruby through a native Rust extension (rust_xlsxwriter): " \
                     "20,000 rows in under 100 ms, and memory that stays flat for large exports. Formats, " \
                     "formulas, tables, charts, conditional formats, data validation and more."
  spec.homepage = "https://github.com/7a6163/fast_xlsx"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 3.3.0"
  spec.metadata["github_repo"] = "ssh://github.com/7a6163/fast_xlsx" # links the GitHub Packages gem to this repo
  spec.metadata["homepage_uri"] = spec.homepage
  spec.metadata["source_code_uri"] = spec.homepage
  spec.metadata["changelog_uri"] = "#{spec.homepage}/blob/main/CHANGELOG.md"
  spec.metadata["rubygems_mfa_required"] = "true"

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  gemspec = File.basename(__FILE__)
  spec.files = IO.popen(%w[git ls-files -z], chdir: __dir__, err: IO::NULL) do |ls|
    ls.readlines("\x0", chomp: true).reject do |f|
      (f == gemspec) ||
        f.start_with?(*%w[bin/ Gemfile .gitignore test/ .github/ .rubocop.yml docs/social-preview])
    end
  end
  spec.bindir = "exe"
  spec.executables = spec.files.grep(%r{\Aexe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]
  spec.extensions = ["ext/fast_xlsx/extconf.rb"]

  # Uncomment to register a new dependency of your gem
  # spec.add_dependency "example-gem", "~> 1.0"
  spec.add_dependency "rb_sys", "~> 0.9.130"

  # For more information and examples about making a new gem, check out our
  # guide at: https://bundler.io/guides/creating_gem.html
end
