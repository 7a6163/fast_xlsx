# frozen_string_literal: true

# Compares fast_xlsx with other Ruby xlsx writers on the same data.
#
#   BUNDLE_GEMFILE=bench/Gemfile bundle install
#   bundle exec rake compile
#   BUNDLE_GEMFILE=bench/Gemfile bundle exec ruby bench/compare.rb [rows]
require "stringio"

require "fast_xlsx"
require "fast_excel"
require "caxlsx"
require "write_xlsx"
require "xlsxtream"
require "rubyXL"
require "rubyXL/convenience_methods"

def elapsed
  start = Process.clock_gettime(Process::CLOCK_MONOTONIC)
  yield
  Process.clock_gettime(Process::CLOCK_MONOTONIC) - start
end

ROWS = Integer(ARGV[0] || 20_000)
DATA = Array.new(ROWS) do |n|
  [n, "String string #{n}" * 5, n * 7 % 1000, Time.at((n * 1000) + 1_492_922_688), n * 100.5]
end

# Each writer builds one worksheet from DATA and returns the .xlsx bytes.
WRITERS = {
  "fast_xlsx" => lambda {
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.concat(DATA)
    wb.to_xlsx
  },
  "fast_xlsx (constant_memory)" => lambda {
    wb = FastXlsx::Workbook.new(constant_memory: true)
    wb.add_worksheet.concat(DATA)
    wb.to_xlsx
  },
  "fast_xlsx (low_memory)" => lambda {
    wb = FastXlsx::Workbook.new(low_memory: true)
    wb.add_worksheet.concat(DATA)
    wb.to_xlsx
  },
  "fast_excel" => lambda {
    wb = FastExcel.open
    ws = wb.add_worksheet
    DATA.each { |r| ws << r }
    wb.read_string
  },
  "fast_excel (constant_memory)" => lambda {
    wb = FastExcel.open(constant_memory: true)
    ws = wb.add_worksheet
    DATA.each { |r| ws << r }
    wb.read_string
  },
  "write_xlsx" => lambda {
    io = StringIO.new
    wb = WriteXLSX.new(io)
    ws = wb.add_worksheet
    DATA.each_with_index { |r, i| ws.write_row(i, 0, r) }
    wb.close
    io.string
  },
  "xlsxtream" => lambda {
    io = StringIO.new
    Xlsxtream::Workbook.open(io) do |xlsx|
      xlsx.write_worksheet("Sheet1") { |ws| DATA.each { |r| ws << r } }
    end
    io.string
  },
  "caxlsx" => lambda {
    package = Axlsx::Package.new
    package.workbook.add_worksheet { |ws| DATA.each { |r| ws.add_row(r) } }
    package.to_stream.read
  },
  "rubyXL" => lambda {
    wb = RubyXL::Workbook.new
    ws = wb[0]
    DATA.each_with_index { |r, i| r.each_with_index { |v, j| ws.add_cell(i, j, v) } }
    wb.stream.read
  }
}.freeze

def measure(writer, runs)
  bytes = writer.call # warm up
  times = Array.new(runs) do
    GC.start
    elapsed { writer.call }
  end
  GC.start
  before = GC.stat(:total_allocated_objects)
  writer.call
  { ms: times.sort[runs / 2] * 1000, allocs: GC.stat(:total_allocated_objects) - before, kb: bytes.bytesize / 1024 }
end

puts "#{ROWS} rows x 5 columns (integer, string, integer, Time, float), Ruby #{RUBY_VERSION}, #{RUBY_PLATFORM}"
puts

results = WRITERS.to_h do |name, writer|
  runs = name == "rubyXL" ? 3 : 7
  [name, measure(writer, runs)]
end

fastest = results.values.map { |r| r[:ms] }.min
puts "| Library | Time (median) | vs fastest | Ruby objects allocated | Output |"
puts "|---|---:|---:|---:|---:|"
results.sort_by { |_, r| r[:ms] }.each do |name, r|
  puts format("| %<name>s | %<ms>.0f ms | %<x>.1fx | %<allocs>s | %<kb>d KB |",
              name: name, ms: r[:ms], x: r[:ms] / fastest,
              allocs: r[:allocs].to_s.reverse.scan(/\d{1,3}/).join(",").reverse, kb: r[:kb])
end
