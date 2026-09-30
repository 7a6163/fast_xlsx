# frozen_string_literal: true

# Usage: bundle exec rake compile && ruby -Ilib bench/write.rb
# Compares against fast_excel when it can be loaded (FAST_EXCEL=/path/to/fast_excel/lib/fast_excel).
require "fast_xlsx"

def elapsed
  start = Process.clock_gettime(Process::CLOCK_MONOTONIC)
  yield
  Process.clock_gettime(Process::CLOCK_MONOTONIC) - start
end

ROWS = 20_000
DATA = Array.new(ROWS) do |n|
  [n, "String string #{n}" * 5, n * 7 % 1000, Time.at((n * 1000) + 1_492_922_688), n * 100.5]
end

def report(label, runs = 7, &block)
  times = Array.new(runs) do
    GC.start
    elapsed(&block)
  end
  puts "  #{label.ljust(26)} #{(times.sort[runs / 2] * 1000).round(1)} ms"
end

begin
  require ENV.fetch("FAST_EXCEL", "fast_excel")
rescue LoadError
  nil
end

[false, true].each do |cm|
  puts "constant_memory=#{cm}, #{ROWS}x5 cells, median of 7"
  report("fast_xlsx <<") do
    wb = FastXlsx::Workbook.new(memory: cm ? :constant : :standard)
    ws = wb.add_worksheet
    DATA.each { |r| ws << r }
    wb.to_xlsx
  end
  report("fast_xlsx concat") do
    wb = FastXlsx::Workbook.new(memory: cm ? :constant : :standard)
    wb.add_worksheet.concat(DATA)
    wb.to_xlsx
  end
  next unless defined?(FastExcel)

  report("fast_excel <<") do
    wb = FastExcel.open(constant_memory: cm)
    ws = wb.add_worksheet
    DATA.each { |r| ws << r }
    wb.read_string
  end
end
