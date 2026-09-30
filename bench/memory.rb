# frozen_string_literal: true

# Writes one workbook in this process so its peak memory can be read by the OS:
#
#   BUNDLE_GEMFILE=bench/Gemfile /usr/bin/time -l bundle exec ruby bench/memory.rb fast_xlsx:low unique  # macOS
#   BUNDLE_GEMFILE=bench/Gemfile /usr/bin/time -v bundle exec ruby bench/memory.rb fast_xlsx:low unique  # Linux
#
# writer: baseline (build the data only), fast_xlsx[:constant|:low], fast_excel[:constant]
# data:   unique (every string differs) or repeated (a few distinct strings)
# Subtract the baseline's peak to get what writing the file adds.
require "fast_xlsx"
require "fileutils"
require "tmpdir"

writer = ARGV[0]
kind = ARGV[1] || "unique"
rows = Integer(ARGV[2] || 200_000)
regions = %w[North South East West Central]
statuses = %w[Open Closed Pending Cancelled]
data = Array.new(rows) do |n|
  label = kind == "unique" ? "String string #{n}" * 5 : regions[n % 5]
  [n, label, statuses[n % 4], Time.at((n * 1000) + 1_492_922_688), n * 100.5]
end
out = File.join(Dir.tmpdir, "fast_xlsx_memory_#{Process.pid}.xlsx")

case writer
when "baseline"
  nil
when /\Afast_xlsx/
  mode = writer.split(":")[1]
  wb = FastXlsx::Workbook.new(constant_memory: mode == "constant", low_memory: mode == "low")
  wb.add_worksheet.concat(data)
  wb.save(out)
when /\Afast_excel/
  require "fast_excel"
  wb = FastExcel.open(out, constant_memory: writer.end_with?(":constant"))
  ws = wb.add_worksheet
  data.each { |r| ws << r }
  wb.close
else
  abort "unknown writer #{writer}"
end
FileUtils.rm_f(out)
