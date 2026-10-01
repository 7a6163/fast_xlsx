require "date"
require "fast_xlsx"

# Sample data; in an app it comes from your database.
Sale = Struct.new(:region, :product, :units, :price, :sold_on)
sales = [
  Sale.new("North", "Widget", 120, 9.5, Date.new(2026, 7, 3)),
  Sale.new("North", "Gadget", 45, 24.0, Date.new(2026, 7, 9)),
  Sale.new("South", "Widget", 80, 9.5, Date.new(2026, 7, 14)),
  Sale.new("South", "Gizmo", 12, 120.0, Date.new(2026, 7, 21))
]

# 1. A workbook with one worksheet.
wb = FastXlsx::Workbook.new
ws = wb.add_worksheet("Sales")

# 2. A header row, bold on a light blue fill.
header = FastXlsx::Format.new(bold: true, bg_color: "#DDEBF7", border_bottom: :thin)
ws.append(["Region", "Product", "Units", "Price", "Sold on", "Revenue"], format: header)

# 3. One row per sale. Dates show as yyyy-mm-dd on their own; a formula is a
#    FastXlsx::Formula. In formulas rows count from 1, so the first sale is row 2.
money = FastXlsx::Format.new(num_format: "#,##0.00")
sales.each.with_index(2) do |sale, row|
  ws.append([sale.region, sale.product, sale.units, sale.price, sale.sold_on,
             FastXlsx::Formula.new("C#{row}*D#{row}")],
            format: [nil, nil, nil, money, nil, money])
end

# 4. A totals row: the header's look, with the money format added.
last = sales.size + 1
ws.append(["Total", nil, FastXlsx::Formula.new("SUM(C2:C#{last})"), nil, nil,
           FastXlsx::Formula.new("SUM(F2:F#{last})")],
          format: [header, header, header, header, header, header.merge(num_format: "#,##0.00")])

# 5. Easier to read: columns sized to fit, the header kept in view while
#    scrolling, and filter buttons on the header (rows and columns count from 0).
ws.autofit
ws.freeze_panes("A2")
ws.autofilter("A1:F#{last}")

# 6. Save it. (In Rails: send_data wb.to_xlsx, filename: "sales.xlsx", ...)
wb.save("sales.xlsx")
