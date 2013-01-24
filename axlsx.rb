require 'rubygems'
require 'axlsx'

p = Axlsx::Package.new

p.use_shared_strings = true

p.workbook do |wb|
	
	# worksheet names
	text_ws = 'Text format'
	chart_ws = 'Simple Pie Chart'
	merge_cells_ws = 'Merging Cells'
	
	# define customized styles 
	styles = wb.styles
	title = styles.add_style :sz => 15, :b => true, :u => true
	header = styles.add_style :bg_color => '00', :fg_color => 'ff', :b => true
	money = styles.add_style :format_code => "#,###,##0", :border => Axlsx::STYLE_THIN_BORDER
	percent = styles.add_style :num_fmt => Axlsx::NUM_FMT_PERCENT, :border => Axlsx::STYLE_THIN_BORDER		
		
	wb.add_worksheet(:name => text_ws) do |ws|
		ws.add_row ["funny"], :style => title	
		ws.add_row
		ws.add_row ['Quarter', 'Profit', '% of Total'], :style => header
		ws.add_row ['Q1-2010', '153451120976', '=B4/SUM(B4:B7)'], :style => [title, money, percent]			
	end
	
	wb.add_worksheet(:name => chart_ws) do |ws|
		ws.add_row ["Simpel Pie Chart"], :style => title

		%w(first second third).each { |label| ws.add_row [label, rand(24)+1] }
		ws.add_chart(Axlsx::Pie3DChart, :start_at => [0,5], :end_at => [10, 20], :title => "example 3: Pie Chart") do |chart|
      		chart.add_series :data => ws["B2:B4"], :labels => ws["A2:A4"],  :colors => ['FF0000', '00FF00', '0000FF']
      	end    
	end
	
	wb.add_worksheet(:name => merge_cells_ws) do |ws|
		# cell level style overides when adding cells
		ws.add_row ["col 1", "col 2", "col 3", "col 4"], :sz => 16
		ws.add_row [1, 2, 3, "=SUM(A2:C2)"]
		ws.add_row [2, 3, 4, "=SUM(A3:C3)"]
		ws.add_row ["total", "", "", "=SUM(D2:D3)"]
		ws.merge_cells("A4:C4")
		ws["A1:D1"].each { |c| c.color = "FF0000"}
		ws["A1:D4"].each { |c| c.style = Axlsx::STYLE_THIN_BORDER }
	end
	
	wb.add_worksheet(:name => "Table") do |ws|
		ws.add_row ["Build Matrix"]
		ws.add_row ["Build", "Duration", "Finished", "Rvm"]
		ws.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
		ws.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
		ws.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
		ws.add_table "A2:D5", :name => 'Build Matrix', :style_info => { :name => "TableStyleMedium23" }
	end
end

p.serialize 'getting_barred.xlsx'

