# Campus Explorer CSV to Excel-ready data

start_time = Time.now
puts "Starting Script..."

require 'csv'
require 'iconv'
require 'date'
require 'roo'

class String
  def string_between_markers marker1, marker2
    self[/#{Regexp.escape(marker1)}(.*?)(#{Regexp.escape(marker2)}|\z)/m, 1]
  end
end

# Processes a CE data CSV or TSV-XLS File into output
def process_ce_data_file (input_filename, output_filename)
	#Check if the input file is xls.  If so, change to CSV
	if input_filename.include? "xls"
		csv_filename = input_filename.gsub "xls", "csv"
		ce_tsv_to_csv(input_filename, csv_filename)
		input_filename = csv_filename
	end

	CSV.open(output_filename, "wb") do |csv|
		# Create Header Row
		csv << ["Date",
				"Widget Impressions",
				"Lead Request Users",
				"Lead Users",
				"Leads",
				"Clickout Impressions",
				"Clickouts",
				"Lead Revenue",
				"Clickout Revenue",
				"Total Revenue",
				"Landing Page",
				"Source",
				"Campaign ID",
				"Device",
				"Device2",
				"Keyword",
				"Match",
				"Ad ID",
				"Ad Page",
				"Ad Top/Side",
				"Ad Position",
				"Network",
				"Widget Location",
				"Original Source"]
		# For each row from Campus Explorer CSV File
		CSV.foreach(input_filename, :headers => true, :return_headers => false, :encoding => 'windows-1251:utf-8') do |row|
			# Process the utm_campaign string as passed through from Source Code into separate values in their own cells
			source_data = process_source_code row["Source Code"]
			# Is there data?
			if has_campusexplorer_data? row, source_data
				# Write ALL values out to processed CSV file
				csv << [row["Grouping"],
						row["Widget Impressions"],
						row["Lead Request Users"],
						row["Lead Users"],
						row["Leads"],
						row["Clickout Impressions"],
						row["Clickouts"],
						row["Unreconciled Publisher Lead Revenue"],
						row["Unreconciled Publisher Clickout Revenue"],
						row["Unreconciled Publisher Total Revenue"],
						source_data[:lp],
						source_data[:source],
						source_data[:campaign_id].gsub("]",""),
						source_data[:device],
						source_data[:device2],
						"'" + source_data[:keyword] + "'",
						source_data[:match],
						source_data[:ad_id],
						source_data[:ad_page],
						source_data[:ad_top_side],
						source_data[:ad_position],
						source_data[:network],
						source_data[:widget_location],
						row["Source Code"]
						]
			end
		end
	end
end

def process_ad_adwords_data_file (input_filename, output_filename)
	# Convert to CSV
	adwords_csv_filename = "adwords-prepped.csv"
	adwords_tsv_to_csv input_filename, adwords_csv_filename
	CSV.open(output_filename, "wb") do |csv|
		# Create Header Row
		csv << ["Date",
				"Impressions",
				"Clicks",
				"Cost",
				"Average Position",
				"Position Weight",
				"Network",
				"Device",
				"Campaign",
				"Ad Group",
				"Ad ID"]
		counter = 0
		CSV.foreach(adwords_csv_filename, :headers => true, :return_headers => false, :encoding => 'utf-8') do |row|			
			csv << [Date.strptime(row["Day"], '%Y-%m-%d').strftime("%Y-%m-%d %a"),
					row["Impressions"],
					row["Clicks"],
					row["Cost"],
					row["Avg. position"],
					position_weight(row["Impressions"], row["Avg. position"]),
					row["Network (with search partners)"],
					device( row["Device"] ),
					row["Campaign"],
					row["Ad group"],
					row["Ad ID"]]
		end
	end
end

def process_ad_bing_data_file (input_filename, output_filename)
	# Convert to CSV
	bing_csv_filename = "bing-prepped.csv"
	bing_xlsx_to_csv input_filename, bing_csv_filename
	CSV.open(output_filename, "wb") do |csv|
		# Create Header Row
		csv << ["Date",
				"Impressions",
				"Clicks",
				"Cost",
				"Average Position",
				"Position Weight",
				"Network",
				"Device",
				"Campaign",
				"Ad Group",
				"Ad ID"]
		counter = 0
		CSV.foreach(bing_csv_filename, :headers => true, :return_headers => false, :encoding => 'utf-8') do |row|			
			csv << [Date.strptime(row["Day"], '%Y-%m-%d').strftime("%Y-%m-%d %a"),
					row["Impressions"],
					row["Clicks"],
					row["Cost"],
					row["Avg. position"],
					position_weight(row["Impressions"], row["Avg. position"]),
					row["Network (with search partners)"],
					device( row["Device"] ),
					row["Campaign"],
					row["Ad group"],
					row["Ad ID"]]
		end
	end
end

def combine_all_files(revenue_data_filename, ad_data_filename, output_filename)
	CSV.open(output_filename, "wb") do |csv|
		# Create Header Row
		csv << ["Date",
						"Day Of Week",
						"Ad ID",
						"Campaign",
						"Adgroup",
						"Impressions",
						"Clicks",
						"Cost",
						"Lead Request Users",
						"Leads",
						"Clickouts",
						"Lead Revenue",
						"Clickout Revenue",
						"Total Revenue",
						"Position Weight",
						"YYYY-MM-DD DDD",
						"W/C Date",
						"Device",
						"Niche",
						"SEED",
						"Lead Users",
						"Network",
						"Original Source"
						]

		
		ad_data = CSV.read(ad_data_filename, :headers => true, :return_headers => false, :encoding => 'utf-8')
		revenue_data = CSV.read(revenue_data_filename, :headers => true, :return_headers => false, :encoding => 'utf-8')

		ad_data.each do |row|
			csv << [row["Date"],
							Date.parse(row["Date"]).strftime('%A'),
							row["Ad ID"],
							row["Campaign"],
							row["Ad Group"],
							row["Impressions"],
							row["Clicks"],
							row["Cost"],
							row["Lead Request Users"],
							row["Leads"],
							row["Clickouts"],
							row["Lead Revenue"],
							row["Clickout Revenue"],
							row["Total Revenue"],
							row["Position Weight"],
							Date.parse(row["Date"]).strftime('%Y-%m-%d %a'),
							Date.parse(row["Date"]).strftime('%Y-%m-%d'),
							row["Device"],
							row["Campaign"].string_between_markers("[", "]") || "{Not Found}", # Niche
							row["Campaign"].string_between_markers("{", " +") || "{Not Found}", # Seed
							row["Lead Users"],
							"adwords", # Network
							row["Original Source"]
						 ]
		end
		
		revenue_data.each do |row|
			ad_id_row = ad_data.find {|ad_row| ad_row['Ad ID'] == row["Ad ID"]}
			campaign =  ad_id_row.nil? ? "{Not Found}" : ad_id_row["Campaign"]
			ad_group =  ad_id_row.nil? ? "{Not Found}" : ad_id_row["Ad Group"]
			niche = campaign.string_between_markers "[", "]"
			seed = campaign == "{Not Found}" ? "{Not Found}" : campaign.string_between_markers("{", " +")

			csv << [row["Date"],
							Date.parse(row["Date"]).strftime('%A'),
							row["Ad ID"],
							campaign,
							ad_group,
							row["Impressions"],
							row["Clicks"],
							row["Cost"],
							row["Lead Request Users"],
							row["Leads"],
							row["Clickouts"],
							row["Lead Revenue"],
							row["Clickout Revenue"],
							row["Total Revenue"],
							row["Position Weight"],
							Date.parse(row["Date"]).strftime('%Y-%m-%d %a'),
							Date.parse(row["Date"]).strftime('%Y-%m-%d'),
							row["Device"],
							niche || "{Not Found}",
							seed || "{Not Found}",
							row["Lead Users"],
							row["Network"],
							row["Original Source"]
						 ]
		end
	end		
end

def estimated_impression_share (impression_share_string)
	case impression_share_string
	when "< 10%"
		0.05
	when " --"
		1
	else
		impression_share_string.to_f / 100
	end
end

def estimated_searches(impressions, est_impression_share)
	impressions.to_f / est_impression_share.to_f
end

def position_weight (impressions, avg_position)
	impressions.to_f * avg_position.to_f
end

def device(device_string)
	case device_string
	when "Mobile devices with full browsers"
		"mb"
	when "Computers"
		"dt"
	when "Tablets with full browsers"
		"dt"
	end

end	

def adwords_tsv_to_csv (tsv_filename, csv_filename)
	CSV.open(csv_filename, "wb:utf-8") do |csv|
		File.open(tsv_filename, "rb:utf-16le") do |file|
			counter = 0
			file.each_line do |tsv|
				#Remove first 5 rows of header data
				if counter > 4
					tsv = tsv.encode('utf-8')
					tsv.chomp!
					tsv.gsub!("\"","")
					row_data = tsv.split(/\t/)
					csv << row_data unless row_data[0] == "Total"
				end
				counter = counter + 1
			end
		end
	end
end	

def bing_xlsx_to_csv (xlsx_filename, csv_filename)
	csv_file = File.open(csv_filename, "w")
	xlsx_file = Roo::Excelx.new(xlsx_filename)
	10.upto(xlsx_file.last_row) do |line|
  	csv_file.write CSV.generate_line xlsx_file.row(line)
	end
end	

def ce_tsv_to_csv (tsv_filename, csv_filename)
	CSV.open(csv_filename, "wb") do |csv|
		File.open(tsv_filename) do |file|
			counter = 0
			file.each_line do |tsv|
				tsv.chomp!
				tsv.gsub!('"','')
				csv << tsv.split(/\t/)
				counter = counter + 1
			end
		end
	end
end

def has_campusexplorer_data? (row, source_data)
	# => YES -> if source parameter is set and is adwords
	# row["Unreconciled Publisher Total Revenue"].to_f > 0 &&
	source_data[:source] == "adwords"
end

def process_source_code (sourcecode)
	if sourcecode.nil?
		sourcecode = ""
	end		

	# Decode Match Type
	match_type = sourcecode.string_between_markers "_m*", "_"
	case match_type
	when "e"
		match_type = "Exact"
	when "p"
		match_type = "Phrase"
	when "b"
		match_type = "Broad"
	end

	# Decode Network Type
	network = sourcecode.string_between_markers "_src*", "_" || "adwords"

	# Break down ad position
	position_data = sourcecode.string_between_markers "_p*", "_"
	device = sourcecode.string_between_markers "_d*", "_"
	unless position_data.nil? || position_data == "none"
		ad_page = position_data[0]
		ad_position = position_data[2]
		ad_top_side = position_data[1]
		case ad_top_side
		when "t"
			ad_top_side = "Top"
		when "s"
			ad_top_side = "Side"
		when "o"			
			ad_top_side = "Other"
			ad_top_side = "Bottom"  if device == "dt"
			ad_top_side = "Mobile" if device == "mb"
		end
	end

	# Set Widget Location
	if sourcecode.include? "RightSidebar"
		widget_location = "Right Sidebar"
	elsif sourcecode.include? "ContentCTA"
		widget_location = "Content CTA Lightbox"
	end

	{ 	
		lp: lp,
		source: (sourcecode.string_between_markers "_src*", "_"),
		campaign_id: (sourcecode.string_between_markers "_x*", "_") || "",
		device: device,
		device2: (sourcecode.string_between_markers "_d2*", "_"),
		keyword: (sourcecode.string_between_markers "_k*", "_"),
		match: match_type,
		ad_id: sourcecode.string_between_markers("_c*", "_"),
		ad_page: ad_page,
		ad_top_side: ad_top_side,
		ad_position: ad_position,
		network: network,
		widget_location: widget_location,
	}
end

#process_ce_data_file("ce-activity-summary.xls", "Campus Explorer Revenue.csv")
#process_ad_adwords_data_file("Ad performance report.csv", "adwords-ads.csv")
process_ad_bing_data_file("Ad_Performance_Report.xlsx", "bing-ads.csv")
#combine_all_files("Campus Explorer Revenue.csv","adwords-ads.csv", "final_output.csv")

puts "Script Complete!"
puts "Time elapsed: #{Time.now - start_time} seconds"