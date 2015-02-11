# Campus Explorer CSV to Excel-ready data

start_time = Time.now
puts "Starting Script..."

require 'time'
require 'csv'
require 'iconv'
require 'date'
require 'roo'
require 'fileutils'

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
				"Organic",
				"Original Source"]
		# For each row from Campus Explorer CSV File
		counter = 0
		CSV.foreach(input_filename, :headers => true, :return_headers => false, :encoding => 'windows-1251:utf-8') do |row|
			# Process the utm_campaign string as passed through from Source Code into separate values in their own cells
			source_data = process_source_code row["Source Code"]
			counter += 1
			# Is there data?
			if has_campusexplorer_data? row
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
						source_data[:keyword],
						source_data[:match],
						source_data[:ad_id],
						source_data[:ad_page],
						source_data[:ad_top_side],
						source_data[:ad_position],
						source_data[:network],
						source_data[:widget_location],
						source_data[:organic],
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

def process_campaign_adwords_data_file (input_filename, output_filename)
	# Convert to CSV
	adwords_csv_filename = "adwords-campaign-prepped.csv"
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
				"Campaign ID",
				"Est. Impression Share",
				"Total Impressions",
				"Search Lost IS (rank)",
				"Search Lost IS (budget)"]

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
					row["Campaign ID"],
					estimated_impression_share( row["Search Impr. share"] ),
					total_impressions( row["Impressions"], row["Search Impr. share"]),
					estimated_lost_impression_share(row["Search Lost IS (rank)"]),
					estimated_lost_impression_share(row["Search Lost IS (budget)"])
				]
		end
	end
end

def total_impressions( impressions, impression_share)
	impressions = impressions.to_f
	impression_share = estimated_impression_share(impression_share)
	puts impressions.inspect
	puts impression_share.inspect
	(impressions/impression_share).round.to_s
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
		CSV.foreach(bing_csv_filename, :headers => true, :return_headers => false, :encoding => 'utf-8') do |row|			
			csv << [Date.strptime(row["Gregorian date"], '%Y-%m-%d').strftime("%Y-%m-%d %a"),
					row["Impressions"],
					row["Clicks"],
					row["Spend"],
					row["Avg. position"],
					position_weight(row["Impressions"], row["Avg. position"]),
					row["Network"],
					device( row["Device type"] ),
					row["Campaign name"],
					row["Ad group"],
					row["Ad ID"].to_i.to_s]
		end
	end
end

def process_campaign_bing_data_file (input_filename, output_filename)
	# Convert to CSV
	bing_csv_filename = "bing-campaign-prepped.csv"
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
				"Campaign ID",
				"Est. Impression Share",
				"Total Impressions",
				"Search Lost IS (rank)",
				"Search Lost IS (budget)"]

		counter = 0
		CSV.foreach(bing_csv_filename, :headers => true, :return_headers => false, :encoding => 'utf-8') do |row|			
			csv << [Date.strptime(row["Gregorian date"], '%Y-%m-%d').strftime("%Y-%m-%d %a"),
					row["Impressions"],
					row["Clicks"],
					row["Spend"],
					row["Avg. position"],
					position_weight(row["Impressions"], row["Avg. position"]),
					row["Network"],
					device( row["Device type"] ),
					row["Campaign name"],
					row["Campaign ID"],
					estimated_impression_share( row["Impression share (%)"] ),
					total_impressions( row["Impressions"], row["Impression share (%)"]),
					estimated_lost_impression_share( (row["Impression share lost to bid (%)"].to_f + row["Impression share lost to rank (%)"].to_f).to_s),
					estimated_lost_impression_share( row["Impression share lost to budget (%)"] )
				]
		end
	end
end

def combine_all_files(revenue_data_filename, adwords_ad_data_filename, bing_ad_data_filename, adwords_imp_share_data_filename, bing_imp_share_data_filename, output_filename)
	# Open CSVs in Memory
	database_data = CSV.read(output_filename, :headers => true, :return_headers => true, :encoding => 'utf-8')
	adwords_ad_data = CSV.read(adwords_ad_data_filename, :headers => true, :return_headers => false, :encoding => 'utf-8')
	revenue_data = CSV.read(revenue_data_filename, :headers => true, :return_headers => false, :encoding => 'utf-8')
	bing_ad_data = CSV.read(bing_ad_data_filename, :headers => true, :return_headers => false, :encoding => 'utf-8')
	adwords_imp_share_data = CSV.read(adwords_imp_share_data_filename, :headers => true, :return_headers => false, :encoding => 'utf-8')
	bing_imp_share_data = CSV.read(bing_imp_share_data_filename, :headers => true, :return_headers => false, :encoding => 'utf-8')

	# Get earliest date from each file type (should be the same in each file)
	earliest_adwords_ad_date = adwords_ad_data.min{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	earliest_adwords_imp_share_date = adwords_imp_share_data.min{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	earliest_revenue_date = revenue_data.min{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	earliest_bing_ad_date = bing_ad_data.min{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	earliest_bing_imp_share_date = bing_imp_share_data.min{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]


	if earliest_adwords_ad_date != earliest_bing_ad_date ||  
		 earliest_adwords_ad_date != earliest_revenue_date ||
		 earliest_adwords_ad_date != earliest_adwords_imp_share_date ||
		 earliest_adwords_ad_date != earliest_bing_imp_share_date
		raise Exception, "All files do not start on the same date"
	end
	
	earliest_date = earliest_adwords_ad_date
	

	# Get latest date from each file type (should be the same in each file)
	latest_adwords_date = adwords_ad_data.max{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	latest_adwords_imp_share_date = adwords_imp_share_data.max{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	latest_revenue_date = revenue_data.max{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	latest_bing_date = bing_ad_data.max{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]
	latest_bing_imp_share_date = bing_imp_share_data.max{ |a_row, b_row| Date.parse(a_row["Date"]) <=> Date.parse(b_row["Date"]) }["Date"]

	if latest_adwords_date != latest_bing_date || 
		 latest_adwords_date != latest_revenue_date ||
		 latest_adwords_date != latest_adwords_imp_share_date ||
		 latest_adwords_date != latest_bing_imp_share_date
		raise Exception, "All files do not end on the same date"
	end
	
	latest_date = latest_adwords_date

	# Select (i.e keep) those rows whose date is earlier than the earliest date in the new file
	# (note, this delete all other rows...so this removed overlap)
	database_data = database_data.select do |row| 
		if row.header_row? 
			true
		else
			row["Date"] ? Date.parse(row["Date"]) < Date.parse(earliest_date) : false
		end
	end

	#Add AdWords ad data to the database data.
	adwords_ad_data.each do |row|
		headers = database_data.first.headers
		puts "Average Position: " + row["Average Position"].inspect
		fields = [row["Date"],
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
						row["Average Position"],
						row["Position Weight"],
						Date.parse(row["Date"]).strftime('%Y-%m-%d %a'),
						Date.parse(row["Date"]).strftime('%Y-%m-%d'),
						row["Device"],
						row["Campaign"].string_between_markers("[", "]") || "{Not Found}", # Niche
						row["Campaign"].string_between_markers("{", " +") || "{Not Found}", # Seed
						row["Lead Users"],
						"adwords", # Network
						row["Original Source"],
						row["Est. Impression Share"],
						row["Total Impressions"],
						row["Search Lost IS (rank)"],
						row["Search Lost IS (budget)"],
						"paid" #row["Organic"]
					 ]
		database_data << CSV::Row.new(headers, fields)
	end

	#Add Bing ad data to the database data.
	bing_ad_data.each do |row|
		headers = database_data.first.headers
		fields = [row["Date"],
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
						row["Average Position"],
						row["Position Weight"],
						Date.parse(row["Date"]).strftime('%Y-%m-%d %a'),
						Date.parse(row["Date"]).strftime('%Y-%m-%d'),
						row["Device"],
						row["Campaign"].string_between_markers("[", "]") || "{Not Found}", # Niche
						row["Campaign"].string_between_markers("{", " +") || "{Not Found}", # Seed
						row["Lead Users"],
						"BingAds", # Network
						row["Original Source"],
						row["Est. Impression Share"],
						row["Total Impressions"],
						row["Search Lost IS (rank)"],
						row["Search Lost IS (budget)"],
						"paid" #row["Organic"]
					 ]
		database_data << CSV::Row.new(headers, fields)
	end

	# Add AdWords Impression Share data to the database data
	adwords_imp_share_data.each do |row|
		headers = database_data.first.headers
		fields = [row["Date"],
						Date.parse(row["Date"]).strftime('%A'),
						row["Ad ID"],
						row["Campaign"],
						row["Ad Group"],
						"", #row["Impressions"],
						"", #row["Clicks"],
						"", #row["Cost"],
						row["Lead Request Users"],
						row["Leads"],
						row["Clickouts"],
						row["Lead Revenue"],
						row["Clickout Revenue"],
						row["Total Revenue"],
						"", #row["Average Position"],
						"", #row["Position Weight"],
						Date.parse(row["Date"]).strftime('%Y-%m-%d %a'),
						Date.parse(row["Date"]).strftime('%Y-%m-%d'),
						row["Device"],
						row["Campaign"].string_between_markers("[", "]") || "{Not Found}", # Niche
						row["Campaign"].string_between_markers("{", " +") || "{Not Found}", # Seed
						row["Lead Users"],
						"adwords", # Network
						row["Original Source"],
						row["Est. Impression Share"],
						row["Total Impressions"],
						row["Search Lost IS (rank)"],
						row["Search Lost IS (budget)"],
						"paid" #row["Organic"]
					 ]
		database_data << CSV::Row.new(headers, fields)
	end

	# Add Bing Impression Share data to the database data
	bing_imp_share_data.each do |row|
		headers = database_data.first.headers
		fields = [row["Date"],
						Date.parse(row["Date"]).strftime('%A'),
						row["Ad ID"],
						row["Campaign"],
						row["Ad Group"],
						"", #row["Impressions"],
						"", #row["Clicks"],
						"", #row["Cost"],
						row["Lead Request Users"],
						row["Leads"],
						row["Clickouts"],
						row["Lead Revenue"],
						row["Clickout Revenue"],
						row["Total Revenue"],
						"", #row["Average Position"],
						"", # row["Position Weight"],
						Date.parse(row["Date"]).strftime('%Y-%m-%d %a'),
						Date.parse(row["Date"]).strftime('%Y-%m-%d'),
						row["Device"],
						row["Campaign"].string_between_markers("[", "]") || "{Not Found}", # Niche
						row["Campaign"].string_between_markers("{", " +") || "{Not Found}", # Seed
						row["Lead Users"],
						"BingAds", # Network
						row["Original Source"],
						row["Est. Impression Share"],
						row["Total Impressions"],
						row["Search Lost IS (rank)"],
						row["Search Lost IS (budget)"],
						"paid" #row["Organic"]
					 ]
		database_data << CSV::Row.new(headers, fields)
	end

	revenue_data.each do |row|
		campaign = "{Not Found}"
		ad_group = "{Not Found}"

		if ad_id_row = database_data.find { |ad_row| ad_row['Ad ID'] == row["Ad ID"] }
			campaign =  ad_id_row["Campaign"]
			ad_group =  ad_id_row["Ad Group"]
		end

		niche = campaign.string_between_markers "[", "]"
		seed = campaign == "{Not Found}" ? "{Not Found}" : campaign.string_between_markers("{", " +")

		headers = database_data.first.headers
		fields = [row["Date"],
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
						row["Avg. Position"],
						row["Position Weight"],
						Date.parse(row["Date"]).strftime('%Y-%m-%d %a'),
						Date.parse(row["Date"]).strftime('%Y-%m-%d'),
						row["Device2"] || row["Device"],
						niche || "{Not Found}",
						seed || "{Not Found}",
						row["Lead Users"],
						row["Network"],
						row["Original Source"],
						row["Est. Impression Share"],
						row["Total Impressions"],
						row["Search Lost IS (rank)"],
						row["Search Lost IS (budget)"],
						row["Organic"]
					 ]
		database_data << CSV::Row.new(headers, fields)
	end

	temp_output_filename = "!!!" + output_filename
	CSV.open(temp_output_filename, "wb") do |csv|
		database_data.each do |row|
			csv << row
		end
	end
end

def estimated_impression_share (impression_share_string)
	case impression_share_string
	when "< 10%" # When calculating for adwords
		0.05
	when "0.0" # When calculating for bing
		0.02
	when " --", nil
		1
	else
		impression_share_string.to_f / 100
	end
end

def estimated_lost_impression_share (impression_share_string)
	puts impression_share_string
	case impression_share_string
	when "> 90%" # When calculating for adwords
		0.95
	when "100.0" # When calculating for bing
		0.98
	when " --", nil
		0
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
	when "Computer"
		"dt"
	when "Tablet"
		"dt"
	when "Smartphone"
		"mb"
	when "Smartphones"
		"mb"
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
	puts "Converting: #{xlsx_filename} from XLSX to CSV"
	csv_file = File.open(csv_filename, "w")
	xlsx_file = Roo::Excelx.new(xlsx_filename)
	10.upto(xlsx_file.last_row - 2) do |line|
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

def has_campusexplorer_data? row
	sourcecode = row["Source Code"]
	
	(row['Lead Users'] != "0" ||
	row['Lead Request Users'] != "0" ||
  row['Unreconciled Publisher Total Revenue'] != "0.00") &&
  sourcecode != nil &&
  sourcecode != ""
end

def process_source_code (sourcecode)
	if sourcecode.nil?
		sourcecode = ""
	end		

	# Decode Match Type
	match_type = sourcecode.string_between_markers("_m*", "_")
	case match_type
	when "e"
		match_type = "Exact"
	when "p"
		match_type = "Phrase"
	when "b"
		match_type = "Broad"
	end

	organic = (organic? sourcecode) ? "organic" : "paid"

	# Decode Network Type
	network = sourcecode.string_between_markers "_src*", "_"
	network = "adwords" if network.nil?
	network.gsub!("-sitelink", "")
	# Clear network if it is detected as organic
	if organic? sourcecode 
		network = ""
	end

	# Break down ad position
	position_data = sourcecode.string_between_markers "_p*", "_"
	device = sourcecode.string_between_markers("_d*", "_") || sourcecode.string_between_markers("-d*", "_")
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

	keyword = (sourcecode.string_between_markers "_k*", "_")
	keyword = "'" + keyword + "'" unless keyword.nil?

	{
		source: (sourcecode.string_between_markers "_src*", "_") || "",
		campaign_id: (sourcecode.string_between_markers "_x*", "_") || "",
		device: device || "",
		device2: (sourcecode.string_between_markers "_d2*", "_") || "",
		keyword: keyword || "",
		match: match_type || "",
		ad_id: sourcecode.string_between_markers("_c*", "_") || "",
		ad_page: ad_page,
		ad_top_side: ad_top_side,
		ad_position: ad_position,
		network: network,
		widget_location: widget_location,
		organic: organic
	}
end

def organic? sourcecode
	# Does it have a valuetrack tag (i.e. d* for desktop/mobile)?
	# If yes, then it is not organic.  If no, then it is.
	if sourcecode.include?("_d*") || sourcecode.include?("-d*")
		return false
	# If no, does it have uscnaclassesonlne.com or lpnprogramshq.com in the source?
	elsif ( sourcecode.include?("uscnaclassesonline.com") || 
			    sourcecode.include?("lpnprogramshq.com") ||
					(sourcecode == "sa-50216D3D-RightSidebarCNA") ||
					(sourcecode == "sa-50216D3D-RightSidebarLPN") ||
					(sourcecode == "sa-50216D3D-ContentCTAButtonCNA") ||
					(sourcecode == "sa-50216D3D-ContentCTAButtonLPN") ||
					(sourcecode == "sa-50216D3D-RightSidebarCNA-") ||
					(sourcecode == "sa-50216D3D-RightSidebarLPN-") ||
					(sourcecode == "sa-50216D3D-ContentCTAButtonCNA-") ||
					(sourcecode == "sa-50216D3D-ContentCTAButtonLPN-") ||
					(sourcecode == "sa-50216D3D") )
	# If yes, then it is organic
		return true
	else 
		# If none of these apply, then it is a summary row.
		return false
	end
end

def clean_up_directory
	# Force overwrite old DB with new file
	# Do this just in case there is a problem while writing and the DB becomes corrupt
	# so the old version will not be re-written with a corrupt version.
	FileUtils.mv(temp_output_filename, output_filename, {:force => true})
	# FileUtils.rm(Dir.glob "*erformance*") # Delete files with "Performance" (Ad Reports)
	# FileUtils.rm(Dir.glob "*ce-activity-summary*") # Delete CE activity summary files

	FileUtils.rm(Dir.glob "adwords*") # Delete files with "adwords" (Temp Files)
	FileUtils.rm(Dir.glob "bing*") # Delete files with "bing" (Temp Files)
	FileUtils.rm(Dir.glob "*Campus Explorer*") # Delete the temp CE Rev File
end

def update_database
	process_ce_data_file("ce-activity-summary.xls", "Campus Explorer Revenue.csv")
	process_ad_adwords_data_file("Ad performance report.csv", "adwords-ads.csv")
	process_campaign_adwords_data_file("Campaign performance report.csv", "adwords-campaigns.csv")

	begin
		process_ad_bing_data_file("Ad_Performance_Report.xlsx", "bing-ads.csv")
	rescue
		puts "ERROR: The Bing Ad Performance XLSX file failed to be read.\nTry opening and saving file in excel first?"
		exit
	end
	
	begin
		process_campaign_bing_data_file("Campaign_Performance_Report.xlsx", "bing-campaigns.csv")
	rescue
		puts "ERROR: The Bing Campaign Performance XLSX file failed to be read.\nTry opening and saving file in excel first?"
		exit
	end

	combine_all_files("Campus Explorer Revenue.csv","adwords-ads.csv", "bing-ads.csv", "adwords-campaigns.csv", "bing-campaigns.csv", "Koodlu Database.csv")
	# clean_up_directory
end

update_database

puts "Script Complete!"
'say "Script Finished!"'
puts "Time elapsed: #{Time.now - start_time} seconds"