# Campus Explorer CSV to Excel-ready data

start_time = Time.now
puts "Starting Script..."

require 'csv'

class String
  def string_between_markers marker1, marker2
    self[/#{Regexp.escape(marker1)}(.*?)(#{Regexp.escape(marker2)}|\z)/m, 1]
  end
end

# Processes a CE data **CSV** File into output
def process_ce_csv (input_filename, output_filename)
	#Check if the input file is xls.  If so, change to CSV
	if input_filename.include? "xls"
		csv_filename = input_filename.gsub "xls", "csv"
		xls_to_csv(input_filename, csv_filename)
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
				"Creative",
				"Ad Page",
				"Ad Top/Side",
				"Ad Position",
				"Network",
				"Widget Location",
				"Niche"]
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
						source_data[:campaign_id],
						source_data[:device],
						source_data[:device2],
						source_data[:keyword],
						source_data[:match],
						source_data[:creative],
						source_data[:ad_page],
						source_data[:ad_top_side],
						source_data[:ad_position],
						source_data[:network],
						source_data[:widget_location],
						source_data[:niche]
						]
			end
		end
	end
end

def xls_to_csv (xls_filename, csv_filename)
	CSV.open(csv_filename, "wb") do |csv|
	  File.open(xls_filename) do |file|
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

def get_input_filename
	filename = ARGV[0]
	if filename.nil?
		puts "Enter CampusExplorer File Name:"
		filename = gets.chomp
	end
	return filename
end

def get_output_filename
	filename = ARGV[1]
	if filename.nil?
		puts "Enter Output File Name:"
		filename = gets.chomp
	end
	return filename
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
	network = sourcecode.string_between_markers "_n*", "_"
	case network
	when "g"
		network = "Google"
	when "s"
		network = "Search"
	when "d"
		network = "Display"
	end

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

	# Decode Niche
	lp = sourcecode.string_between_markers "-lp*", "_"
	unless lp.nil?	
		niche = "CNA" if lp.include? "cna"
		niche = "LPN" if lp.include? "lpn"
	end

	{ 	
		lp: lp,
		source: (sourcecode.string_between_markers "_src*", "_"),
		campaign_id: (sourcecode.string_between_markers "_x*", "_"),
		device: device,
		device2: (sourcecode.string_between_markers "_d2*", "_"),
		keyword: (sourcecode.string_between_markers "_k*", "_"),
		match: match_type,
		creative: (sourcecode.string_between_markers "_c*", "_"),
		ad_page: ad_page,
		ad_top_side: ad_top_side,
		ad_position: ad_position,
		network: network,
		widget_location: widget_location,
		niche: niche
	}
end

input_filename = get_input_filename
output_filename = get_output_filename
process_ce_csv(input_filename, output_filename)

puts "Script Complete!"
puts "Time elapsed: #{Time.now - start_time} seconds"