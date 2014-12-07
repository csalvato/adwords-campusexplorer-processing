# Campus Explorer CSV to Excel-ready data

start_time = Time.now
puts "Starting Script..."

require 'csv'

class String
  def string_between_markers marker1, marker2
    self[/#{Regexp.escape(marker1)}(.*?)#{Regexp.escape(marker2)}/m, 1]
  end
end

# Processes a CE data **CSV** File into output
def process_ce_csv (input_filename, output_filename)
	CSV.open(output_filename, "wb") do |csv|
	# For each row from Campus Explorer CSV File
		CSV.foreach(input_filename, :headers => true, :return_headers => false, :encoding => 'windows-1251:utf-8') do |row|
		# Process the utm_campaign string as passed through from Source Code into separate values in their own cells
		# Is there data?
			if has_campusexplorer_data? row
				puts process_source_code row["Source Code"]
				# Write ALL values out to processed CSV file
			end
		end
	end
end

def has_campusexplorer_data? (row)
	# => YES -> If there's revenue and a source code (maybe it's if there is any utm_campaign string appended at all?)	
	row["Unreconciled Publisher Total Revenue"].to_f > 0 &&
	!row["Source Code"].empty?
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
	
	{ 	
		lp: (sourcecode.string_between_markers "lp*", "_"),
		source: (sourcecode.string_between_markers "src*", "_"),
		campaign_id: (sourcecode.string_between_markers "x*", "_"),
		device: (sourcecode.string_between_markers "d*", "_"),
		device2: (sourcecode.string_between_markers "d2*", "_"),
		keyword: (sourcecode.string_between_markers "k*", "_"),
		match: (sourcecode.string_between_markers "m*", "_"),
		creative: (sourcecode.string_between_markers "c*", "_"),
		ad_position: (sourcecode.string_between_markers "p*", "_"),
		network: (sourcecode[/n\*(.+)/m, 1])
	}

end

input_filename = get_input_filename
output_filename = get_output_filename
process_ce_csv(input_filename, output_filename)

puts "Script Complete!"
puts "Time elapsed: #{Time.now - start_time} seconds"