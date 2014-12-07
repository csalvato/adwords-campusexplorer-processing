# Campus Explorer CSV to Excel-ready data

start_time = Time.now
puts "Starting Script..."

require 'csv'

def process_ce_csv (input_filename, output_filename)
	CSV.open(output_filename, "wb") do |csv|
	# For each row from Campus Explorer CSV File
		CSV.foreach(input_filename, :headers => true, :return_headers => false, :encoding => 'windows-1251:utf-8') do |row|
		# Process the utm_campaign string as passed through from Source Code into separate values in their own cells
		# Is there data?
			# => YES -> If there's revenue (maybe it's if there is any utm_campaign string appended at all?)
				# utm_campaign = _src*adwords_x*205882121_d*mb_d2*{ifmobile:mb}{ifnotmobile:dt}_k*{keyword}_m*{matchtype}_c*{creative}_p*{adposition}_n*{network}&utm_source=Google&utm_medium=cpc
					# => Area of Study
					# => concentration
					# => seed
					# => sublocation
					# => Location Code
					# => headline
					# => Broken data? (Does it have {} in the URL?)
					# => LP Parameter
					# => Source Parameter
					# => Campaign ID (and decode back to campaign name)
					# => Desktop/Mobile values
					# => Second Desktop/Mobile Value
					# => Keyword
					# => Match Type?
					# => creative
					# => Ad position
					# => Network
			# Write ALL values out to processed CSV file
		end
	puts "#{input_filename}"
	puts "#{output_filename}"
	end
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

input_filename = get_input_filename
output_filename = get_output_filename
process_ce_csv(input_filename, output_filename)

puts "Script Complete!"
puts "Time elapsed: #{Time.now - start_time} seconds"