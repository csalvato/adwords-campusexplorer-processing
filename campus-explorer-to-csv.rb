# For each row from Campus Explorer CSV File
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