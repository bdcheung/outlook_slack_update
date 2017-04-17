on run
	launch application "Microsoft Outlook"
	launch application "Slack"
end run

on idle
	global startDate
	global endDate
	global presentStatus
	global cachedStatus
	try
	cachedStatus
	on error
	set cachedStatus to ""
	end try
	global results
	global emailAddress
	try
		emailAddress
	on error
		display dialog "Email address..." default answer "first.last@phishme.com" with title "Email Address"
		set emailAddress to text returned of result
		
	end try
	
	set startDate to (current date)
	set endDate to startDate + (5 * minutes)
	
	tell application "Microsoft Outlook"
		set results to query freebusy exchange account 1 for attendees emailAddress range start time startDate range end time endDate interval 5 as list
		set presentStatus to results's last item as string
		
		if presentStatus = cachedStatus then
			-- Do nothing
		else
			if presentStatus = "free" then
				tell application "System Events"
					tell application "Slack" to activate
					key code 18 using {command down}
					delay 1
					keystroke "/status :free: Free and available!"
					delay 0.5
					key code 36
					delay 0.3
					set visible of process "Slack" to false
				end tell
				set cachedStatus to "free"
			else
				tell application "System Events"
					tell application "Slack" to activate
					key code 18 using {command down}
					delay 1
					keystroke "/status :no_entry_sign: On a call or in a meeting"
					delay 0.5
					key code 36
					delay 0.3
					set visible of process "Slack" to false
				end tell
				set cachedStatus to "busy"
			end if
		end if
	end tell
	
	return 30 -- every 30 seconds
end idle

on quit
	quit application "Microsoft Outlook"
	quit application "Slack"
	continue quit
end quit
