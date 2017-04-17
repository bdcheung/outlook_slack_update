on run
	launch application "Microsoft Outlook"
	launch application "Slack"
end run

on idle
	global startDate
	global endDate
	global presentStatus
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
		
		if presentStatus = "free" then
			-- noop
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
		end if
	end tell
	
	return 900 -- every 15 minutes
end idle

on quit
	quit application "Microsoft Outlook"
	quit application "Slack"
	continue quit
end quit