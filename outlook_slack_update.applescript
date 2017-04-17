launch application "Microsoft Outlook"
launch application "Slack"

property startDate : current date
property endDate : startDate + (5 * minutes)
property cachedStatus : ""
property results : {}
property running : true

global emailAddress
try
	emailAddress
on error
	display dialog "Email address..." default answer "first.last@phishme.com" with title "Email Address"
	set emailAddress to text returned of result
	
end try

repeat while running
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
			else if presentStatus = "tentative" then
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
				set cachedStatus to "tentative"
			else if presentStatus = "busy" then
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
			else if presentStatus = "oof" then
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
				set cachedStatus to "oof"
			end if
		end if
	end tell
	
	delay 60
end repeat
