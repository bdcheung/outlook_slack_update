launch application "Microsoft Outlook"
launch application "Slack"

property startDate : current date
property endDate : startDate + (5 * minutes)
property cachedStatus : ""
property presentStatus : ""
property results : {}
property running : true
property displaySleepState : ""

global emailAddress


repeat while running
	set startDate to current date
	set endDate to startDate + (5 * minutes)
	set displaySleepState to do shell script "python -c 'import sys,Quartz; d=Quartz.CGSessionCopyCurrentDictionary(); print d'"

	if displaySleepState contains "CGSSessionScreenIsLocked = 1" then
		-- do nothing
	else
		try
			emailAddress
		on error
			display dialog "Email address..." default answer "first.last@phishme.com" with title "Email Address"
			set emailAddress to text returned of result
			
		end try
		
		
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
						keystroke "/status :open: Open and available!"
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
						keystroke "/status :question: Questionable availability"
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
						keystroke "/status :beach_with_umbrella: Out of office"
						delay 0.5
						key code 36
						delay 0.3
						set visible of process "Slack" to false
					end tell
					set cachedStatus to "oof"
				end if
			end if
		end tell
	end if
	
	delay 5
end repeat
