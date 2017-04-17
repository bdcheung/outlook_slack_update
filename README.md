# Slack Status Updater

This app works with Outlook and Slack to update your Slack status based on your Outlook calendar availability.  It checks your Outlook calendar every 30" and changes your Slack status if your Outlook availability has changed.

# How to use
1. Download and unzip the latest release.
1. Double-click to run the app (this will automatically open Outlook and Slack if they aren't already open)
1. When prompted, enter the email address of the attendee whose status you want to use.
1. Profit!

# Notes
1. As of version 0.2, the script updates the status of whichever Slack team is at the top of the list.  There's an outstanding issue to allow users to specify which team number should receive the update.
1. It's presently not possible to change the status emoji or message without changing the source code and re-compiling the app.
