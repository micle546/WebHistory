# WebHistory
Small project to pull History from web browsers

Program pulls web browser history from other users on a local computer from an account with admin access

# Usage


	usage: WebHistory.exe [-h] --username USERNAME [--browser {Edge,Chrome,Firefox,All}]
	optional arguments
	-h, --help            show this help message and exit
	--username USERNAME   doin another thing
	--browser {Edge,Chrome,Firefox,All}
# Example
Will pull Chrome history from user Bob, and save it to your desktop in WebHistory\WebHistory.xlsx
	
	WebHistory.exe --username Bob --browser Chrome
