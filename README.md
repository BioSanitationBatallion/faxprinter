# faxprinter
Simple IDLE client for an IMAP server that automatically prints PDFs found in received emails.  Written in Python.

There are many IMAP IDLE scripts to be found around the Internet but they tend to be overcomplicated with multi-threading silliness or contrib libraries that, to my mind, aren't necessary.  Python has good socket support and has handy timeout options so just use those in the main process.

This script is intended to be run on startup and should survive any network hiccups to maintain its IDLE connection.  

Just change the connection info at the head of the script and it's ready to go.

Note that it's intended purpose, to print faxes, can be overridden simply by removing the reference to getnewmessagesandprint() and write your own.
