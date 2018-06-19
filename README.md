# AccuTermClient

AccuTerm client for Sublime Text.

## Features 
* Download & upload files
* Compile source code
* Lock item on the MultiValue server while open in Sublime
* Browse files on MV server
* Global case change (preserves case in comments & strings)

##Requirements
* Sublime Text 3
* Windows Operating System
* AccuTerm terminal emulator running the FTSERVER program

##Installation
1. cd %appdata%\Sublime Text 3\Packages
2. git clone https://github.com/ianharper/AccuTermClient.git
3. Open AccuTermClient.py and change the base_path variable to the default location for saving files from the MV server.

##Usage
This package connects to your MV database using AccuTerm's FTSERVER. To use, launch AccuTerm and run FTSERVER from TCL. This package will connect to the AccuTerm server with no additional configuration. This package expects that all files on your local machine to be contained in a folder that matches the file on the MV server. Files will be named on your local machine with a ".bp" suffix. When uploaded to the MV server this suffix will be removed. 

###Commands
* AccuTermClient Upload (ctrl+alt+u) - Upload current file to MV server.
* AccuTermClient Compile - Compile Current file on MV server.
* AccuTermClient Release - Release lock of current file on MV server.
* AccuTermClient Open (ctrl+alt+o) - Download item from MV server by entering MV file reference.
* AccuTermClient Unlock - unlock item on MV server by entering MV file reference.
* AccuTermClient Refresh - Update currently open file in Sublime from MV server and lock item on MV server.
* AccuTermClient List - Browse files on MV server using Sublime's command palate, select item with enter to download. 
* AccuTermClient Lock - Lock item on MV server by entering MV file reference.
* AccuTermClient Global Upcase - Convert case of currently open file to uppercase while preserving case in strings and comments.
* AccuTermClient Global Downcase - Convert case of currently open file to lowercase while preserving case in strings and comments.

#Todo
* Add oconv/iconv functions.
* Add aditional keyboard shortcuts.
* Add settings for default save location.
* Refactor AccuTermClient.py to remove duplicated code.
* Add more event listener functions (automatic check for changes on server)
* Fix need to unlock & re-lock items on MV server on upload.

