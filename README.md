# AccuTermClient

AccuTerm client is a plugin for Sublime Text allows you to connect to a MultiValue server using AccuTerm. This plugin allows for 
source code editing and compiling on remote MV servers as well as executing commands and using the native data conversion processing codes (OCONV/ICONV). 

## Features 
### Standard Features - Will work with all MV DBMS that support Accuterm
* Download & upload files
* Compile source code
* Lock/Unlock items on the MultiValue server
* Iconv/Oconv from Sublime
* Run currently open file 
* Execute commands on the MV Server and display output in Sublime
* Global case change while preserving case in comments & strings

### Extended Features - Limited availability without additional configuration 
These features require DBMS specific configuration to run. D3, QM, and jBASE are configured automatically. Additional DBMS can be setup manually (see the _Settings_ section below).

* Prev/Next compile error
* Browse command stack
* Browse files on MV server (jBASE Windows not supported)

## Requirements
* Sublime Text 3
* PyWin32 Sublime Package
* Windows Operating System
* AccuTerm terminal emulator running the FTSERVER program


## Installation
### Using Sublime Package Manager
1. Package Control: Install
2. AccuTermClient

### Using Github
1. cd %appdata%\Sublime Text 3\Packages
2. git clone https://github.com/ianharper/AccuTermClient.git
3. Install the PyWin32 Sublime package.


## Usage
This package connects to your MV database using AccuTerm's FTSERVER program. To use, launch AccuTerm and run FTSERVER from TCL. AccuTermClient will connect to the AccuTerm server with no additional configuration. 

This package expects that all files on your local machine to be contained in a folder that matches the file on the MV server. Files will be named on your local machine with a ".bp" suffix. When uploaded to the MV server this suffix will be removed. 
Example: C:\code\BP\HELLO.WORLD.bp will be uploaded to BP HELLO.WORLD in the account that FTSERVER is running in.

### Commands
* Upload - Upload current file to MV server.
* Compile - Compile Current file on MV server.
* Release - Release lock of current file on MV server.
* Open - Download item from MV server by entering MV file reference. Will lock item on MV server if _open_with_readu_ setting is true.
* Open (Read Only) - Download item from MV sever without locking by entering MV file reference.
* Unlock - unlock item on MV server by entering MV file reference.
* Refresh - Update currently open file in Sublime from MV server and lock item on MV server.
* List - Browse files on MV server using Sublime's command palate, select item with enter to download. 
* Lock - Lock item on MV server by entering MV file reference.
* Execute - Run commands on MV server and show output in Sublime (to console, new file, or append to current file).
* Run - Run the currently open file. If the item is in the MD/VOC then the item name will be used to run (enables running PROC, PARAGRAPH, or MACRO commands).
* Iconv/Oconv - Convert data using the MV server's iconv/oconv functions.
* Global Upcase - Convert case of currently open file to uppercase while preserving case in strings and comments.
* Global Downcase - Convert case of currently open file to lowercase while preserving case in strings and comments.

### Settings
The settings can be accessed in the Preferences>Package Settings>AccuTermClient>Settings. The settings are in json format. Each top level key-value pair will be explained below. Some settings are specific to the MV DBMS, they will have a second key that specifies the DBMS. This key for your DBMS can be found in ACCUTERM,ACCUTERMCTRL, KMTCFG<51>. These settings can be set for general editing in Sublime or for specific Sublime projects.

| Setting Key | Description |
| ----------- | ----------- |
| Default Save Location | Location to save MV files by default. When editing a file in a Sublime project the project folder will be used instead.|
| remove_file_extensions | File extensions to remove when uploading to the MV server. | 
| compile_command | Command to execute when the Sublime Build command is run. |
| result_line_regex | Regular expression used to find the line number of compile errors. See [exec Target Options](https://www.sublimetext.com/docs/3/build_systems.html#exec_options) in the Sublime Docs for details. |
| list_files_command | Command to list all the files in the account. Used in the AccuTermClient List command. The output must contain only the file name, one per line. |
| list_command | This command is run after a file is chosen from the List command. The value is appended to limit the output to only the item names. |
| syntax_file_locations | List of MV syntaxes to apply after downloading. The default values come from the MultiValue Basic Sublime package |
| command_history | MV file and item for the command stack. |



# Todo
* Add more event listener functions (automatic check for changes on server)
* Add support for jBASE windows.
* Autofill execute command with selected text.
* Add ability to run multiple commands at once and view the output.
* Enable custom compile compile commands to allow multiple commands (like BASIC %FILE %ITEM & COMPILE %FILE %ITEM)
* Add setting to disable automatic locking of items when downloaded.
* Set MV syntax automatically based on DBMS type and file contents (ex. PQ in line 1 should set PROC).
* Allow file extensions to be set based on DBMS type and file contents (ex. PQ in line 1 should set proc ext.).
