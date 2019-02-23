# Package: AccuTermClient
# A Sublime Text client for the AccuTerm server (from AccuTerm terminal emulator).

import sublime
import sublime_plugin
import Pywin32.setup
from win32com.client import Dispatch
import os
import threading
import pythoncom
import re


# Function: log_output
# Displays text in an output panel.
# 
# Parameters:
#   window - The Sublime window object that will hold the ouput panel.
#   output_text - String containing the text to show in the output panel.
#   panel_name - Name of the output panel (defaults to "AccuTermClient").
# 
# Returns:
#   None
def log_output(window, output_text, panel_name='AccuTermClient'):
    output_text = output_text.replace('\r', '')
    panel = window.find_output_panel(panel_name)
    if panel == None:
        panel = window.create_output_panel(panel_name, False)
    panel.run_command('append', {'characters': output_text + '\n'})
    panel.show(panel.size())
    window.run_command('show_panel', {'panel': 'output.' + panel_name})


# Function: connect
# Connects to an AccuTerm session running the FTSERVER and returns the AccuTerm Server object. 
# 
# Parameters:
#   panel_name - Name of the output panel to send error messages to (Defaults to AccuTermClient).
# 
# Returns:
#   object - AccuTerm Server object.
def connect(panel_name='AccuTermClient'):
    mv_svr = Dispatch('atMVSvr71.Server')
    if mv_svr.Connect():
        # log_output(sublime.active_window(), 'Connected', panel_name) # Ideally the connecct would be passed the window but this is intended for debugging only.
        pass
    else: 
        log_output(sublime.active_window(), 'Unable to connect to AccuTerm\nMake sure AccuTerm is running FTSERVER.', panel_name)
    return mv_svr


# Function: check_error_message
# Checks an AccuTerm Server object for errors resulting from a previous command. 
# If errors are found the error message is sent to an output panel otherwise the status bar
# will show a success message.
# 
# Parameters:
#   window - The Sublime window object that will hold the ouput panel.
#   mv_svr - AccuTerm server object (see <connect>).
#   success_msg - Message to show in status bar if command was successful (defaults to "Success").
# 
# Returns:
#   bool - True for success, False if an error was found.
def check_error_message(window, mv_svr, success_msg='Success'):
    if mv_svr.LastErrorMessage:
        log_output(window, str(mv_svr.LastError) + " " + mv_svr.LastErrorMessage)
        # log_output(window, 'Connection Failed', panel_name)
        # return None
        return False
    else:
        log_output(window, success_msg)
        window.destroy_output_panel('AccuTermClient')
        window.status_message(success_msg)
        return True

# Function: getHostType
# Get the MultiValue host from an AccuTerm server object.
# 
# Parameters:
#   mv_svr - AccuTerm server object (see <connect>).
# 
# Returns:
#   string - MultiValue host type.
def getHostType(mv_svr):
    return mv_svr.Readitem('ACCUTERMCTRL', 'KMTCFG', 51)

# Function:get_setting_for_host
# Gets a setting from the AccuTermClient Sublime settings based on the MV host type.
# 
# Parameters:
#   mv_svr - AccuTerm server object (see <connect>).
#   setting_name - Setting name as a string.
# 
# Returns:
#   string - Setting value.
def get_setting_for_host(mv_svr, setting_name):
    host_type = sublime.load_settings('AccuTermClient.sublime-settings').get('host_type', 'auto')
    if host_type.lower() == 'auto': host_type = getHostType(mv_svr)
    setting_val = sublime.load_settings('AccuTermClient.sublime-settings').get(setting_name, None)
    if bool(host_type): 
        if host_type in setting_val: setting_val = setting_val[host_type]
    return setting_val

# Function: is_mv_syntax
# Returns True if view or settings are in MultiValue syntax. 
# 
# Parameters:
#   entity - Sublime view or settings object.
# 
# Returns:
#   True - View/Settings is set to a MultiValue syntax.
#   False - View/Settings is not set to a MultiValue syntax.
def is_mv_syntax(entity):
    if type(entity) is sublime.View: 
        settings = entity.settings()
    elif type(entity) is sublime.Settings:
        settings = entity
    else:
        return False
    syntax = os.path.splitext(settings.get('syntax').split('/')[-1])[0]
    multivalue_syntaxes = sublime.load_settings('AccuTermClient.sublime-settings').get('multivalue_syntaxes')
    return syntax in multivalue_syntaxes

# Function: get_file_item
# Gets the file item reference from a passed Sublime view object.
# 
# Parameters:
#   view - Sublime view object.
# 
# Returns:
#   tuple - [0] MV filename.
# 
#           [1] MV item
def get_file_item(view):
    file_name = os.sep.join(view.settings().get('AccuTermClient_mv_file_item', []))
    if not bool(file_name): # file item not stored in view settings, get based on file path name.
        file_name = view.file_name()
        if file_name == None: # file not saved locally, spoof a file name
            file_name = ''.join([view.settings().get('default_dir'), os.sep, view.name()])
    mv_file = file_name.split(os.sep)[-2]
    mv_item = file_name.split(os.sep)[-1] 
    remove_file_ext = sublime.load_settings('AccuTermClient.sublime-settings').get('remove_file_extensions')
    if os.path.splitext(mv_item.lower())[1][1:] in remove_file_ext: mv_item = os.path.splitext(mv_item)[0]
    return (mv_file, mv_item)


# Function: get_filename
# Get the windows pathname from the MV file item reference.
# 
# Parameters:
#   window - Sublime window object.
#   mv_file - MV file string.
#   mv_item - MV item string.
# 
# Returns:
#   string - Windows pathname.
def get_filename(window, mv_file, mv_item):
    file_ext = sublime.load_settings('AccuTermClient.sublime-settings').get('default_file_extension', 'bp')
    if file_ext != '': file_ext = '.' + file_ext
    return os.sep.join([get_base_path(window), mv_file, mv_item + file_ext])


# Function: get_base_path
# Gets the base pathname froM a Sublime window object.
# 
# Parameters:
#   window - Sublime window object.
# 
# Returns:
#   string - Windows pathname.
def get_base_path(window=sublime.active_window()):
    project_file_name = window.project_file_name()
    base_path = os.path.expandvars(
        sublime.load_settings('AccuTermClient.sublime-settings').get('default_save_location', '%userprofile%')
        )
    return os.path.dirname(project_file_name) if bool(project_file_name) else base_path


# Function: get_view_lock_state
# Get the lock state of a MV item from the corresponding Sublime view.
# 
# Parameters:
#   view - Sublime view object.
# 
# Returns:
#   string - lock state.
# 
# Lock States:
#   locked - Item is locked on the MV server.
#   released - Item is not locked on MV server.
#   no_locking - Item was not opened with locking.
def get_view_lock_state(view):
    lock_state = view.settings().get('AccuTermClient_lock_state', None)
    if lock_state == None: 
        lock_state = 'released' if sublime.load_settings('AccuTermClient.sublime-settings').get('open_with_readu', True) else 'no_locking'
    return lock_state


# Function: find_view
# Find view based on full filename or view name.
# 
# Parameters:
#   view_name - string.
# 
# Returns:
#   view - Sublime view object.
def find_view(view_name):
    for window in sublime.windows():
        if window.find_open_file(view_name): 
            return window.find_open_file(view_name)

    # If view hasn't been found search based on view name.
    view_name = os.path.basename(view_name)
    for window in sublime.windows():
        for view in window.views():
            if view.name() == view_name: 
                return view
    return None

# Function: download
# Download item from MV server into a Sublime view.
# 
# Parameters:
#   window - Sublime window object.
#   mv_file - Filename on MV server.
#   mv_item - Item ID on MV server.
# 
# Returns:
#   None
def download(window, mv_file, mv_item, file_name=None, readu_flag=None):
    if bool(mv_file) and bool(mv_item):
        if not file_name: file_name = get_filename(window, mv_file, mv_item)
        mv_svr = connect()
        if mv_svr:
            if bool( mv_svr.ItemExists(mv_file, mv_item) ):
                mv_syntaxes = sublime.load_settings('AccuTermClient.sublime-settings').get('syntax_file_locations', {})

                if readu_flag == None: readu_flag = sublime.load_settings('AccuTermClient.sublime-settings').get('open_with_readu', True)
                if readu_flag:
                    mv_svr.UnlockItem(mv_file, mv_item)
                    data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 1)
                    if mv_svr.LastError == 260 and \
                    sublime.ok_cancel_dialog( mv_file + ' ' + mv_item + ' is locked by another port. Do you want to open this as read-only (without a lock)?', ok_title='Yes'):
                        data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 0)
                else:
                    data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 0)

                if check_error_message(window, mv_svr, 'Download success'):
                    new_view = find_view(file_name)
                    if new_view == None:
                        new_view = window.new_file()
                        new_view.set_name( os.path.split(get_filename(window, mv_file, mv_item))[1] )
                        new_view.sel().clear()
                        default_dir = get_base_path(window) + os.sep + mv_file
                        if not os.path.exists(default_dir): os.makedirs(default_dir)
                        new_view.settings().set('default_dir', default_dir)
                        host_type = getHostType(mv_svr)
                        if host_type in mv_syntaxes: new_view.set_syntax_file(mv_syntaxes[host_type])
                    new_view.settings().set('AccuTermClient_mv_file_item', [mv_file, mv_item])
                    new_view.settings().set('AccuTermClient_sync_state', 'skip')
                    if readu_flag:
                        new_view.settings().set('AccuTermClient_lock_state', 'locked')
                    else:
                        new_view.settings().set('AccuTermClient_lock_state', 'no_locking')
                    new_view.run_command('accu_term_replace_file', {"text": data})
                    if new_view.substr(sublime.Region(0,2)).upper() == 'PQ':
                        new_view.set_syntax_file(mv_syntaxes['PROC'])
            else: 
                log_output(window, mv_file + ' ' + mv_item + ' not found.')
            mv_svr.Disconnect()
    else:
        log_output(window, 'Invalid Input: ' + str(mv_file) + ' ' + str(mv_item) + ' (Must be [file] [item])')


# Function: upload
# Upload the contents of a view to the MV server.
# 
# Parameters:
#   view - Sublime view object.
#   mv_server - AccuTerm server object (optional).
# 
# Returns:
#   string - Error message, empty for success.
def upload(view, mv_svr=None):
    if not mv_svr: mv_svr = connect()
    (mv_file, mv_item) = get_file_item(view)
    data = view.substr( sublime.Region(0, view.size()) ).replace('\n', '\xFE')
    if mv_svr.IsConnected():
        if get_view_lock_state(view) == 'locked':
            mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 1)
        else:
            mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 0)
        check_error_message(view.window(), mv_svr, 'Uploaded to ' + mv_file + ' ' + mv_item)
        view.settings().set('AccuTermClient_sync_state', 'check')
    return mv_svr.LastError


# Function: check_sync
# Compare the contents of a view against the corresponding item on the MV server.
# 
# Parameters:
#   view - Sublime view object.
# 
# Returns:
#   sync_state - String state value.
# 
# Sync States:
#   check - proceed with sync on load. When the local copy is changed set to changed.
#   skip  - item just downloaded. Change state to check and return. When the local copy is changed set to changed.
#   changed - Local copy has changed, do not check. On upload set this to check.
#   None  - If syntax is not do not check, otherwise change state to check and then check sync.
def check_sync(view):
    sync_state = view.settings().get('AccuTermClient_sync_state')
    if sync_state == None and is_mv_syntax(view): sync_state = 'check'
    if sync_state == 'skip':
        sync_state = 'check'
        view.settings().set('AccuTermClient_sync_state', sync_state)
    elif sync_state == 'check':
        (mv_file, mv_item) = get_file_item(view)
        mv_svr = connect()
        if mv_svr.IsConnected() and bool( mv_svr.ItemExists(mv_file, mv_item) ):
            data_mv = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 0).replace('\r', '')
            data_local = view.substr( sublime.Region(0, view.size()) )
            if data_mv != data_local:  
                prompt = mv_file + ' ' + mv_item + ' has changed on the MV server. Do you want to download a fresh copy?'
                if sublime.ok_cancel_dialog(prompt, 'Download'):
                    view.run_command('accu_term_refresh')
    return sync_state


# Class: AccuTermUploadCommand
# Upload the current view to the MV server.
class AccuTermUploadCommand(sublime_plugin.TextCommand):
    def run(self, edit, mv_svr=None):
        if not mv_svr: mv_svr = connect()
        (mv_file, mv_item) = get_file_item(self.view)
        data = self.view.substr( sublime.Region(0, self.view.size()) ).replace('\n', '\xFE')
        if mv_svr.IsConnected():
            if get_view_lock_state(self.view) == 'locked':
                mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 1)
            else:
                mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 0)
            check_error_message(self.view.window(), mv_svr, 'Uploaded to ' + mv_file + ' ' + mv_item)


# Class: AccuTermCompileCommand
# Compile the current view on the MV server.
class AccuTermCompileCommand(sublime_plugin.WindowCommand):
    view = None

    def get_result_line_regex(self, mv_svr):
        host_type = sublime.load_settings('AccuTermClient.sublime-settings').get('host_type', 'auto')
        if host_type.lower() == 'auto': host_type = getHostType(mv_svr)
        result_line_regex = sublime.load_settings('AccuTermClient.sublime-settings').get('result_line_regex', None)
        if type(result_line_regex) == type({}) and host_type in result_line_regex:
            result_line_regex = result_line_regex[host_type]  
        else:
            result_line_regex = ''
        return result_line_regex

    def run(self, **kwargs):
        self.window.destroy_output_panel('exec')
        self.view = self.window.active_view()
        if self.view.is_dirty() and bool(self.view.file_name): self.view.run_command('save')
        data = self.view.substr(sublime.Region(0, self.view.size())).replace('\n', '\xFE')        
        sublime.set_timeout_async( lambda: self.upload(self, data = data), 0)

    def upload(self, *args, data=None):
        if threading.currentThread ().getName() != 'MainThread':
           pythoncom.CoInitialize ()

        mv_svr = connect()
        if upload(self.view, mv_svr): 
            log_output(self.window, mv_svr.LastErrorMessage, 'exec')
            return

        file_name = self.view.file_name() if self.view.file_name() else self.view.name()
        (mv_file, mv_item) = get_file_item(self.view)


        panel = self.window.create_output_panel('exec', False)
        panel.settings().set('AccuTermClient_saved_locally', bool(self.view.file_name()))
        self.panel = panel

        if mv_svr.IsConnected():
            if panel:
                panel.settings().set("result_file_regex", r"Compiling:\s(.*)()")
                panel.settings().set("result_line_regex", self.get_result_line_regex(mv_svr))
                panel.settings().set("result_base_dir", self.view.settings().get('default_dir'))

            if mv_svr.LastErrorMessage:
                log_output(self.window, mv_svr.LastErrorMessage, 'exec')
            else:
                self.window.destroy_output_panel('AccuTermClient')
                compile_command = sublime.load_settings('AccuTermClient.sublime-settings').get('compile_command', 'BASIC')
                if type(compile_command) == str:
                    result = mv_svr.Execute(compile_command + ' ' + mv_file + ' ' + mv_item)
                else:
                    result = '\n'.join( map(lambda cmd: mv_svr.Execute(cmd + ' ' + mv_file + ' ' + mv_item), compile_command) )
                log_output(self.window, 'Compiling: ' + file_name + '\n' + result, 'exec')
                if result.split('\n')[-1][:5] == '[241]': 
                    self.window.destroy_output_panel('exec')
                    self.window.status_message(mv_file + ' ' + mv_item + ' compiled')
        else:
            log_output(self.window, 'Not connected to MV server.')


# Class: AccuTermReleaseCommand
# Release the lock on the MV server corresponding to the current view.
class AccuTermReleaseCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        if file_name == None: file_name = os.sep.join([self.view.settings().get('default_dir'), self.view.name()])
        log_output(self.view.window(), file_name)
        (mv_file, mv_item) = get_file_item(self.view)
        mv_svr = connect()
        if mv_svr:
            mv_svr.UnlockItem(mv_file, mv_item)
            if mv_svr.LastError == 0: self.view.settings().set('AccuTermClient_lock_state', 'released')
            check_error_message(self.view.window(), mv_svr, 'Released ' + mv_file + ' ' + mv_item)


# Class: AccuTermReplaceFileCommand
# Replace the contents of the current view with text. Used internally to replace contents of a view
# with data from the MV server.
class AccuTermReplaceFileCommand(sublime_plugin.TextCommand):
    # Function: run
    # Parameters:
    #   self - Sublime TextCommand instance.
    #   edit - Sublime edit object.
    #   text - string
    def run(self, edit, text=''):
        if self.view.is_loading():
            sublime.set_timeout_async(lambda: self.view.run_command('accu_term_replace_file', {"text": text}), 100)
        else:
            text = text.replace(os.linesep, '\n')
            self.view.replace(edit, sublime.Region(0, self.view.size()), text)


# Class: AccuTermDownload
# Download an item from the MV server.
class AccuTermDownload(sublime_plugin.WindowCommand):
    def __init__(self, window):
        self.readu_flag = None
        self.window     = window

    # Function: on_done
    # Run the <Download> function with the mv file item reference.
    # 
    # Parameters:
    #   self
    #   item_ref - string
    def on_done(self, item_ref):
        item_ref = item_ref.split()
        if len(item_ref) == 2:
            [mv_file, mv_item] = item_ref
            download(self.window, mv_file, mv_item, readu_flag=self.readu_flag)
        else:
            log_output(self.window, 'Invalid Input: ' + ' '.join(item_ref) + ' (Must be [file] [item])')

    # Function: run
    # Show a Sublime input panel to get the MV file item reference.
    def run(self, **kwargs):
        if 'readu_flag' in kwargs: self.readu_flag = kwargs['readu_flag']
        self.window.show_input_panel('Enter the MV file and item', '', self.on_done, None, None)


# Class: ConvDataInputHandler
class ConvDataInputHandler(sublime_plugin.TextInputHandler):
    def __init__(self, view):
        self.view = view

    def name(self):
        return 'data'
        
    def placeholder(self):
        return 'Data'

    def initial_text(self):
        if len(self.view.sel()) == 1:
            return self.view.substr(self.view.sel()[0])
        else:
            return ''

    def next_input(self, args):
        return ConvCodeInputHandler(self.view, args)


# Class: ConvCodeInputHandler
class ConvCodeInputHandler(sublime_plugin.TextInputHandler):
    def __init__(self, view, args):
        self.view = view
        self.data = args['data']
        self.conv_type = args['conv_type']
        self.mv_svr = connect()

    def name(self):
        return 'conv_code'
        
    def placeholder(self):
        return 'Conversion Code'

    def preview(self, conv_code=''):
        text = ''
        if conv_code != '' and AccuTermConv.IsValid(conv_code):
            if self.mv_svr.IsConnected(): 
                text = self.mv_svr.Oconv(self.data, conv_code) if self.conv_type == 'oconv' else self.mv_svr.Iconv(self.data, conv_code)
        return text


# Class: AccuTermConv
# Convert data using MV processing codes on the MV server.
class AccuTermConv(sublime_plugin.TextCommand):
    def input(self, args):
        return ConvDataInputHandler(self.view)
 
    def run(self, edit, conv_type='oconv', data='', conv_code=''):
        mv_svr = connect()
        check_error_message(self.view.window(), mv_svr)
        if mv_svr.IsConnected(): 
            mv_svr.Oconv(data, conv_code)

    def IsValid(conv_code):
        status = False
        valid_codes = ['C', 'D', 'G', 'MC', 'MD', 'ML', 'MR', 'MT', 'S']
        if len(conv_code) > 0 and conv_code[0].upper() in valid_codes: status = True
        if len(conv_code) > 1 and conv_code[:2].upper() in valid_codes: status = True
        return status


# Class: AccuTermExecute
# Execute a command on the MV server specifying the output destination:
# 
#   new - Create a new view with the output.
#   append - Append the output to the current view.
#   replace - Replace the current view with the output.
#   console - Append the output to the console.
class AccuTermExecute(sublime_plugin.TextCommand):
    def input(self, args):
        if 'command' in args and args['command']: return None
        return ExecuteInputHandler(self.view)

    def run(self, edit, output_to='console', command=None): 
        # Expand the command with defined environment variables.
        if command:
            (mv_file, mv_item) = get_file_item(self.view)
            command = command.replace('${FILE}', mv_file)
            command = command.replace('${ITEM}', mv_item)

        self.command = command
        self.command_view = self.view
        if self.view.window() == None: output_to = 'console'
        
        def append():
            if threading.currentThread().getName() != 'MainThread': pythoncom.CoInitialize()
            self.command_view.run_command('append', {'characters': self.run_commands(self.command)})
            self.command_view.run_command("accu_term_execute", {"output_to": "append"} )

        def log():
            if threading.currentThread().getName() != 'MainThread': pythoncom.CoInitialize()
            window = self.command_view.window() if bool(self.command_view.window()) else sublime.active_window()
            log_output(window, '\n')
            log_output(window, self.run_commands(self.command) )

        if not(command):
            initial_text = self.view.substr(self.view.sel()[0])
            self.view.window().show_input_panel('Enter command', initial_text, lambda command: 
                self.view.run_command("accu_term_execute", {"output_to": output_to, "command": command} ), 
                None, None)

        elif output_to == 'new':
            new_view = self.view.window().new_file()
            new_view.set_name(command)
            new_view.set_scratch(True)
            self.command_view = new_view
            sublime.set_timeout_async(append, 0)

        elif output_to == 'append':
            sublime.set_timeout_async(append, 0)

        elif output_to == 'replace':
            self.view.set_name(command)
            self.view.replace(edit, sublime.Region(0, self.view.size()), '')
            sublime.set_timeout_async(append, 0)

        else: # output to console
            sublime.set_timeout_async(log, 0)

    def run_commands(self, commands):
        results = ''
        mv_svr = connect()
        if mv_svr.IsConnected():
            commands = commands.split('\n')
            for command in commands:
                results += command + '\n'
                results += mv_svr.Execute(command, '', 1).replace('\x1b', '').replace(os.linesep, '\n') + '\n\n'
                check_error_message(self.view.window(), mv_svr, '')
        return results


# Class: ExecuteInputHandler
class ExecuteInputHandler(sublime_plugin.TextInputHandler):
    def __init__(self, view):
        self.view = view

    def name(self):
        return 'command'

    def initial_text(self):
        if len(self.view.sel()) == 1:
            return self.view.substr(self.view.sel()[0])
        else:
            return ''

    def next_input(self, args):
        if args['command']:
            return None
        else:
            return ExecuteHistoryInputHandler(self.view)


# Class: ExecuteHistoryInputHandler
class ExecuteHistoryInputHandler(sublime_plugin.ListInputHandler):
    def __init__(self, view):
        self.view = view

    def name(self):
        return 'command'

    def list_items(self):
        mv_svr = connect()
        if mv_svr.IsConnected():
            mv_file, mv_item = get_setting_for_host(mv_svr, 'command_history')
            if mv_item == '@USER': mv_item = mv_svr.UserName
            command_history = mv_svr.ReadItem(mv_file, mv_item)
            if not mv_svr.LastError:
                return command_history.replace('\r', '').split('\n')
        else:
            return ['']
        

# Class: AccuTermUnlock
# Unlock an item on the MV server by specifying the file item reference.
class AccuTermUnlock(sublime_plugin.WindowCommand):
    def on_done(self, item_ref):
        item_ref = item_ref.split()
        if len(item_ref) == 2:
            [mv_file, mv_item] = item_ref
            mv_svr = connect()
            if mv_svr:
                mv_svr.UnlockItem(mv_file, mv_item)
                check_error_message(self.window, mv_svr, mv_file + ' ' + mv_item + ' unlocked')
                mv_svr.Disconnect()
        else:
            log_output(self.window, 'Invalid Input: ' + item_ref + ' (Must be [file] [item])')

    def run(self, **kwargs):
        self.window.show_input_panel('Enter the MV file and item', '', self.on_done, None, None)


# Class: AccuTermRefreshCommand
# Update the current view with the item from the MV server and lock the item on the server.
class AccuTermRefreshCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        (mv_file, mv_item) = get_file_item(self.view)
        download(self.view.window(), mv_file, mv_item, self.view.file_name())


# Class: AccuTermListCommand
# Browse files on MV server using Sublime quick panels.
class AccuTermListCommand(sublime_plugin.WindowCommand):
    def run(self, **kwargs):
        self.mv_svr = connect()
        if self.mv_svr:
            list_files_command = get_setting_for_host(self.mv_svr, 'list_files_command')
            if list_files_command:
                self.list = ''.join(self.mv_svr.Execute(list_files_command, '', 1)).split('\r\n')
            elif self.mv_svr.MDName == 'VOC':
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_svr.MDName + ' WITH A1 = "F" "Q" A0 COL-HDR-SUPP ID-SUPP NOPAGE COUNT.SUP', '', 1)).split('\r\n')
            else:
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_svr.MDName + ' WITH A1 = "D" "Q" A0 COL-HDR-SUPP ID-SUPP NOPAGE NI-SUPP', '', 1)).split('\r\n')
            if check_error_message(self.window, self.mv_svr, ''):
                self.window.show_quick_panel(self.list, self.listFile)

    def listFile(self, list_index):
        if list_index > -1:
            self.mv_file = self.list[list_index]
            list_command = get_setting_for_host(self.mv_svr, 'list_command')
            if list_command:
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_file + list_command, '', 1)).split('\r\n')
            elif self.mv_svr.MDName == 'VOC':
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_file + ' A0 COL-HDR-SUPP ID-SUPP NOPAGE COUNT.SUP', '', 1)).split('\r\n')
            else:
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_file + ' A0 COL-HDR-SUPP ID-SUPP NOPAGE NI-SUPP', '', 1)).split('\r\n')
            self.list.insert(0, '..')
            self.window.show_quick_panel(self.list, self.pickItem)

    def pickItem(self, item_index):
        if item_index == 0:
            self.run()
        elif item_index > -1:
            mv_file = self.mv_file
            mv_item = self.list[item_index]
            download(self.window, mv_file, mv_item)


# Class: AccuTermLockCommand
# Lock an item on the MV server with a supplied file item reference.
class AccuTermLockCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        (mv_file, mv_item) = get_file_item(self.view)
        mv_svr = connect()
        if mv_svr:
            data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 1)
            if mv_svr.LastError == 260:
                self.view.window().destroy_output_panel('AccuTermClient')
                self.view.window().status_message(mv_file + ' ' + mv_item + ' is already locked')
                # self.view.window().status_message(mv_svr.LastErrorMessage)
                self.view.settings().set('AccuTermClient_lock_state', 'locked')
            else :
                if mv_svr.LastError == 0: 
                    self.view.settings().set('AccuTermClient_lock_state', 'locked')
                else:
                    self.view.settings().set('AccuTermClient_lock_state', 'released')
                check_error_message(self.view.window(), mv_svr, mv_file + ' ' + mv_item + ' locked')

def changeCase(text, case_funct='upper()'):
    source_code = []
    quotes = '\'"\\'
    comments = "*!"
    lines = text.split('\n')
    for line in lines:
        if line.strip() != '' and line.strip()[0] not in comments:
            idx = -1
            while idx + 1 < len(line):
                idx += 1
                char = line[idx]
                if char in quotes:
                    matching_quote_idx = line.find(char, idx + 1)
                    if matching_quote_idx >= 0: 
                        idx = matching_quote_idx
                        continue
                if char == ";" and len(line) > idx + 1 and line[idx+1:].replace(' ', '').replace('\t', '')[0] in comments: 
                    idx = len(line)
                    continue
                if case_funct == 'lower':
                    line = line[:idx] + line[idx].lower() + line[idx + 1:]
                else: 
                    line = line[:idx] + line[idx].upper() + line[idx + 1:]
        source_code.append(line)
    return '\n'.join(source_code)


# Class: AccuTermGlobalUpcase
# Convert all text in current view to uppercase except text in string quotes and comments.
class AccuTermGlobalUpcase(sublime_plugin.TextCommand):
    def run(self, edit):
        source_upcase = changeCase( self.view.substr(sublime.Region(0, self.view.size())), 'upper')
        sublime.set_timeout_async(lambda: self.view.run_command('accu_term_replace_file', {"text": source_upcase}), 0)


# Class: AccuTermGlobalDowncase
# Convert all text in current view to lowercase except text in string quotes and comments.
class AccuTermGlobalDowncase(sublime_plugin.TextCommand):
    def run(self, edit):       
        source_upcase = changeCase( self.view.substr(sublime.Region(0, self.view.size())), 'lower')
        self.view.run_command('accu_term_replace_file', {"text": source_upcase})        


# Class: AccuTermCheckSyncCommand
# Check the active MV item against the item on the server.
class AccuTermCheckSyncCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        self.view.settings().set('AccuTermClient_sync_state', 'check')
        check_sync(self.view)
        self.view.window().status_message('Sync check completed.')


# Class: EventListener        
# Register event handlers.
class EventListener(sublime_plugin.EventListener):
    def on_pre_close(self, view):
        lock_state = view.settings().get('AccuTermClient_lock_state', None)
        if lock_state == 'locked': view.run_command('accu_term_release')

    def on_window_command(self, window, command_name, args):
        if command_name in ['prev_result', 'next_result']:
            panel = window.find_output_panel('exec')
            saved_locally = panel.settings().get('AccuTermClient_saved_locally', True)
            if not saved_locally: 
                print('disabling prev/next')
                return ('None', '')


# Class: AccuTermClientLoadListener
# When view is loaded it will be checked against the item on the MV server with <check_sync>.
class AccuTermClientLoadListener(sublime_plugin.ViewEventListener):
    def __init__(self, view):
        self.view = view

    def is_applicable(settings):
        if settings.get('AccuTermClient_sync_state') == None:
            return is_mv_syntax(settings)
        else:
            return True

    def applies_to_primary_view_only():
        return True

    def on_load_async(self):
        check_sync(self.view)



# Event: plugin_loaded
# Lock all MV items that were locked previously. Triggered by Sublime during startup.
def plugin_loaded():
    for window in sublime.windows():
        for view in window.views():
            if not is_mv_syntax(view): continue
            check_sync(view)
            if 'locked' == get_view_lock_state(view):
                view.run_command('accu_term_lock')


# Class: AccuTermRunCommand
# Run the currently open file. If the item is in the MD/VOC then the item name will be used to run (enables running PROC, PARAGRAPH, or MACRO commands).
class AccuTermRunCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        (mv_file, mv_item) = get_file_item(self.view)
        mv_svr = connect()
        if mv_svr.IsConnected(): 
            if bool(mv_svr.ItemExists(mv_svr.MDName, mv_item)): 
                command = mv_item 
            else: 
                command = 'RUN ' + mv_file + ' ' + mv_item
            self.view.run_command('accu_term_execute', {"output_to": 'console', "command": command})
