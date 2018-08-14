import sublime
import sublime_plugin
import Pywin32.setup
from win32com.client import Dispatch
import os
import threading
import pythoncom
import re


def log_output(window, output_text, panel_name='AccuTermClient'):
    output_text = output_text.replace('\r', '')
    panel = window.find_output_panel(panel_name)
    if panel == None:
        panel = window.create_output_panel(panel_name, False)
    panel.run_command('append', {'characters': output_text + '\n'})
    panel.show(panel.size())
    window.run_command('show_panel', {'panel': 'output.' + panel_name})


def connect(panel_name='AccuTermClient'):
    mv_svr = Dispatch('atMVSvr71.Server')
    if mv_svr.Connect():
        # log_output(sublime.active_window(), 'Connected', panel_name) # Ideally the connecct would be passed the window but this is intended for debugging only.
        pass
    else: 
        log_output(sublime.active_window(), 'Unable to connect to AccuTerm\nMake sure AccuTerm is running FTSERVER.', panel_name)
    return mv_svr


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

def getHostType(mv_svr):
    return mv_svr.Readitem('ACCUTERMCTRL', 'KMTCFG', 51)

def get_setting_for_host(mv_svr, setting_name):
    host_type = sublime.load_settings('AccuTermClient.sublime-settings').get('host_type', 'auto')
    if host_type.lower() == 'auto': host_type = getHostType(mv_svr)
    setting_val = sublime.load_settings('AccuTermClient.sublime-settings').get(setting_name, None)
    if bool(host_type): 
        if host_type in setting_val: setting_val = setting_val[host_type]
    return setting_val


def get_file_item(file_name):
    if type(file_name) == sublime.View: file_name = file_name.file_name()
    mv_file   = file_name.split(os.sep)[-2]
    mv_item   = file_name.split(os.sep)[-1] 
    remove_file_ext = sublime.load_settings('AccuTermClient.sublime-settings').get('remove_file_extensions')
    if os.path.splitext(mv_item.lower())[1][1:] in remove_file_ext: mv_item = os.path.splitext(mv_item)[0]
    return (mv_file, mv_item)

def get_filename(window, mv_file, mv_item):
    file_ext = sublime.load_settings('AccuTermClient.sublime-settings').get('default_file_extension', 'bp')
    if file_ext != '': file_ext = '.' + file_ext
    return os.sep.join([get_base_path(window), mv_file, mv_item + file_ext])

def get_base_path(window=sublime.active_window()):
    project_file_name = window.project_file_name()
    base_path = sublime.load_settings('AccuTermClient.sublime-settings').get('default_save_location', '%userprofile%')
    return os.path.dirname(project_file_name) if bool(project_file_name) else base_path

def download(window, mv_file, mv_item):
    if bool(mv_file) and bool(mv_item):
        file_name = get_filename(window, mv_file, mv_item)
        mv_svr = connect()
        if mv_svr:
            if bool( mv_svr.ItemExists(mv_file, mv_item) ):
                mv_svr.UnlockItem(mv_file, mv_item)
                data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 1)
                if check_error_message(window, mv_svr, 'Download success'):
                    if not os.path.exists(file_name): 
                        new_view = window.new_file()
                        new_view.set_name( get_filename(window, mv_file, mv_item) )
                        default_dir = get_base_path(window) + os.sep + mv_file
                        if not os.path.exists(default_dir): os.makedirs(default_dir)
                        new_view.settings().set('default_dir', default_dir)
                        host_type = getHostType(mv_svr)
                        mv_syntaxes = sublime.load_settings('AccuTermClient.sublime-settings').get('syntax_file_locations', {})
                        if host_type in mv_syntaxes: new_view.set_syntax_file(mv_syntaxes[host_type])
                    else: 
                        new_view = window.open_file(file_name)
                    new_view.run_command('accu_term_replace_file', {"text": data})
            else: 
                log_output(window, mv_file + ' ' + mv_item + ' not found.')
            mv_svr.Disconnect()
    else:
        log_output(window, 'Invalid Input: ' + str(mv_file) + ' ' + str(mv_item) + ' (Must be [file] [item])')

class AccuTermUploadCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        (mv_file, mv_item) = get_file_item(self.view)
        data = self.view.substr( sublime.Region(0, self.view.size()) ).replace('\n', '\xFE')
        mv_svr = connect()
        if mv_svr:
            mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 1)
            check_error_message(self.view.window(), mv_svr, 'Uploaded to ' + mv_file + ' ' + mv_item)


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
        self.view = self.window.active_view()
        if self.view.is_dirty(): self.view.run_command('save')
        data = self.view.substr(sublime.Region(0, self.view.size())).replace('\n', '\xFE')        
        sublime.set_timeout_async( lambda: self.upload(self, data = data), 0)

    def upload(self, *args, data=None):
        if threading.currentThread ().getName() != 'MainThread':
           pythoncom.CoInitialize ()
        file_name = self.view.file_name()
        (mv_file, mv_item) = get_file_item(self.view)
        panel = self.window.create_output_panel('exec', False)
        self.panel = panel

        mv_svr = connect()
        if mv_svr.IsConnected():
            if panel:
                panel.settings().set("result_file_regex", r"Compiling:\s(.*)()")
                panel.settings().set("result_line_regex", self.get_result_line_regex(mv_svr))

            mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 1)
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
                    # self.window.focus_view( self.window.find_open_file(file_name) )
                # else:
                #     sublime.active_window().run_command('show_panel', {'panel': 'output.exec'})


class AccuTermReleaseCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        if file_name == None: file_name = os.sep.join([self.view.settings().get('default_dir'), self.view.name()])
        log_output(self.view.window(), file_name)
        (mv_file, mv_item) = get_file_item(file_name)
        mv_svr = connect()
        if mv_svr:
            mv_svr.UnlockItem(mv_file, mv_item)
            check_error_message(self.view.window(), mv_svr, 'Released ' + mv_file + ' ' + mv_item)


class AccuTermReplaceFileCommand(sublime_plugin.TextCommand):
    def run(self, edit, text=''):
        if self.view.is_loading():
            sublime.set_timeout_async(lambda: self.view.run_command('accu_term_replace_file', {"text": text}), 100)
        else:
            text = text.replace(os.linesep, '\n')
            self.view.replace(edit, sublime.Region(0, self.view.size()), text)


class AccuTermDownload(sublime_plugin.WindowCommand):
    def on_done(self, item_ref):
        item_ref = item_ref.split()
        if len(item_ref) == 2:
            [mv_file, mv_item] = item_ref
            download(self.window, mv_file, mv_item)
        else:
            log_output(self.window, 'Invalid Input: ' + ' '.join(item_ref) + ' (Must be [file] [item])')

    def run(self, **kwargs):
        self.window.show_input_panel('Enter the MV file and item', '', self.on_done, None, None)


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


class AccuTermExecute(sublime_plugin.TextCommand):
    def input(self, args):
        return ExecuteInputHandler(self.view)

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

    def run(self, edit, output_to='console', command=None): 
        if not(command):
            self.view.window().show_input_panel('Enter command', '', lambda command: 
                self.view.run_command("accu_term_execute", {"output_to": output_to, "command": command} ), 
                None, None)

        elif output_to == 'new':
            new_view = self.view.window().new_file()
            new_view.set_name(command)
            new_view.run_command('append', {"characters": self.run_commands(command)})
            new_view.set_scratch(True)
            new_view.run_command("accu_term_execute", {"output_to": "append"} )
            self.view.window().show_input_panel('Enter command', '', lambda command: 
                new_view.run_command("accu_term_execute", {"output_to": "append", "command": command} ), 
                None, None)

        elif output_to == 'append':
            self.view.run_command('append', {"characters": self.run_commands(command)})
            self.view.window().show_input_panel('Enter command', '', lambda command: 
                self.view.run_command("accu_term_execute", {"output_to": "append", "command": command} ), 
                None, None)

        elif output_to == 'replace':
            self.view.set_name(command)
            self.view.replace(edit, sublime.Region(0, self.view.size()), self.run_commands(command))

        else:
            log_output(self.view.window(), self.run_commands(command) )


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


class AccuTermRefreshCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        if bool(file_name):
            (mv_file, mv_item) = get_file_item(self.view)
            download(self.view.window(), mv_file, mv_item)
        else:
            log_output(self.view.window(), 'Unable to determine MV file reference. Make sure this file is saved locally first.')


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
            self.window.show_quick_panel(self.list, self.pickItem)

    def pickItem(self, item_index):
        if item_index > -1:
            mv_file = self.mv_file
            mv_item = self.list[item_index]
            download(self.window, mv_file, mv_item)

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
            else :
                check_error_message(self.view.window(), mv_svr, mv_file + ' ' + mv_item + ' locked')
            mv_svr.Disconnect()

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

class AccuTermGlobalUpcase(sublime_plugin.TextCommand):
    def run(self, edit):
        source_upcase = changeCase( self.view.substr(sublime.Region(0, self.view.size())), 'upper')
        sublime.set_timeout_async(lambda: self.view.run_command('accu_term_replace_file', {"text": source_upcase}), 0)

class AccuTermGlobalDowncase(sublime_plugin.TextCommand):
    def run(self, edit):       
        source_upcase = changeCase( self.view.substr(sublime.Region(0, self.view.size())), 'lower')
        self.view.run_command('accu_term_replace_file', {"text": source_upcase})        
        
class EventListener(sublime_plugin.EventListener):
    def on_pre_close(self, view):
        if view.scope_name(0).split('.')[-1].strip() in ['d3-basic', 'qm-basic', 'jbase-basic']:
            view.run_command('accu_term_release')
    

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

