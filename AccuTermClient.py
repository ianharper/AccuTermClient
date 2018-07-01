import sublime
import sublime_plugin
import Pywin32.setup
from win32com.client import Dispatch
import os
import threading
import pythoncom
import re


def log_output(output_text, panel_name='AccuTermClient'):
    output_text = output_text.replace('\r', '')
    panel = sublime.active_window().find_output_panel(panel_name)
    if panel == None:
        panel = sublime.active_window().create_output_panel(panel_name, False)
    panel.run_command('append', {'characters': output_text + '\n'})
    panel.show(panel.size())
    sublime.active_window().run_command('show_panel', {'panel': 'output.' + panel_name})


def connect(panel_name='AccuTermClient'):
    mv_svr = Dispatch('atMVSvr71.Server')
    if mv_svr.Connect():
        # log_output('Connected', panel_name)
        return mv_svr


def check_error_message(mv_svr, success_msg='Success'):
    if mv_svr.LastErrorMessage:
        log_output(str(mv_svr.LastError) + " " + mv_svr.LastErrorMessage)
        log_output('Connection Failed', panel_name)
        return None
        return False
    else:
        log_output(success_msg)
        sublime.active_window().destroy_output_panel('AccuTermClient')
        sublime.active_window().status_message(success_msg)
        return True

def get_file_item(file_name):
    if type(file_name) == sublime.View: file_name = file_name.file_name()
    mv_file   = file_name.split(os.sep)[-2]
    mv_item   = file_name.split(os.sep)[-1] 
    remove_file_ext = sublime.load_settings('AccuTermClient.sublime-settings').get('remove_file_extensions')
    if os.path.splitext(mv_item.lower())[1][1:] in remove_file_ext: mv_item = os.path.splitext(mv_item)[0]
    return (mv_file, mv_item)

def get_base_path():
    project_file_name = sublime.active_window().project_file_name()
    base_path = sublime.load_settings('AccuTermClient.sublime-settings').get('default_save_location', '%userprofile%')
    return os.path.dirname(project_file_name) if bool(project_file_name) else base_path

class AccuTermUploadCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        (mv_file, mv_item) = get_file_item(self.view)
        data = self.view.substr( sublime.Region(0, self.view.size()) ).replace('\n', '\xFE')
        mv_svr = connect()
        if mv_svr:
            mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 1)
            check_error_message(mv_svr, 'Uploaded to ' + mv_file + ' ' + mv_item)


class AccuTermCompileCommand(sublime_plugin.WindowCommand):
    panel = None

    def run(self, **kwargs):
        cur_view = sublime.active_window().active_view()
        if cur_view.is_dirty(): cur_view.run_command('save')
        data = cur_view.substr(sublime.Region(0, cur_view.size())).replace('\n', '\xFE')        
        sublime.set_timeout_async( lambda: self.upload(self, cur_view=cur_view, data = data), 0)

    def upload(self, *args, cur_view=None, data=None):
        if threading.currentThread ().getName() != 'MainThread':
           pythoncom.CoInitialize ()
        file_name = cur_view.file_name()
        (mv_file, mv_item) = get_file_item(cur_view)
        # panel = sublime.Window.create_output_panel(sublime.active_window(), 'exec', True)
        panel = sublime.active_window().create_output_panel('exec', False)
        self.panel = panel
        if panel:
            panel.settings().set("result_file_regex", r"Compiling.\s(.*)\s([0-9]*)")
            panel.settings().set("result_line_regex", r"Line.([0-9]+).()\s+(.*)")

        mv_svr = connect()
        if mv_svr:
            mv_svr.WriteItem(mv_file, mv_item, data, 0, 0, 0, 1)
            if mv_svr.LastErrorMessage:
                log_output(mv_svr.LastErrorMessage, 'exec')
            else:
                sublime.active_window().destroy_output_panel('AccuTermClient')
                compile_command = sublime.load_settings('AccuTermClient.sublime-settings').get('compile_command', 'BASIC')
                result = mv_svr.Execute(compile_command + ' ' + mv_file + ' ' + mv_item)
                log_output('Compiling: ' + file_name + ' 0' + '\n' + result, 'exec')
                if result.split('\n')[-1][:5] == '[241]': 
                    sublime.active_window().destroy_output_panel('exec')
                    sublime.active_window().status_message(mv_file + ' ' + mv_item + ' compiled')
                    sublime.active_window().focus_view( sublime.active_window().find_open_file(file_name) )
                else:
                    sublime.active_window().run_command('show_panel', {'panel': 'output.exec'})


class AccuTermReleaseCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        if file_name == None: file_name = os.sep.join([self.view.settings().get('default_dir'), self.view.name()])
        log_output(file_name)
        (mv_file, mv_item) = get_file_item(file_name)
        mv_svr = connect()
        if mv_svr:
            mv_svr.UnlockItem(mv_file, mv_item)
            check_error_message(mv_svr, 'Released ' + mv_file + ' ' + mv_item)


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
            file_name = os.sep.join([get_base_path(), mv_file, mv_item + '.bp'])
            mv_svr = connect()
            if mv_svr:
                if bool( mv_svr.ItemExists(mv_file, mv_item) ):
                    mv_svr.UnlockItem(mv_file, mv_item)
                    data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 1)
                    if check_error_message(mv_svr, 'Download success'):
                        if not os.path.exists(file_name): 
                            new_view = self.window.new_file()
                            new_view.set_name(mv_item + '.bp')
                            default_dir = get_base_path() + os.sep + mv_file
                            if not os.path.exists(default_dir): os.makedirs(default_dir)
                            new_view.settings().set('default_dir', default_dir)
                        else: 
                            new_view = self.window.open_file(file_name)
                        new_view.run_command('accu_term_replace_file', {"text": data})
                else: 
                    log_output(mv_file + ' ' + mv_item + ' not found.')
                mv_svr.Disconnect()
        else:
            log_output('Invalid Input: ' + item_ref + ' (Must be [file] [item])')

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

    def name(self):
        return 'conv_code'
        
    def placeholder(self):
        return 'Conversion Code'

    def preview(self, conv_code=''):
        text = ''
        if conv_code != '' and AccuTermOconv.IsValid(conv_code):
            mv_svr = connect()
            if mv_svr.IsConnected(): 
                text = mv_svr.Oconv(self.data, conv_code) if self.conv_type == 'oconv' else mv_svr.Iconv(self.data, conv_code)
        return text


class AccuTermConv(sublime_plugin.TextCommand):
    def input(self, args):
        return ConvDataInputHandler(self.view)
 
    def run(self, edit, conv_type='oconv', data='', conv_code=''):
        mv_svr = connect()
        check_error_message(mv_svr)
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
                check_error_message(mv_svr, '')
        return results

    def run(self, edit, output_to='console', command=None): 
        if not(command): return

        if output_to == 'new':
            new_view = sublime.active_window().new_file()
            new_view.set_name(command)
            new_view.run_command('append', {"characters": self.run_commands(command)})
            new_view.set_scratch(True)
            new_view.run_command("accu_term_execute", {"output_to": "append"} )
            sublime.active_window().show_input_panel('Enter command', '', lambda command: 
                self.view.run_command("accu_term_execute", {"output_to": "append", "command": command} ), 
                None, None)

        elif output_to == 'append':
            self.view.run_command('append', {"characters": self.run_commands(command)})
            sublime.active_window().show_input_panel('Enter command', '', lambda command: 
                self.view.run_command("accu_term_execute", {"output_to": "append", "command": command} ), 
                None, None)

        elif output_to == 'replace':
            self.view.set_name(command)
            self.view.replace(edit, sublime.Region(0, self.view.size()), self.run_commands(command))

        else:
            log_output( self.run_commands(command) )


class AccuTermUnlock(sublime_plugin.WindowCommand):
    def on_done(self, item_ref):
        item_ref = item_ref.split()
        if len(item_ref) == 2:
            [mv_file, mv_item] = item_ref
            mv_svr = connect()
            if mv_svr:
                mv_svr.UnlockItem(mv_file, mv_item)
                check_error_message(mv_svr, mv_file + ' ' + mv_item + ' unlocked')
                mv_svr.Disconnect()
        else:
            log_output('Invalid Input: ' + item_ref + ' (Must be [file] [item])')

    def run(self, **kwargs):
        self.window.show_input_panel('Enter the MV file and item', '', self.on_done, None, None)


class AccuTermRefreshCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        (mv_file, mv_item) = get_file_item(self.view)
        mv_svr = connect()
        if mv_svr:
            mv_svr.UnlockItem(mv_file, mv_item)
            data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 1)
            if check_error_message(mv_svr, mv_file + ' ' + mv_item + ' refreshed and locked'):
                self.view.window().open_file(file_name).run_command('accu_term_replace_file', {"text": data})
            mv_svr.Disconnect()



class AccuTermListCommand(sublime_plugin.WindowCommand):
    # def __init__(self, window):
    #     super(AccuTermListCommand, self).__init__()
    #     mv_svr = connect()
    #     if mv_svr:
    #         self.md_list = mv_svr.Execute('sort only ' + mv_svr.MDName + ' = "A" "D" *A0 (JICN')

    def run(self, **kwargs):
        self.mv_svr = connect()
        if self.mv_svr:
            if self.mv_svr.MDName == 'VOC':
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_svr.MDName + ' WITH A1 = "F" "Q" A0 COL-HDR-SUPP ID-SUPP NOPAGE COUNT.SUP', '', 1)).split('\r\n')
            else:
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_svr.MDName + ' WITH A1 = "D" "Q" A0 COL-HDR-SUPP ID-SUPP NOPAGE NI-SUPP', '', 1)).split('\r\n')
            if check_error_message(self.mv_svr, ''):
                sublime.active_window().show_quick_panel(self.list, self.listFile)

    def listFile(self, list_index):
        if list_index > -1:
            self.mv_file = self.list[list_index]
            if self.mv_svr.MDName == 'VOC':
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_file + ' A0 COL-HDR-SUPP ID-SUPP NOPAGE COUNT.SUP', '', 1)).split('\r\n')
            else:
                self.list = ''.join(self.mv_svr.Execute('SORT ' + self.mv_file + ' A0 COL-HDR-SUPP ID-SUPP NOPAGE NI-SUPP', '', 1)).split('\r\n')
            sublime.active_window().show_quick_panel(self.list, self.pickItem)

    def pickItem(self, item_index):
        if item_index > -1:
            mv_file = self.mv_file
            mv_item = self.list[item_index]
            file_name = os.sep.join([get_base_path(), mv_file, mv_item + '.bp'])
            if self.mv_svr.IsConnected():
                self.mv_svr.UnlockItem(mv_file, mv_item)
                data = self.mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 1)
                if check_error_message(self.mv_svr, 'Download success'):
                    if not os.path.exists(file_name): 
                        new_view = self.window.new_file()
                        new_view.set_name(mv_item + '.bp')
                        default_dir = get_base_path() + os.sep + mv_file
                        if not os.path.exists(default_dir): os.makedirs(default_dir)
                        new_view.settings().set('default_dir', default_dir)
                    else: 
                        new_view = self.window.open_file(file_name)
                    new_view.run_command('accu_term_replace_file', {"text": data})
                self.mv_svr.Disconnect()


class AccuTermLockCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        (mv_file, mv_item) = get_file_item(self.view)
        mv_svr = connect()
        if mv_svr:
            data = mv_svr.Readitem(mv_file, mv_item, 0, 0, 0, 1)
            if mv_svr.LastError == 260:
                sublime.active_window().destroy_output_panel('AccuTermClient')
                sublime.active_window().status_message(mv_file + ' ' + mv_item + ' is already locked')
            else :
                check_error_message(mv_svr, mv_file + ' ' + mv_item + ' locked')
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
        cur_view = sublime.active_window().active_view()
        source_upcase = changeCase( cur_view.substr(sublime.Region(0, cur_view.size())), 'upper')
        sublime.set_timeout_async(lambda: cur_view.run_command('accu_term_replace_file', {"text": source_upcase}), 0)

class AccuTermGlobalDowncase(sublime_plugin.TextCommand):
    def run(self, edit):       
        cur_view = sublime.active_window().active_view()
        source_upcase = changeCase( cur_view.substr(sublime.Region(0, cur_view.size())), 'lower')
        cur_view.run_command('accu_term_replace_file', {"text": source_upcase})        
        
class EventListener(sublime_plugin.EventListener):
    def on_pre_close(self, view):
        if view.scope_name(0).split('.')[-1].strip() == 'd3-basic':
            view.run_command('accu_term_release')
    