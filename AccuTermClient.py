import sublime
import sublime_plugin
import Pywin32.setup
from win32com.client import Dispatch
import os
import threading
import pythoncom
import re

base_path = 'H:\\Programs\\MV Basic'
# todo: move this into some kind of settings file

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
        log_output('Connected', panel_name)
        return mv_svr
    else: 
        log_output('Connection Failed', panel_name)
        return None


def check_error_message(mv_svr, success_msg='Success'):
    if mv_svr.LastErrorMessage:
        log_output(str(mv_svr.LastError) + " " + mv_svr.LastErrorMessage)
        return False
    else:
        log_output(success_msg)
        sublime.active_window().destroy_output_panel('AccuTermClient')
        sublime.active_window().status_message(success_msg)
        return True

def get_base_path():
    project_file_name = sublime.active_window().project_file_name()
    return os.path.dirname(project_file_name) if bool(project_file_name) else base_path

class AccuTermUploadCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        file_name = self.view.file_name()
        mv_file   = file_name.split(os.sep)[-2]
        mv_item   = file_name.split(os.sep)[-1]
        if mv_item[-3:].lower() == '.bp': mv_item = mv_item[:-3]
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
        file_name = cur_view.file_name()
        mv_file   = file_name.split(os.sep)[-2]
        mv_item   = file_name.split(os.sep)[-1]
        if mv_item[-3:].lower() == '.bp': mv_item = mv_item[:-3]
        data = cur_view.substr(sublime.Region(0, cur_view.size())).replace('\n', '\xFE')        
        sublime.set_timeout_async(lambda: self.upload(self, file_name = file_name, 
                                                        mv_file = mv_file, 
                                                        mv_item = mv_item, 
                                                        data = data), 
                                                        0)

    def upload(self, *args, file_name=None, mv_file=None, mv_item=None, data=None):
        if threading.currentThread ().getName() != 'MainThread':
           pythoncom.CoInitialize ()

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
                result = mv_svr.Execute('compile ' + mv_file + ' ' + mv_item)
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
        mv_file   = file_name.split(os.sep)[-2]
        mv_item   = file_name.split(os.sep)[-1]
        if mv_item[-3:].lower() == '.bp': mv_item = mv_item[:-3]
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
            mv_file = item_ref[0]
            mv_item = item_ref[1]
            if mv_item[-3:].lower() == '.bp': mv_item = mv_item[:-3]
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


class AccuTermUnlock(sublime_plugin.WindowCommand):
    def on_done(self, item_ref):
        item_ref = item_ref.split()
        if len(item_ref) == 2:
            mv_file = item_ref[0]
            mv_item = item_ref[1]
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
        mv_file   = file_name.split(os.sep)[-2]
        mv_item   = file_name.split(os.sep)[-1]
        if mv_item[-3:].lower() == '.bp': mv_item = mv_item[:-3]
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
            self.list = ''.join(self.mv_svr.Execute('sort only ' + self.mv_svr.MDName + ' with *A1 = "Q]" or with *a1 = "D]" *A0 FMT"LX" (JICN', '', 1)).split('\r\n')
            if check_error_message(self.mv_svr, ''):
                sublime.active_window().show_quick_panel(self.list, self.listFile)

    def listFile(self, list_index):
        if list_index > -1:
            self.mv_file = self.list[list_index]
            self.list = ''.join(self.mv_svr.Execute('sort only ' + self.mv_file + ' *A0 FMT"LX" (JICN', '', 1)).split('\r\n')
            sublime.active_window().show_quick_panel(self.list, self.pickItem)

    def pickItem(self, item_index):
        if item_index > -1:
            mv_file = self.mv_file
            mv_item = self.list[item_index]
            if mv_item[-3:].lower() == '.bp': mv_item = mv_item[:-3]
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
        mv_file   = file_name.split(os.sep)[-2]
        mv_item   = file_name.split(os.sep)[-1]
        if mv_item[-3:].lower() == '.bp': mv_item = mv_item[:-3]
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
    