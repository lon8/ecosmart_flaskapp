import PySimpleGUI as sg

class LoguiError(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
        return f"Ошибка: {self.message}"


class Logui:
    def __init__(self, title):
        outputwin = [
            [sg.Output(size=(160, 30))]
        ]
        layout = [
            # [sg.Frame('Output', layout=outputwin)]
            [outputwin]
        ]
        self.window = sg.Window(title, layout)
        self.processing = False
        self.after_processing_complete = True

    def close(self):
        self.window.close()

    def proceed(self):
        event, values = self.window.read(timeout=10)
        if event in (None, 'Exit'):
            self.window.close()
            # raise LoguiError("Процесс остановлен.")
            return False
        return True

    def can_processing(self):
        if not self.processing and self.after_processing_complete:
            self.processing = True
            return True
        return False

    def complete_processing(self):
        self.processing = False
        self.after_processing_complete = False

    def perform_long_operation(self, func):
        self.window.perform_long_operation(func, 'CONVERT_EVENT')

    def can_after_processing(self):
        return not self.processing and not self.after_processing_complete

    def complete_after_processing(self):
        self.after_processing_complete = True
