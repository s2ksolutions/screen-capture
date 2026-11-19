"""Launcher for capture_gallery.py. Keep both files together and build this one.
API key storage is PLAINTEXT in local settings.json by user choice.
Calculator view changes presentation only, not process visibility or capture protections.
"""
import ast
import math
import operator
import time
import tkinter as tk
from tkinter import ttk, messagebox
from capture_gallery import App, settings_read, settings_save


class PlaintextKeyStore:
    def get_password(self, service, account):
        value = settings_read().get('gemini_api_key', '')
        return value if isinstance(value, str) else ''

    def set_password(self, service, account, password):
        settings_save(gemini_api_key=password)

    def delete_password(self, service, account):
        settings_save(gemini_api_key='')


class CalculatorViewApp(App):
    def __init__(self, root):
        self.warmup_pending = False
        self.warmup_after = None
        self.warmup_deadline = None
        self.saved_geometry = None
        self.cover_visible = False
        self.missing_key_error = False
        self.key_error_paused = False
        super().__init__(root)
        try:
            settings = settings_read()
        except Exception:
            settings = {}
        self.calculator_view = tk.BooleanVar(value=bool(settings.get('calculator_view', False)))
        try:
            seconds = float(settings.get('warmup_seconds', 5))
            if not math.isfinite(seconds) or not 0 <= seconds <= 600:
                seconds = 5
        except (TypeError, ValueError):
            seconds = 5
        self.warmup_seconds = tk.StringVar(value=f'{seconds:g}')
        self.countdown_text = tk.StringVar(value='Ready')
        self.expression = tk.StringVar()
        self.build_options()
        self.build_calculator()
        self.controls()
        self.key.trace_add('write', self.on_api_key_changed)
        if self.key.get():
            self.api_status.set('API key loaded from local plaintext settings.')

    def vault(self):
        return PlaintextKeyStore()

    def save_key(self):
        value = self.key.get().strip()
        if not value:
            self.fail('API key not entered.')
            return
        try:
            self.vault().set_password(self.SERVICE, self.ACCOUNT, value)
            self.api_status.set('API key saved in local plaintext settings.')
        except Exception as exc:
            self.fail('Could not save API key: ' + str(exc))

    def delete_key(self):
        if not messagebox.askyesno(
            'Delete saved API key?',
            'Clear the local settings key and key field? This does not revoke '
            'the key at Google, erase backups, or cancel an active request.',
            parent=self.root, default='no'):
            return
        try:
            self.vault().delete_password(self.SERVICE, self.ACCOUNT)
            self.key.set('')
            self.stop()
            self.api_status.set('Saved plaintext API key cleared.')
        except Exception as exc:
            self.fail('Could not clear saved API key: ' + str(exc))

    def on_api_key_changed(self, *args):
        if not self.key.get().strip() or not self.missing_key_error:
            return
        # Only the exact known missing-key error may be cleared.
        # Do not erase a newer error that replaced it.
        if getattr(self, 'missing_key_message', None) != self.error.get():
            self.missing_key_error = False
            return
        self.missing_key_error = False
        self.error.set('')
        self.api_status.set('API key entered. Start capture or explicitly resend to continue.')
        self.status.set('Missing-key error cleared. Capture and queued sends remain stopped.')
        if hasattr(self, 'countdown_text'):
            self.countdown_text.set('Ready. Click Start capture when ready.')
        self.controls()

    def require_key(self):
        if self.key.get().strip():
            return True
        self.fail('API key not entered.')
        return False

    def confirm_send(self, cid, text):
        if not self.require_key():
            return
        # Temporarily permit scheduling for this explicit action only.
        previous = self.key_error_paused
        before = set(self.allowed)
        super().confirm_send(cid, text)
        if set(self.allowed) != before:
            self.key_error_paused = False
        else:
            self.key_error_paused = previous

    def refresh_models(self):
        if self.require_key():
            return super().refresh_models()

    def toggle_auto(self):
        if self.auto_var.get() and not self.require_key():
            self.auto_var.set(False)
            return
        return super().toggle_auto()

    def build_options(self):
        self.options = ttk.Frame(self.root, padding=8)
        # Pack before the expanding notebook so the footer always gets space.
        self.options.pack(side='bottom', fill='x', before=self.tabs)
        self.view_check = ttk.Checkbutton(
            self.options,
            text='Calculator view when Start capture is clicked',
            variable=self.calculator_view)
        self.view_check.pack(side='left')
        ttk.Label(self.options, text='Warm-up (0–600 seconds):').pack(side='left', padx=(15, 5))
        self.warmup_input = ttk.Spinbox(
            self.options, from_=0, to=600, increment=5,
            width=8, textvariable=self.warmup_seconds)
        self.warmup_input.pack(side='left')
        self.save_options_button = ttk.Button(
            self.options, text='Save view + warm-up', command=self.save_view_options)
        self.save_options_button.pack(side='left', padx=8)
        ttk.Label(self.options, textvariable=self.countdown_text, wraplength=320).pack(side='left', padx=8)

    def build_calculator(self):
        self.cover = ttk.Frame(self.root, padding=14)
        for column in range(4):
            self.cover.columnconfigure(column, weight=1)
        ttk.Entry(self.cover, textvariable=self.expression, justify='right',
                  font=('Segoe UI', 22)).grid(
            row=0, column=0, columnspan=4, sticky='ew', pady=(0, 10))
        rows = (('7','8','9','/'), ('4','5','6','*'),
                ('1','2','3','-'), ('0','.','=','+'))
        for row, values in enumerate(rows, 1):
            self.cover.rowconfigure(row, weight=1)
            for column, label in enumerate(values):
                ttk.Button(self.cover, text=label,
                    command=lambda label=label: self.keypress(label)).grid(
                    row=row, column=column, sticky='nsew', padx=3, pady=3, ipady=10)
        ttk.Button(self.cover, text='Clear', command=lambda:self.expression.set('')).grid(
            row=5, column=0, columnspan=2, sticky='ew', padx=3, pady=6)
        ttk.Button(self.cover, text='Backspace',
            command=lambda:self.expression.set(self.expression.get()[:-1])).grid(
            row=5, column=2, columnspan=2, sticky='ew', padx=3, pady=6)
        ttk.Label(self.cover, textvariable=self.countdown_text, wraplength=440).grid(
            row=6, column=0, columnspan=4, sticky='w', pady=(8, 3))
        tk.Label(self.cover, textvariable=self.error, fg='#b00020',
                 anchor='w', justify='left', wraplength=440).grid(
            row=7, column=0, columnspan=4, sticky='ew', pady=5)
        ttk.Button(self.cover, text='Stop', command=self.stop).grid(
            row=8, column=0, columnspan=2, sticky='ew', padx=3, pady=6)
        ttk.Button(self.cover, text='Show full interface', command=self.show_main).grid(
            row=8, column=2, columnspan=2, sticky='ew', padx=3, pady=6)

    def keypress(self, key):
        if key != '=':
            value = self.expression.get()
            self.expression.set((('' if value == 'Error' else value) + key)[:120])
            return
        try:
            tree = ast.parse(self.expression.get(), mode='eval')
            if sum(1 for _ in ast.walk(tree)) > 80:
                raise ValueError('Expression too long')
            binary = {ast.Add:operator.add, ast.Sub:operator.sub,
                      ast.Mult:operator.mul, ast.Div:operator.truediv}
            unary = {ast.UAdd:operator.pos, ast.USub:operator.neg}
            def evaluate(node):
                if isinstance(node, ast.Expression):
                    return evaluate(node.body)
                if isinstance(node, ast.Constant) and type(node.value) in (int, float):
                    value = node.value
                elif isinstance(node, ast.BinOp) and type(node.op) in binary:
                    value = binary[type(node.op)](evaluate(node.left), evaluate(node.right))
                elif isinstance(node, ast.UnaryOp) and type(node.op) in unary:
                    value = unary[type(node.op)](evaluate(node.operand))
                else:
                    raise ValueError('Unsupported expression')
                if not math.isfinite(value) or abs(value) > 1e100:
                    raise ValueError('Result out of range')
                return value
            self.expression.set(f'{evaluate(tree):.12g}')
        except Exception:
            self.expression.set('Error')

    def save_view_options(self):
        try:
            seconds = float(self.warmup_seconds.get())
            if not math.isfinite(seconds) or not 0 <= seconds <= 600:
                raise ValueError('Warm-up must be between 0 and 600 seconds.')
            settings_save(warmup_seconds=seconds, calculator_view=self.calculator_view.get())
            self.countdown_text.set('View and warm-up settings saved.')
            return seconds
        except Exception as exc:
            self.fail('Saving view/warm-up: ' + str(exc))
            return None

    def show_calculator(self):
        if self.cover_visible:
            return
        self.saved_geometry = self.root.geometry()
        self.tabs.pack_forget()
        self.options.pack_forget()
        self.root.minsize(420, 520)
        self.root.geometry('480x610')
        self.cover.pack(fill='both', expand=True)
        self.cover_visible = True

    def show_main(self):
        if not self.cover_visible:
            return
        self.cover.pack_forget()
        self.options.pack(side='bottom', fill='x')
        self.tabs.pack(fill='both', expand=True)
        self.root.minsize(1050, 740)
        self.root.geometry(self.saved_geometry or '1320x980')
        self.cover_visible = False

    def start_capture(self, mode):
        if mode != 'capture':
            self.show_main()
            return super().start_capture(mode)
        if self.warmup_pending or self.busy or self.request_busy or self.selector or self.closing:
            return
        if self.auto_active and not self.require_key():
            return
        seconds = self.save_view_options()
        if seconds is None:
            return
        self.error.set('')
        self.missing_key_error = False
        self.warmup_pending = True
        self.warmup_deadline = time.monotonic() + seconds
        self.controls()
        if self.calculator_view.get():
            self.show_calculator()
        self.warmup_tick()

    def warmup_tick(self):
        self.warmup_after = None
        if not self.warmup_pending or self.closing:
            return
        remaining = self.warmup_deadline - time.monotonic()
        if remaining > 0:
            value = f'Capture starts in {math.ceil(remaining)} seconds. Stop cancels.'
            self.countdown_text.set(value)
            self.status.set(value)
            self.warmup_after = self.root.after(100, self.warmup_tick)
            return
        self.warmup_pending = False
        self.key_error_paused = False
        self.countdown_text.set('Initializing capture/OCR...')
        super().start_capture('capture')
        if not self.busy and self.error.get():
            self.countdown_text.set('Stopped after error.')

    def cancel_warmup(self):
        self.warmup_pending = False
        if self.warmup_after is not None:
            self.root.after_cancel(self.warmup_after)
            self.warmup_after = None

    def controls(self):
        super().controls()
        if not hasattr(self, 'options'):
            return
        locked = self.warmup_pending or self.busy or self.closing
        for widget in (self.view_check, self.warmup_input, self.save_options_button):
            widget.configure(state='disabled' if locked else 'normal')
        if self.warmup_pending:
            for widget in self.locked + [self.start_btn]:
                widget.configure(state='disabled')

    def fail(self, detail):
        self.cancel_warmup()
        normalized = str(detail).casefold()
        key_variable = getattr(self, 'key', None)
        missing = key_variable is not None and not key_variable.get().strip()
        phrases = ('api key not entered', 'missing gemini key',
                   'enter api key and model', 'enter a key before refreshing models')
        self.missing_key_error = missing and any(p in normalized for p in phrases)
        if self.missing_key_error:
            self.key_error_paused = True
        super().fail(detail)
        self.missing_key_message = self.error.get() if self.missing_key_error else None
        if hasattr(self, 'countdown_text'):
            self.countdown_text.set('Stopped after error. See the red message.')

    def stop(self):
        self.cancel_warmup()
        super().stop()
        if hasattr(self, 'countdown_text'):
            self.countdown_text.set('Stopped.')

    def poll(self):
        super().poll()
        if hasattr(self, 'countdown_text') and self.busy and not self.error.get():
            self.countdown_text.set('Capture active. Stop ends capture and pending sends.')

    def scheduler(self):
        if self.warmup_pending or self.key_error_paused:
            return
        return super().scheduler()

    def close(self):
        super().close()
        if self.closing:
            self.cancel_warmup()


if __name__ == '__main__':
    root = tk.Tk()
    CalculatorViewApp(root)
    root.mainloop()
