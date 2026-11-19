"""Windows capture/OCR/Gemini base application.
Keep beside calculator_launcher.py and build the LAUNCHER for calculator view,
warm-up and plaintext API-key storage. Automatic mode starts OFF.
Model discovery does not verify pricing or free quota. Retries can duplicate usage.
"""
import json
import logging
import math
import os
import queue
import sqlite3
import sys
import threading
import time
import tkinter as tk
import unicodedata
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk


def now():
    return datetime.now().isoformat(timespec='milliseconds')


def app_dir():
    return Path(sys.executable if getattr(sys,'frozen',False) else __file__).resolve().parent


def settings_path():
    location=os.environ.get('LOCALAPPDATA')
    if not location:
        raise RuntimeError('Windows LOCALAPPDATA is unavailable')
    return Path(location)/'CalculatorScreenshotOCR'/'settings.json'


def settings_read():
    path=settings_path()
    data=json.loads(path.read_text(encoding='utf-8')) if path.exists() else {}
    if not isinstance(data,dict):
        raise ValueError('Settings must be a JSON object')
    return data


def settings_save(**changes):
    data=settings_read()
    data.update(changes)
    path=settings_path()
    path.parent.mkdir(parents=True,exist_ok=True)
    temp=path.with_suffix('.tmp')
    temp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    temp.replace(path)


def clean(text,ascii_only=False):
    text=unicodedata.normalize('NFKC',text).replace('\r\n','\n').replace('\r','\n')
    text=''.join(c for c in text if c in '\n\t' or (c!='\ufffd' and not unicodedata.category(c).startswith('C')))
    if ascii_only:
        text=text.translate(str.maketrans({'“':'"','”':'"','‘':"'",'’':"'",'—':'-','–':'-','…':'...'}))
        text=unicodedata.normalize('NFKD',text).encode('ascii','ignore').decode('ascii')
    return '\n'.join(' '.join(line.split()) for line in text.splitlines() if line.strip())


def connect(base):
    base.mkdir(parents=True,exist_ok=True)
    con=sqlite3.connect(base/'gallery.sqlite3',timeout=15)
    con.row_factory=sqlite3.Row
    con.executescript("""
    CREATE TABLE IF NOT EXISTS captures (
      id INTEGER PRIMARY KEY, stamp TEXT NOT NULL, session TEXT NOT NULL,
      image TEXT NOT NULL, cleaned TEXT NOT NULL, raw TEXT NOT NULL,
      search_text TEXT NOT NULL, error TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS jobs (
      id INTEGER PRIMARY KEY, capture_id INTEGER, created TEXT NOT NULL,
      model TEXT NOT NULL, prompt TEXT NOT NULL, status TEXT NOT NULL,
      attempts INTEGER NOT NULL DEFAULT 0, max_attempts INTEGER NOT NULL,
      answer TEXT NOT NULL DEFAULT '', error TEXT NOT NULL DEFAULT '',
      completed TEXT NOT NULL DEFAULT '', automatic INTEGER NOT NULL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS attempts (
      id INTEGER PRIMARY KEY, job_id INTEGER NOT NULL, number INTEGER NOT NULL,
      started TEXT NOT NULL, finished TEXT NOT NULL DEFAULT '',
      outcome TEXT NOT NULL DEFAULT 'running', error TEXT NOT NULL DEFAULT '');
    CREATE INDEX IF NOT EXISTS jobs_capture ON jobs(capture_id);
    """)
    con.commit()
    return con


def text_set(widget,value):
    widget.configure(state='normal')
    widget.delete('1.0','end')
    widget.insert('1.0',value)
    widget.configure(state='disabled')


def bounded(value,low,high,default):
    try:
        value=float(value)
        return value if math.isfinite(value) and low<=value<=high else default
    except (TypeError,ValueError):
        return default


class Speech:
    """All SAPI calls live on one COM-initialized worker thread."""
    def __init__(self,events):
        self.events=events
        self.lock=threading.Lock()
        self.wake=threading.Event()
        self.revision=0
        self.text=''
        self.enabled=True
        self.loop=False
        self.repeats=1
        self.rate=0
        self.shutdown=False
        threading.Thread(target=self.run,daemon=True).start()

    def configure(self,enabled,loop,repeats,rate):
        with self.lock:
            self.enabled,self.loop,self.repeats,self.rate=enabled,loop,repeats,rate
            if not enabled:
                self.text=''
                self.revision+=1
        self.wake.set()

    def say(self,text):
        with self.lock:
            self.revision+=1
            self.text=text if self.enabled else ''
        self.wake.set()

    def stop(self):
        self.say('')

    def close(self):
        with self.lock:
            self.shutdown=True
            self.revision+=1
            self.text=''
        self.wake.set()

    def run(self):
        voice=None
        com_ready=False
        revision=-1
        count=0
        try:
            while True:
                with self.lock:
                    shutdown,rev,text=self.shutdown,self.revision,self.text
                    enabled,loop,repeats,rate=self.enabled,self.loop,self.repeats,self.rate
                if shutdown:
                    break
                try:
                    if rev!=revision:
                        if voice is not None:
                            voice.Speak('',3)
                        revision,count=rev,0
                    if not text or not enabled or (not loop and count>=repeats):
                        self.wake.wait(.15)
                        self.wake.clear()
                        continue
                    if voice is None:
                        import pythoncom
                        import win32com.client
                        if not com_ready:
                            pythoncom.CoInitialize()
                            com_ready=True
                        voice=win32com.client.Dispatch('SAPI.SpVoice')
                    voice.Rate=rate
                    voice.Speak(text,1|2|16)
                    while not voice.WaitUntilDone(100):
                        with self.lock:
                            changed=self.shutdown or self.revision!=revision or not self.enabled
                        if changed:
                            voice.Speak('',3)
                            break
                    count+=1
                    self.wake.wait(.5)
                    self.wake.clear()
                except Exception as exc:
                    with self.lock:
                        if self.revision==rev:
                            self.text=''
                            self.revision+=1
                    voice=None
                    self.events.put(('speech_error',str(exc)))
                    self.wake.wait(.2)
                    self.wake.clear()
        finally:
            if voice is not None:
                try:
                    voice.Speak('',3)
                except Exception:
                    pass
            if com_ready:
                import pythoncom
                pythoncom.CoUninitialize()


class App:
    SERVICE='CalculatorScreenshotOCR.Gemini'
    ACCOUNT='api-key'
    MODELS=('gemini-3.5-flash-lite',)
    DEFAULT=('Answer the question below. Give only the correct answer. '
             'If the question is incomplete or cannot be answered, say so rather than guessing.')

    def __init__(self,root):
        self.root=root
        root.title('Calculator')
        root.geometry('1320x980')
        root.minsize(1050,740)
        settings_error=None
        try:
            settings=settings_read()
        except Exception as exc:
            settings,settings_error={},str(exc)
        self.events=queue.Queue()
        self.stop_event=threading.Event()
        self.lock=threading.Lock()
        self.worker=self.selector=None
        self.busy=self.closing=self.request_busy=self.models_busy=False
        self.active_job=None
        self.allowed,self.due=set(),{}
        self.last_start=self.last_complete=self.last_saved=None
        self.manual_seconds=self.capture_seconds=2.0
        self.auto_active=False
        self.crop=None
        self.generation=0
        self.rows,self.thumbs={},[]
        self.selected_image=self.photo=None
        self.auto_seconds=bounded(settings.get('auto_interval',30),10,3600,30)
        self.cooldown_seconds=bounded(settings.get('auto_cooldown',10),0,3600,10)
        self.folder=tk.StringVar(value=str(settings.get('library_folder',app_dir()/'capture_library')))
        self.query,self.error=tk.StringVar(),tk.StringVar()
        self.status=tk.StringVar(value='Ready')
        self.api_status=tk.StringVar(value='No request sent. Verify model pricing and quota.')
        self.interval=tk.DoubleVar(value=2)
        self.interval_label=tk.StringVar(value='2 seconds')
        self.auto_var=tk.BooleanVar(value=False)
        self.auto_input=tk.StringVar(value=f'{self.auto_seconds:g}')
        self.cooldown=tk.StringVar(value=f'{self.cooldown_seconds:g}')
        self.ascii=tk.BooleanVar(value=False)
        self.region=tk.StringVar(value='Full display')
        self.key=tk.StringVar()
        self.show_key=tk.BooleanVar(value=False)
        model=settings.get('gemini_model')
        self.model=tk.StringVar(value=model if model in self.MODELS else self.MODELS[0])
        self.max_attempts=tk.StringVar(value=str(int(bounded(settings.get('max_attempts',3),1,5,3))))
        self.tts_on=tk.BooleanVar(value=bool(settings.get('tts_enabled',True)))
        self.tts_loop=tk.BooleanVar(value=bool(settings.get('tts_loop',False)))
        self.tts_repeats=tk.StringVar(value=str(int(bounded(settings.get('tts_repeats',1),1,20,1))))
        self.tts_rate=tk.DoubleVar(value=bounded(settings.get('tts_rate',0),-10,10,0))
        self.speech=Speech(self.events)
        self.build(settings)
        self.apply_speech(False)
        try:
            saved=self.vault().get_password(self.SERVICE,self.ACCOUNT)
            if saved:
                self.key.set(saved)
        except Exception:
            self.api_status.set('Saved key unavailable; enter one for this session.')
        self.recover_library()
        self.refresh()
        if settings_error:
            self.fail('Loading settings: '+settings_error)
        root.protocol('WM_DELETE_WINDOW',self.close)
        root.after(150,self.poll)

    def build(self,settings):
        self.tabs=ttk.Notebook(self.root)
        self.tabs.pack(fill='both',expand=True)
        gallery,api=ttk.Frame(self.tabs),ttk.Frame(self.tabs,padding=12)
        self.tabs.add(gallery,text='Gallery / capture')
        self.tabs.add(api,text='Gemini / speech')
        top=ttk.Frame(gallery,padding=10)
        top.pack(fill='x')
        top.columnconfigure(1,weight=1)
        ttk.Label(top,text='Library folder').grid(row=0,column=0)
        self.folder_entry=ttk.Entry(top,textvariable=self.folder)
        self.folder_entry.grid(row=0,column=1,sticky='ew',padx=8)
        self.browse_btn=ttk.Button(top,text='Choose / open library',command=self.browse)
        self.browse_btn.grid(row=0,column=2)
        ttk.Label(top,text='Manual capture interval').grid(row=1,column=0,pady=8)
        self.slider=ttk.Scale(top,from_=.5,to=60,variable=self.interval,command=self.manual_changed)
        self.slider.grid(row=1,column=1,sticky='ew')
        ttk.Label(top,textvariable=self.interval_label).grid(row=1,column=2)
        tk.Label(gallery,textvariable=self.error,fg='#b00020',anchor='w',justify='left',wraplength=1240,padx=10,pady=4).pack(fill='x')
        controls=ttk.Frame(gallery,padding=(10,0,10,8))
        controls.pack(fill='x')
        self.start_btn=ttk.Button(controls,text='Start capture',command=lambda:self.start_capture('capture'))
        self.start_btn.pack(side='left')
        self.stop_btn=ttk.Button(controls,text='Stop capture / pending sends',command=self.stop)
        self.stop_btn.pack(side='left',padx=5)
        self.region_btn=ttk.Button(controls,text='Preview / draw area',command=lambda:self.start_capture('preview'))
        self.region_btn.pack(side='left',padx=5)
        self.full_btn=ttk.Button(controls,text='Full display',command=self.full)
        self.full_btn.pack(side='left')
        ttk.Label(controls,textvariable=self.region).pack(side='left',padx=8)
        self.ascii_btn=ttk.Checkbutton(controls,text='English/ASCII cleanup',variable=self.ascii)
        self.ascii_btn.pack(side='left')
        timing=ttk.Frame(gallery,padding=(10,0,10,8))
        timing.pack(fill='x')
        self.auto_check=ttk.Checkbutton(timing,text='Automatic answer',variable=self.auto_var,command=self.toggle_auto)
        self.auto_check.pack(side='left')
        ttk.Label(timing,text='Interval (10–3600s):').pack(side='left',padx=(12,4))
        ttk.Spinbox(timing,from_=10,to=3600,increment=5,width=7,textvariable=self.auto_input).pack(side='left')
        ttk.Label(timing,text='Cooldown after API result (0–3600s):').pack(side='left',padx=(12,4))
        ttk.Spinbox(timing,from_=0,to=3600,increment=5,width=7,textvariable=self.cooldown).pack(side='left')
        ttk.Label(timing,text='Max attempts:').pack(side='left',padx=(12,4))
        ttk.Spinbox(timing,from_=1,to=5,width=4,textvariable=self.max_attempts).pack(side='left')
        ttk.Button(timing,text='Apply + save',command=self.apply_timing).pack(side='left',padx=8)
        search=ttk.Frame(gallery,padding=(10,0,10,8))
        search.pack(fill='x')
        ttk.Label(search,text='Search OCR + answers:').pack(side='left')
        entry=ttk.Entry(search,textvariable=self.query)
        entry.pack(side='left',fill='x',expand=True,padx=8)
        entry.bind('<Return>',lambda event:self.refresh())
        ttk.Button(search,text='Search / refresh',command=self.refresh).pack(side='left')
        ttk.Button(search,text='Clear',command=self.clear_search).pack(side='left',padx=5)
        self.delete_btn=ttk.Button(search,text='Delete all...',command=self.delete_all)
        self.delete_btn.pack(side='left')
        panes=ttk.Panedwindow(gallery,orient='horizontal')
        panes.pack(fill='both',expand=True,padx=10)
        left,right=ttk.Frame(panes),ttk.Frame(panes)
        panes.add(left,weight=1)
        panes.add(right,weight=2)
        self.tree=ttk.Treeview(left,columns=('when','status'),show='tree headings',selectmode='browse')
        for col,label,width in (('#0','Screenshot',95),('when','Captured',145),('status','Gemini status / OCR',230)):
            self.tree.heading(col,text=label)
            self.tree.column(col,width=width,stretch=col=='status')
        ttk.Style().configure('Gallery.Treeview',rowheight=66)
        self.tree.configure(style='Gallery.Treeview')
        self.tree.pack(side='left',fill='both',expand=True)
        bar=ttk.Scrollbar(left,orient='vertical',command=self.tree.yview)
        bar.pack(side='right',fill='y')
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.bind('<<TreeviewSelect>>',self.select)
        self.tree.bind('<Button-3>',self.context)
        self.tree.bind('<Delete>',lambda event:self.delete_selected())
        self.menu=tk.Menu(self.root,tearoff=False)
        self.menu.add_command(label='Resend to Gemini...',command=self.resend)
        self.menu.add_command(label='Speak latest saved answer',command=self.speak_selected)
        self.menu.add_separator()
        self.menu.add_command(label='Delete screenshot and related files...',command=self.delete_selected)
        nav=ttk.Frame(right)
        nav.pack(fill='x')
        ttk.Button(nav,text='Previous',command=lambda:self.move(-1)).pack(side='left')
        ttk.Button(nav,text='Next',command=lambda:self.move(1)).pack(side='left',padx=5)
        ttk.Button(nav,text='Open image',command=self.open_image).pack(side='left')
        self.image_label=ttk.Label(right,text='Select a screenshot',anchor='center')
        self.image_label.pack(fill='both',expand=True,pady=5)
        self.image_label.bind('<Configure>',lambda event:self.render_image())
        detail=ttk.Notebook(right)
        detail.pack(fill='both',expand=True)
        combined=ttk.Frame(detail)
        detail.add(combined,text='OCR + Gemini answers')
        ttk.Label(combined,text='OCR text').pack(anchor='w')
        self.ocr_view=ScrolledText(combined,height=5,wrap='word',state='disabled')
        self.ocr_view.pack(fill='both',expand=True)
        ttk.Label(combined,text='Gemini responses / request history').pack(anchor='w',pady=(5,0))
        self.answer_view=ScrolledText(combined,height=8,wrap='word',state='disabled')
        self.answer_view.pack(fill='both',expand=True)
        self.raw_view=ScrolledText(detail,height=12,wrap='word',state='disabled')
        detail.add(self.raw_view,text='Original OCR / errors')
        ttk.Label(gallery,textvariable=self.status,wraplength=1240,padding=10).pack(fill='x')
        self.locked=[self.folder_entry,self.browse_btn,self.region_btn,self.full_btn,self.ascii_btn,self.delete_btn]
        api.columnconfigure(1,weight=1)
        api.rowconfigure(7,weight=1)
        api.rowconfigure(10,weight=1)
        ttk.Label(api,text='API key').grid(row=0,column=0)
        self.key_entry=ttk.Entry(api,textvariable=self.key,show='*')
        self.key_entry.grid(row=0,column=1,sticky='ew',padx=8)
        ttk.Checkbutton(api,text='Show',variable=self.show_key,command=lambda:self.key_entry.configure(show='' if self.show_key.get() else '*')).grid(row=0,column=2)
        keys=ttk.Frame(api)
        keys.grid(row=1,column=0,columnspan=3,sticky='w',pady=6)
        ttk.Button(keys,text='Save key',command=self.save_key).pack(side='left')
        ttk.Button(keys,text='Delete saved key',command=self.delete_key).pack(side='left',padx=6)
        ttk.Button(keys,text='Get API key',command=lambda:self.url('https://aistudio.google.com/apikey')).pack(side='left')
        ttk.Button(keys,text='Pricing / quotas',command=lambda:self.url('https://ai.google.dev/gemini-api/docs/pricing')).pack(side='left',padx=6)
        ttk.Label(api,text='Model candidates').grid(row=2,column=0)
        self.model_box=ttk.Combobox(api,textvariable=self.model,values=self.MODELS,state='readonly')
        self.model_box.grid(row=2,column=1,sticky='ew',padx=8)
        ttk.Button(api,text='Refresh available',command=self.refresh_models).grid(row=2,column=2)
        ttk.Label(api,text='Model list is not exhaustive. Discovery does not verify free pricing or remaining quota.',wraplength=1100).grid(row=3,column=0,columnspan=3,sticky='w')
        defaults=ttk.Frame(api)
        defaults.grid(row=4,column=0,columnspan=3,sticky='ew',pady=(8,4))
        ttk.Label(defaults,text='Default prompt').pack(side='left')
        ttk.Button(defaults,text='Save prompt + model',command=self.save_defaults).pack(side='right')
        self.default=ScrolledText(api,height=3,wrap='word')
        self.default.grid(row=5,column=0,columnspan=3,sticky='ew')
        value=settings.get('default_prompt',self.DEFAULT)
        self.default.insert('1.0',value if isinstance(value,str) else self.DEFAULT)
        promptrow=ttk.Frame(api)
        promptrow.grid(row=6,column=0,columnspan=3,sticky='ew',pady=(8,4))
        ttk.Label(promptrow,text='Manual question (linked to selected screenshot if one is selected)').pack(side='left')
        ttk.Button(promptrow,text='Append selected OCR',command=self.append_ocr).pack(side='right')
        self.prompt=ScrolledText(api,height=5,wrap='word')
        self.prompt.grid(row=7,column=0,columnspan=3,sticky='nsew')
        send=ttk.Frame(api)
        send.grid(row=8,column=0,columnspan=3,sticky='ew',pady=8)
        ttk.Button(send,text='Queue send to Gemini',command=self.manual_send).pack(side='left')
        ttk.Label(send,textvariable=self.api_status,wraplength=1000).pack(side='left',padx=8)
        ttk.Label(api,text='Latest response (also saved with its linked screenshot)').grid(row=9,column=0,columnspan=3,sticky='w')
        self.response=ScrolledText(api,height=6,wrap='word',state='disabled')
        self.response.grid(row=10,column=0,columnspan=3,sticky='nsew')
        speechrow=ttk.LabelFrame(api,text='Local Windows speech',padding=8)
        speechrow.grid(row=11,column=0,columnspan=3,sticky='ew',pady=8)
        ttk.Checkbutton(speechrow,text='Read received answers',variable=self.tts_on).pack(side='left')
        ttk.Checkbutton(speechrow,text='Loop until next answer',variable=self.tts_loop).pack(side='left',padx=8)
        ttk.Label(speechrow,text='Otherwise repeats (1–20):').pack(side='left')
        ttk.Spinbox(speechrow,from_=1,to=20,width=4,textvariable=self.tts_repeats).pack(side='left',padx=5)
        ttk.Label(speechrow,text='Slow').pack(side='left',padx=(8,0))
        ttk.Scale(speechrow,from_=-10,to=10,variable=self.tts_rate,length=130).pack(side='left')
        ttk.Label(speechrow,text='Fast').pack(side='left')
        ttk.Button(speechrow,text='Apply + save',command=self.apply_speech).pack(side='left',padx=8)
        ttk.Button(speechrow,text='Stop speech',command=self.speech.stop).pack(side='left')

    def base(self):
        value=self.folder.get().strip()
        if not value:
            raise ValueError('Choose a library folder')
        return Path(value).expanduser().resolve()

    def sql(self,statement,params=(),fetch=False):
        con=connect(self.base())
        try:
            result=con.execute(statement,params)
            value=result.fetchall() if fetch else result.lastrowid
            con.commit()
            return value
        finally:
            con.close()

    def recover_library(self):
        try:
            self.sql("UPDATE attempts SET outcome='interrupted',finished=? WHERE outcome='running'",(now(),))
            self.sql("UPDATE jobs SET status='interrupted',error='Application ended before a response was recorded. Resend explicitly to retry.' WHERE status='inflight'")
        except Exception as exc:
            self.fail('Opening library: '+str(exc))

    def fail(self,detail):
        value=str(detail)
        key=self.key.get().strip()
        if key:
            value=value.replace(key,'[REDACTED]')
        self.error.set('ERROR — '+value[:3000])
        self.stop_event.set()
        self.speech.stop()
        self.status.set('Capture and queued sending paused. Fix error, then Start capture or explicitly resend.')
        self.api_status.set('Paused after error; see red message on Gallery / capture.')
        self.controls()

    def vault(self):
        from keyring.backends.Windows import WinVaultKeyring
        return WinVaultKeyring()

    def url(self,value):
        import webbrowser
        webbrowser.open(value)

    def save_key(self):
        if not self.key.get().strip():
            self.fail('API key not entered.')
            return
        try:
            self.vault().set_password(self.SERVICE,self.ACCOUNT,self.key.get().strip())
            self.api_status.set('Saved key in Windows credentials.')
        except Exception:
            self.fail('Could not save key; no plaintext fallback used.')

    def delete_key(self):
        if not messagebox.askyesno('Delete saved key?','Remove local credential? This does not revoke key or cancel an active request.',parent=self.root,default='no'):
            return
        try:
            store=self.vault()
            if store.get_password(self.SERVICE,self.ACCOUNT) is not None:
                store.delete_password(self.SERVICE,self.ACCOUNT)
            self.key.set('')
            self.stop()
        except Exception:
            self.fail('Could not delete saved credential')

    def save_defaults(self):
        try:
            settings_save(default_prompt=self.default.get('1.0','end-1c'),gemini_model=self.model.get())
            self.api_status.set('Default prompt and model saved.')
        except Exception as exc:
            self.fail('Saving prompt: '+str(exc))

    def apply_speech(self,save=True):
        try:
            repeats=int(self.tts_repeats.get())
            rate=int(round(self.tts_rate.get()))
            if not 1<=repeats<=20:
                raise ValueError('Speech repeats must be 1–20')
            if not -10<=rate<=10:
                raise ValueError('Speech rate must be -10 to 10')
            if save:
                settings_save(tts_enabled=self.tts_on.get(),tts_loop=self.tts_loop.get(),tts_repeats=repeats,tts_rate=rate)
            self.speech.configure(self.tts_on.get(),self.tts_loop.get(),repeats,rate)
        except Exception as exc:
            self.fail('Speech settings: '+str(exc))

    def apply_timing(self):
        try:
            interval=float(self.auto_input.get())
            cooldown=float(self.cooldown.get())
            attempts=int(self.max_attempts.get())
            if not math.isfinite(interval) or not 10<=interval<=3600:
                raise ValueError('Automatic interval must be 10–3600 seconds')
            if not math.isfinite(cooldown) or not 0<=cooldown<=3600:
                raise ValueError('Cooldown must be 0–3600 seconds')
            if not 1<=attempts<=5:
                raise ValueError('Max attempts must be 1–5')
            settings_save(auto_interval=interval,auto_cooldown=cooldown,max_attempts=attempts)
            self.auto_seconds,self.cooldown_seconds=interval,cooldown
            self.update_interval()
            return True
        except Exception as exc:
            self.fail('Timing settings: '+str(exc))
            return False

    def manual_changed(self,value):
        self.manual_seconds=round(float(value),1)
        self.update_interval()

    def update_interval(self):
        seconds=max(10.,self.auto_seconds) if self.auto_active else self.manual_seconds
        with self.lock:
            self.capture_seconds=seconds
        self.interval_label.set(f'{seconds:g} seconds'+(' (automatic)' if self.auto_active else ''))

    def toggle_auto(self):
        if not self.auto_var.get():
            self.auto_active=False
            self.update_interval()
            self.controls()
            return
        if not self.key.get().strip():
            self.auto_var.set(False)
            self.fail('API key not entered.')
            return
        if not self.model.get():
            self.auto_var.set(False)
            self.fail('Choose a Gemini model.')
            return
        if not self.apply_timing():
            self.auto_var.set(False)
            return
        question=(f'Send each nonempty captured OCR text plus default prompt to Google automatically?\n\n'
                  f'Capture/request minimum: {self.auto_seconds:g}s. Cooldown after API result: {self.cooldown_seconds:g}s. '
                  f'Up to {self.max_attempts.get()} attempts per request. Retries can duplicate usage. Speech may repeat aloud.\n\n'
                  'No screenshot images uploaded. Free-only billing is not enforced. Up to 25 pending requests; capture stops if queue fills.')
        if not messagebox.askyesno('Enable automatic sending?',question,parent=self.root,default='no'):
            self.auto_var.set(False)
            return
        self.auto_active=True
        self.update_interval()
        self.controls()

    def controls(self):
        locked=self.busy or self.selector is not None or self.closing
        for widget in self.locked:
            widget.configure(state='disabled' if locked or self.request_busy or bool(self.allowed) else 'normal')
        self.start_btn.configure(state='disabled' if locked or self.request_busy else 'normal')
        self.slider.configure(state='disabled' if self.auto_active or self.closing else 'normal')
        self.stop_btn.configure(state='disabled' if self.closing else 'normal')

    def stop(self):
        self.stop_event.set()
        self.speech.stop()
        self.auto_active=False
        self.auto_var.set(False)
        self.generation+=1
        self.update_interval()
        for jid in list(self.allowed):
            if jid!=self.active_job:
                try:
                    self.sql("UPDATE jobs SET status='paused' WHERE id=? AND status IN ('queued','retry')",(jid,))
                except Exception:
                    pass
        self.allowed.clear()
        self.due.clear()
        self.controls()
        if not self.error.get():
            self.status.set('Stopped. Pending sends paused; active server request may finish and be saved.')

    def browse(self):
        value=filedialog.askdirectory(parent=self.root)
        if value:
            self.folder.set(value)
            try:
                settings_save(library_folder=value)
            except Exception as exc:
                self.fail(str(exc))
            self.recover_library()
            self.refresh()

    def full(self):
        self.crop=None
        self.region.set('Full display')

    def slot(self):
        announced=False
        while not self.stop_event.is_set():
            with self.lock:
                interval,last,automatic=self.capture_seconds,self.last_saved,self.auto_active
            if not automatic or last is None:
                return True
            remaining=max(10.,interval)-(time.monotonic()-last)
            if remaining<=0:
                return True
            if not announced:
                self.events.put(('wait',math.ceil(remaining)))
                announced=True
            self.stop_event.wait(min(.1,remaining))
        return False

    def start_capture(self,mode):
        if self.busy or self.request_busy or self.selector or self.closing:
            return
        try:
            base=self.base()
            con=connect(base)
            con.close()
        except Exception as exc:
            self.fail('Opening capture folder: '+str(exc))
            return
        self.error.set('')
        self.stop_event.clear()
        self.busy=True
        self.controls()
        self.status.set('Preview in 2 seconds...' if mode=='preview' else 'Initializing OCR/capture; first capture follows initialization.')
        if mode=='preview':
            self.root.withdraw()
        self.worker=threading.Thread(target=self.capture_worker,args=(mode,base,self.crop,self.ascii.get()),daemon=False)
        self.worker.start()

    def capture_worker(self,mode,base,crop,ascii_only):
        camera=con=handler=logger=None
        preview=error=None
        stage='loading capture dependencies'
        try:
            import dxcam
            if mode=='preview':
                if self.stop_event.wait(2):
                    return
            else:
                import numpy as np
                from rapidocr import RapidOCR
                session=datetime.now().strftime('%Y-%m-%d/%H%M%S_%f')
                folder=base/'sessions'/session
                for name in ('images','text','raw','answers'):
                    (folder/name).mkdir(parents=True,exist_ok=True)
                logger=logging.getLogger(session)
                logger.setLevel(logging.INFO)
                logger.propagate=False
                handler=logging.FileHandler(folder/'capture.log',encoding='utf-8')
                handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
                logger.addHandler(handler)
                (folder/'settings.json').write_text(json.dumps({'crop':crop,'ascii_cleanup':ascii_only},indent=2),encoding='utf-8')
                con=connect(base)
                stage='initializing OCR'
                ocr=RapidOCR()
            if self.stop_event.is_set():
                return
            stage='initializing capture'
            camera=dxcam.create(output_idx=0,output_color='RGB')
            if camera is None:
                raise RuntimeError('No capture device')
            deadline=time.monotonic()+5
            while not self.stop_event.is_set():
                if mode=='capture' and not self.slot():
                    break
                started=time.monotonic()
                stage='capturing screenshot'
                frame=camera.grab()
                if frame is None:
                    if mode=='preview' and started>deadline:
                        raise RuntimeError('No preview frame')
                    self.stop_event.wait(.15)
                    continue
                image=Image.fromarray(frame)
                if mode=='preview':
                    preview=image.copy()
                    break
                if crop:
                    left,top,right,bottom=crop
                    if not (0<=left<right<=image.width and 0<=top<bottom<=image.height):
                        raise ValueError('Selected region no longer fits; reselect')
                    image=image.crop(crop)
                if not self.slot():
                    break
                stamp=now()
                stem=datetime.now().strftime('%H%M%S_%f')
                path=folder/'images'/(stem+'.png')
                stage='saving screenshot'
                image.save(path)
                with self.lock:
                    self.last_saved=time.monotonic()
                raw=cleaned=ocr_error=''
                try:
                    stage='OCR'
                    result=ocr(np.asarray(image)[:,:,::-1].copy())
                    lines=list(result.txts) if result is not None and result.txts is not None else []
                    raw='\n'.join(lines)
                    cleaned=clean(raw,ascii_only)
                except Exception as exc:
                    ocr_error=str(exc)
                    logger.exception('OCR failed')
                stage='saving text/index'
                (folder/'text'/(stem+'.txt')).write_text(cleaned,encoding='utf-8')
                (folder/'raw'/(stem+'.txt')).write_text(raw,encoding='utf-8')
                if ocr_error:
                    (folder/'raw'/(stem+'.error.txt')).write_text(ocr_error,encoding='utf-8')
                result=con.execute('INSERT INTO captures (stamp,session,image,cleaned,raw,search_text,error) VALUES (?,?,?,?,?,?,?)',
                    (stamp,session,str(path.relative_to(base)),cleaned,raw,clean(cleaned).casefold(),ocr_error))
                cid=result.lastrowid
                con.commit()
                logger.info('Saved capture %s',cid)
                self.events.put(('capture_saved',(cid,cleaned,not bool(ocr_error))))
                if ocr_error:
                    raise RuntimeError(ocr_error)
                while not self.stop_event.is_set():
                    with self.lock:
                        remaining=self.capture_seconds-(time.monotonic()-started)
                    if remaining<=0:
                        break
                    self.stop_event.wait(min(.1,remaining))
        except Exception as exc:
            error=stage+': '+str(exc)
            self.stop_event.set()
            if logger:
                logger.exception('Capture failed')
        finally:
            if camera is not None:
                try:
                    camera.release()
                except Exception as exc:
                    if error is None:
                        error='Capture cleanup: '+str(exc)
            if con is not None:
                con.close()
            if handler:
                handler.close()
                logger.removeHandler(handler)
            self.events.put(('capture_done',(mode,preview,error)))

    def selected_id(self):
        selected=self.tree.selection()
        return int(selected[0]) if selected else None

    def append_ocr(self):
        row=self.rows.get(str(self.selected_id()))
        if row:
            self.prompt.insert('end','\n\n--- OCR TEXT ---\n'+row['cleaned'])
            self.prompt.see('end')

    def manual_send(self):
        text=self.prompt.get('1.0','end-1c').strip()
        if not text:
            self.fail('Enter a question or append OCR text')
            return
        self.confirm_send(self.selected_id(),text)

    def resend(self):
        cid=self.selected_id()
        row=self.rows.get(str(cid))
        if not row:
            return
        if not row['cleaned'].strip():
            self.fail('Selected screenshot has no OCR text')
            return
        self.confirm_send(cid,row['cleaned'])

    def confirm_send(self,cid,text):
        if not self.key.get().strip():
            self.fail('API key not entered.')
            return
        if not self.model.get():
            self.fail('Choose a Gemini model.')
            return
        question=(f'Send default prompt and this question/OCR text to Google? Linked screenshot: {cid if cid else "none"}.\n\n'
                  f'Up to {self.max_attempts.get()} attempts. Retries can duplicate processing/usage. '
                  'No screenshot image upload. Verify pricing/quota.')
        if not messagebox.askyesno('Send to Gemini?',question,parent=self.root,default='no'):
            return
        self.error.set('')
        self.queue_job(cid,text,False)

    def queue_job(self,cid,text,automatic):
        if not text.strip():
            return
        try:
            if len(self.allowed)>=25:
                raise RuntimeError('Pending request queue reached 25. Capture paused; finish or stop pending sends before restarting.')
            attempts=int(self.max_attempts.get())
            if not 1<=attempts<=5:
                raise ValueError('Max attempts must be 1–5')
            prompt='\n\n'.join(part for part in (self.default.get('1.0','end-1c').strip(),'QUESTION / OCR TEXT:\n'+text) if part)
            jid=self.sql('INSERT INTO jobs(capture_id,created,model,prompt,status,max_attempts,automatic) VALUES(?,?,?,?,?,?,?)',
                (cid,now(),self.model.get(),prompt,'queued',attempts,int(automatic)))
            self.allowed.add(jid)
            self.due[jid]=0.
            self.export_capture(cid)
            self.controls()
            self.refresh()
        except Exception as exc:
            self.fail('Queueing request: '+str(exc))

    def scheduler(self):
        if self.request_busy or self.closing or self.error.get() or self.models_busy or not self.allowed:
            return
        clock=time.monotonic()
        minimum=max(10.,self.auto_seconds) if self.auto_active else 10.
        if self.last_start is not None and clock-self.last_start<minimum:
            return
        if self.last_complete is not None and clock-self.last_complete<self.cooldown_seconds:
            return
        for jid in sorted(self.allowed):
            if clock<self.due.get(jid,0):
                continue
            rows=self.sql('SELECT * FROM jobs WHERE id=?',(jid,),True)
            if not rows:
                self.allowed.discard(jid)
                continue
            job=dict(rows[0])
            if job['status'] not in ('queued','retry'):
                self.allowed.discard(jid)
                continue
            key=self.key.get().strip()
            if not key:
                self.fail('Missing Gemini key; sending paused')
                return
            attempt=job['attempts']+1
            self.sql("UPDATE jobs SET status='inflight',attempts=?,error='' WHERE id=?",(attempt,jid))
            aid=self.sql('INSERT INTO attempts(job_id,number,started) VALUES(?,?,?)',(jid,attempt,now()))
            # Complete fallible sidecar/UI work before marking a worker active.
            self.export_capture(job['capture_id'])
            base=self.base()
            self.request_busy=True
            self.active_job=jid
            self.last_start=clock
            self.api_status.set(f'Request #{jid}: attempt {attempt}/{job["max_attempts"]}, awaiting response...')
            self.refresh()
            self.controls()
            threading.Thread(target=self.request_worker,args=(base,job,aid,attempt,key,self.generation),daemon=True).start()
            return

    def request_worker(self,base,job,aid,attempt,key,generation):
        client=None
        answer,error='',''
        retryable,delay=False,0.
        stage='loading Gemini SDK'
        try:
            import google.genai as genai
            from google.genai import types
            stage='calling Gemini Interactions API'
            client=genai.Client(api_key=key,http_options=types.HttpOptions(timeout=60000,retry_options=types.HttpRetryOptions(attempts=1)))
            result=client.interactions.create(model=job['model'],input=job['prompt'],store=False)
            answer=getattr(result,'output_text',None)
            if not isinstance(answer,str) or not answer.strip():
                pieces=[]
                for output in (getattr(result,'outputs',None) or []):
                    if isinstance(output,dict):
                        if output.get('type')=='text' and isinstance(output.get('text'),str):
                            pieces.append(output['text'])
                    elif getattr(output,'type',None)=='text' and isinstance(getattr(output,'text',None),str):
                        pieces.append(output.text)
                answer='\n'.join(pieces)
            if not answer or not answer.strip():
                answer=''
                raise RuntimeError('No text response received. Response may be empty or blocked.')
        except Exception as exc:
            answer=''
            error=(stage+': '+str(exc)).replace(key,'[REDACTED]')[:3500]
            status=getattr(exc,'code',None) or getattr(exc,'status_code',None)
            try:
                status=int(status)
            except (TypeError,ValueError):
                status=None
            lowered=str(exc).lower()
            retryable=(status in (408,429,500,502,503,504) or isinstance(exc,(TimeoutError,ConnectionError)) or
                any(word in lowered for word in ('timed out','timeout','connection reset','connection aborted','no text response')))
            if any(word in lowered for word in ('daily quota','per day','per_day','billing disabled','insufficient credit')):
                retryable=False
            delay=max(10.,10.*(2**(attempt-1)))
            if status==429:
                delay=max(delay,60.)
            response=getattr(exc,'response',None)
            headers=getattr(response,'headers',{}) or {}
            try:
                delay=max(delay,float(headers.get('Retry-After',headers.get('retry-after',0))))
            except (ValueError,TypeError):
                pass
        finally:
            if client is not None:
                try:
                    client.close()
                except Exception:
                    pass
            save_error=None
            try:
                con=connect(base)
                try:
                    con.execute('UPDATE attempts SET finished=?,outcome=?,error=? WHERE id=?',(now(),'received' if answer else 'error',error,aid))
                    if answer:
                        con.execute("UPDATE jobs SET status='done',answer=?,error='',completed=? WHERE id=?",(answer,now(),job['id']))
                    else:
                        con.execute("UPDATE jobs SET status='failed',error=?,completed=? WHERE id=?",(error,now(),job['id']))
                    con.commit()
                finally:
                    con.close()
            except Exception as exc:
                save_error='Could not persist Gemini result: '+str(exc)
            self.events.put(('request_result',{'id':job['id'],'cid':job['capture_id'],'answer':answer,'error':error,
                'retryable':retryable,'delay':delay,'attempt':attempt,'max':job['max_attempts'],
                'generation':generation,'save_error':save_error,'finished':time.monotonic()}))

    def handle_result(self,result):
        self.request_busy=False
        self.active_job=None
        self.last_complete=result['finished']
        jid,cid=result['id'],result['cid']
        if result['save_error']:
            if result['answer']:
                text_set(self.response,result['answer'])
            self.fail(result['save_error'])
            self.controls()
            return
        self.export_capture(cid)
        if result['answer']:
            self.allowed.discard(jid)
            self.due.pop(jid,None)
            text_set(self.response,result['answer'])
            self.api_status.set(f'Request #{jid}: response received and saved.')
            if result['generation']==self.generation and not self.error.get() and not self.closing:
                self.speech.say(result['answer'])
        elif result['retryable'] and result['attempt']<result['max'] and jid in self.allowed and not self.closing:
            self.sql("UPDATE jobs SET status='retry' WHERE id=?",(jid,))
            self.due[jid]=time.monotonic()+result['delay']
            self.api_status.set(f'Request #{jid}: no answer; retry {result["attempt"]+1}/{result["max"]} in at least {result["delay"]:g}s. '+result['error'][:180])
            self.export_capture(cid)
        else:
            self.allowed.discard(jid)
            self.due.pop(jid,None)
            self.fail(f'Request #{jid} failed after {result["attempt"]} attempt(s): '+result['error'])
        self.refresh()
        self.controls()

    def capture_paths(self,row):
        base=self.base()
        sessions=(base/'sessions').resolve()
        relative=Path(row['image'])
        if relative.is_absolute():
            raise ValueError('Absolute indexed path refused')
        image=(base/relative).resolve()
        image.relative_to(sessions)
        if image.parent.name!='images' or image.suffix.lower()!='.png':
            raise ValueError('Unexpected image path')
        folder,stem=image.parent.parent,image.stem
        paths=[image,folder/'text'/(stem+'.txt'),folder/'raw'/(stem+'.txt'),folder/'raw'/(stem+'.error.txt'),
               folder/'answers'/(stem+'.json'),folder/'answers'/(stem+'.txt')]
        result=[]
        for path in paths:
            path=path.resolve()
            path.relative_to(sessions)
            result.append(path)
        return result

    def export_capture(self,cid):
        if cid is None:
            return
        rows=self.sql('SELECT * FROM captures WHERE id=?',(cid,),True)
        if not rows:
            return
        paths=self.capture_paths(rows[0])
        json_path,text_path=paths[-2:]
        jobs=[dict(row) for row in self.sql('SELECT * FROM jobs WHERE capture_id=? ORDER BY id',(cid,),True)]
        for job in jobs:
            job['attempt_history']=[dict(row) for row in self.sql('SELECT * FROM attempts WHERE job_id=? ORDER BY id',(job['id'],),True)]
        json_path.parent.mkdir(parents=True,exist_ok=True)
        data={'capture_id':cid,'image':rows[0]['image'],'requests':jobs}
        temp=json_path.with_suffix('.json.tmp')
        temp.write_text(json.dumps(data,indent=2,ensure_ascii=False),encoding='utf-8')
        temp.replace(json_path)
        blocks=[f'REQUEST {job["id"]} | {job["created"]} | {job["model"]} | {job["status"]}\n{job["answer"] or job["error"] or "Awaiting response"}' for job in jobs]
        text_path.write_text('\n\n'.join(blocks),encoding='utf-8')

    def refresh_models(self):
        if self.models_busy or self.request_busy:
            return
        key=self.key.get().strip()
        if not key:
            self.fail('Enter a key before refreshing models')
            return
        self.models_busy=True
        threading.Thread(target=self.models_worker,args=(key,),daemon=True).start()

    def models_worker(self,key):
        client=None
        try:
            import google.genai as genai
            from google.genai import types
            client=genai.Client(api_key=key,http_options=types.HttpOptions(timeout=30000))
            available={(item.name or '').removeprefix('models/') for item in client.models.list()}
            self.events.put(('models',[item for item in self.MODELS if item in available]))
        except Exception as exc:
            self.events.put(('models_error',str(exc).replace(key,'[REDACTED]')))
        finally:
            if client:
                try:
                    client.close()
                except Exception:
                    pass

    def refresh(self):
        try:
            terms=clean(self.query.get()).casefold().split()
            rows=self.sql('SELECT * FROM captures ORDER BY id DESC',fetch=True)
            matches=[]
            for row in rows:
                jobs=self.sql('SELECT status,answer FROM jobs WHERE capture_id=? ORDER BY id',(row['id'],),True)
                combined=clean(row['cleaned']+'\n'+'\n'.join(job['answer'] for job in jobs)).casefold()
                if all(term in combined for term in terms):
                    state=jobs[-1]['status'] if jobs else ('empty OCR' if not row['cleaned'].strip() else 'not sent')
                    matches.append((row,state))
                if len(matches)>=300:
                    break
            selected=self.tree.selection()
            self.tree.delete(*self.tree.get_children())
            self.rows,self.thumbs={},[]
            for row,state in matches:
                ident=str(row['id'])
                path=self.base()/row['image']
                thumb=None
                try:
                    with Image.open(path) as image:
                        image.thumbnail((88,58))
                        thumb=ImageTk.PhotoImage(image.copy())
                        self.thumbs.append(thumb)
                except Exception:
                    pass
                options={'image':thumb} if thumb else {}
                self.tree.insert('','end',iid=ident,text='',values=(row['stamp'].replace('T',' '),state+' | '+row['cleaned'].replace('\n',' ')[:90]),**options)
                self.rows[ident]=row
            chosen=selected[0] if selected and selected[0] in self.rows else next(iter(self.rows),None)
            if chosen:
                self.tree.selection_set(chosen)
                self.select()
            else:
                self.selected_image=self.photo=None
                self.image_label.configure(image='',text='No matches')
                for widget in (self.ocr_view,self.answer_view,self.raw_view):
                    text_set(widget,'')
            if not self.error.get():
                self.status.set(f'Showing {len(matches)} newest matches (max 300); active queue {len(self.allowed)}. Search includes OCR and answers.')
        except Exception as exc:
            self.fail('Loading gallery: '+str(exc))

    def clear_search(self):
        self.query.set('')
        self.refresh()

    def select(self,event=None):
        cid=self.selected_id()
        row=self.rows.get(str(cid))
        if not row:
            return
        self.selected_image=self.base()/row['image']
        text_set(self.ocr_view,row['cleaned'])
        text_set(self.raw_view,row['raw']+ ('\n\nOCR ERROR:\n'+row['error'] if row['error'] else ''))
        try:
            jobs=self.sql('SELECT * FROM jobs WHERE capture_id=? ORDER BY id',(cid,),True)
            blocks=[]
            for job in jobs:
                attempts=self.sql('SELECT number,started,finished,outcome,error FROM attempts WHERE job_id=? ORDER BY id',(job['id'],),True)
                block=f'Request #{job["id"]} | {job["model"]} | {job["status"]} | attempts {job["attempts"]}/{job["max_attempts"]}\n'
                block+=job['answer'] or job['error'] or 'Awaiting response'
                for attempt in attempts:
                    block+=f'\n  Attempt {attempt["number"]}: {attempt["outcome"]} | {attempt["started"]}'
                    if attempt['error']:
                        block+='\n    '+attempt['error']
                blocks.append(block)
            text_set(self.answer_view,'\n\n'.join(blocks) or 'No Gemini request has been queued for this screenshot.')
            self.render_image()
        except Exception as exc:
            self.fail('Loading saved responses: '+str(exc))

    def render_image(self):
        if not self.selected_image:
            return
        try:
            with Image.open(self.selected_image) as image:
                image.thumbnail((max(100,self.image_label.winfo_width()-10),max(100,self.image_label.winfo_height()-10)))
                self.photo=ImageTk.PhotoImage(image.copy())
            self.image_label.configure(image=self.photo,text='')
        except Exception as exc:
            self.image_label.configure(image='',text='Preview unavailable: '+str(exc))

    def move(self,offset):
        ids=self.tree.get_children()
        if not ids:
            return
        selected=self.tree.selection()
        index=ids.index(selected[0]) if selected else 0
        target=ids[max(0,min(len(ids)-1,index+offset))]
        self.tree.selection_set(target)
        self.tree.see(target)
        self.select()

    def open_image(self):
        if self.selected_image:
            try:
                os.startfile(str(self.selected_image))
            except Exception as exc:
                self.fail('Opening image: '+str(exc))

    def context(self,event):
        item=self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        self.select()
        try:
            self.menu.tk_popup(event.x_root,event.y_root)
        finally:
            self.menu.grab_release()

    def speak_selected(self):
        cid=self.selected_id()
        if cid is None:
            return
        try:
            rows=self.sql("SELECT answer FROM jobs WHERE capture_id=? AND status='done' ORDER BY id DESC LIMIT 1",(cid,),True)
            if rows:
                self.speech.say(rows[0]['answer'])
        except Exception as exc:
            self.fail('Loading answer for speech: '+str(exc))

    def delete_selected(self):
        cid=self.selected_id()
        if cid is not None:
            self.delete_captures([cid])

    def delete_all(self):
        try:
            self.delete_captures([row['id'] for row in self.sql('SELECT id FROM captures',fetch=True)])
        except Exception as exc:
            self.fail(str(exc))

    def delete_captures(self,ids):
        if self.busy or self.request_busy or self.allowed:
            messagebox.showinfo('Calculator','Stop capture/pending sends and wait for active request before deleting.',parent=self.root)
            return
        if not ids:
            return
        question=(f'Delete {len(ids)} screenshot(s), OCR files, saved answers, request/attempt histories and gallery entries?\n\n'
                  'No Recycle Bin. Session logs/settings and unindexed files remain.')
        if not messagebox.askyesno('Permanently delete?',question,parent=self.root,default='no'):
            return
        try:
            self.speech.stop()
            for cid in ids:
                rows=self.sql('SELECT * FROM captures WHERE id=?',(cid,),True)
                if not rows:
                    continue
                paths=self.capture_paths(rows[0])
                paths.append(paths[-2].with_suffix('.json.tmp'))
                for path in paths:
                    path.unlink(missing_ok=True)
                con=connect(self.base())
                try:
                    con.execute('DELETE FROM attempts WHERE job_id IN (SELECT id FROM jobs WHERE capture_id=?)',(cid,))
                    con.execute('DELETE FROM jobs WHERE capture_id=?',(cid,))
                    con.execute('DELETE FROM captures WHERE id=?',(cid,))
                    con.commit()
                finally:
                    con.close()
            self.selected_image=None
            self.refresh()
        except Exception as exc:
            self.fail('Deletion partially failed; some files may already be removed: '+str(exc))
            self.refresh()

    def select_region(self,image):
        window=tk.Toplevel(self.root)
        self.selector=window
        window.title('Calculator — Draw capture area')
        scale=min(1,(window.winfo_screenwidth()-100)/image.width,(window.winfo_screenheight()-220)/image.height)
        width,height=max(1,round(image.width*scale)),max(1,round(image.height*scale))
        ttk.Label(window,text='Drag a rectangle, then Use selection.',padding=8).pack()
        canvas=tk.Canvas(window,width=width,height=height,highlightthickness=0,cursor='crosshair')
        canvas.pack()
        canvas.photo=ImageTk.PhotoImage(image.resize((width,height),Image.Resampling.LANCZOS))
        canvas.create_image(0,0,anchor='nw',image=canvas.photo)
        state={'start':None,'rect':None,'box':None}
        def close():
            self.selector=None
            window.destroy()
            self.controls()
        def start(event):
            if state['rect']:
                canvas.delete(state['rect'])
            x,y=min(width,max(0,event.x)),min(height,max(0,event.y))
            state['start'],state['box']=(x,y),None
            state['rect']=canvas.create_rectangle(x,y,x,y,outline='red',width=3)
            use.configure(state='disabled')
        def drag(event):
            if state['start'] is None:
                return
            x,y=state['start']
            xx,yy=min(width,max(0,event.x)),min(height,max(0,event.y))
            canvas.coords(state['rect'],x,y,xx,yy)
            left,right=sorted((x,xx))
            top,bottom=sorted((y,yy))
            if right-left>=3 and bottom-top>=3:
                state['box']=(math.floor(left*image.width/width),math.floor(top*image.height/height),min(image.width,math.ceil(right*image.width/width)),min(image.height,math.ceil(bottom*image.height/height)))
                use.configure(state='normal')
            else:
                state['box']=None
                use.configure(state='disabled')
        def apply():
            if state['box']:
                self.crop=state['box']
                left,top,right,bottom=self.crop
                self.region.set(f'Area: {right-left} × {bottom-top}')
                close()
        buttons=ttk.Frame(window,padding=8)
        buttons.pack(fill='x')
        use=ttk.Button(buttons,text='Use selection',command=apply,state='disabled')
        use.pack(side='left')
        ttk.Button(buttons,text='Cancel',command=close).pack(side='left',padx=8)
        canvas.bind('<ButtonPress-1>',start)
        canvas.bind('<B1-Motion>',drag)
        canvas.bind('<ButtonRelease-1>',drag)
        window.protocol('WM_DELETE_WINDOW',close)
        window.bind('<Escape>',lambda event:close())
        window.grab_set()
        self.controls()

    def capture_finished(self,result):
        if self.worker and self.worker.is_alive():
            self.root.after(50,lambda:self.capture_finished(result))
            return
        self.worker=None
        self.busy=False
        if self.closing:
            return
        self.root.deiconify()
        self.controls()
        mode,preview,error=result
        if error:
            self.fail(error)
        elif mode=='preview' and preview is not None and not self.error.get():
            try:
                self.select_region(preview)
            except Exception as exc:
                if self.selector is not None:
                    self.selector.destroy()
                    self.selector=None
                self.fail('Region selector: '+str(exc))
        self.refresh()

    def poll(self):
        try:
            refresh=False
            while True:
                try:
                    kind,value=self.events.get_nowait()
                except queue.Empty:
                    break
                if kind=='capture_saved':
                    cid,text,ocr_ok=value
                    refresh=True
                    if ocr_ok and self.auto_active and not self.stop_event.is_set() and not self.error.get() and not self.closing:
                        self.queue_job(cid,text,True)
                elif kind=='capture_done':
                    self.capture_finished(value)
                elif kind=='request_result':
                    self.handle_result(value)
                elif kind=='speech_error':
                    self.fail('Local speech: '+value)
                elif kind=='wait' and not self.error.get():
                    self.status.set(f'Waiting about {value}s to preserve screenshot spacing')
                elif kind=='models':
                    self.models_busy=False
                    self.model_box.configure(values=value)
                    if self.model.get() not in value:
                        self.model.set(value[0] if value else '')
                    self.api_status.set(f'{len(value)} candidate model(s) discovered; pricing/free quota not verified.')
                elif kind=='models_error':
                    self.models_busy=False
                    self.fail('Model discovery: '+value)
            if refresh and not self.closing:
                self.refresh()
            if not self.closing:
                self.scheduler()
        except Exception as exc:
            self.fail('Application operation: '+str(exc))
        if self.closing and (self.worker is None or not self.worker.is_alive()) and not self.request_busy:
            self.root.destroy()
            return
        self.root.after(150,self.poll)

    def close(self):
        if self.closing:
            return
        if self.request_busy:
            question=('An API call is active. Capture/pending sends and speech will stop. '
                      'The app will wait for active call to finish or time out and save its result before closing. Continue?')
            if not messagebox.askyesno('Close after saving active response?',question,parent=self.root,default='no'):
                return
        self.closing=True
        self.stop()
        self.speech.close()
        self.status.set('Closing after active capture/API work is saved...')
        if (self.worker is None or not self.worker.is_alive()) and not self.request_busy:
            self.root.destroy()


if __name__=='__main__':
    root=tk.Tk()
    App(root)
    root.mainloop()
