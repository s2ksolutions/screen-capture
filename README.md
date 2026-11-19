# Calculator — Screenshot OCR and Gemini Gallery

A Windows desktop application that captures a selected display region, reads its text locally, organizes screenshots in a searchable gallery, and optionally sends OCR text to Gemini for answers. Includes local Windows speech and an optional calculator-style interface.

> Status: generated application code. Syntax and selected helper/storage tests have been checked. The complete Windows GUI, model downloads, Gemini integration, speech, and packaged executable have not been end-to-end verified. Test before relying on it.

## Features

- Start/stop screenshot capture and choose a library folder.
- Manual capture interval slider.
- Automatic capture interval: **10–3,600 seconds**.
- Screenshot preview with a drag-to-select OCR region.
- Configurable warm-up: **0–600 seconds**, followed by OCR/capture initialization.
- Optional working calculator view while capture runs, with Stop and Show full interface buttons.
- Local CPU OCR using RapidOCR and ONNX Runtime.
- Unicode cleanup, optional English/ASCII cleanup, and preserved original OCR text.
- Thumbnail gallery, previous/next browsing, and search across OCR text and saved Gemini answers.
- Gemini responses and request history displayed below each screenshot's OCR text.
- Saved request prompts, models, answers, attempt history, and errors.
- Right-click resend, speak latest saved answer, or permanently delete a capture and associated files.
- Delete all indexed screenshots with confirmation.
- Default prompt and model selection saved between launches.
- API-key entry with local saving through the launcher.
- Tracked request queue, one active API request at a time, maximum 25 pending jobs.
- Configurable total attempts: **1–5**. Selected transient failures use bounded retries with backoff.
- Configurable cooldown after an API attempt finishes: **0–3,600 seconds**.
- Local Windows SAPI speech: adjustable speed, **1–20** repetitions, or loop until the next answer.
- Red operational error message above Start capture; terminal errors stop capture and pause queued sending.
- Known missing-key errors clear when a nonempty key is entered; this does not validate the key or automatically resume queued uploads.

## Files

Keep these files in the same project folder:

```text
project/
├── capture_gallery.py
├── calculator_launcher.py
├── requirements.txt
├── README.md
└── instructions.txt          # Optional short usage guide
```

**Run/build `calculator_launcher.py`.** It imports the base app and adds plaintext key storage, missing-key handling, calculator view, and warm-up. Running `capture_gallery.py` directly does not enable those overrides.

## Prerequisites

- Windows with a normal interactive desktop session.
- Python 3.10 or later, with Tkinter, in an architecture supported by all dependencies. A 64-bit Python installation is the intended starting point; the entire supported-version matrix has not been tested.
- An available Gemini API key/project and a model that accepts requests for that project.
- Internet access for Gemini and potentially initial OCR model downloads.
- A functioning Windows SAPI voice for speech.

Check Tkinter if needed:

```powershell
py -m tkinter
```

A GPU is not required for the intended CPU OCR setup. Capture restrictions imposed by another application are not bypassed by this app.

## Install

Open PowerShell in the project folder. These commands do not require activating the virtual environment:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

The requirements are intentionally unpinned because no complete dependency/version combination has been end-to-end verified. After verifying your environment, record its exact versions if you need reproducible builds:

```powershell
.\.venv\Scripts\python.exe -m pip freeze > requirements-lock.txt
```

Do not publish a lock file as tested unless you have actually verified that environment.

## Run from source

```powershell
.\.venv\Scripts\python.exe calculator_launcher.py
```

## Build the Windows executable

Close any running copy of the previous executable. From the folder containing both scripts, run:

```powershell
.\.venv\Scripts\python.exe -m PyInstaller --clean --onefile --noconsole --hidden-import=cv2 --hidden-import=google.genai --hidden-import=google.genai.types --collect-all rapidocr --collect-all onnxruntime --collect-all google.genai --copy-metadata google-genai --hidden-import=win32com.client --hidden-import=pythoncom --hidden-import=pywintypes --name Calculator calculator_launcher.py
```

Expected build output:

```text
dist\Calculator.exe
```

The executable bundles its Python dependencies, but it is not guaranteed to bundle every first-run OCR asset. Test model initialization before assuming offline readiness.

### Verify the Gemini SDK before building

```powershell
.\.venv\Scripts\python.exe -c "import google.genai as genai; from google.genai import types; print(genai.__file__)"
```

If this fails, fix the source environment first. Check for a project file named `google.py` or a conflicting local `google` folder.

### Diagnose an executable that fails silently

For a diagnostic build, replace `--noconsole` with `--console`. Use a different build name if you want to keep the release build. Launch from a terminal while the target application is closed, and retain the exact traceback.

### Inspect bundle size

For easier inspection, replace `--onefile` with `--onedir`, keeping the other options unchanged. Distribute the entire resulting `dist\Calculator` folder, not just its EXE.

List the largest files:

```powershell
Get-ChildItem .\dist\Calculator -Recurse -File |
    Sort-Object Length -Descending |
    Select-Object -First 20 FullName,
        @{Name="MB";Expression={[math]::Round($_.Length / 1MB, 2)}}
```

OCR runtimes, OpenCV, NumPy, and model assets can make the bundle large. Broad collection flags can also include optional modules. Remove collection flags only after testing the affected features. An `onedir` build does not necessarily reduce total distribution size.

## Quick start

1. Launch the executable built from `calculator_launcher.py`.
2. Choose a library folder.
3. Click **Preview / draw area**, expose the desired content within two seconds, drag a rectangle, and click **Use selection**.
4. Open **Gemini / speech**. Enter the API key and click **Save key** if you want persistence.
5. Choose a model and verify its availability, pricing, and quota in your Google project. Refresh is model discovery, not a free-tier check.
6. Edit the default prompt and click **Save prompt + model**.
7. Set speech options and click the speech **Apply + save** button.
8. Return to **Gallery / capture**. Set automatic interval, cooldown, and total attempts; click **Apply + save**.
9. Enable **Automatic answer** if you want captured OCR text uploaded automatically. Review the confirmation.
10. Optionally enable **Calculator view when Start capture is clicked**, set warm-up, and save those settings.
11. Click **Start capture**. Capture begins after warm-up, initialization, and any remaining minimum-spacing wait.
12. Browse screenshots. The saved Gemini history appears directly below OCR text. Search checks both.
13. Right-click a screenshot to resend its OCR or speak its latest saved answer.
14. Use **Stop capture / pending sends** or the calculator view's **Stop** button to stop capture and pending sends.

For a manual question, use the Gemini tab's question box and **Queue send to Gemini**. If a screenshot is selected, that manual request is linked to it. Automatic requests use the newly captured OCR, not whichever gallery item is selected.

## Timing, queue, and retry behavior

- The first capture in a fresh session follows initialization; the interval is not an extra mandatory startup delay.
- In automatic mode, saved screenshots are spaced by at least the configured interval, never below 10 seconds. Restarting capture does not reset the last screenshot timestamp within the running app.
- All queued request starts, including retries/resends, share a minimum 10-second spacing. Automatic mode uses its configured interval if higher.
- Cooldown is measured after the API attempt finishes, **not after speech**. An infinite speech loop must not prevent the next answer.
- OCR/API processing can make actual intervals longer.
- Each nonempty automatically captured OCR result can create a job, even when its text matches a previous screenshot. Do not assume duplicate-screen suppression.
- The queue is capped at 25 allowed jobs. A full queue pauses capture with an error.
- Request state and attempt history are stored locally. There is no exactly-once delivery guarantee.
- Retries are intended for selected transient failures and empty text responses, not invalid keys, permissions, or missing models. A timeout may occur after server processing, so retrying can duplicate usage or charges.
- A retry does not necessarily run before every other eligible queued job.
- Stop pauses queued work. A server request already running may finish and be saved.
- After restart, previous jobs are not automatically authorized/resumed. Use explicit resend for a screenshot.
- A terminal operational error pauses capture/sending but does not silently switch automatic capture to a faster manual rate.

## Speech

- Apply/save speech settings before relying on them.
- Select either a fixed repeat count or looping until the next successful answer.
- A new accepted answer replaces prior speech.
- Stop speech silences the current answer without stopping capture.
- The slider uses Windows SAPI's native rate range, -10 to +10; it is not a words-per-minute guarantee.
- Rate changes take effect when a new speech stream starts.

## Saved data

Default layout inside the selected library folder:

```text
capture_library/
├── gallery.sqlite3
└── sessions/
    └── YYYY-MM-DD/
        └── HHMMSS_microseconds/
            ├── capture.log
            ├── settings.json
            ├── images/
            │   └── timestamp.png
            ├── text/
            │   └── timestamp.txt
            ├── raw/
            │   ├── timestamp.txt
            │   └── timestamp.error.txt     # Only when applicable
            └── answers/
                ├── timestamp.json
                └── timestamp.txt
```

Answer sidecars contain request/response history, including prompt snapshots, model IDs, and errors. The SQLite index is also required for gallery operation. Keep the entire library together when moving/backing it up.

Manual requests without an associated screenshot are saved in SQLite but have no screenshot-linked answer sidecar or gallery item. Older unindexed screenshots are not imported automatically.

Application preferences and the launcher's plaintext key are stored in:

```text
%LOCALAPPDATA%\CalculatorScreenshotOCR\settings.json
```

## API-key and privacy warning

- The launcher intentionally stores the key in **plaintext** because that behavior was requested. Anyone who can read the file or its backups can copy the key.
- Do not commit settings, keys, screenshots, OCR, or answer histories to source control. Do not embed a real key in the script or executable.
- The base script's default storage uses Windows credentials; the launcher overrides this. Build the launcher for the selected plaintext behavior.
- Deleting a locally saved key does not revoke it at Google or erase copies/backups. Revoke exposed keys through Google.
- Capture and OCR run locally. Manual/automatic Gemini requests transmit default prompt plus question/OCR text to Google. This code does not upload screenshot images.
- `store=False` is used on Interactions requests; this is not a blanket guarantee about all provider logging, retention, or data-use policies. Review Google's current terms for your project.
- Model candidates are a maintained list, not an exhaustive list of free models. Pricing and eligibility can change. The application does not enforce free-only billing.
- Use capture/upload only for content you are authorized to process.
- Calculator view changes UI presentation only. It does not hide the process, defeat capture restrictions, or guarantee what monitoring software sees.

## Deleting screenshots

Stop capture/pending sends and wait for any active request. Right-click a capture or click **Delete all** and confirm.

Deletion permanently removes indexed screenshot images, associated OCR/answer files, and their database request/attempt history. It does not use the Recycle Bin. Delete all includes entries hidden by search or the 300-result display limit. Session logs/settings, unindexed files, backups, and provider-side data are not removed.

Deletion is not transactional across files and SQLite. A partial error may leave some files already removed; inspect the error and retry carefully.

## Troubleshooting

| Symptom | Next check |
|---|---|
| No calculator-view/warm-up controls | Run/build `calculator_launcher.py`, and use the latest launcher. |
| Missing-key error remains | Use the latest launcher; only recognized missing-key errors clear on entry. Invalid-key/permission/quota errors stay visible. |
| `cannot import ... genai` | Upgrade `google-genai` in the build environment; run the SDK import test above; rebuild cleanly. |
| SDK missing `interactions` or `HttpRetryOptions` | Upgrade `google-genai`; retain the exact error if it persists. No full version matrix has been verified. |
| `cv2` missing | Install `opencv-python` in the build environment and keep the explicit hidden import. |
| Model `404` | Verify model access for your project. The model candidate list may require maintenance; do not retry a permanent model error indefinitely. |
| `429` or temporary server failure | Review quota and attempt status. Increase interval/cooldown and avoid repeated manual resends. |
| No new capture frame | Inspect target/display state and session log. A protected, frozen, or unavailable capture path is not solved by renaming the EXE. |
| OCR starts slowly | First-run model initialization/download may be required. Test on normal desktop text first. |
| Red error message | Capture and queued sending are paused. Fix the stated stage; explicitly Start or resend. |
| Answer exists but no speech | Check Read received answers and Apply + save; test Speak latest saved answer; inspect any SAPI error. |
| Large EXE | Inspect an `onedir` build; OCR/native dependencies dominate many bundles. |

## Verification checklist

- [ ] Both scripts are present and the launcher starts.
- [ ] Missing-key errors clear on key entry without automatically resuming uploads.
- [ ] Save/reload key and default prompt work.
- [ ] Region drawing maps to the intended screenshot area.
- [ ] Warm-up can be cancelled before capture starts.
- [ ] Calculator view, full-interface return, and Stop work.
- [ ] OCR captures normal desktop text correctly.
- [ ] One manual Gemini response is received, saved, and visible below its screenshot.
- [ ] Resend creates another history entry without overwriting an older answer.
- [ ] Automatic screenshots respect the configured minimum interval.
- [ ] Retries are visible and bounded; permanent errors pause capture.
- [ ] Speech speed, fixed repeats, loop replacement, and Stop speech work.
- [ ] Search finds a distinctive word in a saved answer.
- [ ] Deleting a disposable capture removes its associated files/index entries.
- [ ] Restart preserves library/history but does not automatically restart sending.

Do not distribute a build as verified until you have completed the relevant checks on the intended Windows environment.
