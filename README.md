# Two Degrees dashboard

A Python Streamlit proof of concept for tracking a milk batch from milking to delivery. It uses three simulated batches, six stages and 355 readings. No physical sensor or separate backend is required.

## Start on Windows

```powershell
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

## Start on macOS or Linux

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Open the localhost address printed in the terminal, normally `http://localhost:8501`.

## Main flow

1. Choose a batch.
2. Press **Track batch**, use **Next update**, or move the replay slider.
3. Select a journey stage to inspect its temperature graph and reading log.
4. Open **Incidents** to see automatically generated alerts.
5. Acknowledge an alert, then complete its incident report.
6. Open **Records** to export a readable Excel workbook or technical JSON file.

The alert field is binary: `0` means no active alert and `1` means the system currently detects a problem. Ideal temperatures are green, warnings are orange, and critical values are red. The supplied temperature-excursion scenario stays in the orange warning range.

## Demo data

Open **Demo data** in the top navigation to upload another journey JSON file, restore the bundled sample, download the sample or change playback speed.

| Batch | Volume | Scenario |
| --- | ---: | --- |
| MILK-001 | 500 L | Normal journey |
| MILK-002 | 800 L | Transport cooling interruption and recovery |
| MILK-003 | 600 L | Transport sensor connection lost and restored |

## Files

- `app.py`: dashboard, replay, alert workflow and exports.
- `sample_data.json`: three synthetic batch journeys.
- `logo-mark-transparent.png`: transparent logo used in the header.
- `requirements.txt`: pinned Python dependencies.
- `.streamlit/config.toml`: light interface theme.

## Scope

All data, people and organisations are fictional. Thresholds are demonstration assumptions, not product-safety certification. The prototype has no login, persistent database, live sensor connection or tamper-proof audit store.
