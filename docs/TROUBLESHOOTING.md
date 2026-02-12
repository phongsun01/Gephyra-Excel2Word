# Troubleshooting Guide

## Common Issues with Gephyra (pywin32 Edition)

### 1. "Object is not connected to server"
**Error:** `(-2147220995, 'Object is not connected to server', None, None)`
**Cause:** The Microsoft Word or Excel application instance was closed unexpectedly or disconnected.
**Solution:**
- Ensure no other script is force-closing Word/Excel while Gephyra is running.
- This version of Gephyra manages shared application instances. If you manually close the hidden Word window, the script will fail.
- Restart Gephyra.

### 2. "Cannot initialize Word COM"
**Error:** `RuntimeError: Cannot initialize Word COM.`
**Cause:**
- Microsoft Office is not installed.
- `pywin32` was installed but the post-install script wasn't run.
**Solution:**
1. Verify you can open Word normally.
2. Run the registration script:
   ```bash
   python Scripts/pywin32_postinstall.py -install
   ```
   (You may need to run Command Prompt as Administrator).

### 3. Excel Table Copy Fails
**Error:** `Error copying table Sheet1!A1:D10`
**Cause:**
- The sheet name in `config.yaml` doesn't match the Excel file.
- The range is invalid.
- Excel is showing a popup dialog (e.g., Activation Wizard, Recovery Pane).
**Solution:**
- Open Excel manually and ensure there are no popups blocking automation.
- Check `config.yaml` spelling.

### 4. Application Hangs / Zombie Processes
**Symptom:** The script finishes but `WINWORD.EXE` or `EXCEL.EXE` are still running in Task Manager.
**Solution:**
- This version creates shared instances for performance. They should close automatically.
- If they persist after a crash, manually close them via Task Manager or run:
  ```cmd
  taskkill /F /IM WINWORD.EXE
  taskkill /F /IM EXCEL.EXE
  ```

### 5. "Visible property cannot be set"
**Cause:** Sometimes COM fails to set `Visible=False` if the application is in a specific state (e.g. starting up).
**Solution:** Gephyra handles this automatically now. Warnings in the log can be ignored if the process continues.
