Here’s a `README.md` file for your GitHub project:

---

# Word File Generation Microservice

This microservice processes an Excel file, maps data according to specific rules, and generates a Word document. It is designed to be deployed as a Windows service and supports specific configurations to run smoothly in various environments.

## Installation

### 1. Install the Microservice as a Windows Service

To package the Python script (`Wordsplit.py`) and install it as a Windows service, use **pyinstaller**:

```bash
pyinstaller -F --icon=word.ico --hidden-import=win32timezone Wordsplit.py
```

### 2. Common Error: Service Not Starting

If you encounter the following error:

```
The service did not respond to the start or control request in a timely fashion.
```

**Solution**: 

Copy the file `pywintypes36.dll` from the following directory:

```
Python36\Lib\site-packages\pywin32_system32
```

And paste it into this directory:

```
Python36\Lib\site-packages\win32
```

### 3. Install Required Dependencies

To work with Excel files, you'll need `xlrd`. Install it using these commands:

```bash
python3 -m pip install --user xlrd
python3 -m pip install xlrd
```

### 4. Configuring Microsoft Word for DCOM

Ensure that Word is configured properly by following these steps:

1. Open **Console Root** → **Component Services** → **Computers** → **My Computer** → **DCOM Config** → **Microsoft Word Document**.
2. Right-click on **Microsoft Word Document** and select **Properties**.
3. Go to the **Identity** tab and select **Interactive user** instead of **Launching user**. This ensures MS Word will run with the rights of the currently logged-on user.

### 5. 32-bit Word Configuration

If you are using a 32-bit version of Microsoft Word, you can open the DCOM configuration console using this command:

```bash
C:\WINDOWS\SysWOW64\mmc comexp.msc /32
```

### 6. Install Required Python Packages

Install the necessary Python libraries for the project:

```bash
pip install python-docx xlrd watchdog servicemanager pypiwin32
```

## Current Remarks

- **Cancel Executor**: Remove or comment out the executor if no longer needed.
- **Cancel Service**: Disable or remove the service registration if the service is not required.

---

This file should provide clear instructions for installing and configuring the service on Windows.
