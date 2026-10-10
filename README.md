# Legacy Windows Automation Archive

A collection of early systems administration, file-system auditing, and OS automation utilities written in VBScript and Batch. 

This repository serves as a historical archive of early script development—focusing on raw COM API interactions, native Windows scripting interfaces, and portable local automation before moving on to modern compiled systems programming (C++/C#).

---

## Included Utilities

### 1. `ScanAll` (File System Auditor)
* **Files:** `ScanAll.vbs`, `ScanAll.bat`
* **Description:** A self-contained, dependency-free VBScript utility designed to recursively traverse local drives, log file paths and creation dates, and output a clean text report directly to the user's desktop.
* **Features:**
  * Auto-incrementing filename collision handling (`Results1.txt`, `Results2.txt`, etc.).
  * Graceful error trapping (`On Error Resume Next`) to bypass permission-locked directories without crashing.
  * Accompanying batch script wrapper utilizing `%~dp0` path resolution and automated UAC elevation (`-Verb RunAs`) for administrative audits.
* *Note: A compiled standalone `.exe` iteration of this utility is maintained in a separate repository.*

### 2. `update.vbs` (WUA Orchestration Script)
* **Files:** `update.vbs`
* **Description:** A legacy administrative utility leveraging the native Windows Update Agent (WUA) COM interface (`Microsoft.Update.Session`) to search, filter, prompt for EULA acceptance, download, and install applicable software updates directly from a script session.
* **Context:** Built as an exercise in low-level Windows COM object navigation and manual state management without relying on modern PowerShell module abstractions.

---

## Usage / Disclaimer
* **Environment:** Designed for legacy and modern Windows environments supporting `wscript.exe` and native scripting runtimes.
* **Purpose:** Maintained strictly for archival and reference purposes.
