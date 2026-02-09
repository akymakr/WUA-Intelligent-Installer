# WUA-Intelligent-Installer

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://microsoft.com/powershell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/)

**WUA-Intelligent-Installer** is a robust, enterprise-grade PowerShell script designed for autonomous Windows Update management. Unlike standard update commands, this tool is built for **DevOps and SRE workflows**, featuring multi-pass scanning, asynchronous progress tracking, and intelligent reboot state detection.

---

## 🌟 Key Features

* **Multi-Pass Execution:** Automatically handles complex update chains (e.g., Servicing Stack Updates followed by Cumulative Updates) by performing multiple scan/install cycles in a single run.
* **Asynchronous Progress:** Provides real-time feedback during download and installation phases, ensuring visibility even during long-running updates.
* **Intelligent Reboot Logic:**
  * Detects actual reboot requirements from both the Installer and System Registry.
  * Prevents reboot loops via a **JSON Restart Marker** system.
  * Schedules graceful shutdowns with customizable delays.
* **DevOps Ready:** Returns standard exit codes (e.g., `3010` for success with reboot) for seamless integration with CI/CD pipelines and orchestration tools.
* **Advanced Filtering:** Support for `Delayed` mode to defer updates by a set number of days, mitigating the risk of "Day 0" patch issues.

---

## 🚀 Detailed Usage Guide

### Prerequisites
* **OS:** Windows Server 2012 R2+ / Windows 10+.
  * **Tested on:** Windows Server 2019, Windows 10 24H2, Windows 11 25H2, including my own PC.
* **Privileges:** Must be executed as **Administrator**.
* **Execution Policy:** Ensure script execution is allowed (`Set-ExecutionPolicy RemoteSigned`).
* **Folder:** Ensure `C:\Windows\Logs\Updates` is created and writable.
  * By default, Windows Server should have this folder created but not Windows 10 or Windows 11.

### Basic Examples

**1. Default Production Run**
Install updates older than 28 days and schedule a reboot in 2 hours:
```powershell
.\Install-WindowsUpates.ps1 -UpdateMode Delayed -DeferDays 28 -RestartDelayMinutes 120
```

**2. Security-Only Audit Only install security/critical updates and do not reboot:**
```powershell
.\Install-WindowsUpates.ps1 -UpdateMode SecurityOnly -NoRestart
```

3. Preview Mode (Dry Run) Scan for available updates based on filters without installing anything:
```powershell
.\Install-WindowsUpates.ps1 -ScanOnly
```
