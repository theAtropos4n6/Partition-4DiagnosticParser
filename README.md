# Partition%4Diagnostic.evtx Log Parser

**Partition%4DiagnosticParser** is a Python tool that parses the Windows 10 **Microsoft-Windows-Partition%4Diagnostic.evtx** log file (Path:C:\Windows\System32\winevt\Logs) and reports information about all the connected devices and their Volume Serial Numbers, both currently present on the device and previously existed. It accomplices this task by doing the following:

- parsing all the records in the log,

- analyzing the Vbr elements of each record (so as to translate them into meaningful Volume Serial Numbers) and

- attribute interpreted VSNs to the specific connected device they belong to.

Taking into consideration the fact that malicious actors would often try to cover their illicit activities and files, by performing either a format or wipe action to their device, is what gives this information an added value. Utilizing this info, an investigator can attribute LNK and Jump List files,  to a specific removable device, even after the original files were deleted and the device formatted. 

## Installation

This is a tool written in Python (version 3.8.5 used). The .exe file (**Partition%4DiagnosticParser.exe**) works on Microsoft Windows based machines by just double clicking.

The source code file (**Partition%4DiagnosticParser.py**) can be run on a system with python 3 installed. It only needs two additional libraries to run successfully. Use the package manager [pip](https://pip.pypa.io/en/stable/) to install them.

```bash
pip install evtx
```
for the [evtx](https://pypi.org/project/evtx/) module from Omer Ben-Amram at https://pypi.org/project/evtx/ 

```bash
pip install PySimpleGUI
```
for the [PySimpleGUI](https://pypi.org/project/PySimpleGUI/) module from  MikeTheWatchGuy at https://pypi.org/project/PySimpleGUI/

## Usage

The tool comes with a GUI interface. User has **to provide the tool with a Partition%4Diagnostic.evtx log file** (when in live systems, bear in mind that this Event Log is locked from the OS and needs to be copied elsewere first) and then has the following options:

**1.** Run a **full report action** to get a report (HTML or CSV) showing all of the connected devices, their manufacturer, model, First Connected Timestamp in UTC, Last Connected Timestamp in UTC and finally **every Volume Serial Number that ever existed on these devices** historically throughout the whole log.

![GitHub Logo](/CaptureFULL.PNG)

**2.** Run a **targeted analysis action** for a specific removable device by providing the **device's S/N** (embedded by the manufacturer and usually written on the device's outer case. See the command Cheatsheet included within the first few lines of the .py file, for obtaining a device's S/N via CMD). With this option the user will get a report (HTML or CSV) providing an actual **timeline** of all the times the device was connected to the OS, along with:

- the Volume Serial Number of each volume that was present at that point on the device and

- a flag indicating when a VSN change was detected, indicating a possible format action on the device (which appears every time a device connected to the system had a different Volume Serial Number than the previous time).

![GitHub Logo](/CaptureSN.PNG)

In any case, the tool provides a verbose output pane providing extra info for the analysis.

## Caveats

- Given the fact that the **Partition%4Diagnostic.evtx** log file is a Windows file, it only provides information about Volumes formatted with a Microsoft File System (e.g. NTFS, ExFAT, FAT32). If a removable device contains Volumes with other File Systems (non-Windows FS), the log file can provide no information about the Volume Serial Numbers of its Volumes. The tool however will parse the log and provide the remaining information described in the **Usage** section.

- Due to the Event Log's structure, extraction of VSNs is available only for Unpartitioned (usually USB thumb-drives) or MBR partitioned devices. However, this tool will parse the log and provide the remaining information described in the **Usage** section, even for devices with GPT partition scheme.


- An **"Unknown Volume Type"** result in the report (**targeted analysis action**) indicates that the Volume could either have been formatted with a non-Windows File System, or it is the result of a format action by the user. The latter action, would result in an Event Log entry with 512 zeroes in the Vbr element of the device (in such cases the log can not offer any data that could be parsed and translated into a Volume Serial Number). The tool will however parse the log and provide the remaining information described in the **Usage** section.

- An empty result in the **"Volume Serial Number"** section within the report (**full report action**), means that the device had only Volumes with non-Microsoft File Systems, so no Volume Serial Number is stored and therefore extracted from it.


## License
[MIT](https://github.com/theAtropos4n6/Partition-4DiagnosticParser/blob/main/LICENSE)
