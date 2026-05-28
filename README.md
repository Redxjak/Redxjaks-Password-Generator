# Redxjak's Password Generator

A Python/Tkinter password generator that can save generated password records to Excel, CSV, text, JSON, or Markdown.

## Community

Join the Discord server for project updates and feedback: [Redxjak Discord](https://discord.gg/9C4npSfNQd)

## Download

Download the Windows executable here:

[Download Redxjak's Password Generator](https://github.com/Redxjak/Redxjaks-Password-Generator/releases/download/v1.1.2/PasswordGenerator.exe)

## Features

- Generates random passwords with configurable minimum and maximum length
- Optional minimum counts for uppercase letters, lowercase letters, and special characters
- Lets you exclude confusing characters such as `O`, `0`, `l`, and `1`
- Generates one password or multiple passwords at the same time
- Copies generated passwords directly to the clipboard
- Saves entries with website/application, username, password, and timestamp
- Supports `.xlsx`, `.csv`, `.txt`, `.json`, and `.md` exports
- Includes a command-line mode for systems without a GUI
- Warns the user if the selected output file is open and cannot be saved

## Requirements

- Python 3
- openpyxl

Install the dependency with:

```powershell
pip install openpyxl
```

Run the app with:

```powershell
python PWGenV2.py
```

Show the version:

```powershell
python PWGenV2.py -v
```

Run from the command line with options:

```powershell
python PWGenV2.py --site Example --username redxjak --count 5 --min-length 12 --max-length 20 --min-uppercase 2 --min-symbols 2 --exclude O0l1 --output passwords.csv
```

Copy generated passwords from the command line:

```powershell
python PWGenV2.py --count 3 --copy
```

If the app is started without command-line options on a computer that cannot open a GUI, it prints a message explaining that the program should be run with a GUI and suggests `--help` for command-line usage.
