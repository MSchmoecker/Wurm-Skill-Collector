# Wurm Online Skill Collector
## Overview
This is a simple python program which collects the skills from a Wurm Online Account and writes those to a Google Docs table. 
The table must be adjusted to work with the tool.

## Usage
### Program setup
The program is a single-file-executable and therefore nothing must be installed. The `WurmSkillCollector.exe` has to be downloaded
and put in one of the following folders:
- If the Wurm Online Steam version is used: `C:\Program Files\Steam\steamapps\common\Wurm Online`
- If the old Wurm client is used: `C:\Users\USER\Wurm` or AppData if it doesn’t exists

A `SkillCollectorConfig.json` must be created and all values must be passed. An example file is in the repository (`SkillCollectorConfigSample.json`).
It is recommended to link the tool to a desktop shortcut to the .exe file.

It is possible to put those both files (.exe and .json) directly in the Wurm folder or in a subfolder.

### Wurm Online setup
In order to collect the skills from a file, the option `Save Skills On Exit` has to be activated in Wurm. It only saves the skills when Wurm closes.

![](https://github.com/MSchmoecker/Wurm-Skill-Collector/blob/master/Docs/WurmSaveSettings.png?raw=true)


### Google table setup
All skills have to be in the column `$C:$D` (only one skill per row, you can use merged cells for parent skills).
All names have to be in the row `5`. Example:

<img src="https://github.com/MSchmoecker/Wurm-Skill-Collector/blob/master/Docs/GoogleTable.png?raw=true" width="50%" />

### Execution
Before the tool runs, the player name has to be written **exactly** like the ingame name to the Google table.

![](https://github.com/MSchmoecker/Wurm-Skill-Collector/blob/master/Docs/ProgramSample.png?raw=true)

The program can now be executed. It tries to find the newest log and write the the skill values to the table.
The tool only runs once and has to be started manually every time new skills are collected in Wurm.

### Multiple Characters
If you want to upload skills of more than one char, you have to tell the program to explicit find them. To do this
you have to edit the `SkillCollectorConfig.json` file. Note that the trailing comma is not allowed if it is the last line with data.

```
  "players": ["PlayerNameA", "PlayerNameB"],
```

## Development
### Setup
1. The repository must be cloned into the Wurm Online folder as described above
2. A `SkillCollectorConfig.json` must be created and all values must be passed. An example file is in the repository (`SkillCollectorConfigSample.json`).
	- the `sheet_id` can be copied from the url line.
	- the `worksheet_name` is the worksheet where all skills are stored.
	- under https://console.developers.google.com should a new project be created.
	- under https://console.developers.google.com/apis/credentials should a new service-worker be created. A new key should be 
	generated and the json must be downloaded. The data of the service-worker must be passed into the `SkillCollectorConfig.json` file.
    **The document must be shared to the e-mail of the worker.** A public link is not sufficient.

The `main.py` may now be executed.

### Building
To build the project, pyinstaller is used (https://pyinstaller.readthedocs.io).
The command is:
```
pyinstaller WurmSkillCollector.spec -n WurmSkillCollector --onefile
```

## Troubleshooting
### Google
The program uses a service-worker which only grands 100 read requests in 100 seconds. It uses 2 read requests for one execution, 
so only 50 Players can write to the table in 100 seconds. The daily limit, however, is not limited.

### Other errors
Please use the Github bug tracker if an error occurs.
