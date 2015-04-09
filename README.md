# Word2SettingDBs #

Word2SettingDBs.py is a script I developed/am developing at NPPD to assist with documenting relay settings. At NPPD we develop relay setting calculations in a Microsoft Word document. Settings are then entered into Aspen Relay Database as our official record of the settings that are issued to the field to be entered in the relay. We usually also create an SEL AcSELerator database for technicians to use in batch uploading the settings to the relay. It is quite tedious to take setting calculations from the Word document and then enter them one-by-one into the Aspen Relay Database and the AcSELerator database, and this process is prone to having the three setting documentation locations get out of sync if a change is made but all three documents/databases are not updated.

The script assists in this process by scanning the Word document to find relay settings and then creating files that can be imported into Aspen Database and AcSELerator. Some NPPD-specific information is hard-coded in the script, and the script depends on some NPPD-specific conventions as to how our Word templates for setting calculations are set up. Most likely the script will be useful only to NPPD since it is tailored specifically to our relay setting calculation and documentation workflow.
