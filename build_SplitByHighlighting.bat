REM Build using Pyinstaller (http://www.pyinstaller.org/) to create a standalone executable.
pyinstaller.exe --distpath="T:\T&DElectronicFiling\ProtCntrl\ProtectionMaster\Software\Word2SettingDBs" --noconfirm --onefile SplitByHighlighting.py
REM Also save source code to file server.
copy SplitByHighlighting.py "T:\T&DElectronicFiling\ProtCntrl\ProtectionMaster\Software\Word2SettingDBs"
