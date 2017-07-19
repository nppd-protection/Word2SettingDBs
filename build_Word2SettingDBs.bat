REM Build using Pyinstaller (http://www.pyinstaller.org/) to create a standalone executable.
pyinstaller.exe --distpath="T:\T&DElectronicFiling\ProtCntrl\ProtectionMaster\Software\Word2SettingDBs" --noconfirm --onefile Word2SettingDBs.py
REM Also save source code to file server.
copy Word2SettingDBs.py "T:\T&DElectronicFiling\ProtCntrl\ProtectionMaster\Software\Word2SettingDBs"
