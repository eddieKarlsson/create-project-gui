Remove-Item dist/create-mc-project.exe
Remove-Item create-mc-project.exe
pyinstaller create-mc-project.py --onefile
Copy-Item dist/create-mc-project.exe .
