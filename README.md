
```
rm -r dist
python -m PyInstaller --onedir --contents-directory "." main.py
cp -r result settings upload dist/main
cp -r dist/main /c/CLAIMX/newProgram
explorer C:\CLAIMX
```