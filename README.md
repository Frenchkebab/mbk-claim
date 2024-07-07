
```
rm -r dist
python -m PyInstaller --onedir --contents-directory "." main.py
cp -r result ./dist/main
cp -r settings ./dist/main
cp -r upload ./dist/main
cp -r dist/main /c/CLAIMX/newProgram
explorer C:\CLAIMX
```