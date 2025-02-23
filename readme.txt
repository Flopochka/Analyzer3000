pyinstaller --onefile --add-data ".\;." main.py
Remove-Item -Recurse -Force .\build, .\dist
.\venv\Scripts\Activate