1. pyinstaller --windowed --collect-submodules=app --hidden-import pydantic --hidden-import tkcalendar --hidden-import babel --hidden-import babel.numbers --hidden-import fastexcel main.py
2. Mover las carpetas build y dist a otra parte
3. Copiar y pegar el main.py a la carpeta anterior
4. Copiar y pegar la carpeta app en dist/main/_internal/
5. comprimir los archivos en un .zip
6. extarer los arhcivo y probar que funcione haciendo clic en el main.exe de la ruta dist/main/

