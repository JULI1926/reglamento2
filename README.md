pyinstaller --noconfirm --onefile --windowed --icon "I:\reglament.ico" --add-data "C:\Users\julia\Desktop\reglamento-MOD\template.docx;." "C:\Users\julia\Desktop\reglamento-MOD\main.py"


Con este comando se genera el ejecutable de la aplicación. El archivo se encuentra en la carpeta dist. Para ejecutar la aplicación se debe hacer doble clic en el archivo main.exe, se debe 
tener en cuenta las rutas (ejemplo --add-data "C:\Users\julia\Desktop\reglamento-MOD\template.docx)cambian dependiendo de la ubicación de los archivos en el computador.