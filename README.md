# pdf2word

Herramienta en Python para convertir documentos PDF a Word y extraer sus gráficas
como imágenes PNG. El resultado se guarda junto al PDF original con el mismo
nombre base.

## Uso básico

```bash
python pdf2word.py archivo.pdf
```

Opcionalmente se pueden indicar la fuente, tamaño, interlineado y márgenes:

```bash
python pdf2word.py archivo.pdf --font Arial --size 12 --spacing 1.5 --margin 1
```

Solo se conservan las imágenes que parecen gráficas. Si se quiere exportar todas
las imágenes del PDF puede usarse `--include-all-images`.

El script genera además un fichero `*_process.log` con el detalle de la
operación y las imágenes extraídas se guardan como `*_p<pagina>_chart<idx>.png`.

## Ejecutable para usuarios no técnicos

Es posible crear un archivo `pdf2word.exe` para Windows usando
[PyInstaller](https://www.pyinstaller.org/). Tras instalar PyInstaller::

    pip install pyinstaller

Se genera el ejecutable con::

    pyinstaller --onefile pdf2word.py

El fichero resultante se encuentra en `dist/pdf2word.exe`. Basta arrastrar un
PDF encima de él para convertirlo. Al finalizar se mostrará un mensaje
informativo.
