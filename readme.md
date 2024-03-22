# Repositorio recopilatorio de scripts VBA para Excel

Recopilatorio de scripts VBA

## Dividir Excel en Filas
Este script lo que hace es dividir un excel de varios miles de filas en el número de filas que desees.
Por ejemplo si tienes un excel con sus encabezados para importar a una base de datos que tiene 12.000 registros y debes dividirlo para poder cargarlo a dicha base de datos, los pasos son:

1. Abre el archivo excel a dividir
2. Si están las macros deshabilitados, habilitalos (te sale un mensaje)
3. Presiona ALT + F11 para abrir la consola de VBA
4. En la venta de proyecto donde aparece el nombre de las hojas del archivo excel, le damos a boton derecho y elegimos "Ver Código"
5. Se nos abrirá una ventana donde deberemos pegar el script indicado y una vez pegado le damos al icono verde en forma de triangulo (Ejecutar)
6. Nos pedirá el nº de filas que contendrá cada fichero y uan vez indicado creará dichos ficheros en el directorio donde se encuentra el fichero principal.

Control de Errores
El dato indicado debe cumplir las siguientes reglas:
- Mayor a 0 celdas
- Menor al nº de máximo de celdas
- Que no se indique ningún dato
- Que no sea numérico
