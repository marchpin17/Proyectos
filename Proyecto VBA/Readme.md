# [Proyecto Excel-VBA]: Consolidador Automático de Reportes Operativos
## 📋 Descripción
Este proyecto surge de la necesidad de optimizar el procesamiento de datos en entornos operativos, donde la información suele llegar fragmentada en múltiples archivos CSV (como reportes de producción o inventarios de diferentes aplicativos).

La herramienta automatiza la extracción, limpieza y consolidación de estos datos en un archivo maestro de Excel, gestionando además el flujo de archivos procesados para mantener el orden del repositorio.

## 🛠️ Tecnologías y Habilidades
**Lenguaje:**  VBA (Excel)

**Conceptos aplicados:**

* Manipulación del Sistema de Archivos (FileSystemObject).
* Automatización de tareas repetitivas (ETL básico).
* Manejo de errores y validación de rutas.
* Optimización de procesos operativos.

## 🚀 Situación y Reto (Caso de Negocio)
**Problema:** El equipo de operaciones recibía diariamente múltiples archivos CSV de diferentes fuentes. La consolidación manual tomaba aproximadamente [X] minutos/horas al día, con un alto riesgo de error humano al copiar y pegar.

**Solución:** Desarrollé una macro robusta que:

* Escanea una ruta específica en búsqueda de nuevos archivos.
* Importa y formatea el contenido automáticamente.
* Mueve los archivos procesados a una carpeta histórica (/Cargado) para evitar duplicidad de datos en la siguiente ejecución.

## 📈 Impacto Operativo
**Ahorro de Tiempo:** Reducción del tiempo de consolidación en un [X]%.

**Integridad de Datos:** Eliminación de errores por copiado manual.

**Escalabilidad:** Capacidad de procesar cientos de archivos en segundos.

## 📝 Instrucciones de Uso
1. Descargar el archivo .xlsm
2. Abrir el archivo y clickear el boton "Introducir ruta"
3. Copiar la ruta donde se encuentran los archivos (Debe existir una carpeta llamada *Cargados*)
4. Macro ejecutada
