# 📝 Gestor de Cuestionarios a Moodle XML
![Python](https://img.shields.io/badge/Python-3.8+-blue.svg?logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-lightgrey.svg)
![Moodle](https://img.shields.io/badge/Target-Moodle_XML-orange.svg?logo=moodle&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Status](https://img.shields.io/badge/Status-Stable_Release-success.svg)

Una herramienta de escritorio inteligente (ETL) diseñada para automatizar la transcripción, auditoría y conversión de bancos de preguntas desde documentos de Microsoft Word (`.docx`) hacia el formato estandarizado Moodle XML.

Este proyecto elimina horas de trabajo manual para el cuerpo docente, asegurando una **integridad de datos del 100%** mediante un motor de validación estricto basado en Expresiones Regulares (RegEx).

---

## 📑 Tabla de Contenidos
- [✨ Características Principales](#-características-principales)
- [🛠️ Arquitectura y Tecnologías](#️-arquitectura-y-tecnologías)
- [📋 Reglas de Formato (Documentos Word)](#-reglas-de-formato-documentos-word)
- [🚀 Instalación y Uso](#-instalación-y-uso)
- [🛡️ Lógica de Validación (Fail Loudly)](#️-lógica-de-validación-fail-loudly)
- [👨‍💻 Autor](#-autor)

---

## ✨ Características Principales
* **📂 Procesamiento en Lote (Batch):** Carga y analiza múltiples archivos `.docx` simultáneamente.
* **🧠 Parser Inteligente:** Identifica automáticamente preguntas de Opción Múltiple y Verdadero/Falso descartando textos residuales (Garbage Collection).
* **👁️ Auditoría Maestro-Detalle:** Previsualización estructurada con interfaz al estilo Moodle `( • )` para verificar respuestas correctas antes de exportar.
* **🛡️ Validación Estricta:** Las preguntas sin respuesta detectada son marcadas como `⚠️ Rechazadas` y excluidas del XML final para evitar corrupción en la base de datos de Moodle.
* **📝 Corrección en un Clic:** Botón integrado que abre automáticamente el documento Word original exacto para corregir los errores detectados en segundos.
* **📦 Cero-Configuración:** Disponible como ejecutable standalone, sin necesidad de que el usuario final instale Python.

---

## 🛠️ Arquitectura y Tecnologías
El sistema implementa el principio de **Separación de Responsabilidades (SoC)**, dividiendo el código en dos capas aisladas:

1. **Capa de Dominio (`moodle_logic.py`):** Motor algorítmico independiente. Utiliza `python-docx` para análisis del DOM, `re` para procesamiento de lenguaje natural y `xml.sax` para la inyección sanitizada de datos (CDATAs).
2. **Capa de Presentación (`app_gui.py`):** Interfaz gráfica construida con `tkinter` y `ttk`. Utiliza `threading` para evitar bloqueos durante el análisis asíncrono de documentos masivos.

---

## 📋 Reglas de Formato (Documentos Word)
Para que el motor procese exitosamente las preguntas, el documento origen debe cumplir con un estándar tipográfico sencillo:

**1. Enunciados:** Deben comenzar con un número seguido de un punto, guion o paréntesis.
> `1. ¿Qué es la Arquitectura Empresarial?`

**2. Opciones:** Letras seguidas de punto o paréntesis. 

**3. Marcado de Respuesta (Aceptación):** Existen dos métodos soportados:

* **Método A (Marca en línea):** Un asterisco o 'x' entre corchetes antes de la opción.

    ```text
    [x] a) Opción correcta
    [ ] b) Opción incorrecta
    ```

* **Método B (Declaración final):** Escribir la respuesta al final de la pregunta.

    ```text
    a) Primera opción
    b) Segunda opción
    Respuesta: b
    ```

---

## 🚀 Instalación y Uso

### Para Usuarios Finales (Profesores/Administrativos)
Simplemente descargue la última versión compilada (Ejecutable `.exe` / `.app`) desde la pestaña **Releases**. No requiere instalación de bases de datos ni lenguajes de programación.

### Para Desarrolladores
Si deseas ejecutar el código fuente o contribuir al proyecto:

1. Clona este repositorio:

    ```bash
    git clone https://github.com/ShanderSandoval/Generador-Preguntas.git
    ```

2. Crea y activa un entorno virtual:

    **En Windows:**
    ```cmd
    python -m venv .venv
    .\.venv\Scripts\activate
    ```
    **En macOS / Linux:**
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate
    ```

3. Instala las dependencias necesarias desde el archivo de requerimientos:

    ```bash
    pip install -r requirements.txt
    ```

4. Ejecuta la aplicación:

    **En Windows:**
    ```cmd
    python app_gui.py
    ```
    **En macOS / Linux:**
    ```bash
    python3 app_gui.py
    ```

---

## 🛡️ Lógica de Validación (Fail Loudly)
Este software prioriza la **integridad referencial** de la base de datos de destino. Aplicando el principio de diseño *Fail Loudly* (Falla de forma visible), si una pregunta carece de una respuesta válida:

1. No se intentan forzar respuestas predeterminadas (No Fallback).
2. Se aísla visualmente en la tabla de auditoría de la interfaz con la etiqueta `⚠️ Rechazada`.
3. Se excluye algorítmicamente de la función exportadora `generate_valid_xml()`.

---

## 👨‍💻 Autor
* **Shander Sandoval** - *Estudiante Ingeniería de Sistemas de Información* - Universidad Central del Ecuador.

---
*Si este proyecto te ayudó a automatizar tus clases, no olvides darle una ⭐ al repositorio.*