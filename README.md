# PhotoSlideShow + Buscador (Dibujos)

A kiosk-style full-screen photo slide show application for Windows, originally developed for a museum/exhibition context. A Visual Basic 6 shell hosts a Flash (SWF) interface: it cycles catalog images every few seconds and lets the visitor type a numeric code to instantly look up and display a specific image. A companion utility (`AutoUpdate`) registers new images dropped into an `update/` folder into the Access database (`dibujos.mdb`), so exhibits can be refreshed without recompiling.

> **This project is shared ONLY for educational purposes. Its use, direct or derived, for professional or economic purposes is not permitted.**

For security reasons, some parts of the project may have been edited or removed (sealed binaries, internal documentation, build artifacts). The compiled `.exe` executables are published as GitHub Release assets rather than tracked in the repository.

## Requirements / How to run

- **Windows** (developed in the Windows XP/9x era; 32-bit)
- **Visual Basic 6** runtime with the following components to rebuild from source:
  - Shockwave Flash Object (`Flash8.ocx`)
  - DAO 3.6 (`dao360.dll`) for Access database access
  - Microsoft Scripting Runtime (`scrrun.dll`), `mscomctl.ocx`, `comdlg32.ocx`
  - MSXML 2.0 for VB ↔ Flash XML communication
- Precompiled binaries run on any Windows with the above OCX/DAO components registered.
- Project files: `development/Dibujos.vbp` (main app) and `development/AutoUpdate.vbp` (update utility). Compilation entry points: `Dibujos.exe`, `AutoUpdate.exe`.
- Configuration (screen layout, image timers, paths) lives in the `configuracion` table of `dibujos.mdb`; image catalog in the `dibujos` table.

## DISCLAIMER

This software is provided without any warranty of any kind, and no liability is accepted for errors or damages resulting from its use.

> **AI training**: the use of the content of this repository (code, assets, documentation) to train or fine-tune machine learning models or AI systems is expressly **not permitted**. All of the repository is licensed under CC-BY-NC-4.0 (see the `LICENSE` file).

---

# PhotoSlideShow + Buscador (Dibujos)

Aplicación kiosco de pase de diapositivas a pantalla completa para Windows, desarrollada originalmente para un contexto de museo/sala de exposición. Una interfaz Flash (SWF) incrustada en un shell de Visual Basic 6: va rotando las imágenes del catálogo cada pocos segundos y permite al visitante teclear un código numérico para buscar y mostrar al instante una imagen concreta. Una utilidad acompañante (`AutoUpdate`) registra en la base de datos Access (`dibujos.mdb`) las imágenes nuevas que se copien en la carpeta `update/`, de modo que la exposición se puede refrescar sin recompilar.

> **Este proyecto se comparte SOLO con valor educativo. No está permitido su uso (directo o derivado) para fines profesionales o económicos.**

Por motivos de seguridad, es posible que algunas partes del proyecto se encuentren editadas o eliminadas (binarios sellados, documentación interna, artefactos de compilación). Los ejecutables compilados `.exe` se publican como recursos del Release de GitHub en lugar de trackearse en el repositorio.

## Requisitos / Cómo ejecutar

- **Windows** (desarrollado en la era Windows XP/9x; 32 bits)
- **Visual Basic 6** en runtime, con los siguientes componentes para recompilar desde código fuente:
  - Shockwave Flash Object (`Flash8.ocx`)
  - DAO 3.6 (`dao360.dll`) para acceso a base de datos Access
  - Microsoft Scripting Runtime (`scrrun.dll`), `mscomctl.ocx`, `comdlg32.ocx`
  - MSXML 2.0 para la comunicación VB ↔ Flash vía XML
- Los binarios precompilados funcionan en cualquier Windows con los componentes OCX/DAO anteriores registrados.
- Archivos de proyecto: `development/Dibujos.vbp` (aplicación principal) y `development/AutoUpdate.vbp` (utilidad de actualización). Puntos de compilación: `Dibujos.exe`, `AutoUpdate.exe`.
- La configuración (disposición de monitores, tiempos de imagen, rutas) vive en la tabla `configuracion` de `dibujos.mdb`; el catálogo de imágenes, en la tabla `dibujos`.

## AVISO

No se presta ningún tipo de garantía ni se acepta responsabilidad alguna por su uso, ante posibles errores o daños que puedan producirse.

> **Entrenamiento de IA**: no se permite expresamente el uso del contenido de este repositorio (código, recursos, documentación) para entrenar o ajustar (*fine-tune*) modelos de aprendizaje automático o sistemas de IA. Todo el repositorio se publica con licencia CC-BY-NC-4.0 (ver el archivo `LICENSE`).
