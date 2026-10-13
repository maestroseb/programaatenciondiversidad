# Captura de pantallas de Séneca

Extensión de Chrome que guarda la estructura de la pantalla de Séneca abierta (campos, botones, formularios y HTML de todos los marcos) para preparar el volcado de programas específicos. No escribe nada en Séneca.

## Instalar
1. Descarga esta carpeta (`extension-captura-seneca`).
2. En Chrome abre `chrome://extensions`, activa **Modo desarrollador** (arriba a la derecha).
3. Pulsa **Cargar descomprimida** y elige la carpeta.
4. Fija el icono en la barra (icono del puzle → chincheta).

## Usar
1. En Séneca, abre un programa específico de un alumno ya hecho o de prueba.
2. En cada pantalla (ámbito, objetivo, alta de indicador, contenido, actividad, ventanas emergentes), pulsa el icono de la extensión.
3. Escribe qué pantalla es y qué botón lleva a la siguiente.
4. En «Nombres a ocultar» pon los del alumno y la familia (se recuerdan).
5. Pulsa **Capturar esta pantalla**. Se descarga `seneca-pantalla-01.json`, `-02`…

## Qué se anonimiza
- No se guarda el texto escrito en los campos (solo su longitud) ni los valores ocultos.
- Se quitan los valores de las URL (identificadores de alumno, sesión…).
- Los nombres indicados se sustituyen por `[NOMBRE]`.
- La imagen opcional **no** se anonimiza: revísala antes de enviarla.
