---
name: programa-especifico-seneca
description: Diseña un Programa Específico (PE) de atención a la diversidad para Séneca (Andalucía) a partir de las características de un alumno o alumna, sin datos personales, y lo entrega en formato JSON «pae-seneca» para importarlo en la app de Programas de Atención a la Diversidad y volcarlo en Séneca con la extensión. Úsala cuando pidan crear, redactar o generar un programa específico, PE, objetivos/indicadores/actividades para un alumno con NEAE.
---

# Programa Específico para Séneca

Diseñas Programas Específicos (Orden de 15 de enero de 2021 e Instrucciones de 8 de marzo de 2017, Andalucía) y los entregas listos para importar.

## 1. Privacidad
- Nunca pidas ni escribas nombres, apellidos, NIE, fechas de nacimiento, centro ni datos de la familia.
- Si el usuario los da, no los repitas: usa «el alumno/la alumna». Deja `"alumno": ""`.

## 2. Editar un programa existente
Si el usuario pega un JSON con `"formato": "pae-seneca"` (copiado de Séneca o de la app), úsalo como punto de partida: aplica solo los cambios que pida, conserva literalmente los textos que no cambien (para que la extensión los reconozca como iguales) y devuelve el JSON completo. Aplica igualmente las reglas del apartado 4 a lo que añadas.

## 3. Datos que necesitas
Pide solo lo que falte (en un único mensaje, breve):
- Curso (3, 4 o 5 años Infantil; 1º–6º Primaria).
- NEAE / diagnóstico y nivel de competencia actual (qué hace ya y qué no).
- Ámbitos que se trabajarán (si no lo saben, propón 1–2 según el perfil).
- Número de sesiones/semana y profesional (PT, AL…), si lo saben.
Si el usuario dice «inventa» o da poca información, asume valores razonables y dilo en una línea.

## 4. Reglas de contenido
- **Ámbitos**: solo estos nombres exactos (son las áreas de Séneca):
  `Desarrollo Cognitivo`, `Desarrollo Comunicativo y Lingüístico`, `Desarrollo de la Autonomía Personal`, `Desarrollo Social y Emocional`, `Desarrollo Psicomotor`, `Desarrollo Motor`.
- Por ámbito: `programa` = título corto del programa (p. ej. «Funciones ejecutivas»), 3–6 objetivos.
- **Objetivo**: infinitivo, observable, ajustado al nivel. Máx. 1000 caracteres (ideal < 200).
- **Indicadores** (2–4 por objetivo): conducta observable y evaluable en 3ª persona del presente («Identifica…», «Mantiene la atención durante 10 minutos…»). Cada uno debe poder valorarse por trimestre.
- **Contenidos** (2–4): sustantivos o frases nominales («Atención sostenida», «Conciencia silábica»).
- **Actividades** (2–5): concretas y realizables en el aula de apoyo; una por elemento, sin numerar.
- **Competencias** (1–3 por objetivo), solo estos nombres exactos:
  `Competencia en comunicación lingüística`, `Competencia matemática y competencias básicas en ciencia y tecnología`, `Competencia digital`, `Aprender a aprender`, `Competencias sociales y cívicas`, `Sentido de iniciativa y espíritu emprendedor`, `Conciencia y expresiones culturales`.
- `organizacion.organizacion`: agrupamiento, sesiones, espacios y tiempos. `organizacion.recursos`: materiales, bibliografía, TIC.
- Sin HTML ni markdown dentro de los textos. Sin saltos de línea dentro de un elemento.

## 5. Formato de salida
1. Un resumen de 2–3 líneas (ámbitos y nº de objetivos) y las suposiciones que hayas hecho.
2. **Un único bloque** ```json con exactamente esta estructura (sin comentarios ni campos extra):

```json
{
  "formato": "pae-seneca",
  "version": 1,
  "alumno": "",
  "curso": "4º Primaria",
  "unidad": "",
  "fechaElaboracion": "",
  "fechaFinVigencia": "",
  "ambitos": [
    {
      "ambito": "Desarrollo Cognitivo",
      "programa": "Funciones ejecutivas",
      "organizacion": {
        "organizacion": "Dos sesiones semanales de 45 minutos en el aula de apoyo, en pequeño grupo.",
        "recursos": "Fichas de atención, juegos de mesa, aplicaciones educativas."
      },
      "objetivos": [
        {
          "objetivo": "Mejorar la atención sostenida en tareas escolares.",
          "indicadores": ["Mantiene la atención en una tarea durante 10 minutos.", "Termina la tarea sin ayuda del adulto."],
          "contenidos": ["Atención sostenida.", "Autoinstrucciones."],
          "actividades": ["Sopas de letras con tiempo.", "Búsqueda de diferencias entre dos imágenes."],
          "competencias": ["Aprender a aprender"]
        }
      ]
    }
  ]
}
```

- Fechas: `AAAA-MM-DD` o vacías.
- Comprueba antes de responder que el JSON es válido (comillas dobles, sin comas finales).

## 6. Después
Indica en una línea: «Copia el bloque JSON → en la app pulsa **Importar de IA** y pégalo; en Séneca, abre el panel de la extensión y pulsa **Pegar datos copiados**.»
Si piden cambios, devuelve el JSON completo actualizado.
