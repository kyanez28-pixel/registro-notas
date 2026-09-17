# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
import copy

def get_dua_subtable_xml():
    return """<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tblPr>
    <w:tblStyle w:val="Tablaconcuadrcula"/>
    <w:tblW w:w="0" w:type="auto"/>
    <w:jc w:val="center"/>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="683"/>
    <w:gridCol w:w="683"/>
    <w:gridCol w:w="683"/>
  </w:tblGrid>
  <w:tr>
    <w:trPr>
      <w:trHeight w:val="278"/>
      <w:jc w:val="center"/>
    </w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="683" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="7030A0"/>
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Sinespaciado"/>
          <w:jc w:val="center"/>
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:color w:val="FFFFFF"/>
            <w:sz w:val="20"/>
          </w:rPr>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:color w:val="FFFFFF"/>
            <w:sz w:val="20"/>
          </w:rPr>
          <w:t>R</w:t>
        </w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="683" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="00B0F0"/>
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Sinespaciado"/>
          <w:jc w:val="center"/>
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:color w:val="FFFFFF"/>
            <w:sz w:val="20"/>
          </w:rPr>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:color w:val="FFFFFF"/>
            <w:sz w:val="20"/>
          </w:rPr>
          <w:t>A-E</w:t>
        </w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="683" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="538135"/>
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Sinespaciado"/>
          <w:jc w:val="center"/>
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:color w:val="FFFFFF"/>
            <w:sz w:val="20"/>
          </w:rPr>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:color w:val="FFFFFF"/>
            <w:sz w:val="20"/>
          </w:rPr>
          <w:t>C-M</w:t>
        </w:r>
      </w:p>
    </w:tc>
  </w:tr>
</w:tbl>"""

def build_plan():
    doc = docx.Document('MATRIZ DE PLANIFICACIÓN MICROCURRICULAR 26-27_ORIGINAL_BACKUP.docx')
    table = doc.tables[0]

    # 1. Encabezado institucional (Row 0)
    for p in table.rows[0].cells[0].paragraphs:
        if '2025' in p.text:
            p.text = p.text.replace('2025 – 2026', '2026 – 2027').replace('2025 - 2026', '2026 – 2027')
        if 'PLANIFICACIÓN MICROCURRICULAR TRIMESTRAL' in p.text:
            p.text = 'UNIDAD EDUCATIVA FISCAL “CALDERÓN 2”\nPLANIFICACIÓN MICROCURRICULAR TRIMESTRAL\n2026 – 2027'

    # 2. Datos Informativos (Row 2 y Row 3)
    # Row 2
    r2 = table.rows[2]
    r2.cells[1].text = "LIC. KLEVER YANEZ / MSC. MERY DÍAZ / LIC. MONICA VINUEZA / LIC. ROCIO CHAUCA / MSC. EDGAR SARANGO"
    r2.cells[7].text = "LENGUA Y LITERATURA"
    r2.cells[12].text = "LENGUA Y LITERATURA"
    r2.cells[16].text = "05/10/2026"

    # Row 3
    r3 = table.rows[3]
    r3.cells[1].text = "SÉPTIMO"
    r3.cells[7].text = "A, B, C, D Y E"
    r3.cells[12].text = "1°"
    r3.cells[18].text = "05/10/2026"
    r3.cells[20].text = "23 DE NOVIEMBRE DEL 2026"

    # Row 5: Objetivos de Aprendizaje
    r5 = table.rows[5]
    r5.cells[3].text = (
        "O.LL.3.1. Interactuar con diversas expresiones culturales para acceder, participar y apropiarse de la cultura escrita.\n"
        "O.LL.3.3. Comprender discursos orales en diversos contextos de la actividad social y cultural y analizarlos con sentido crítico.\n"
        "O.LL.3.6. Leer de manera autónoma textos no literarios, con fines de recreación, información y aprendizaje, y aplicar estrategias cognitivas de comprensión.\n"
        "O.LL.3.8. Escribir relatos y textos expositivos, descriptivos e instructivos, adecuados a una situación comunicativa determinada para aprender y comunicarse.\n"
        "O.LL.3.11. Seleccionar y disfrutar textos literarios para realizar interpretaciones personales y construir significados compartidos con otros lectores."
    )

    # Row 6: Inserciones Curriculares
    r6 = table.rows[6]
    r6.cells[3].text = "CÍVICA, ÉTICA E INTEGRIDAD | EDUCACIÓN SOCIOEMOCIONAL | EDUCACIÓN PARA LA SEGURIDAD VIAL Y MOVILIDAD SOSTENIBLE"

    # 3. Preparar las 7 semanas
    r9 = table.rows[9]
    r10 = table.rows[10]

    for _ in range(6):
        clone = copy.deepcopy(r9._tr)
        r10._tr.addprevious(clone)

    weeks = [
        {
            "sem_title": "SEMANA 1",
            "dcd": "LL.3.1.1. Participar en contextos y situaciones que evidencien la funcionalidad de la lengua escrita como herramienta cultural.",
            "tema": "Tema: Orígenes de la escritura (de las pinturas rupestres a la era digital)\n\nMateriales:\nTexto de Lengua 7mo (Páginas 12 a 17)\nPizarra y marcadores de colores\nCuadernos de trabajo\nGuía Docente Maya/MinEduc pág. 12",
            "indicador": "Reconoce la funcionalidad de la lengua escrita como manifestación cultural y de identidad en diferentes contextos y situaciones, atendiendo a la diversidad lingüística del Ecuador. I.LL.3.1.1.",
            "evaluacion": "• Instrumentos: Rúbrica de participación oral y taller de aplicación del texto (pág. 17).\n• Actividades: Descifrado de códigos de sustitución alfabética, línea de tiempo en el cuaderno y respuestas a preguntas reflexivas.",
            "metodologia": "GAMIFICACIÓN & APRENDIZAJE POR DESCUBRIMIENTO",
            "actividad_nombre": "“Misión: Arqueólogos del Tiempo y el Enigma del Alfabeto Perdido” (Texto Pág. 12-17)",
            "fases": [
                {
                    "titulo": "Fase 1: Motivación – “Desbloqueando la Misión Arqueológica”",
                    "actividad": "El docente presenta una historia gamificada a partir de las imágenes de la página 12: 'La Gran Biblioteca de Alejandría y los primeros escribas de la humanidad necesitan recuperar los mensajes ocultos que explican el nacimiento de las civilizaciones. Los estudiantes serán Arqueólogos de la Cultura Escrita y deberán descifrar enigmas para avanzar en el mapa del tiempo.' Preguntas detonantes del libro: ¿Qué conoces sobre las primeras formas de escritura? ¿Cómo ha evolucionado en el tiempo? ¿Para qué usas la palabra escrita hoy?",
                    "mecanica": "Gamificación: Cada estudiante elige su 'Avatar Histórico' (Pintor Rupestre, Escriba Sumerio, Filósofo del Papiro o Escriba Digital) y recibe su bitácora de viaje con 100 Puntos de Sabiduría iniciales.",
                    "dua_rep": "Representación: Infografía visual de la pág. 12 con pinturas rupestres, tablillas sumerias, jeroglíficos y tecnología digital moderna.",
                    "dua_acc": "Acción y expresión: Selección de su avatar mediante dibujo o distintivo; participación oral en lluvia de ideas.",
                    "dua_mot": "Compromiso-Motivación: Desafío de exploración histórica inmersiva con insignias y niveles de arqueólogo."
                },
                {
                    "titulo": "Fase 2: Comprensión – “La Ruleta de los Soportes y la Línea del Tiempo”",
                    "actividad": "Lectura guiada de las páginas 13 a 17 sobre los soportes de la escritura (arcilla cuneiforme, papiro egipcio, pergamino, papel de trapo, imprenta de Gutenberg y soportes digitales). Dinámica de relevos por equipos: relacionar cada material con su civilización y función social para registrar cosechas, leyes y memorias.",
                    "mecanica": "Gamificación: Cada asociación correcta otorga 10 puntos y desbloquea la carta de poder cívico 'Memoria Colectiva' (valor de la preservación de acuerdos y derechos de los pueblos).",
                    "dua_rep": "Representación: Muestras táctiles/visuales de papel, arcilla y pergamino; organizadores gráficos cromáticos en pizarra.",
                    "dua_acc": "Acción y expresión: Respuestas orales, construcción de una línea de tiempo cronológica en el cuaderno o dramatización rápida.",
                    "dua_mot": "Compromiso-Motivación: Trabajo cooperativo en clanes con metas claras y recompensas inmediatas de equipo."
                },
                {
                    "titulo": "Fase 3: Aplicación – “El Gran Código Criptográfico de los Escribas”",
                    "actividad": "Los estudiantes asumen el reto de la página 17 del texto escolar: descifrar el mensaje secreto utilizando el código de sustitución alfabética. Luego, responden a situaciones reales: ¿Por qué la palabra escrita ofrece mayor seguridad jurídica y personal que un acuerdo verbal?",
                    "mecanica": "Gamificación: Los equipos que decodifican el mensaje en tiempo récord ganan 'Estrellas de Criptografía' y se convierten en 'Descifradores Maestros'.",
                    "dua_rep": "Representación: Clave criptográfica visual de la pág. 17 proyectada y en hojas de trabajo.",
                    "dua_acc": "Acción y expresión: Resolución escrita individual o en parejas; opción de verbalizar o graficar el mensaje.",
                    "dua_mot": "Compromiso-Motivación: Reto de resolución lógica de acertijos con retroalimentación formativa y lúdica."
                },
                {
                    "titulo": "Fase 4: Cierre – “El Papiro de la Identidad y la Cultura Escrita”",
                    "actividad": "Cada estudiante redacta un breve manifiesto personal en su bitácora: '¿Por qué la lengua escrita es una herramienta cultural imprescindible en mi vida cotidiana?' Se pegan en el gran 'Muro de los Escribas de 7mo'. (Viernes 09/10: Feriado Independencia de Guayaquil).",
                    "mecanica": "Gamificación: Condecoración colectiva con el diploma y título de 'Guardianes de la Memoria Humana'.",
                    "dua_rep": "Representación: Formatos variados en tarjetas pergamino impresas o cuadernos.",
                    "dua_acc": "Acción y expresión: Libertad de plasmar la conclusión en prosa, verso o cartel ilustrado.",
                    "dua_mot": "Compromiso-Motivación: Reconocimiento del logro individual dentro de la producción colectiva del aula."
                }
            ]
        },
        {
            "sem_title": "SEMANA 7 (12/10 al 16/10/2026) · COMUNICACIÓN ORAL",
            "dcd": "LL.3.2.1. Escuchar discursos orales y formular juicios de valor con respecto a su contenido y forma, y participar de manera respetuosa frente a las intervenciones de los demás.",
            "tema": "Tema: Discursos y juicios de valor / Taller de Comunicación Oral: “Hoy me convertiré en vendedor de sueños”\n\nMateriales:\nTexto de Lengua 7mo (Páginas 19 a 25)\nPizarra y marcadores de colores\nGuía Docente pág. 13\nCuadernos",
            "indicador": "Escucha discursos orales, parafrasea su contenido, formula juicios de valor sobre su forma y participa de manera respetuosa frente a las intervenciones de los demás. I.LL.3.2.1.",
            "evaluacion": "• Instrumentos: Rúbrica de expresión oral y lista de cotejo de juicios de valor respetuosos.\n• Actividades: Presentación del taller 'Vendedor de sueños' (pág. 25) y análisis crítico del discurso modelo.",
            "metodologia": "AULA INVERTIDA & GAMIFICACIÓN",
            "actividad_nombre": "“El Gran Torneo de los Vendedores de Sueños y Maestros de la Oratoria” (Texto Pág. 19-25)",
            "fases": [
                {
                    "titulo": "Fase 1: Motivación – “El Desafío de la Voz en Público”",
                    "actividad": "Apertura con la escena dialogada de la página 19 ('¿Prepararon ya su discurso para la Semana de la Expresión Oral?'). Audición de un fragmento del histórico discurso 'Yo tengo un sueño' de Martin Luther King. Preguntas de reflexión socioemocional: ¿Cómo te sientes al hablar frente al público? ¿Cómo podemos transformar los nervios en fuerza comunicativa?",
                    "mecanica": "Gamificación: 'Desafío del Micrófono Dorado': Cada estudiante recibe una tarjeta de orador y acumula puntos de elocuencia y empatía auditiva.",
                    "dua_rep": "Representación: Audio claro del discurso, viñeta gráfica del texto y preguntas guía en pizarra.",
                    "dua_acc": "Acción y expresión: Lluvia de ideas oral y dramatización de gestos corporales de confianza.",
                    "dua_mot": "Compromiso-Motivación: Identificación con emociones reales y superación guiada del miedo escénico."
                },
                {
                    "titulo": "Fase 2: Comprensión – “Desarmando el Discurso: El Rompecabezas Retórico”",
                    "actividad": "Lectura comentada de 'Discurso de un niño' (págs. 20-21). En equipos cooperativos, los estudiantes desarman y reconstruyen la estructura formal del discurso oral: Introducción (saludo y captura de atención), Desarrollo (argumentos fundamentados) y Conclusión (llamado a la acción). Análisis de cómo emitir juicios de valor asertivos sin agredir.",
                    "mecanica": "Gamificación: Dinámica 'Detectives de Argumentos': Ganan 15 puntos quienes identifiquen la idea principal y separen opiniones de hechos comprobables.",
                    "dua_rep": "Representación: Tarjetas de colores para cada parte del discurso (Verde: Inicio, Azul: Desarrollo, Naranja: Cierre).",
                    "dua_acc": "Acción y expresión: Ordenamiento físico de tarjetas en pizarra y resumen esquemático en cuaderno.",
                    "dua_mot": "Compromiso-Motivación: Aprendizaje activo y deductivo mediante el análisis de discursos reales."
                },
                {
                    "titulo": "Fase 3: Aplicación – “Torneo en Vivo: Vendedores de Sueños”",
                    "actividad": "Puesta en práctica del Taller de Comunicación Oral de la página 25: en parejas, eligen una temática de impacto social o ecológico, redactan su guion en 3 partes y presentan una disertación de 2 minutos ante la clase aplicando modulación vocal, contacto visual y pausas efectivas.",
                    "mecanica": "Gamificación: Los compañeros del auditorio actúan como jurado calificador otorgando 'Estrellas de Juicio Constructivo' mediante una rúbrica sencilla.",
                    "dua_rep": "Representación: Guía paso a paso de la pág. 25 y cronómetro visual proyectado.",
                    "dua_acc": "Acción y expresión: Exposición oral en parejas con apoyo de una tarjeta síntesis.",
                    "dua_mot": "Compromiso-Motivación: Empoderamiento de la voz del estudiante en un ambiente de respeto mutuo."
                },
                {
                    "titulo": "Fase 4: Cierre – “El Podio de la Oratoria y Acuerdos de Escucha”",
                    "actividad": "Coevaluación formativa grupal: cada equipo recibe retroalimentación afectiva y constructiva. Firma colectiva del 'Decálogo del Auditorio Respetuoso' en el aula.",
                    "mecanica": "Gamificación: Entrega de la insignia digital 'Maestro de la Elocuencia' a todos los participantes.",
                    "dua_rep": "Representación: Cartel síntesis con los compromisos de escucha asertiva.",
                    "dua_acc": "Acción y expresión: Formulación oral de compromisos personales de respeto a la palabra ajena.",
                    "dua_mot": "Compromiso-Motivación: Celebración del crecimiento comunicativo grupal."
                }
            ]
        },
        {
            "sem_title": "SEMANA 8 (19/10 al 23/10/2026) · LECTURA",
            "dcd": "LL.3.3.5. Valorar los aspectos de forma y el contenido de un texto científico, a partir de criterios preestablecidos.\n\nLL.3.3.8. Leer con fluidez y entonación en diversos contextos y con diferentes propósitos.",
            "tema": "Tema: Leo para acercarme a la ciencia (textos de divulgación científica) y Uso de la biblioteca escolar\n\nMateriales:\nTexto de Lengua 7mo (Páginas 26 a 37)\nPizarra y marcadores de colores\nCuadernos\nFichas de biblioteca escolar",
            "indicador": "Realiza inferencias fundamentales y proyectivo-valorativas, valora los contenidos y aspectos de forma a partir de criterios preestablecidos, al monitorear su comprensión mediante el uso de estrategias cognitivas. I.LL.3.3.2.",
            "evaluacion": "• Instrumentos: Cuestionario de comprensión lectora y ficha de préstamo bibliotecario evaluada.\n• Actividades: Resolución de actividades págs. 34-37 del texto de Lengua y Literatura.",
            "metodologia": "APRENDIZAJE BASADO EN PROBLEMAS (ABP) & COOPERATIVO",
            "actividad_nombre": "“Expedición Paleontológica: Cazadores de Fósiles y de la Verdad Científica” (Texto Pág. 26-37)",
            "fases": [
                {
                    "titulo": "Fase 1: Motivación – “El Dilema del Hallazgo Científico”",
                    "actividad": "El docente plantea un reto de la vida real: 'En redes sociales se afirma que se encontró un dinosaurio vivo en la Amazonía ecuatoriana. ¿Cómo sabemos si una noticia científica es real o un engaño?' Apertura con la ilustración de la pág. 26 (paleontólogos y fósiles). Preguntas detonantes: ¿Dónde buscamos información científica verídica? ¿Qué es un artículo de divulgación?",
                    "mecanica": "Gamificación: Misión 'Cazadores de Fake News': Cada equipo recibe su pasaporte científico y el rango de 'Investigador Junior'.",
                    "dua_rep": "Representación: Imágenes paratextuales de fósiles de la pág. 26 y comparación visual entre una noticia falsa y un artículo formal.",
                    "dua_acc": "Acción y expresión: Lluvia de ideas y debates breves en mesas redondas.",
                    "dua_mot": "Compromiso-Motivación: Despertar del espíritu crítico frente a la información digital."
                },
                {
                    "titulo": "Fase 2: Comprensión – “Disección del Artículo Científico y la Biblioteca”",
                    "actividad": "Lectura secuencial y comentada del artículo de divulgación científica de las páginas 28 a 31 ('Huellas en el tiempo'). Análisis de su silueta textual: título atrayente, introducción, datos comprobables, glosario técnico y conclusiones. Exploración de las normas de la biblioteca escolar y fichas de préstamo (pág. 37).",
                    "mecanica": "Gamificación: 'Rally Bibliográfico': Puntos acumulables al identificar datos cuantitativos, citas de expertos y fuentes del texto.",
                    "dua_rep": "Representación: Organizador visual de llaves en pizarra y subrayado por colores de ideas clave.",
                    "dua_acc": "Acción y expresión: Registro en cuaderno y diseño de una ficha técnica de préstamo bibliotecario.",
                    "dua_mot": "Compromiso-Motivación: Familiarización con la biblioteca como espacio de investigación autónoma."
                },
                {
                    "titulo": "Fase 3: Aplicación – “Laboratorio de Comprensión y Fichaje Científico”",
                    "actividad": "Resolución de las actividades de comprensión lectora, inferencias y valoración crítica de las páginas 34 a 36 del libro escolar. En parejas cooperativas, seleccionan un libro de ciencias de la biblioteca del aula y elaboran una ficha técnica de registro bibliográfico completa (pág. 37).",
                    "mecanica": "Gamificación: Los equipos que completan el fichaje sin errores de formato reciben la 'Insignia de Curador Científico'.",
                    "dua_rep": "Representación: Plantillas impresas de fichas bibliográficas y cuestionario del libro.",
                    "dua_acc": "Acción y expresión: Trabajo en duplas, argumentación escrita y validación mutua.",
                    "dua_mot": "Compromiso-Motivación: Aplicación directa de los métodos de búsqueda y archivo documental."
                },
                {
                    "titulo": "Fase 4: Cierre – “Publicación del Boletín de Paleontología”",
                    "actividad": "Socialización en plenaria: cada dupla presenta un dato científico curioso comprobado en su lectura. Registro del avance en el portafolio escolar.",
                    "mecanica": "Gamificación: Ascenso de nivel a 'Investigador Científico de 7mo'.",
                    "dua_rep": "Representación: Mural de aula con las fichas de libros de divulgación investigados.",
                    "dua_acc": "Acción y expresión: Exposición oral rápida (1 minuto) por equipo.",
                    "dua_mot": "Compromiso-Motivación: Celebración de la lectura comprensiva y rigurosa."
                }
            ]
        },
        {
            "sem_title": "SEMANA 9 (26/10 al 30/10/2026) · ESCRITURA & GRAMÁTICA",
            "dcd": "LL.3.4.1. Relatar textos con secuencia lógica, manejo de conectores y coherencia en el uso de la persona y el tiempo verbal.\n\nLL.3.4.6. Autorregular la producción escrita mediante planificación, redacción y revisión.\n\nLL.3.4.10. Integrar en las producciones escritas el verbo (accidentes, modos) y uso del punto y coma (;).",
            "tema": "Tema: Escritura: Escribo una nota científica (planificación y redacción) y Gramática: El verbo (accidentes, modos) y Uso del punto y coma (;)\n\nMateriales:\nTexto de Lengua 7mo (Páginas 38 a 53)\nPizarra y marcadores de colores\nCuadernos de trabajo\nPlantilla de redacción",
            "indicador": "Produce textos informativos utilizando el proceso de escritura, elementos gramaticales apropiados (verbos, modos y tiempos) y signos de puntuación (punto y coma). I.LL.3.6.1. / I.LL.3.6.3.",
            "evaluacion": "• Instrumentos: Rúbrica de producción escrita y prueba de aplicación gramatical.\n• Actividades: Redacción de la nota científica y resolución de talleres págs. 46-53 del texto.",
            "metodologia": "GAMIFICACIÓN & APRENDIZAJE POR RETOS",
            "actividad_nombre": "“Sala de Redacción: Los Detectives del Verbo y el Punto y Coma” (Texto Pág. 38-53)",
            "fases": [
                {
                    "titulo": "Fase 1: Motivación – “Convocatoria Editorial Escolar”",
                    "actividad": "El docente transforma el aula en una Sala de Prensa Científica: 'La revista escolar necesita redactores para publicar notas sobre innovaciones sostenibles y movilidad activa (pág. 38)'. Preguntas: ¿Cómo redactar con exactitud? ¿Por qué los verbos son el motor de las oraciones en una investigación científica?",
                    "mecanica": "Gamificación: Cada estudiante recibe su credencial de 'Periodista Científico' y un tablero de retos gramaticales.",
                    "dua_rep": "Representación: Modelo visual de nota científica breve en pág. 39 con titulares y párrafos numerados.",
                    "dua_acc": "Acción y expresión: Lluvia de ideas y elección del tema de investigación individual o en duplas.",
                    "dua_mot": "Compromiso-Motivación: Sentido de propósito real: escribir para ser leídos por la comunidad escolar."
                },
                {
                    "titulo": "Fase 2: Comprensión – “El Tablero de Conjugación y el Laboratorio Gramatical”",
                    "actividad": "Planificación textual con la matriz de la página 40 (propósito, destinatario y fuentes). Análisis de los accidentes verbales (persona, número, tiempo) y de los Modos Verbales: indicativo (hechos reales), subjuntivo (deseos/dudas) e imperativo (instrucciones) págs. 42-45. Reglas del uso del punto y coma (;) en enumeraciones complejas (pág. 52).",
                    "mecanica": "Gamificación: Reto 'Detectives del Verbo': Acumulan puntos al transformar oraciones a modo subjuntivo y colocar correctamente el punto y coma.",
                    "dua_rep": "Representación: Tablas cromáticas de conjugación y esquemas sinópticos de modos verbales.",
                    "dua_acc": "Acción y expresión: Ejercicios de conjugación en pizarra y transcripción organizada en cuaderno.",
                    "dua_mot": "Compromiso-Motivación: Descubrimiento lógico y lúdico de la estructura gramatical del idioma."
                },
                {
                    "titulo": "Fase 3: Aplicación – “Redacción en Cabina y Laboratorio de Corrección”",
                    "actividad": "Redacción del primer borrador de la nota científica (páginas 46-47) estructurada en 3 párrafos coherentes, empleando verbos precisos y conectores de causa-efecto. Revisión cruzada en parejas aplicando la lista de cotejo del texto escolar.",
                    "mecanica": "Gamificación: 'Sello Editorial de Calidad': Obtienen la insignia de 'Editor Riguroso' al corregir concordancias y puntuación.",
                    "dua_rep": "Representación: Plantilla guía de redacción estructurada y rúbrica de autoevaluación pág. 47.",
                    "dua_acc": "Acción y expresión: Escritura autónoma, corrección colaborativa entre pares y edición.",
                    "dua_mot": "Compromiso-Motivación: Orgullo por la creación de un texto informativo de calidad."
                },
                {
                    "titulo": "Fase 4: Cierre – “Edición Final y Publicación del Boletín”",
                    "actividad": "Paso a limpio de las notas científicas corregidas y compilación en el primer número del 'Boletín Científico de Séptimo'. Resolución de actividades de consolidación págs. 50-53.",
                    "mecanica": "Gamificación: Publicación oficial de artículos y premiación con la medalla 'Pluma de Oro'.",
                    "dua_rep": "Representación: Ejemplar físico o digital compilado en el aula.",
                    "dua_acc": "Acción y expresión: Lectura compartida en asamblea de aula.",
                    "dua_mot": "Compromiso-Motivación: Cierre exitoso del ciclo de producción escrita."
                }
            ]
        },
        {
            "sem_title": "SEMANA 10 (02/11 al 06/11/2026) · LITERATURA",
            "dcd": "LL.3.5.1. Reconocer en un texto literario los elementos característicos que le dan sentido (personajes míticos, orígenes del mundo, cosmovisión).",
            "tema": "Tema: Literatura: Disfruto de los mitos (concepto, características, mitos de creación, dioses y héroes griegos y andinos)\n\nMateriales:\nTexto de Lengua 7mo (Páginas 54 a 65)\nPizarra y marcadores\nCuadernos\nAntología de mitos andinos y universales",
            "indicador": "Reconoce en textos de la literatura oral y escrita (mitos) los elementos característicos que les dan sentido; participa en discusiones sobre textos literarios aportando interpretaciones personales. I.LL.3.7.1.",
            "evaluacion": "• Instrumentos: Rúbrica de análisis literario y cuadro comparativo de mitos.\n• Actividades: Resolución de actividades págs. 62-65 del libro y participación en debates literarios. (Feriados 02 y 03 Noviembre).",
            "metodologia": "PEDAGOGÍA CRÍTICA & GAMIFICACIÓN",
            "actividad_nombre": "“La Liga Mitológica: El Despertar de los Dioses y la Cosmovisión Ancestral” (Texto Pág. 54-65)",
            "fases": [
                {
                    "titulo": "Fase 1: Motivación – “El Llamado de los Dioses y el Origen del Fuego”",
                    "actividad": "Narración dramatizada del mito de la creación del Sol y la Luna (pág. 54). Preguntas detonantes: ¿Por qué los pueblos antiguos crearon relatos sagrados sobre seres sobrenaturales? ¿Qué explicaciones buscaban ante los truenos, volcanes o la muerte? ¿Qué mitos ecuatorianos conocen (Chimborazo y Tungurahua, el mito Cañari)?",
                    "mecanica": "Gamificación: Los estudiantes se agrupan en 'Panteones Mitológicos' (Panteón Andino, Panteón Amazónico, Panteón Griego y Panteón Nórdico).",
                    "dua_rep": "Representación: Relato oral ambiental con efectos sonoros e ilustraciones míticas págs. 54-55.",
                    "dua_acc": "Acción y expresión: Comentarios orales de leyendas y relatos transmitidos por abuelos o familias.",
                    "dua_mot": "Compromiso-Motivación: Conexión con lo sagrado, la fantasía y la identidad intercultural."
                },
                {
                    "titulo": "Fase 2: Comprensión – “El Círculo de la Cosmovisión y el Mito”",
                    "actividad": "Lectura dialogada de los mitos seleccionados en las páginas 56 a 61. Sistematización de la diferencia entre mito (relato sagrado cosmogónico), leyenda (tradición popular local) y cuento común. Clasificación: mitos cosmogónicos, teogónicos y etiológicos. Análisis de los valores éticos de los héroes y dioses.",
                    "mecanica": "Gamificación: 'El Oráculo de la Sabiduría': Desafíos de preguntas inferenciales sobre símbolos y enseñanzas éticas que otorgan gemas de sabiduría.",
                    "dua_rep": "Representación: Cuadro comparativo visual entre mito y realidad; mapas conceptuales de tipos de mitos.",
                    "dua_acc": "Acción y expresión: Lectura por roles y transcripción creativa del mapa mental en cuaderno.",
                    "dua_mot": "Compromiso-Motivación: Respeto profundo por las diversas cosmovisiones y memoria histórica.",
                },
                {
                    "titulo": "Fase 3: Aplicación – “El Ágora Mitológica y Talleres de Interpretación”",
                    "actividad": "Resolución de los talleres de comprensión, interpretación de personajes y análisis de símbolos de las páginas 62 a 65 del texto. Creación de una ficha de personaje mítico (poderes, debilidades, enseñanza moral y dibujo simbólico).",
                    "mecanica": "Gamificación: Duelo de Panteones: Presentación breve del personaje mítico para ganar 'Puntos de Honor'.",
                    "dua_rep": "Representación: Cuestionarios estructurados del libro y fichas de personajes ilustradas.",
                    "dua_acc": "Acción y expresión: Producción plástica y argumentación escrita individual.",
                    "dua_mot": "Compromiso-Motivación: Disfrute estético y recreación de la literatura tradicional."
                },
                {
                    "titulo": "Fase 4: Cierre – “El Círculo de la Memoria Ancestral”",
                    "actividad": "Plenaria y asamblea de aula: reflexión sobre la importancia de cuidar las narraciones orales de nuestras nacionalidades indígenas. (Lunes 02 y Martes 03/11: Feriados Día de Difuntos e Independencia de Cuenca).",
                    "mecanica": "Gamificación: Consagración de todos los panteones con el título de 'Custodios de los Mitos del Mundo'.",
                    "dua_rep": "Representación: Mural colectivo de personajes mitológicos.",
                    "dua_acc": "Acción y expresión: Socialización oral y apreciación estética entre pares.",
                    "dua_mot": "Compromiso-Motivación: Valoración de la interculturalidad como tesoro vivo del Ecuador."
                }
            ]
        },
        {
            "sem_title": "SEMANA 11 (09/11 al 13/11/2026) · ESCRITURA CREATIVA & CIERRE UNIDAD 1",
            "dcd": "LL.3.5.1. Recrear textos literarios leídos o escuchados mediante el uso de diversos formatos.\n\nLL.3.4.13. Producir textos de acuerdo con la situación comunicativa en formatos variados (historieta mitológica).",
            "tema": "Tema: Escritura creativa: Historieta mitológica y Evaluación sumativa de la Unidad 1\n\nMateriales:\nTexto de Lengua 7mo (Páginas 66 a 71)\nCartulinas, reglas y lápices de colores\nPizarra y marcadores\nCuestionario sumativo de Unidad 1",
            "indicador": "Recrea textos literarios leídos mediante adaptaciones creativas (historieta) y demuestra el dominio de destrezas comunicativas de la Unidad 1. I.LL.3.7.2. / I.LL.3.6.1.",
            "evaluacion": "• Instrumentos: Rúbrica de la Historieta Mitológica y Evaluación sumativa de Unidad 1 (págs. 70-71).\n• Actividades: Exposición de historietas en el mural y prueba escrita de base estructurada. (Cierre oficial de la Planificación Microcurricular: 13/11/2026).",
            "metodologia": "JUEGO-TRABAJO & APRENDIZAJE COLABORATIVO",
            "actividad_nombre": "“Estudio Creativo de Cómics y Gran Desafío Sumativo de Unidad 1” (Texto Pág. 66-71)",
            "fases": [
                {
                    "titulo": "Fase 1: Motivación – “El Taller de los Historietistas”",
                    "actividad": "El docente transforma el aula en un estudio de diseño gráfico y cómic: exposición de muestras de historietas (pág. 66). Exploración de saberes previos: ¿Cómo cuenta una historia la combinación de imagen y texto? ¿Qué tipos de globos de diálogo conocen?",
                    "mecanica": "Gamificación: Cada dupla recibe el rol de 'Guionista' e 'Ilustrador' y desbloquea el 'Kit de Viñetas'.",
                    "dua_rep": "Representación: Modelos visuales de tiras cómicas e historietas míticas en págs. 66-67.",
                    "dua_acc": "Acción y expresión: Análisis oral interactivo de bocadillos, cartelas y onomatopeyas sonoras.",
                    "dua_mot": "Compromiso-Motivación: Entusiasmo infantil por el lenguaje dinámico del cómic."
                },
                {
                    "titulo": "Fase 2: Comprensión – “El Laboratorio del Cómic y Repaso Conceptual”",
                    "actividad": "Análisis técnico de los elementos de la historieta (págs. 67-68): planos visuales, viñetas, globos de diálogo (habla, susurro, pensamiento, grito), cartelas del narrador y onomatopeyas. Repaso de los conceptos nodales de la Unidad 1 (lengua y cultura, discursos, divulgación científica, gramática del verbo y mitos).",
                    "mecanica": "Gamificación: 'Trivia Sumativa de Repaso': Preguntas relámpago que otorgan 'Escudos del Conocimiento' para la prueba.",
                    "dua_rep": "Representación: Tipologías visuales de viñetas y mapa síntesis en pizarra.",
                    "dua_acc": "Acción y expresión: Boceto de storyboard previo en el cuaderno y resolución de dudas conceptuales.",
                    "dua_mot": "Compromiso-Motivación: Seguridad y preparación integral para el cierre de la unidad."
                },
                {
                    "titulo": "Fase 3: Aplicación – “Mesa de Dibujantes y Desafío Sumativo”",
                    "actividad": "En duplas, los estudiantes adaptan un mito leído a una historieta ilustrada de 4 a 6 viñetas (págs. 68-69) en cartulina. Aplicación individual de la Evaluación Sumativa estructurada de la Unidad 1 (págs. 70-71) para evidenciar el dominio de las destrezas con criterio de desempeño.",
                    "mecanica": "Gamificación: Premiación con 'Insignias de Maestría Gráfica' por originalidad, ortografía y expresividad plástica.",
                    "dua_rep": "Representación: Formato de evaluación sumativa del texto escolar y plantillas de historieta.",
                    "dua_acc": "Acción y expresión: Producción plástica/literaria colaborativa y resolución autónoma de la prueba.",
                    "dua_mot": "Compromiso-Motivación: Demostración tangible de las competencias adquiridas en el ciclo."
                },
                {
                    "titulo": "Fase 4: Cierre – “Inauguración de la Galería Mitológica”",
                    "actividad": "Exposición pública de las historietas en el mural del pasillo ('Galería Mitológica de Séptimo'). Retroalimentación y felicitación colectiva. (Cierre oficial de las 6 semanas de Planificación Microcurricular al 13 de Noviembre de 2026).",
                    "mecanica": "Gamificación: Clausura con diploma de 'Autores Consagrados de 7mo'.",
                    "dua_rep": "Representación: Mural escolar visitado por otros grados y docentes.",
                    "dua_acc": "Acción y expresión: Explicación oral de la historieta a los visitantes.",
                    "dua_mot": "Compromiso-Motivación: Culminación exitosa y satisfacción por el producto final terminado."
                }
            ]
        },
        {
            "sem_title": "SEMANA 12 (16/11 al 23/11/2026) · PROYECTO INTEGRADOR",
            "dcd": "LL.3.4.13. Producir textos de diversa índole de acuerdo con la situación comunicativa, mediante formatos y recursos diversos.\n\nLL.3.2.2. Proponer intervenciones orales estructuradas para sustentar el producto del proyecto interdisciplinar.",
            "tema": "Tema: EJECUCIÓN, CONSOLIDACIÓN Y SOCIALIZACIÓN DEL PROYECTO INTEGRADOR DEL 1ER TRIMESTRE (1 SEMANA · FIN: 23 DE NOVIEMBRE DEL 2026)\n\nMateriales:\nGuía oficial del Proyecto Interdisciplinario\nRúbricas ministeriales de evaluación\nMateriales para estands del aula, dípticos y carteles ilustrados",
            "indicador": "Publica y socializa los resultados de proyectos integradores mediante textos estructurados e intervenciones orales claras, coherentes y fundamentadas. I.LL.3.6.4.",
            "evaluacion": "• Instrumentos: Rúbrica oficial ministerial del Proyecto Integrador (30% sumativo del proyecto).\n• Actividades: Sustentación oral del producto final y exposición en feria de aula interdisciplinar (FIN: 23 DE NOVIEMBRE DEL 2026).",
            "metodologia": "APRENDIZAJE BASADO EN PROYECTOS (ABP) & COOPERATIVO",
            "actividad_nombre": "“Feria Interdisciplinar Comunitaria: Mi Huella Sostenible y Ciudadana” (Fin Oficial: 23 de Noviembre de 2026)",
            "fases": [
                {
                    "titulo": "Fase 1: Motivación – “El Desafío Integrador Comunitario”",
                    "actividad": "El docente activa la semana culminante del Proyecto Integrador del 1er Trimestre: 'La comunidad de Calderón necesita soluciones prácticas e interdisciplinares para fomentar la lectura, la convivencia pacífica y el cuidado ambiental'. Revisión de las metas y articulación de Lengua y Literatura con Ciencias, Estudios Sociales y Matemática.",
                    "mecanica": "Gamificación: Cada equipo consolida su rango de 'Comité de Proyecto Interdisciplinar' y revisa su panel de hitos.",
                    "dua_rep": "Representación: Rúbrica ministerial proyectada en cartelera y cronograma semanal visible.",
                    "dua_acc": "Acción y expresión: Asamblea de equipo y distribución equitativa de roles para la sustentación.",
                    "dua_mot": "Compromiso-Motivación: Entusiasmo por visibilizar el producto elaborado durante el trimestre."
                },
                {
                    "titulo": "Fase 2: Comprensión – “Auditoría de Calidad y Retroalimentación entre Pares”",
                    "actividad": "Talleres de coevaluación y retroalimentación entre equipos de estudiantes: revisión del informe escrito, claridad de las conclusiones, ortografía y preparación del material de soporte (trípticos, carteles, maquetas o presentaciones digitales).",
                    "mecanica": "Gamificación: Dinámica 'Control de Calidad': Obtienen sellos de validación al solventar observaciones de mejora.",
                    "dua_rep": "Representación: Lista de chequeo coevaluativa con criterios objetivos.",
                    "dua_acc": "Acción y expresión: Diálogo asertivo y sugerencias fundamentadas entre pares.",
                    "dua_mot": "Compromiso-Motivación: Cultura de automejora y responsabilidad compartida."
                },
                {
                    "titulo": "Fase 3: Aplicación – “Feria de Aula y Sustentación Pública”",
                    "actividad": "Montaje de estands en el aula y presentación oficial del Proyecto Integrador ante la comunidad escolar (docentes, directivos y compañeros). Cada equipo realiza una sustentación oral aplicando las técnicas de oratoria aprendidas (contacto visual, modulación vocal, respuestas claras a preguntas del público).",
                    "mecanica": "Gamificación: Evaluación con rúbrica ministerial y entrega de insignias de 'Impacto Comunitario'.",
                    "dua_rep": "Representación: Estands interactivos, dípticos informativos y paneles visuales.",
                    "dua_acc": "Acción y expresión: Disertación oral fluida y demostración práctica del producto final.",
                    "dua_mot": "Compromiso-Motivación: Reconocimiento público del esfuerzo colectivo y aprendizaje colaborativo."
                },
                {
                    "titulo": "Fase 4: Cierre – “Consolidación Final y Calificación Ministerial”",
                    "actividad": "Evaluación sumativa oficial mediante la Rúbrica Ministerial del Proyecto Integrador (30% de la calificación sumativa trimestral). Registro y entrega del informe consolidado. Cierre de la etapa académica al **23 de Noviembre del 2026**.",
                    "mecanica": "Gamificación: Clausura institucional del proyecto con certificación de honor a los equipos destacados.",
                    "dua_rep": "Representación: Cuadro de calificaciones e informes cuantitativos y cualitativos.",
                    "dua_acc": "Acción y expresión: Reflexión individual escrita sobre los aprendizajes logrados y metas futuras.",
                    "dua_mot": "Compromiso-Motivación: Coronación exitosa del Proyecto Integrador del Primer Trimestre."
                }
            ]
        }
    ]

    for idx, week_info in enumerate(weeks):
        row = table.rows[9 + idx]

        # Col 0: DCD
        row.cells[0].text = week_info["dcd"]

        # Col 2: Tema y Recursos
        row.cells[2].text = week_info["tema"]

        # Col 6: Estrategias Metodológicas Activas
        c6 = row.cells[6]
        c6.text = "" # limpiar

        # Título de la semana
        p_title = c6.add_paragraph()
        r_title = p_title.add_run(week_info["sem_title"])
        r_title.bold = True
        r_title.font.name = "Times New Roman"
        r_title.font.size = Pt(11)
        p_title.paragraph_format.space_before = Pt(2)
        p_title.paragraph_format.space_after = Pt(2)

        # Metodología y Nombre de la actividad
        p_met = c6.add_paragraph()
        r_met_label = p_met.add_run(week_info["metodologia"] + "\n")
        r_met_label.bold = True
        r_met_label.font.name = "Times New Roman"
        r_met_label.font.size = Pt(10.5)
        
        r_act = p_met.add_run("Nombre de la actividad:\n" + week_info["actividad_nombre"])
        r_act.font.name = "Times New Roman"
        r_act.font.size = Pt(10)
        p_met.paragraph_format.space_after = Pt(4)

        # Fases con DUA y Subtabla Badge
        for fase in week_info["fases"]:
            # Título de la fase
            p_fase = c6.add_paragraph()
            r_fase = p_fase.add_run(fase["titulo"])
            r_fase.bold = True
            r_fase.font.name = "Times New Roman"
            r_fase.font.size = Pt(10)
            p_fase.paragraph_format.space_before = Pt(3)
            p_fase.paragraph_format.space_after = Pt(1)

            # Actividad
            p_act = c6.add_paragraph()
            r_act_lbl = p_act.add_run("Actividad: ")
            r_act_lbl.bold = True
            r_act_lbl.font.name = "Times New Roman"
            r_act_lbl.font.size = Pt(9.5)
            r_act_txt = p_act.add_run(fase["actividad"])
            r_act_txt.font.name = "Times New Roman"
            r_act_txt.font.size = Pt(9.5)
            p_act.paragraph_format.space_after = Pt(1)

            # Mecánica activa
            p_mec = c6.add_paragraph()
            r_mec_txt = p_mec.add_run(fase["mecanica"])
            r_mec_txt.font.name = "Times New Roman"
            r_mec_txt.font.size = Pt(9.5)
            p_mec.paragraph_format.space_after = Pt(2)

            # DUA aplicado encabezado
            p_dua = c6.add_paragraph()
            r_dua_lbl = p_dua.add_run("DUA aplicado:")
            r_dua_lbl.bold = True
            r_dua_lbl.font.name = "Times New Roman"
            r_dua_lbl.font.size = Pt(9.5)
            p_dua.paragraph_format.space_after = Pt(1)

            # DUA detalles
            for label, text in [
                ("Representación: ", fase["dua_rep"].replace("Representación: ", "")),
                ("Acción y expresión: ", fase["dua_acc"].replace("Acción y expresión: ", "")),
                ("Compromiso-Motivación: ", fase["dua_mot"].replace("Compromiso-Motivación: ", ""))
            ]:
                p_det = c6.add_paragraph()
                r_lbl = p_det.add_run(label)
                r_lbl.bold = True
                r_lbl.font.name = "Times New Roman"
                r_lbl.font.size = Pt(9)
                r_txt = p_det.add_run(text)
                r_txt.font.name = "Times New Roman"
                r_txt.font.size = Pt(9)
                p_det.paragraph_format.space_after = Pt(1)

            # Insertar la subtabla DUA badge [ R | A-E | C-M ]
            sub_tbl_xml = get_dua_subtable_xml()
            sub_tbl_element = parse_xml(sub_tbl_xml)
            c6._tc.append(sub_tbl_element)

            # Párrafo separador después del badge
            p_sep = c6.add_paragraph()
            p_sep.paragraph_format.space_before = Pt(2)
            p_sep.paragraph_format.space_after = Pt(2)

        # Col 13: Indicadores
        row.cells[13].text = week_info["indicador"]

        # Col 18: Actividades Evaluativas
        row.cells[18].text = week_info["evaluacion"]

    # Firmas de Responsabilidad
    r_firmas_cargos = table.rows[18]
    r_firmas_dates = table.rows[20]

    r_firmas_cargos.cells[0].text = "Docente:\nLic. Klever Yánez\nDocente de Lengua y Literatura"
    r_firmas_cargos.cells[4].text = "Coordinador/a del área/año:\nLic. Rocío Chauca"
    r_firmas_cargos.cells[10].text = "Junta de Directores:\nComisión Técnico Pedagógica"
    r_firmas_cargos.cells[15].text = "Vicerrector:\nMSc. Mauricio Guachamín"

    r_firmas_dates.cells[0].text = "Fecha: 05/10/2026"
    r_firmas_dates.cells[4].text = "Fecha: 09/10/2026"
    r_firmas_dates.cells[10].text = "Fecha: 11/10/2026"
    r_firmas_dates.cells[15].text = "Fecha: 23/11/2026"

    # Guardar en archivos
    target_file = 'MATRIZ DE PLANIFICACIÓN MICROCURRICULAR 26-27.docx'
    alt_file = 'MATRIZ DE PLANIFICACIÓN MICROCURRICULAR 26-27_ACTUALIZADA.docx'
    backup_target = 'PLANIFICACION_MICROCURRICULAR_7MO_LENGUA_1ER_TRIMESTRE.docx'
    
    doc.save(alt_file)
    print(f"Archivo guardado exitosamente: {alt_file}")
    
    doc.save(backup_target)
    print(f"Archivo guardado exitosamente: {backup_target}")
    
    try:
        doc.save(target_file)
        print(f"Archivo principal guardado exitosamente: {target_file}")
    except PermissionError:
        print(f"AVISO: {target_file} está actualmente abierto en Microsoft Word. Se guardó como {alt_file}")

if __name__ == "__main__":
    build_plan()
