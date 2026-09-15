# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import copy

def create_7mo_plan():
    doc = docx.Document('LENGUA TERCER  TRIMESTRE CONTINUACION.docx')
    table = doc.tables[0]
    
    # 1. Modificar encabezado de institución y periodo (Row 0)
    for p in table.rows[0].cells[0].paragraphs:
        if '2025' in p.text or 'PLANIFICACI' in p.text:
            p.text = p.text.replace('2025 – 2026', '2026 – 2027').replace('2025 - 2026', '2026 – 2027')
            p.text = p.text.replace('PLANIFICACIÓN MICROCURRICULAR TRIMESTRAL', 'PLANIFICACIÓN MICROCURRICULAR DEL PRIMER TRIMESTRE (8 SEMANAS)')

    # 2. Modificar Datos Informativos (Row 2 y Row 3)
    # Row 2: Docentes
    r2_c1 = table.rows[2].cells[1]
    r2_c1.text = "LIC. KLEVER YANEZ / MSC. MERY DÍAZ / MSC. LUIS ROBLES / MSC. MARÍA ISABEL RIVAS / MSC. LAURA TORRES"
    
    # Row 3: Grado, trimestre y fechas oficiales
    r3 = table.rows[3]
    r3.cells[1].text = "SÉPTIMO"
    r3.cells[7].text = "A, B, C, D Y E (Asignado: 7mo \"C\")"
    r3.cells[12].text = "1er Trimestre"
    r3.cells[17].text = "05/10/2026"
    r3.cells[21].text = "30/11/2026 (Microcurricular: 05/10 al 13/11 · Proyecto: 16/11 al 23/11 · Exámenes: 24/11 al 30/11)"

    # 3. Aprendizaje Disciplinar (Row 5: Objetivos de Séptimo)
    r5_cell = table.rows[5].cells[3]
    r5_cell.text = (
        "O.LL.3.1. Interactuar con diversas expresiones culturales para acceder, participar y apropiarse de la cultura escrita.\n"
        "O.LL.3.3. Comprender discursos orales en diversos contextos de la actividad social y cultural y analizarlos con sentido crítico.\n"
        "O.LL.3.6. Leer de manera autónoma textos no literarios, con fines de recreación, información y aprendizaje, y aplicar estrategias cognitivas de comprensión.\n"
        "O.LL.3.8. Escribir relatos y textos expositivos, descriptivos e instructivos, adecuados a una situación comunicativa determinada para aprender y comunicarse.\n"
        "O.LL.3.11. Seleccionar y disfrutar textos literarios para realizar interpretaciones personales y construir significados compartidos con otros lectores."
    )

    # 4. Semanas de planificación: 6 microcurriculares + 1 proyecto integrador + 1 exámenes = 8 semanas
    # Insertar 2 filas clonando row 14 para semanas 12 y 13 antes de las firmas (row 15)
    r14 = table.rows[14]
    r15 = table.rows[15]
    
    # Clon 1 para semana 12
    clone1 = copy.deepcopy(r14._tr)
    r15._tr.addprevious(clone1)
    
    # Clon 2 para semana 13
    clone2 = copy.deepcopy(r14._tr)
    r15._tr.addprevious(clone2)

    weeks_data = [
        {
            # SEMANA 6 (Semana 1 de la Microcurricular)
            "row_idx": 9,
            "dcd": "LL.3.3.5. Valorar los aspectos de forma y el contenido de un texto, a partir de criterios preestablecidos.\n\nLL.3.3.8. Leer con fluidez y entonación en diversos contextos y con diferentes propósitos.",
            "tema": "Tema: Leo para acercarme a la ciencia (textos de divulgación científica) y Uso de la biblioteca escolar\n\nMateriales:\nTexto de Lengua 7mo (Páginas 26 a 37)\nPizarra y marcadores de colores\nCuadernos\nFichas de préstamo de biblioteca escolar\nGuía Docente pág. 12-14",
            "estrategias": [
                ("SEMANA 6 (Semana 1 de la Planificación Microcurricular · 05/10 al 09/10/2026)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("El docente solicita a los estudiantes abrir el libro en la página 26 y observar la ilustración de apertura (\"Leo para acercarme a la ciencia: huellas del pasado y fósiles\"). Lee en voz alta un fragmento sobre hallazgos paleontológicos y plantea preguntas: ¿Qué estudia la paleontología?, ¿cómo sabemos que existieron los dinosaurios?, ¿dónde encontramos artículos científicos confiables?", False, 10),
                ("Inserción Educación Cívica, Ética e Integridad:", True, 10),
                ("El docente reflexiona con los estudiantes sobre la honestidad académica, el valor de la verdad científica frente a noticias falsas (fake news) y el respeto a la propiedad intelectual al consultar fuentes bibliográficas.", False, 10),
                ("DUA", True, 10),
                ("Representación: ilustración del texto pág. 26, lectura modelo guiada con entonación y paratextos.\nAcción/expresión: lluvia de ideas oral, comentarios sobre fósiles y documentales.\nMotivación: curiosidad científica y activación de saberes previos del entorno natural.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("El docente guía la lectura comentada del artículo de divulgación científica de las páginas 28 a 31. Formula preguntas de análisis: ¿Cuál es el propósito comunicativo de un texto de divulgación?, ¿en qué se diferencia de un cuento o mito?, ¿qué vocabulario técnico utiliza? Los estudiantes identifican que no todos los textos comunican lo mismo y que los textos científicos buscan informar con rigor.", False, 10),
                ("DUA", True, 10),
                ("Representación: lectura paratextual (títulos, subtítulos, fotografías, pies de foto, glosario).\nAcción/expresión: respuestas a preguntas inferenciales de la pág. 32.\nMotivación: debate grupal sobre descubrimientos fósiles en el Ecuador y América Latina.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("Con la mediación del docente, los estudiantes identifican la silueta y estructura del artículo científico: título, introducción, desarrollo con datos comprobables, conclusiones y fuentes consultadas (pág. 33).\nEn la pizarra se construye un organizador gráfico (esquema de llaves) que los estudiantes completan en sus cuadernos:\n1. Definición de texto de divulgación científica.\n2. Características (claridad, objetividad, rigor científico).\n3. Estructura textual.\n4. Uso y normas de la biblioteca escolar y fichas de préstamo (pág. 37).", False, 10),
                ("DUA", True, 10),
                ("Representación: organizador gráfico estructurado en pizarra con marcadores de colores.\nAcción/expresión: registro ordenado en el cuaderno y elaboración de una ficha modelo de préstamo bibliotecario.\nMotivación: rol activo como pequeños investigadores científicos.", False, 10),
                ("APLICACIÓN", True, 11),
                ("Los estudiantes resuelven las actividades de comprensión lectora y valoración crítica de las páginas 34 a 36 del libro.\nEn parejas, seleccionan un libro informativo de la biblioteca escolar del aula y llenan una ficha de registro bibliotecario (pág. 37). Coevaluación guiada y socialización en plenaria. (Nota: Viernes 09 de Octubre: Feriado por la Independencia de Guayaquil).", False, 10),
                ("DUA", True, 10),
                ("Representación: consignas impresas del texto y plantilla de ficha bibliográfica.\nAcción/expresión: trabajo cooperativo en parejas, redacción de respuestas y socialización.\nMotivación: autogestión y fomento del hábito lector autónomo.", False, 10),
            ],
            "indicador": "Realiza inferencias fundamentales y proyectivo-valorativas, valora los contenidos y aspectos de forma a partir de criterios preestablecidos, al monitorear y autorregular su comprensión mediante el uso de estrategias cognitivas. I.LL.3.3.2.",
            "evaluacion": "Técnica: Análisis del desempeño y observación directa.\nInstrumento: Rúbrica de comprensión lectora y ficha de préstamo bibliotecario evaluada (actividades págs. 34-37)."
        },
        {
            # SEMANA 7 (Semana 2 de la Microcurricular)
            "row_idx": 10,
            "dcd": "LL.3.4.1. Relatar textos con secuencia lógica, manejo de conectores y coherencia en el uso de la persona y el tiempo verbal.\n\nLL.3.4.6. Autorregular la producción escrita mediante el uso habitual del proceso de planificación, redacción y revisión.",
            "tema": "Tema: Escribo una nota científica (estructura, planificación y redacción) y El verbo (persona, número, tiempo)\n\nMateriales:\nTexto de Lengua 7mo (Páginas 38 a 47)\nPizarra y marcadores de colores\nCuadernos\nFichas de planificación textual",
            "estrategias": [
                ("SEMANA 7 (Semana 2 de la Planificación Microcurricular · 12/10 al 16/10/2026)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("El docente ambienta la clase motivando al grupo a imaginar la preparación de una revista escolar para la Feria de Ciencias (pág. 38). Plantea preguntas: ¿Sobre qué fenómeno natural o avance tecnológico les gustaría informar?, ¿qué pasos seguimos antes de escribir un artículo para que sea claro y confiable?", False, 10),
                ("Inserción Educación para la Seguridad Vial y Movilidad Sostenible:", True, 10),
                ("Se propone como tema modelo de indagación científica el impacto del uso de la bicicleta y los vehículos de cero emisiones en la reducción de la huella de carbono y la movilidad activa en la ciudad.", False, 10),
                ("DUA", True, 10),
                ("Representación: modelo de nota científica breve en el texto pág. 39.\nAcción/expresión: lluvia de ideas oral sobre innovaciones sostenibles.\nMotivación: vinculación con la ecología y el cuidado del planeta.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("El docente analiza junto a los estudiantes el esquema de planificación de la página 40: propósito, destinatario, tipo de lenguaje y recopilación de datos objetivos. Se reflexiona sobre el papel del verbo en la redacción: ¿Qué palabras expresan las acciones y descubrimientos en una investigación? (ej: descubrieron, investigan, transformará).", False, 10),
                ("DUA", True, 10),
                ("Representación: tabla guía de planificación textual del texto.\nAcción/expresión: debate y selección del tema individual o en parejas.\nMotivación: libertad de elección temática con relevancia escolar.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("Explicación gramatical del verbo y sus accidentes gramaticales (págs. 42-45):\n- Persona (primera, segunda, tercera).\n- Número (singular y plural).\n- Tiempo (presente, pretérito/pasado, futuro).\nEn la pizarra se construye un organizador gráfico (mapa conceptual y cuadro de doble entrada de conjugación verbal) que los estudiantes transcriben a sus cuadernos, identificando raíz y desinencia verbal.", False, 10),
                ("DUA", True, 10),
                ("Representación: esquemas cromáticos en pizarra que diferencian raíz y terminaciones.\nAcción/expresión: completar tablas de conjugación y ejercicios de identificación de verbos.\nMotivación: dinámica lúdica \"detectives de verbos\".", False, 10),
                ("APLICACIÓN", True, 11),
                ("Los estudiantes redactan el primer borrador de su nota científica (págs. 46-47) aplicando la concordancia obligatoria entre sujeto y verbo. Desarrollan la revisión en parejas con una lista de cotejo básica (título claro, párrafos ordenados, verbos en tiempo adecuado).", False, 10),
                ("DUA", True, 10),
                ("Representación: plantilla estructurada de redacción en 3 párrafos.\nAcción/expresión: producción escrita individual y coevaluación formativa.\nMotivación: satisfacción de crear un texto científico propio.", False, 10),
            ],
            "indicador": "Produce textos informativos utilizando el proceso de escritura, elementos gramaticales apropiados (verbos en tiempo y persona adecuados) y estructura lógica. I.LL.3.6.1.",
            "evaluacion": "Técnica: Análisis de la producción escrita.\nInstrumento: Lista de cotejo de planificación y borrador de la nota científica (págs. 40-47)."
        },
        {
            # SEMANA 8 (Semana 3 de la Microcurricular)
            "row_idx": 11,
            "dcd": "LL.3.4.10. Expresar sus ideas con precisión e integrar en las producciones escritas los diferentes tipos de verbo, modos y tiempos.\n\nLL.3.4.12. Comunicar ideas con eficiencia a partir de la aplicación de las reglas de puntuación (uso del punto y coma).",
            "tema": "Tema: Modos verbales (indicativo, subjuntivo, imperativo), tiempos compuestos y Uso del punto y coma (;)\n\nMateriales:\nTexto de Lengua 7mo (Páginas 48 a 53)\nPizarra y marcadores\nTarjetas didácticas de conjugación\nCuadernos",
            "estrategias": [
                ("SEMANA 8 (Semana 3 de la Planificación Microcurricular · 19/10 al 23/10/2026)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("El docente escribe 3 oraciones en la pizarra con diferentes intenciones: \"Los científicos estudian el clima\" (certeza), \"Ojalá encontremos soluciones al cambio climático\" (deseo/duda), \"Cuiden los recursos naturales del aula\" (orden/exhortación). Plantea preguntas: ¿Qué actitud refleja el emisor en cada oración?, ¿en cuál afirma un hecho real?, ¿en cuál expresa un deseo?", False, 10),
                ("Inserción Educación Socioemocional:", True, 10),
                ("Se reflexiona sobre cómo expresamos nuestros sentimientos, metas y acuerdos en el aula empleando el modo subjuntivo (\"Espero que trabajemos en equipo\") y el modo indicativo con asertividad.", False, 10),
                ("DUA", True, 10),
                ("Representación: contraste visual en pizarra de las oraciones con colores distintos.\nAcción/expresión: dramatización de intenciones expresivas y respuestas orales.\nMotivación: conexión con las emociones y la comunicación respetuosa.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("Los estudiantes analizan las explicaciones y oraciones modelo de las páginas 48 y 49 del libro. Diferencian entre hechos reales (indicativo), deseos/posibilidades (subjuntivo) y mandatos (imperativo). Asimismo, analizan la necesidad del punto y coma (;) para pausar oraciones compuestas extensas.", False, 10),
                ("DUA", True, 10),
                ("Representación: cuadros comparativos del texto pág. 50.\nAcción/expresión: transformación oral de oraciones y justificación gramatical.\nMotivación: descubrimiento inductivo de la regla sintáctica.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("Construcción guiada de organizadores gráficos en el cuaderno:\n1. Cuadro sinóptico de los Modos Verbales: Indicativo, Subjuntivo e Imperativo.\n2. Tiempos compuestos con el verbo auxiliar haber (pretérito perfecto compuesto, pluscuamperfecto).\n3. Reglas de uso del punto y coma (;) para separar oraciones yuxtapuestas y enumeraciones complejas (pág. 52).", False, 10),
                ("DUA", True, 10),
                ("Representación: mapa conceptual estructurado con código de colores.\nAcción/expresión: elaboración de fichas síntesis en el cuaderno.\nMotivación: apropiación del conocimiento para mejorar la redacción.", False, 10),
                ("APLICACIÓN", True, 11),
                ("Desarrollo de las actividades prácticas de las páginas 51 a 53 del texto escolar. Ejercicios de aplicación del punto y coma en oraciones compuestas y corrección de modos verbales en producciones escritas propias.", False, 10),
                ("DUA", True, 10),
                ("Representación: ejercicios del libro y fichas de aplicación.\nAcción/expresión: resolución escrita individual y revisión cruzada entre pares.\nMotivación: autoedición y superación de errores comunes.", False, 10),
            ],
            "indicador": "Aplica elementos gramaticales (modos verbales, tiempos compuestos) y ortográficos (punto y coma) en producciones escritas con coherencia y cohesión. I.LL.3.6.3.",
            "evaluacion": "Técnica: Pruebas escritas y observación directa.\nInstrumento: Cuestionario de aplicación de modos verbales y uso del punto y coma (actividades págs. 51-53)."
        },
        {
            # SEMANA 9 (Semana 4 de la Microcurricular)
            "row_idx": 12,
            "dcd": "LL.3.5.1. Reconocer en un texto literario los elementos característicos que le dan sentido (personajes míticos, orígenes del mundo).\n\nLL.3.4.13. Producir escritos de acuerdo con la situación comunicativa, mediante formatos diversos (historieta mitológica).",
            "tema": "Tema: Disfruto de los mitos, Escritura creativa: Historieta mitológica y Evaluación sumativa Unidad 1\n\nMateriales:\nTexto de Lengua 7mo (Páginas 54 a 71)\nPizarra y marcadores\nCartulinas y lápices de colores\nRúbrica de evaluación sumativa",
            "estrategias": [
                ("SEMANA 9 (Semana 4 de la Planificación Microcurricular · 26/10 al 30/10/2026)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("El docente narra brevemente el mito de la creación del Sol y la Luna (pág. 54). Plantea preguntas detonantes: ¿Por qué las civilizaciones antiguas creaban mitos?, ¿qué fenómenos intentaban explicar?, ¿conocen mitos de los pueblos originarios del Ecuador (ej: mito Cañari de la Guacamaya)?", False, 10),
                ("Inserción Educación Cívica, Ética e Integridad:", True, 10),
                ("Valoración de la memoria ancestral, cosmovisión de las culturas andinas y amazónicas, y respeto profundo a la diversidad intercultural del Ecuador.", False, 10),
                ("DUA", True, 10),
                ("Representación: narración oral con modulación de voz e ilustraciones del libro págs. 54-55.\nAcción/expresión: participación oral comentando relatos transmitidos por sus familias.\nMotivación: fascinación por los relatos cosmogónicos y legendarios.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("Lectura guiada y dialogada de los mitos seleccionados en las páginas 56 a 61 (mitos griegos y andinos). Los estudiantes contrastan: ¿En qué se diferencia un mito de una leyenda o un cuento común?, ¿qué poderes poseen los personajes?, ¿cómo intentaban dar sentido al universo?", False, 10),
                ("DUA", True, 10),
                ("Representación: lectura coral y análisis de paratextos míticos.\nAcción/expresión: cuadro comparativo entre mito y realidad.\nMotivación: intercambio de interpretaciones personales y juicios estéticos.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("Construcción de un organizador gráfico en la pizarra:\n1. Definición de Mito.\n2. Tipos de mitos: cosmogónicos, teogónicos y etiológicos.\n3. Elementos del cómic o historieta (viñetas, bocadillos de diálogo, cartelas, onomatopeyas) pág. 68.\nRegistro en el cuaderno con esquemas creativos ilustrados.", False, 10),
                ("DUA", True, 10),
                ("Representación: esquema conceptual y modelos visuales de viñetas de cómic.\nAcción/expresión: síntesis teórica en cuaderno y bocetos de personajes.\nMotivación: integración de la plástica visual con el lenguaje literario.", False, 10),
                ("APLICACIÓN", True, 11),
                ("Taller creativo: En parejas, los estudiantes adaptan un mito a una historieta ilustrada de 4 a 6 viñetas (págs. 68-69).\nAplicación de la Evaluación Sumativa estructurada de la Unidad 1 (págs. 70-71) para consolidar los aprendizajes del primer bloque temático.", False, 10),
                ("DUA", True, 10),
                ("Representación: rúbrica de autoevaluación pág. 71 y formatos de historietas.\nAcción/expresión: creación plástica y textual, resolución de prueba sumativa.\nMotivación: exposición en el mural del aula (\"Galería Mitológica\").", False, 10),
            ],
            "indicador": "Reconoce en textos de la literatura oral y escrita los elementos característicos que les dan sentido; recrea textos literarios mediante adaptaciones creativas (historietas). I.LL.3.7.1.",
            "evaluacion": "Técnica: Análisis de producciones artísticas/escritas y prueba sumativa.\nInstrumento: Rúbrica de la Historieta Mitológica y Prueba sumativa de Unidad 1 (págs. 70-71)."
        },
        {
            # SEMANA 10 (Semana 5 de la Microcurricular)
            "row_idx": 13,
            "dcd": "LL.3.1.2. Indagar sobre las influencias lingüísticas y culturales que explican los dialectos del castellano en el Ecuador, y valorar la diversidad lingüística del país.",
            "tema": "Tema: Unidad 2: Los dialectos del Ecuador (variedades lingüísticas de Costa, Sierra y Amazonía) y Quichuismos\n\nMateriales:\nTexto de Lengua 7mo (Páginas 74 a 79)\nPizarra y marcadores\nMapa lingüístico del Ecuador\nGuía Docente pág. 15",
            "estrategias": [
                ("SEMANA 10 (Semana 5 de la Planificación Microcurricular · 02/11 al 06/11/2026)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("El docente inicia la clase reproduciendo expresiones y giros lingüísticos de diferentes regiones del país (\"¡Qué bacán la caleta!\", \"Pásame la llacta, guagua\", \"Vamos a la chagra\"). Plantea preguntas: ¿Por qué en la Costa, en la Sierra y en la Amazonía hablamos el castellano con distintos tonos y palabras?, ¿qué palabras quichuas usamos a diario en Quito?", False, 10),
                ("Inserción Educación Cívica, Ética e Integridad:", True, 10),
                ("Convivencia armónica, erradicación de la discriminación por motivos de acento o procedencia geográfica, y orgullo de nuestra identidad ecuatoriana pluricultural.", False, 10),
                ("DUA", True, 10),
                ("Representación: dramatización de diálogos regionales e ilustraciones del mapa pág. 74.\nAcción/expresión: los estudiantes identifican giros lingüísticos que escuchan en su entorno.\nMotivación: conexión con sus raíces familiares e identidad comunitaria.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("Análisis de la lectura del texto págs. 75-76 sobre cómo la historia y las lenguas originarias enriquecieron el castellano. Reflexión guiada: ¿Un dialecto es un castellano \"mal hablado\"? (Superación de prejuicios lingüísticos). ¿Cómo términos como achachay, arrarray, guambra o chapa enriquecen nuestro lenguaje?", False, 10),
                ("DUA", True, 10),
                ("Representación: glosario de términos quichuas incorporados al castellano.\nAcción/expresión: discusión socrática en asamblea de aula.\nMotivación: valoración de los saberes ancestrales y respeto intercultural.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("Explicación docente sobre las variedades diatópicas del castellano ecuatoriano.\nElaboración en la pizarra de un organizador gráfico (matriz comparativa regional):\n1. Dialecto costeño (características fonéticas y léxicas).\n2. Dialecto serrano (influencia quichua, entonación).\n3. Dialecto amazónico (influencia shuar, kichwa y lenguas locales).\n4. Decálogo del respeto a la diversidad lingüística del Ecuador.\nLos estudiantes copian y enriquecen la matriz en sus cuadernos.", False, 10),
                ("DUA", True, 10),
                ("Representación: mapa conceptual visual con el mapa del Ecuador y flechas regionales.\nAcción/expresión: transcripción creativa en el cuaderno.\nMotivación: sentido de pertenencia y riqueza cultural nacional.", False, 10),
                ("APLICACIÓN", True, 11),
                ("Resolución de las actividades de indagación de las páginas 77 a 79 del texto escolar.\nCreación en parejas del \"Mini Diccionario Ilustrado de Ecuatorianismos\", definiendo 5 palabras regionales con su significado y una oración contextualizada. Puesta en común en plenaria. (Nota: Lunes 02 y Martes 03 de Noviembre: Feriados por Día de Difuntos e Independencia de Cuenca).", False, 10),
                ("DUA", True, 10),
                ("Representación: fichas léxicas ilustradas y preguntas del texto.\nAcción/expresión: elaboración manual del mini diccionario y exposición oral.\nMotivación: trabajo cooperativo formativo y lúdico.", False, 10),
            ],
            "indicador": "Indaga sobre las influencias lingüísticas y culturales que explican los diferentes dialectos del castellano en el Ecuador, valorando la diversidad lingüística como factor de identidad. I.LL.3.1.2.",
            "evaluacion": "Técnica: Análisis del desempeño y observación directa.\nInstrumento: Mini Diccionario de Ecuatorianismos evaluado con escala estimativa y actividades del texto págs. 77-79."
        },
        {
            # SEMANA 11 (Semana 6 de la Microcurricular)
            "row_idx": 14,
            "dcd": "LL.3.2.2. Proponer intervenciones orales con una intención comunicativa, organizar el discurso según las estructuras de la lengua oral y utilizar vocabulario adecuado.",
            "tema": "Tema: Unidad 2: Diálogo formal, escucha activa, búsqueda de acuerdos y Taller: Participemos de un cine-foro\n\nMateriales:\nTexto de Lengua 7mo (Páginas 80 a 87)\nPizarra y marcadores\nCortometraje o fragmento audiovisual formativo\nProyector / Audio\nCuadernos",
            "estrategias": [
                ("SEMANA 11 (Semana 6 de la Planificación Microcurricular · 09/11 al 13/11/2026)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("El docente ambienta el aula como una sala de cine-foro (\"Cine-Club Calderón 2\"). Se proyecta un cortometraje breve sobre resolución pacífica de conflictos y empatía escolar. Plantea preguntas exploratorias: ¿Qué mensaje nos transmite la historia?, ¿cómo reaccionaron los personajes frente a las diferencias?, ¿es fácil escuchar a quien piensa distinto?", False, 10),
                ("Inserción Educación Socioemocional y Cívica:", True, 10),
                ("Autorregulación emocional, empatía, escucha activa sin interrupciones y búsqueda de consensos pacíficos a través del diálogo civilizado.", False, 10),
                ("DUA", True, 10),
                ("Representación: cortometraje proyectado con audio claro y subtítulos.\nAcción/expresión: respuestas orales libres expresando sentimientos e ideas.\nMotivación: ambiente lúdico del cine-foro y relevancia cotidiana.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("El docente guía el análisis de la escena y preguntas de las páginas 80 y 81 del libro. Los estudiantes reflexionan sobre los elementos del diálogo constructivo: contacto visual, modulación de la voz, respeto a los turnos de habla y fórmulas de cortesía. Contrastan situaciones de imposición frente a acuerdos dialogados.", False, 10),
                ("DUA", True, 10),
                ("Representación: preguntas orientadoras en pizarra y texto.\nAcción/expresión: debate moderado por turnos.\nMotivación: aula como espacio seguro de opinión respetuosa.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("El docente explica la metodología del Cine-foro y las fases de un diálogo formal o debate (págs. 82-84):\n1. Rol del moderador y del auditorio.\n2. Argumentos fundamentados frente a opiniones sin sustento.\n3. Estructura de la intervención oral (saludo, postura, argumento y propuesta de acuerdo).\nConstrucción en pizarra de un organizador gráfico (esquema del Diálogo Asertivo y Cine-foro) que los estudiantes registran en sus cuadernos.", False, 10),
                ("DUA", True, 10),
                ("Representación: esquema conceptual con recomendaciones para hablar en público.\nAcción/expresión: completar esquema y redactar fichas de argumentos breves.\nMotivación: desarrollo de habilidades de oratoria y seguridad personal.", False, 10),
                ("APLICACIÓN", True, 11),
                ("Desarrollo del taller práctico \"Participemos de un Cine-foro\" (págs. 86-87). Los estudiantes participan por turnos expresando su juicio crítico sobre el cortometraje, aplicando las fórmulas de cortesía aprendidas y consensuando un acuerdo grupal de convivencia. Coevaluación mediante rúbrica de expresión oral. (Cierre oficial de la Planificación Microcurricular: 13 de Noviembre de 2026).", False, 10),
                ("DUA", True, 10),
                ("Representación: rúbrica visual de autoevaluación y coevaluación oral.\nAcción/expresión: intervención oral en el cine-foro y registro de compromisos.\nMotivación: celebración del cierre del ciclo microcurricular de 6 semanas.", False, 10),
            ],
            "indicador": "Propone intervenciones orales organizadas con intención comunicativa, adapta el vocabulario, respeta turnos de habla y busca acuerdos en cine-foros o debates. I.LL.3.2.2.",
            "evaluacion": "Técnica: Observación directa y registro del desempeño oral.\nInstrumento: Rúbrica de participación en el Cine-foro y Diálogo formal (actividades pág. 87)."
        },
        {
            # SEMANA 12 (Semana de Proyecto Integrador)
            "row_idx": 15,
            "dcd": "LL.3.4.13. Producir textos de diversa índole de acuerdo con la situación comunicativa, mediante formatos y recursos diversos.\n\nLL.3.2.2. Proponer intervenciones orales estructuradas para sustentar el producto del proyecto interdisciplinar.",
            "tema": "Tema: EJECUCIÓN, CONSOLIDACIÓN Y SOCIALIZACIÓN DEL PROYECTO INTEGRADOR DEL 1ER TRIMESTRE (1 SEMANA)\n\nMateriales:\nGuía oficial del Proyecto Interdisciplinario\nRúbricas ministeriales de evaluación\nMateriales para estands del aula, dípticos y carteles ilustrados",
            "estrategias": [
                ("SEMANA 12 (Semana de Proyecto Integrador · 16/11 al 23/11/2026 · 1 Semana)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("El docente activa la fase culminante del Proyecto Integrador del 1er Trimestre. Los equipos de trabajo revisan las metas planteadas y el borrador de su producto final. Preguntas de orientación: ¿Cómo nuestro producto final responde a la necesidad comunitaria planteada?, ¿qué saberes de Lengua, Matemática, Ciencias y Estudios Sociales articulamos?", False, 10),
                ("Inserción Educación Cívica y Trabajo Colaborativo:", True, 10),
                ("Liderazgo compartido, equidad de género en los roles de equipo, solidaridad y responsabilidad ética en la entrega de compromisos académicos.", False, 10),
                ("DUA", True, 10),
                ("Representación: rúbrica ministerial del proyecto en cartelera.\nAcción/expresión: verificación en equipo mediante lista de chequeo de avances.\nMotivación: entusiasmo por la culminación del proyecto tangible.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("Talleres de coevaluación y retroalimentación entre equipos de estudiantes. Análisis reflexivo sobre los desafíos encontrados durante la investigación y cómo se resolvieron. Revisión de la presentación visual y ortografía del producto final.", False, 10),
                ("DUA", True, 10),
                ("Representación: lista de cotejo coevaluativa de 5 criterios.\nAcción/expresión: diálogo constructivo entre pares y sugerencias de mejora.\nMotivación: cultura de retroalimentación formativa y superación continua.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("Estructuración del guion formal de exposición del proyecto:\n1. Saludo protocolario e introducción del problema investigado.\n2. Explicación del proceso y articulación de asignaturas.\n3. Demostración práctica del producto final.\n4. Conclusiones y recomendaciones para la comunidad escolar.", False, 10),
                ("DUA", True, 10),
                ("Representación: esquema guía para la sustentación oral.\nAcción/expresión: elaboración de fichas síntesis para los expositores del grupo.\nMotivación: seguridad y dominio del tema frente al público.", False, 10),
                ("APLICACIÓN", True, 11),
                ("Feria de Aula y Socialización del Proyecto Integrador frente a docentes, directivos y compañeros. Cada equipo sustenta su producto aplicando las habilidades de comunicación oral y escrita adquiridas. Entrega oficial del informe y producto final. (Cierre del Proyecto Integrador: 23 de Noviembre de 2026).", False, 10),
                ("DUA", True, 10),
                ("Representación: estands, carteles, trípticos y medios audiovisuales.\nAcción/expresión: exposición oral pública y demostración práctica.\nMotivación: reconocimiento de la comunidad educativa y valoración del esfuerzo colectivo.", False, 10),
            ],
            "indicador": "Publica y socializa los resultados de proyectos integradores mediante textos estructurados e intervenciones orales claras, coherentes y fundamentadas. I.LL.3.6.4.",
            "evaluacion": "Técnica: Evaluación auténtica de proyectos y observación estructurada.\nInstrumento: Rúbrica oficial ministerial del Proyecto Integrador (30% sumativo del proyecto)."
        },
        {
            # SEMANA 13 (Semana de Exámenes Trimestrales)
            "row_idx": 16,
            "dcd": "Evaluación sumativa de las destrezas con criterio de desempeño priorizadas trabajadas en la Unidad 1 y Bloque inicial de Unidad 2 de Lengua y Literatura para 7mo EGB.",
            "tema": "Tema: SEMANA DE EVALUACIONES DEL PRIMER TRIMESTRE Y CIERRE ACADÉMICO (1 SEMANA)\n\nMateriales:\nCuestionarios de base estructurada aprobados por Vicerrectorado\nHojas de respuestas y rúbricas de calificación\nRegistros de notas en el Sistema SIEE",
            "estrategias": [
                ("SEMANA 13 (Semana de Exámenes del Primer Trimestre · 24/11 al 30/11/2026 · 1 Semana)", True, 11),
                ("ERCA / DUA", True, 10),
                ("EXPERIENCIA", True, 11),
                ("Generación de un ambiente de serenidad, confianza y concentración en el aula. Dinámica breve de respiración y relajación previa a la entrega de las evaluaciones. El docente explica las instrucciones generales y los tiempos asignados.", False, 10),
                ("Inserción Educación Socioemocional y Ética:", True, 10),
                ("Gestión de la ansiedad frente a evaluaciones, honestidad individual e integridad académica al responder las pruebas sin recurrir al plagio.", False, 10),
                ("DUA", True, 10),
                ("Representación: lectura pausada y modelada de las instrucciones de la prueba.\nAcción/expresión: preguntas aclaratorias de los estudiantes antes de iniciar.\nMotivación: refuerzo de la autoconfianza y mentalidad de crecimiento.", False, 10),
                ("REFLEXIÓN", True, 11),
                ("Los estudiantes realizan una lectura selectiva de los ítems de la evaluación, planificando el orden de resolución y recordando las estrategias cognitivas trabajadas durante el trimestre.", False, 10),
                ("DUA", True, 10),
                ("Representación: formato de examen claro, tipografía legible y espaciado adecuado.\nAcción/expresión: subrayado de palabras clave en las consignas de la prueba.\nMotivación: autonomía en la gestión del tiempo de examen.", False, 10),
                ("CONCEPTUALIZACIÓN", True, 11),
                ("Consolidación y evocación de los aprendizajes desarrollados en el trimestre: identificación de artículos científicos, el verbo y sus accidentes, modos verbales, puntuación con punto y coma, el mito, los dialectos del Ecuador y las normas del diálogo formal.", False, 10),
                ("DUA", True, 10),
                ("Representación: ítems diversificados (selección múltiple, relación de columnas, completación, análisis sintáctico y producción breve).\nAcción/expresión: resolución estructurada del instrumento sumativo.\nMotivación: oportunidad de evidenciar el progreso académico trimestral.", False, 10),
                ("APLICACIÓN", True, 11),
                ("Desarrollo autónomo de la evaluación sumativa trimestral (del 24 al 30 de noviembre). Entrega de la prueba, revisión inicial de aciertos y errores, retroalimentación individual formativa y consolidación de calificaciones trimestrales (70% formativo + 30% sumativo) para las Juntas de Curso del 01 de Diciembre.", False, 10),
                ("DUA", True, 10),
                ("Representación: escala cuantitativa y cualitativa ministerial visible.\nAcción/expresión: resolución completa del examen y entrega ordenada.\nMotivación: cierre exitoso del 1er Trimestre escolar.", False, 10),
            ],
            "indicador": "Demuestra el dominio de los indicadores de evaluación prioritarios de Lengua y Literatura del subnivel Media: I.LL.3.1.2., I.LL.3.2.2., I.LL.3.3.2., I.LL.3.6.1., I.LL.3.6.3., I.LL.3.7.1.",
            "evaluacion": "Técnica: Pruebas estructuradas formales.\nInstrumento: Cuestionario trimestral de base estructurada aprobado por Vicerrectorado (30% sumativo institucional)."
        }
    ]

    for item in weeks_data:
        r = table.rows[item["row_idx"]]
        
        # Col 0 (DCD)
        r.cells[0].text = item["dcd"]
        
        # Col 2 (Tema y Recursos)
        r.cells[2].text = item["tema"]
        
        # Col 6 (Estrategias ERCA / DUA)
        c6 = r.cells[6]
        c6.text = "" # limpiar
        for text, is_bold, size_pt in item["estrategias"]:
            p = c6.add_paragraph()
            run = p.add_run(text)
            run.bold = is_bold
            run.font.name = "Times New Roman"
            if size_pt:
                run.font.size = Pt(size_pt)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(1)
        
        # Col 13 (Indicador)
        r.cells[13].text = item["indicador"]
        
        # Col 19 (Actividades Evaluativas)
        r.cells[19].text = item["evaluacion"]

    # Guardar documento
    output_filename = "PLANIFICACION_MICROCURRICULAR_7MO_LENGUA_1ER_TRIMESTRE.docx"
    doc.save(output_filename)
    print(f"Documento completo guardado exitosamente: {output_filename}")

if __name__ == "__main__":
    create_7mo_plan()
