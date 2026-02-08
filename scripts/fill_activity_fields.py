"""
Fill empty Activity objective and Activity description in Actividades TBox 25-26.csv.
Sets Source to "Cursor" for rows that are filled.
Uses proper CSV parsing to handle quoted fields with commas and newlines.
"""
import csv
import os
import sys

# Column indices (0-based)
COL_GRADE = 0
COL_PROJECT_NUM = 1
COL_PROJECT_NAME = 2
COL_TOOL = 3
COL_ACTIVITY_NUM = 4
COL_ACTIVITY_NAME = 5
COL_OBJECTIVE = 6
COL_DESCRIPTION = 7
COL_DURATION = 8
COL_SOURCE = 9

# Spanish activity name keywords
SPANISH_ACTIVITY_NAMES = {
    "investigar", "explorar", "construir", "aplicar", "aprender",
    "diseño", "más que mil", "ciudadanía", "primero", "pasos",
}
# Spanish project name patterns (lowercased substrings)
SPANISH_PROJECT_PATTERNS = [
    "diseño web", "más que mil", "ciudadanía digital", "primero pasos",
    "marketing e inteligencia", "animaciones", "eco apps",
]


def is_spanish(project_name: str, activity_name: str) -> bool:
    """Determine if the row content should be in Spanish."""
    p = (project_name or "").strip().lower()
    a = (activity_name or "").strip().lower()
    if any(pat in p for pat in SPANISH_PROJECT_PATTERNS):
        return True
    if any(k in a for k in SPANISH_ACTIVITY_NAMES):
        return True
    # Grade DC3 is Spanish
    return False


def _obj(*lines: str) -> str:
    """Join objective lines with newlines."""
    return "\n".join(lines)


def _get_max_activity_per_project(rows: list) -> dict:
    """Return dict (grade, project_num, project_name) -> max activity number."""
    key_to_max = {}
    for row in rows:
        if len(row) <= COL_ACTIVITY_NUM:
            continue
        try:
            anum = int(row[COL_ACTIVITY_NUM])
        except (ValueError, TypeError):
            continue
        key = (row[COL_GRADE].strip(), row[COL_PROJECT_NUM].strip(), row[COL_PROJECT_NAME].strip())
        key_to_max[key] = max(key_to_max.get(key, 0), anum)
    return key_to_max


def predict_objective(
    tool: str,
    activity_num: int,
    max_activity: int,
    activity_name: str,
    project_name: str,
    lang: str,
) -> str:
    """Generate Activity objective following TBox style (bullet-style learning outcomes)."""
    is_first = activity_num == 1
    is_last = activity_num >= max_activity

    # Tool-normalize for template lookup
    t = (tool or "").strip()
    aname = (activity_name or "").strip().lower()
    pname = (project_name or "").strip()
    is_build = aname in ("construir", "build")

    if lang == "es":
        if is_build and not is_last:
            return _obj(
                "Integrar lo aprendido en un producto o proyecto.",
                "Aplicar las funciones de la herramienta al tema del proyecto.",
            )
        if is_first:
            return _obj(
                "Reconocer las principales opciones de una plataforma educativa.",
                "Usar una fuente electrónica para buscar información.",
                "Describir conceptos relacionados con el tema del proyecto.",
            )
        if is_last:
            return _obj("Poner en práctica lo aprendido.")
        return _obj(
            "Explorar las herramientas del programa.",
            "Aplicar las funciones aprendidas al tema del proyecto.",
        )

    # English
    if is_build and not is_last:
        return _obj(
            "Integrate what has been learned into a product or project.",
            "Apply the tool features to the project theme.",
        )
    if "word 365" in t.lower():
        if is_first:
            return _obj(
                "List the main tools of a word processor.",
                "Describe the interface of a word processor.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Enter and format text in a document.",
            "Insert and modify elements such as tables or images.",
        )

    if "excel 365" in t.lower():
        if is_first:
            return _obj(
                "Describe the main elements of a spreadsheet.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Enter data in a spreadsheet.",
            "Use formulas and format cells.",
            "Create and customize charts.",
        )

    if "office 365" in t.lower():
        if is_first:
            return _obj(
                "Use an electronic source to access information.",
                "Identify the main options of Office applications.",
            )
        if is_last:
            return _obj("Put into practice what has been learned in the project.")
        return _obj(
            "Enter and format text in a document.",
            "Organize data and use images in documents.",
        )

    if "scratch" in t.lower():
        if is_first:
            return _obj(
                "Use an electronic source to search for information.",
                "Describe the main blocks and options of Scratch.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Add and program sprites and backdrops.",
            "Use control and event blocks to create a project.",
        )

    if "pixlr" in t.lower():
        if is_first:
            return _obj(
                "List the basic tools of an image editor.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Edit and enhance images using the editor tools.",
            "Add text and effects to images.",
        )

    if "access 365" in t.lower():
        if is_first:
            return _obj(
                "Describe what a database is and its uses.",
                "Identify tables, records and fields.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Create and modify tables and relationships.",
            "Build queries, forms and reports.",
        )

    if "canva" in t.lower():
        if is_first:
            return _obj(
                "Describe the main options of a graphic design tool.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Create a design using templates and elements.",
            "Add text, images and export the design.",
        )

    if "soundtrap" in t.lower():
        if is_first:
            return _obj(
                "Describe the main options of an audio editing tool.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Record and mix audio tracks.",
            "Add effects and export the project.",
        )

    if "gamefroot" in t.lower():
        if is_first:
            return _obj(
                "Use an electronic source to search for information about game design.",
                "Describe the main options of the game creation tool.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Design levels and characters.",
            "Program game logic and publish the game.",
        )

    if "data analysis" in t.lower():
        if is_first:
            return _obj(
                "Use an electronic source to search for information.",
                "Describe basic concepts of data analysis.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Explore and organize data.",
            "Build and interpret visualizations.",
        )

    if "powerpoint" in t.lower():
        if is_first:
            return _obj(
                "List the main options of a presentation editor.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Create and format slides.",
            "Insert images, text and apply transitions.",
        )

    if "headliner" in t.lower():
        if is_first:
            return _obj(
                "Describe the main options of a video editing tool.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Share the video with the community.")
        return _obj(
            "Plan and create a storyboard.",
            "Edit video, add narration and music.",
        )

    if "animaker" in t.lower():
        if is_first:
            return _obj(
                "Describe the main options of an animation tool.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Create an animation with characters and scenes.",
            "Add text, voice and export the animation.",
        )

    if "gimp" in t.lower():
        if is_first:
            return _obj(
                "Definir qué es la edición de imágenes." if lang == "es" else "Describe what image editing is.",
                "Explorar el entorno de Gimp." if lang == "es" else "Explore the Gimp environment.",
            )
        if is_last:
            return _obj("Poner en práctica lo aprendido." if lang == "es" else "Put into practice what has been learned.")
        return _obj(
            "Usar herramientas de selección y capas." if lang == "es" else "Use selection and layer tools.",
            "Aplicar efectos y texto a la imagen." if lang == "es" else "Apply effects and text to the image.",
        )

    if "python" in t.lower():
        if is_first:
            return _obj(
                "Describe the main features of Python.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Write and run Python code.",
            "Use variables, control structures and functions.",
        )

    if "clipchamp" in t.lower():
        if is_first:
            return _obj(
                "Describe the main options of a video editing tool.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Import and arrange clips.",
            "Edit video, add titles and export.",
        )

    if "vs code" in t.lower() or "vs code" in t:
        if is_first:
            return _obj(
                "Describe the main options of a code editor.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Create and edit code files.",
            "Use the editor features for CSS and web development.",
        )

    if "oracle live sql" in t.lower() or "sql" in t.lower():
        if is_first:
            return _obj(
                "Definir qué es SQL y para qué se usa." if lang == "es" else "Define what SQL is and its uses.",
                "Explorar el entorno de Oracle Live SQL." if lang == "es" else "Explore the Oracle Live SQL environment.",
            )
        if is_last:
            return _obj("Poner en práctica lo aprendido." if lang == "es" else "Put into practice what has been learned.")
        return _obj(
            "Escribir consultas SELECT, INSERT y UPDATE." if lang == "es" else "Write SELECT, INSERT and UPDATE queries.",
        )

    if "javascript" in t.lower():
        if is_first:
            return _obj(
                "Describe the main features of JavaScript.",
                "Use an electronic source to search for information.",
            )
        if is_last:
            return _obj("Put into practice what has been learned.")
        return _obj(
            "Write and run JavaScript code.",
            "Use variables, functions and control structures.",
        )

    if "web browser" in t.lower() or "chatgpt" in t.lower() or "gemini" in t.lower() or "firefly" in t.lower():
        if is_first:
            return _obj(
                "Usar fuentes en línea para investigar sobre marketing e IA." if lang == "es" else "Use online sources to research marketing and AI.",
            )
        if is_last:
            return _obj("Poner en práctica lo aprendido." if lang == "es" else "Put into practice what has been learned.")
        return _obj(
            "Explorar herramientas de IA para contenido visual y texto." if lang == "es" else "Explore AI tools for visual and text content.",
        )

    # Moqups, Squarespace, Notepad++, Bootstrap (web design)
    if any(x in t.lower() for x in ["moqups", "squarespace", "notepad++", "bootstrap"]):
        if is_first:
            return _obj(
                "Explorar el entorno de la herramienta de diseño web." if lang == "es" else "Explore the web design tool environment.",
            )
        if is_last:
            return _obj("Poner en práctica lo aprendido." if lang == "es" else "Put into practice what has been learned.")
        return _obj(
            "Crear y modificar elementos de una página web." if lang == "es" else "Create and modify web page elements.",
        )

    # Default
    if is_first:
        return _obj(
            "Use an electronic source to search for information.",
            "Describe the main options of the tool.",
        )
    if is_last:
        return _obj("Put into practice what has been learned.")
    return _obj(
        "Explore the tool and apply its features to the project theme.",
    )


# Project-specific description overrides (key: project name lowercased substring)
_PROJECT_DESCRIPTIONS_EN = {
    "nature protectors": {
        "first": "Students use the website of the project to research about nature protection and recycling. In addition, they learn the main tools of a word processor to create documents about the environment.",
        "last": "Students participate in a contest to share their literary pieces about nature. They put into practice what they have learned with the word processor and present their work to their classmates.",
        "middle": "Students create literary pieces about nature protection using the word processor. They write poems, riddles, tongue twisters or songs and format their documents with the tools learned.",
    },
    "create your own zoo": {
        "first": "Students research about animals and zoo management. In addition, they learn the main elements of a spreadsheet to organize data.",
        "last": "Students present their virtual zoo project. They use the spreadsheet to show the animals, budget and statistics they have worked on throughout the project.",
        "middle": "Students build their virtual zoo using a spreadsheet. They organize animal data, use functions, create charts and prepare a budget as they progress through the activities.",
    },
    "let's explore operating systems": {
        "first": "Students use the website of the project to learn about operating systems. They identify the main options of a word processor and how to search for information.",
        "last": "Students create a quiz game about operating systems. They put into practice what they have learned with the word processor and share it with their classmates.",
    },
    "culture hunters": {
        "first": "Students research about cultural heritage and tourism. In addition, they learn how to use Office applications to organize and present information.",
        "last": "Students participate in a tourism day to share what they have discovered about cultural heritage. They use the materials created with Office 365 during the project.",
        "middle": "Students work as culture hunters using Office 365. They exchange clues, design a brochure and discover cultural heritage while applying the tools learned.",
    },
    "artistic show": {
        "first": "Students research about music and musical instruments. They learn the main blocks and options of Scratch to create an animated project.",
        "last": "Students enjoy the artistic show and share their Scratch projects about music. They put into practice what they have learned and present their work.",
        "middle": "Students use Scratch to create an artistic show about musical instruments. They program percussion, wind and string instruments and build their own show.",
    },
    "eath healthy": {
        "first": "Students research about nutrition and healthy eating. They learn the main elements of a spreadsheet to organize nutritional data.",
        "last": "Students create and share an online survey about healthy habits. They use the spreadsheet to analyze the results and put into practice what they have learned.",
        "middle": "Students use a spreadsheet to explore energy, protein and carbohydrate requirements. They analyze the nutritional value of dishes and plan a healthy diet.",
    },
    "healthy families": {
        "first": "Students research about healthy habits for families. In addition, they learn the basic tools of an image editor.",
        "last": "Students create an animated image about healthy living. They use the image editor to put into practice what they have learned and share tips with their families.",
        "middle": "Students use the image editor to create materials about sleep, water, sun protection, exercise and important health advice for families.",
    },
    "innovations to solve problems": {
        "first": "Students research about innovations that solve everyday problems. They learn the main options of Scratch to create interactive projects.",
        "last": "Students keep programming and share their innovation projects. They put into practice what they have learned with Scratch and present their solutions.",
        "middle": "Students use Scratch to explore innovations such as technology in the kitchen, the evolution of phones, Hawk-eye, solar energy and helping their community.",
    },
    "technology for a better life": {
        "first": "Students learn about databases and how they support innovation. They identify tables, records and fields in Access 365.",
        "last": "Students create a tutorial about innovations using the database. They put into practice what they have learned and share their work.",
        "middle": "Students work with Access 365 to manage records, use tables, create a trivia on innovations, prepare reports and explore innovations in the future.",
    },
    "tourist destinations": {
        "first": "Students learn about email and tourist destinations. They explore the main options of Canva for creating graphic materials.",
        "last": "Students publish their designs on the web and share them. They put into practice what they have learned with Canva for the tourist destinations project.",
        "middle": "Students use Canva to create a magazine article, tourism poster, infographic, and other materials about tourist attractions and city treasures.",
    },
    "the world of music": {
        "first": "Students research about music and audio production. They learn the main options of Soundtrap for creating podcasts.",
        "last": "Students create and share educational podcasts about music. They put into practice what they have learned with Soundtrap.",
        "middle": "Students use Soundtrap to create their first podcast, mix audios, add effects, publish episodes and explore traditional music.",
    },
    "programming your video game": {
        "first": "Students research about video games and game design. They learn the main options of Gamefroot to create their own game.",
        "last": "Students present their educational video game. They put into practice what they have learned and share their game with others.",
        "middle": "Students use Gamefroot to plan screens, design the main character, add score and time, publish the game and create another level.",
    },
    "mobile fun": {
        "first": "Students learn more about databases and their uses. They explore Access 365 to create a database related to mobile applications or games.",
        "last": "Students demonstrate how the database works and share their project. They put into practice what they have learned with Access 365.",
        "middle": "Students work with Access 365 to create tables, relate them with queries, create forms, design reports and build a video game ratings database.",
    },
    "data here, data there": {
        "first": "Students research about data analysis and its importance. They learn basic concepts and tools for exploring data.",
        "last": "Students apply what they have learned to a real data analysis task. They put into practice the skills developed in the project.",
        "middle": "Students explore data sets and build visualizations or analyses using the data analysis tools introduced in the project.",
    },
    "digital marketing": {
        "first": "Students research about digital marketing. They learn the main options of PowerPoint and how to search for information.",
        "last": "Students share the digital marketing materials they have created. They put into practice what they have learned and present their work.",
        "middle": "Students use PowerPoint to create presentations on basic concepts, online and offline digital marketing, social media facts and marketing materials.",
    },
    "red card to bullying": {
        "first": "Students learn more about video editing and its role in raising awareness. They explore Headliner and research the topic.",
        "last": "Students participate in the Red Card event and share their promotional videos. They put into practice what they have learned about video and the campaign.",
        "middle": "Students use Headliner to create a storyboard, gather materials, edit the video, add narration and music, and create a promotional video against bullying.",
    },
    "animated digital citizenship": {
        "first": "Students learn about digital animation and digital citizenship topics. They explore Animaker and its main options.",
        "last": "Students create an animated science or citizenship piece and share it. They put into practice what they have learned with Animaker.",
        "middle": "Students use Animaker to create short animations about netiquette, ergonomics, shopping online, security tips and digital reputation.",
    },
}


def predict_description(
    tool: str,
    activity_num: int,
    max_activity: int,
    activity_name: str,
    project_name: str,
    lang: str,
) -> str:
    """Generate Activity description (narrative, third person, project- and tool-specific)."""
    is_first = activity_num == 1
    is_last = activity_num >= max_activity
    aname = (activity_name or "").strip()
    pname = (project_name or "").strip()
    t = (tool or "").strip()
    pname_lower = pname.lower()

    # Project-specific (English)
    if lang == "en":
        for key, descs in _PROJECT_DESCRIPTIONS_EN.items():
            if key in pname_lower:
                if is_first and "first" in descs:
                    return descs["first"]
                if is_last and "last" in descs:
                    return descs["last"]
                if not is_first and not is_last and "middle" in descs:
                    return descs["middle"]
                break

    if lang == "es":
        if is_first:
            return (
                f"En esta actividad los estudiantes utilizan el sitio web del proyecto para investigar sobre el tema. "
                f"Además, se familiarizan con las opciones principales de la herramienta {t}."
            )
        if is_last:
            return (
                f"Los estudiantes ponen en práctica lo aprendido en el proyecto «{pname}». "
                f"Utilizan la herramienta {t} para completar un producto final relacionado con la temática."
            )
        return (
            f"En esta actividad los estudiantes trabajan con {t} en el marco del proyecto «{pname}». "
            f"Realizan la actividad «{aname}» aplicando las funciones de la herramienta al contenido del proyecto."
        )
    # English generic
    if is_first:
        return (
            f"Students use the website of the project to research about the topic. "
            f"In addition, they learn about the main options of {t} and how it supports the project theme."
        )
    if is_last:
        return (
            f"Students put into practice what they have learned in the project. "
            f"They use {t} to create or share a final product related to «{pname}»."
        )
    return (
        f"Students work with {t} as part of the project «{pname}». "
        f"They complete the activity «{aname}» using the tool to apply their skills to the project theme."
    )


def predict_row(
    grade: str,
    project_num: str,
    project_name: str,
    tool: str,
    activity_num: int,
    activity_name: str,
    max_activity: int,
) -> tuple[str, str]:
    """Return (objective, description) for a row with empty fields."""
    lang = "es" if is_spanish(project_name, activity_name) else "en"
    try:
        anum = int(activity_num)
    except (ValueError, TypeError):
        anum = 1
    obj = predict_objective(tool, anum, max_activity, activity_name, project_name, lang)
    desc = predict_description(tool, anum, max_activity, activity_name, project_name, lang)
    return obj, desc


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(script_dir)
    csv_path = os.path.join(repo_root, "Actividades TBox 25-26.csv")
    if not os.path.isfile(csv_path):
        print(f"File not found: {csv_path}", file=sys.stderr)
        sys.exit(1)

    with open(csv_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        print("CSV is empty.", file=sys.stderr)
        sys.exit(1)

    header = rows[0]
    if len(header) < 10:
        print("CSV does not have expected columns.", file=sys.stderr)
        sys.exit(1)

    max_activity = _get_max_activity_per_project(rows)
    filled_count = 0

    for i in range(1, len(rows)):
        row = rows[i]
        if len(row) <= COL_SOURCE:
            continue
        objective = (row[COL_OBJECTIVE] or "").strip()
        description = (row[COL_DESCRIPTION] or "").strip()
        if objective != "" or description != "":
            continue
        grade = row[COL_GRADE].strip()
        project_num = row[COL_PROJECT_NUM].strip()
        project_name = row[COL_PROJECT_NAME].strip()
        tool = row[COL_TOOL].strip()
        activity_num = row[COL_ACTIVITY_NUM].strip()
        activity_name = row[COL_ACTIVITY_NAME].strip()
        key = (grade, project_num, project_name)
        max_act = max_activity.get(key, 7)
        try:
            anum = int(activity_num)
        except (ValueError, TypeError):
            anum = 1
        obj, desc = predict_row(
            grade, project_num, project_name, tool, anum, activity_name, max_act
        )
        row[COL_OBJECTIVE] = obj
        row[COL_DESCRIPTION] = desc
        row[COL_SOURCE] = "Cursor"
        filled_count += 1

    # Second pass: fix "Construir"/"Build" objectives for rows we filled (Source already Cursor)
    for i in range(1, len(rows)):
        row = rows[i]
        if len(row) <= COL_SOURCE or (row[COL_SOURCE] or "").strip() != "Cursor":
            continue
        activity_name = (row[COL_ACTIVITY_NAME] or "").strip().lower()
        if activity_name not in ("construir", "build"):
            continue
        grade = row[COL_GRADE].strip()
        project_num = row[COL_PROJECT_NUM].strip()
        project_name = row[COL_PROJECT_NAME].strip()
        key = (grade, project_num, project_name)
        max_act = max_activity.get(key, 10)
        try:
            anum = int((row[COL_ACTIVITY_NUM] or "").strip())
        except (ValueError, TypeError):
            continue
        if anum >= max_act:
            continue
        lang = "es" if is_spanish(project_name, row[COL_ACTIVITY_NAME] or "") else "en"
        if lang == "es":
            row[COL_OBJECTIVE] = _obj(
                "Integrar lo aprendido en un producto o proyecto.",
                "Aplicar las funciones de la herramienta al tema del proyecto.",
            )
        else:
            row[COL_OBJECTIVE] = _obj(
                "Integrate what has been learned into a product or project.",
                "Apply the tool features to the project theme.",
            )

    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        for r in rows:
            writer.writerow(r)

    print(f"Filled {filled_count} rows. Source set to 'Cursor' for those rows.")
    print(f"Updated: {csv_path}")


if __name__ == "__main__":
    main()
