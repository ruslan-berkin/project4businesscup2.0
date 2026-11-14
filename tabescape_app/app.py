import base64
import io
import json
import os
import random
import re
import shutil
import zipfile
import zlib
import xml.etree.ElementTree as ET

import requests
from flask import Flask, render_template, request, jsonify, send_file, Response
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX_NAME = "Книга1.xlsx"
DEFAULT_XLSX_PATH = os.path.join(BASE_DIR, "static", "data", DEFAULT_XLSX_NAME)

GROQ_MODEL = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

SYSTEM_PROMPT = """Ты — игровой ведущий для «TabEscape», увлекательной фэнтезийной текстовой RPG-приключения, которая замаскирована под таблицу Excel.

СТИЛЬ И АТМОСФЕРА (ЛЁГКОЕ И ЖИВОЕ ФЭНТЕЗИ):
- Мир светлый, живой и уютный, с лёгким юмором в духе классических приключений: солнечный тракт, шумный городок Оукхейвен, таверна «Золотой Жёлудь», Гильдия авантюристов, чудаковатые маги, говорящие коты и сонная драконица в очках для чтения.
- ПИШИ ПРОСТО И ЛЕГКО ДЛЯ ЧТЕНИЯ: максимум 2–3 коротких, понятных предложения без перегруженных метафор и мрачного пафоса.

МЕХАНИКА ДЕЙСТВИЙ И ПАРАМЕТРОВ:
У героя три параметра (от 0 до 100):
1. "ROI" (Здоровье / ОЗ): 100 на старте.
2. "KPI" (Бодрость / Мана): 100 на старте.
3. "Anxiety" (Неприятности / Розыск): 10 на старте.

ВАЖНЕЙШЕЕ ПРАВИЛО БАЛАНСА:
- НЕ КАЖДОЕ действие должно менять статы! Обычный выбор направления («Пойти к ярмарке», «Свернуть на лесную тропу») или простой вопрос персонажу («Спросить дорогу у стражника», «Узнать слухи у трактирщицы») НЕ отнимают ни ОЗ, ни Ману (оставляй ROI и KPI без изменений!).
- Отдых или удачная находка немного восстанавливают силы (+10..+15).
- А вот безрассудные глупости (прыгнуть в водоворот в броне, пнуть стражника, выпить подозрительное зелье, ткнуть дракона) бьют ОЧЕНЬ больно (-45..-55 ОЗ), чтобы умереть от глупых решений или слепого спама было легко!
- При ROI <= 0, KPI <= 0 или Anxiety >= 100 ставь "game_over": true и возвращай пустой список options: [].

ПРАВИЛО ВАРИАНТОВ (1, 2, 3):
- Давай ровно 3 варианта в "options".
- Меняй их роли местами на каждом ходу! Не делай 1-й вариант всегда хорошим, а 3-й плохим. Пусть среди вариантов будут: спокойный выбор маршрута/вопрос (без траты статов), находчивое действие и явно безрассудная выходка (с большим уроном), расположенные в случайном порядке.
- В "audit_log" пиши короткую забавную псевдоформулу Excel на русском.
- ВЕСЬ ТЕКСТ — ТОЛЬКО НА РУССКОМ ЯЗЫКЕ.

ФОРМАТ ОТВЕТА (строго валидный JSON без markdown):
{
  "narrative": "Короткое и живое описание ситуации (2-3 простых предложения).",
  "stats": {"ROI": 100, "KPI": 100, "Anxiety": 10},
  "options": ["Вариант 1", "Вариант 2", "Вариант 3"],
  "audit_log": "=ЕСЛИ(ВОПРОС_ТРАКТИРЩИКУ;СЛУХИ_ПОЛУЧЕНЫ;ПУТЬ_ОТКРЫТ)",
  "inventory_add": "Название предмета или null",
  "game_over": false
}"""


# --------------------------------------------------------------------------- #
#  LLM
# --------------------------------------------------------------------------- #

def _parse_json(text):
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
    return json.loads(text)


def call_llm(user_prompt, api_key=None):
    gemini_key = api_key if (api_key and api_key.startswith("AIza")) else os.getenv("GEMINI_API_KEY")
    groq_key = api_key if (api_key and api_key.startswith("gsk_")) else os.getenv("GROQ_API_KEY")
    if gemini_key and "your_" in gemini_key:
        gemini_key = None
    if groq_key and "your_" in groq_key:
        groq_key = None

    if groq_key:
        try:
            resp = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {groq_key}", "Content-Type": "application/json"},
                json={
                    "model": GROQ_MODEL,
                    "messages": [
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": user_prompt},
                    ],
                    "temperature": 0.75,
                    "response_format": {"type": "json_object"},
                },
                timeout=12,
            )
            if resp.status_code == 200:
                return _parse_json(resp.json()["choices"][0]["message"]["content"])
        except Exception as e:  # noqa: BLE001
            print(f"Ошибка Groq: {e}")

    if gemini_key:
        try:
            url = (
                f"https://generativelanguage.googleapis.com/v1beta/models/"
                f"{GEMINI_MODEL}:generateContent?key={gemini_key}"
            )
            payload = {
                "system_instruction": {"parts": [{"text": SYSTEM_PROMPT}]},
                "contents": [{"parts": [{"text": user_prompt}]}],
                "generationConfig": {"temperature": 0.75, "responseMimeType": "application/json"},
            }
            resp = requests.post(url, json=payload, timeout=12)
            if resp.status_code == 200:
                return _parse_json(resp.json()["candidates"][0]["content"]["parts"][0]["text"])
        except Exception as e:  # noqa: BLE001
            print(f"Ошибка Gemini: {e}")

    return None


def clamp(v, lo=0, hi=100):
    try:
        v = int(float(v))
    except (TypeError, ValueError):
        v = lo
    return max(lo, min(hi, v))


def normalize_llm(res, prev_stats):
    if not isinstance(res, dict) or "narrative" not in res:
        return None
    st = res.get("stats") or {}
    stats = {
        "ROI": clamp(st.get("ROI", prev_stats["ROI"])),
        "KPI": clamp(st.get("KPI", prev_stats["KPI"])),
        "Anxiety": clamp(st.get("Anxiety", prev_stats["Anxiety"])),
    }
    over = bool(res.get("game_over")) or stats["ROI"] <= 0 or stats["KPI"] <= 0 or stats["Anxiety"] >= 100
    options = [str(o) for o in (res.get("options") or [])][:3]
    while not over and len(options) < 3:
        options.append([
            "Спросить у прохожего дорогу к площади",
            "Свернуть на тропинку к таверне",
            "Прыгнуть через забор в чужой сад",
        ][len(options)])
    item = res.get("inventory_add")
    if not item or str(item).lower() in ("null", "none"):
        item = None
    return {
        "narrative": str(res["narrative"]),
        "stats": stats,
        "options": [] if over else options,
        "audit_log": str(res.get("audit_log") or "=ПУТЬ_ОТКРЫТ(ОУКХЕЙВЕН)"),
        "inventory_add": item,
        "game_over": over,
        "scene_id": None,
    }


# --------------------------------------------------------------------------- #
#  Сюжетный движок: лёгкое приключенческое фэнтези
#  Мирные вопросы и выбор маршрута имеют дельту (0, 0, 0) — не тратят ОЗ и Ману!
#  Опасные глупости наносят большой урон (-45..-55 ОЗ), поэтому погибнуть легко.
#  Формат опции: (текст, результат, (dROI, dKPI, dAnxiety), предмет, следующая_сцена)
# --------------------------------------------------------------------------- #

SCENES = {
    "start": {
        "narrative": "Вы просыпаетесь на мягком сене в повозке, которая бодро катится по солнечной дороге к городку Оукхейвен. Весёлый возница насвистывает мелодию и говорит, что сегодня в городе ярмарка и день приёма в Гильдию искателей. У вас в кармане пара монет и рекомендательное письмо к мастеру {boss}.",
        "audit_log": "=ЕСЛИ(УТРО=\"СОЛНЕЧНОЕ\";НАЧАТЬ_ПРИКЛЮЧЕНИЕ();СПАТЬ)",
        "options": [
            # В прологе 1-й вариант — опасная глупость, а 2-й и 3-й — спокойные действия без потери ОЗ и маны!
            (
                "На полном ходу выпрыгнуть из повозки в придорожные кусты, чтобы срезать путь",
                "За кустами оказался глубокий колючий овраг! Вы кубарем скатились вниз и сильно ушиблись.",
                (-45, -20, 15),
                None,
                "woods",
            ),
            (
                "Расспросить возницу, чем живёт город и куда лучше сходить в первую очередь",
                "Возница с улыбкой советует заглянуть на площадь у фонтана: оттуда легко дойти и до Гильдии, и до таверны.",
                (0, 0, 0),
                None,
                "crossroads",
            ),
            (
                "Спокойно доехать до городских ворот и выйти на главную площадь",
                "Повозка мягко останавливается у резных ворот Оукхейвена, и вы выходите навстречу утренней суете.",
                (0, 0, 0),
                "Письмо в Гильдию",
                "crossroads",
            ),
        ],
    },
    "crossroads": {
        "narrative": "Вы стоите у красивого каменного фонтана в центре Оукхейвена. Налево ведёт улочка к уютной таверне «Золотой Жёлудь», прямо высится дом Гильдии искателей, а направо шумят пёстрые шатры ярмарки.",
        "audit_log": "=ВЫБОР(МАРШРУТ;\"ТАВЕРНА\";\"ГИЛЬДИЯ\";\"ЯРМАРКА\")",
        "options": [
            (
                "Пойти налево в таверну «Золотой Жёлудь» и узнать свежие новости",
                "Вы толкаете дубовую дверь и заходите в прохладный зал таверны.",
                (0, 0, 0),
                None,
                "tavern",
            ),
            (
                "Направиться прямо в Гильдию искателей к доске поручений",
                "Вы поднимаетесь по широким ступеням и входите в светлый зал Гильдии.",
                (0, 0, 0),
                None,
                "guild",
            ),
            (
                "Свернуть направо и прогуляться по рядам утренней ярмарки",
                "Вы ныряете в весёлую толпу между лавками с диковинками и выпечкой.",
                (0, 0, 0),
                None,
                "market",
            ),
        ],
    },
    "tavern": {
        "narrative": "В таверне вкусно пахнет свежим хлебом и печёными яблоками. Рыжеволосая хозяйка Марта протирает кружки за стойкой, у камина отдыхает бывалый следопыт с картой, а в углу хмурый громила спорит с наёмниками.",
        "audit_log": "=ВПР(\"ТАВЕРНА\";НОВОСТИ_ГОРОДА;1;ЛОЖЬ)",
        "options": [
            (
                "Подойти к следопыту у камина и просто спросить, какие места вокруг стоит посетить",
                "Старик охотно рассказывает о говорящем коте в башне мага и дарит вам свою дорожную карту.",
                (0, 0, 0),
                "Карта долины",
                "wizard_tower",
            ),
            (
                "Спросить у хозяйки Марты, нет ли у неё несложного поручения по пути",
                "Марта просит отнести свёрток специй на ярмарку и угощает вас горячим яблочным пирожком.",
                (15, 10, -5),
                "Тёплый пирожок",
                "market",
            ),
            (
                "Пнуть стул хмурого громилы и вызвать всю его компанию на кулачный бой",
                "Три вышибалы одновременно хватают вас за шиворот и выкидывают в окно прямо на бочки!",
                (-50, -25, 25),
                None,
                "alley",
            ),
        ],
    },
    "guild": {
        "narrative": "В зале Гильдии кипит работа: новички рассматривают карты, а мастер {boss} разбирает утренние отчёты. За стойкой улыбчивая девушка-секретарь готова подсказать подходящее задание.",
        "audit_log": "=ЕСЛИ(НОВИЧОК;КОНСУЛЬТАЦИЯ();ВЫБОР_КВЕСТА)",
        "options": [
            (
                "Подойти к секретарю и спросить, с какого простого задания лучше начать",
                "Девушка советует проведать башню мага Корнелиуса у реки — там как раз нужен толковый помощник.",
                (0, 0, 0),
                "Значок Гильдии",
                "wizard_tower",
            ),
            (
                "Посмотреть карту маршрутов на стене и пойти к речной переправе",
                "Вы запоминаете дорогу и выходите из города по живописной тропинке к реке.",
                (0, 0, 0),
                None,
                "river_ferry",
            ),
            (
                "Без подготовки схватить учебный магический посох со стойки и ударить им об пол",
                "Посох детонирует учебной молнией прямо у вас в руках! Половина зала сбегается тушить ваш плащ.",
                (-45, -35, 30),
                None,
                "crossroads",
            ),
        ],
    },
    "market": {
        "narrative": "На ярмарке полно интересного: травница разложила душистые сборы, часовой мастер чинит карманные компасы, а у края площади стоит указатель в сторону Зелёной Рощи.",
        "audit_log": "=СУММ(ТОВАРЫ_ЯРМАРКИ)",
        "options": [
            (
                "Поговорить с травницей и расспросить о полезных растениях долины",
                "Травница бесплатно даёт вам карманную памятку и советует прогуляться к Зелёной Роще.",
                (0, 0, 0),
                "Справочник трав",
                "woods",
            ),
            (
                "Спросить у часового мастера дорогу к старой ветряной мельнице",
                "Мастер указывает на холм за рекой и дарит вам исправный латунный компас.",
                (0, 0, 0),
                "Старинный компас",
                "windmill",
            ),
            (
                "Залпом выпить кипящую синюю колбу с прилавка алхимика без спроса",
                "Это оказался концентрированный клей для подков! Желудок скручивает так, что в глазах темнеет!",
                (-50, -25, 20),
                None,
                "crossroads",
            ),
        ],
    },
    "wizard_tower": {
        "narrative": "У крыльца аккуратной круглой башни на солнышке греется толстый рыжий кот в маленькой шляпе мага. Из открытого окна доносится весёлое посвистывание чародея Корнелиуса.",
        "audit_log": "=ЕСЛИ(КОТ_В_ШЛЯПЕ=ИСТИНА;ВЕЖЛИВОСТЬ();РАЗРЯД)",
        "options": [
            (
                "Поздороваться с котом и вежливо спросить, можно ли войти в башню",
                "Кот поправляет лапой шляпу, человеческим голосом говорит: «Добро пожаловать!» и катит вам амулет.",
                (0, 0, -5),
                "Кошачий амулет",
                "ruins",
            ),
            (
                "Спросить у выглянувшего в окно мага дорогу к речной переправе",
                "Корнелиус машет вам рукой, показывает короткую тропинку к реке и кидает фляжку яблочного тоника.",
                (10, 10, 0),
                "Зелье бодрости",
                "river_ferry",
            ),
            (
                "Схватить кота за хвост и попытаться выбить дверь башни с разбега",
                "Защитная руна на двери вспыхивает и отбрасывает вас через весь двор прямо в колючий крыжовник!",
                (-50, -25, 25),
                None,
                "woods",
            ),
        ],
    },
    "woods": {
        "narrative": "Утоптанная тропинка идёт через светлую берёзовую рощу. На пеньке у развилки сидит странствующий музыкант с лютней, а деревянный указатель показывает две дороги — к Речной переправе и к Тёплому Утёсу.",
        "audit_log": "=ЕСЛИ(РАЗВИЛКА_В_ЛЕСУ;ВЫБРАТЬ_ТРОПУ();ПРИВАЛ)",
        "options": [
            (
                "Остановиться и спросить у музыканта, куда лучше пойти дальше",
                "Музыкант играет бодрый аккорд и советует подняться на Тёплый Утёс — там отличные виды и совсем не страшно.",
                (0, 0, -5),
                None,
                "dragon_hill",
            ),
            (
                "Пойти по нижней лесной дорожке в сторону Речной переправы",
                "Вы спокойно идёте под шелест листвы и вскоре слышите плеск речной воды.",
                (0, 0, 0),
                None,
                "river_ferry",
            ),
            (
                "Залезть на самое тонкое сухое дерево над обрывом, чтобы достать дикий улей",
                "Ветка с треском ломается! Вы падаете с высоты прямо в рой разъярённых диких шершней!",
                (-50, -25, 15),
                None,
                "camp",
            ),
        ],
    },
    "river_ferry": {
        "narrative": "Широкая река искрится на солнце. На пристани добродушный паромщик готовит лодку к отплытию, а чуть правее через реку перекинут крепкий каменный мост, ведущий к белым колоннам старой обсерватории.",
        "audit_log": "=ЕСЛИ(ПЕРЕПРАВА;ПАРОМ_ИЛИ_МОСТ();ОПАСНОСТЬ)",
        "options": [
            (
                "Попросить паромщика перевезти вас на другой берег к ветряной мельнице",
                "Паромщик охотно берёт вас на борт и за пару минут перевозит через тихую заводь.",
                (0, 0, 0),
                None,
                "windmill",
            ),
            (
                "Перейти по каменному мосту и направиться к старой обсерватории на холме",
                "Вы неспешно переходите мост, любуясь рекой, и выходите к саду с белыми колоннами.",
                (0, 0, 0),
                None,
                "ruins",
            ),
            (
                "Разбежаться и прыгнуть с пирса прямо под лопасти водяного колеса",
                "Бурное течение затягивает вас под колесо! Паромщик чудом вылавливает вас багром на мелководье.",
                (-55, -35, 20),
                None,
                "camp",
            ),
        ],
    },
    "windmill": {
        "narrative": "У старой ветряной мельницы пахнет свежей мукой и клевером. Мельник сидит на лавочке и смеётся: маленькие лесные еноты утащили его связку ключей в беседку у сада.",
        "audit_log": "=ЕСЛИ(ЕНОТЫ_С_КЛЮЧАМИ;УГОСТИТЬ();ПОИСК)",
        "options": [
            (
                "Протянуть енотам корочку хлеба и по-доброму обменять её на ключи",
                "Еноты с удовольствием берут угощение, отдают связку, а мельник дарит вам серебряный талисман!",
                (5, 5, -10),
                "Серебряный ключ",
                "ruins",
            ),
            (
                "Спросить у мельника дорогу к Тёплому Утёсу и продолжить прогулку",
                "Мельник показывает удобную тропинку через луг, и вы отправляетесь к предгорьям.",
                (0, 0, 0),
                None,
                "dragon_hill",
            ),
            (
                "Ухватиться руками за огромное вращающееся крыло ветряной мельницы",
                "Крыло поднимает вас на десять метров вверх и с размаху швыряет в стог колючего чертополоха!",
                (-45, -30, 15),
                None,
                "camp",
            ),
        ],
    },
    "ruins": {
        "narrative": "Белые колонны старой Эльфийской Обсерватории утопают в цветах. На скамейке девушка-учёный делает зарисовки солнечных часов, а от сада расходятся две аккуратные дорожки.",
        "audit_log": "=ЕСЛИ(ОБСЕРВАТОРИЯ;ОСМОТР_ЧАСОВ();ВЫБОР_ПУТИ)",
        "options": [
            (
                "Поинтересоваться у девушки-учёного, как работают старинные солнечные часы",
                "Она с радостью показывает секретный поворот стрелки — из тайника выпадает красивая солнечная линза!",
                (0, 0, 0),
                "Линза рассвета",
                "dragon_hill",
            ),
            (
                "Пройти по садовой аллее в сторону уютного привала путников",
                "Дорожка мягко спускается с холма к поляне, где вьётся дымок костра.",
                (0, 0, 0),
                None,
                "camp",
            ),
            (
                "Изо всех сил пнуть тяжёлую каменную колонну у края террасы",
                "Сверху срывается каменный карниз и больно прилетает вам прямо по макушке!",
                (-45, -20, 20),
                None,
                "woods",
            ),
        ],
    },
    "dragon_hill": {
        "narrative": "На согретой солнцем вершине Тёплого Утёса лежит золотистая драконица Аурелия в больших очках для чтения и листает сборник сказок. Рядом с ней стоит самовар с горным чаем и указатель обратно в долину.",
        "audit_log": "=ЕСЛИ(ДРАКОН_ЧИТАЕТ;ВЕЖЛИВЫЙ_РАЗГОВОР();ОЖОГ)",
        "options": [
            (
                "Вежливо поздороваться с драконицей и спросить, интересная ли книга",
                "Аурелия радостно угощает вас чаем с мёдом и дарит на память блестящую золотую чешуйку!",
                (15, 15, -10),
                "Чешуйка удачи",
                "camp",
            ),
            (
                "Полюбоваться видом на долину со смотровой площадки и спуститься к привалу",
                "Вы делаете пару глубоких вдохов горного воздуха и легко спускаетесь по тропинке.",
                (0, 0, 0),
                None,
                "camp",
            ),
            (
                "Подкрасться сзади, выхватить у драконицы книгу и заорать ей прямо в ухо",
                "От неожиданности Аурелия чихает облаком пламени! С дымящимися бровями вы летите кувырком с холма!",
                (-55, -35, 30),
                None,
                "woods",
            ),
        ],
    },
    "camp": {
        "narrative": "На лесной поляне у ручья горит костёр, вокруг которого отдыхают торговцы и странники. В котелке аппетитно булькает похлёбка, а широкая дорога от поляны ведёт обратно в Оукхейвен.",
        "audit_log": "=ЕСЛИ(ПРИВАЛ_У_КОСТРА;ОТДЫХ();В_ДОРОГУ)",
        "options": [
            (
                "Присесть к костру, послушать весёлые истории путников и передохнуть",
                "Тёплый обед и дружеская компания возвращают вам силы и отличное настроение.",
                (15, 15, -10),
                None,
                "crossroads",
            ),
            (
                "Спросить у караванщика дорогу и неспешно вернуться на площадь Оукхейвена",
                "Вы идёте вместе с попутчиками по ровной дороге и возвращаетесь к городскому фонтану.",
                (0, 0, 0),
                None,
                "crossroads",
            ),
            (
                "Прыгнуть обеими ногами прямо в костёр, чтобы показать всем свою удаль",
                "Сапоги вспыхивают! Путники окатывают вас ледяной водой из ведра, но ожоги весьма чувствительны.",
                (-45, -25, 20),
                None,
                "tavern",
            ),
        ],
    },
    "alley": {
        "narrative": "Вы оказываетесь в тихом солнечном переулке за пекарней. На лавочке добродушный городской стражник кормит голубей булочкой, а за углом виден выход на главную площадь.",
        "audit_log": "=ЕСЛИ(ПЕРЕУЛОК;СПРОСИТЬ_ДОРОГУ();ШТРАФ)",
        "options": [
            (
                "Спросить у стражника, как быстрее пройти к зданию Гильдии",
                "Стражник с улыбкой указывает на короткий проход через арку прямо к дверям Гильдии.",
                (0, 0, -5),
                None,
                "guild",
            ),
            (
                "Спокойно выйти из переулка обратно к фонтану на главной площади",
                "Через полминуты вы уже снова стоите в центре Оукхейвена.",
                (0, 0, 0),
                None,
                "crossroads",
            ),
            (
                "Попытаться отобрать у стражника его шлем и алебарду",
                "Стражник ловким подсечкой отправляет вас на мостовую и выписывает внушительный подзатыльник!",
                (-45, -30, 35),
                None,
                "crossroads",
            ),
        ],
    },
}

SCENE_POOL = [k for k in SCENES if k != "start"]

GENERIC_RESULTS = [
    "Вы спокойно продолжаете прогулку по солнечной долине.",
    "Прохожие дружелюбно кивают вам, и вы идёте дальше.",
    "Лёгкий ветерок шелестит листвой над дорогой.",
]


def _delta_for_custom(text):
    """Дельты для произвольного действия игрока из строки формул."""
    t = text.lower()
    if t.startswith("использовать предмет"):
        return (25, 25, -15), "Вы используете предмет из сумки: здоровье и бодрость заметно восстанавливаются!"
    if any(w in t for w in ("спрос", "узна", "пойти", "идти", "осмотр", "куда", "дорог", "поговор", "бесед")):
        return (0, 0, 0), "Вы спокойно разузнаёте обстановку и выбираете удобный маршрут без потери сил."
    if any(w in t for w in ("отдых", "перекус", "чай", "пирож", "зелье", "сон", "костер", "костёр")):
        return (15, 15, -5), "Небольшая передышка возвращает вам силы и хорошее настроение."
    if any(w in t for w in ("удар", "пнуть", "прыг", "драк", "атак", "украс")):
        return (-40, -25, 25), "Безрассудная выходка заканчивается ощутимыми шишками и ссадинами!"
    return (0, 0, 0), random.choice(GENERIC_RESULTS)


def _fmt(text, role, boss):
    return text.replace("{boss}", boss).replace("{role}", role)


def _encode_scene_state(scene_key, perm, last_btn=-1, streak=0):
    perm_str = "".join(str(i) for i in perm)
    return f"{scene_key}:{perm_str}:{last_btn}:{streak}"


def _decode_scene_state(raw_scene_id):
    if not raw_scene_id or not isinstance(raw_scene_id, str):
        return "start", [0, 1, 2], -1, 0
    parts = raw_scene_id.split(":")
    scene_key = parts[0] if parts[0] in SCENES else "start"
    perm = [0, 1, 2]
    if len(parts) > 1 and len(parts[1]) == 3 and set(parts[1]) == {"0", "1", "2"}:
        perm = [int(ch) for ch in parts[1]]
    last_btn = -1
    streak = 0
    if len(parts) > 3:
        try:
            last_btn = int(parts[2])
            streak = int(parts[3])
        except ValueError:
            pass
    return scene_key, perm, last_btn, streak


def generate_fallback_turn(action_text, stats, role, boss, raw_scene_id, choice_idx, seen):
    roi = clamp(stats.get("ROI", 100))
    kpi = clamp(stats.get("KPI", 100))
    anx = clamp(stats.get("Anxiety", 10))

    scene_key, perm, last_btn, streak = _decode_scene_state(raw_scene_id)
    prev = SCENES.get(scene_key)

    item = None
    next_target = None
    spam_penalty_msg = ""

    if prev and 0 <= choice_idx < 3 and not action_text.lower().startswith("использовать предмет"):
        real_opt_idx = perm[choice_idx]
        opt_tuple = prev["options"][real_opt_idx]
        _, result, delta, item = opt_tuple[0], opt_tuple[1], opt_tuple[2], opt_tuple[3]
        next_target = opt_tuple[4] if len(opt_tuple) > 4 else None

        if choice_idx == last_btn:
            streak += 1
        else:
            last_btn = choice_idx
            streak = 1

        # Если игрок бездумно спамит одну и ту же кнопку 2+ раза подряд, наказываем за невнимательность
        if streak >= 2:
            extra_dmg = 25 * streak
            delta = (delta[0] - extra_dmg, delta[1] - 20 * streak, delta[2] + 20 * streak)
            item = None
            spam_penalty_msg = (
                f" [ШТРАФ ЗА СПАМ: Вы {streak}-й раз подряд нажимаете кнопку №{choice_idx + 1} не глядя — "
                f"по невнимательности вы оступаетесь на ровном месте и теряете {extra_dmg} ОЗ!]"
            )
    else:
        delta, result = _delta_for_custom(action_text)
        streak = 0

    roi = clamp(roi + delta[0])
    kpi = clamp(kpi + delta[1])
    anx = clamp(anx + delta[2])
    result = _fmt(result, role, boss) + spam_penalty_msg

    # Условия поражения
    if roi <= 0:
        return {
            "narrative": (
                f"💫 ВЫ ПОТЕРЯЛИ СОЗНАНИЕ (Здоровье 0%): {result} "
                f"После «{action_text}» вас на тележке отвозят в городской лазарет Оукхейвена отсыпаться до следующего утра."
            ),
            "stats": {"ROI": 0, "KPI": kpi, "Anxiety": anx},
            "options": [],
            "audit_log": "=ОШИБКА(#В_ЛАЗАРЕТ!)",
            "inventory_add": None,
            "game_over": True,
            "scene_id": None,
        }
    if kpi <= 0:
        return {
            "narrative": (
                f"💫 СИЛЫ ИССЯКЛИ (Бодрость/Мана 0%): {result} "
                f"Вы так устали после «{action_text}», что уснули прямо на лавочке у дороги, пропустив все приключения дня."
            ),
            "stats": {"ROI": roi, "KPI": 0, "Anxiety": anx},
            "options": [],
            "audit_log": "=ОШИБКА(#НЕТ_СИЛ!)",
            "inventory_add": None,
            "game_over": True,
            "scene_id": None,
        }
    if anx >= 100:
        return {
            "narrative": (
                f"💫 ВАС ЗАДЕРЖАЛА СТРАЖА (Неприятности 100%): {result} "
                f"После ваших шумных проделок мастер {boss} и городская стража отправили вас чистить картошку на кухню Гильдии до вечера!"
            ),
            "stats": {"ROI": roi, "KPI": kpi, "Anxiety": 100},
            "options": [],
            "audit_log": "=НАРЯД_НА_КУХНЮ(ИСКАТЕЛЬ)",
            "inventory_add": None,
            "game_over": True,
            "scene_id": None,
        }

    if next_target and next_target in SCENES and next_target != scene_key:
        nxt_id = next_target
    else:
        clean_seen = [s.split(":")[0] for s in (seen or [])]
        unseen = [s for s in SCENE_POOL if s not in clean_seen]
        if not unseen:
            unseen = [s for s in SCENE_POOL if s != scene_key]
        nxt_id = random.choice(unseen)

    nxt = SCENES[nxt_id]
    nxt_perm = [0, 1, 2]
    random.shuffle(nxt_perm)
    shuffled_options = [_fmt(nxt["options"][i][0], role, boss) for i in nxt_perm]

    return {
        "narrative": f"{result} {_fmt(nxt['narrative'], role, boss)}",
        "stats": {"ROI": roi, "KPI": kpi, "Anxiety": anx},
        "options": shuffled_options,
        "audit_log": nxt["audit_log"],
        "inventory_add": item,
        "game_over": False,
        "scene_id": _encode_scene_state(nxt_id, nxt_perm, last_btn, streak),
    }


# --------------------------------------------------------------------------- #
#  Экспорт и модификация .xlsx с сохранением 100% цветов, тем, рамок и формул
# --------------------------------------------------------------------------- #

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")


def _col_to_letters(col_idx):
    s = ""
    n = int(col_idx)
    while n > 0:
        m = (n - 1) % 26
        s = chr(65 + m) + s
        n = (n - 1) // 26
    return s


def _letters_to_col(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _cell_col_index(cell_el):
    ref = cell_el.get("r", "")
    m = re.match(r"^([A-Z]+)", ref)
    return _letters_to_col(m.group(1)) if m else 0


def _clean_hex_argb(color_str):
    if not color_str:
        return None
    s = str(color_str).strip()
    if s.startswith("#"):
        h = s[1:].upper()
        if len(h) == 3:
            h = "".join(ch * 2 for ch in h)
        if len(h) == 6:
            return "FF" + h
        if len(h) == 8:
            return h
    m = re.match(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", s)
    if m:
        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"FF{r:02X}{g:02X}{b:02X}"
    return None


def _ru_formula_to_en(formula_str):
    s = formula_str
    if s.startswith("="):
        s = s[1:]
    repl = [
        (r"\bСУММ\s*\(", "SUM("),
        (r"\bСРЗНАЧ\s*\(", "AVERAGE("),
        (r"\bМИН\s*\(", "MIN("),
        (r"\bМАКС\s*\(", "MAX("),
        (r"\bСЧЁТ\s*\(", "COUNT("),
        (r"\bСЧЕТ\s*\(", "COUNT("),
        (r"\bОКРУГЛ\s*\(", "ROUND("),
        (r"\bЕСЛИ\s*\(", "IF("),
    ]
    for pat, sub in repl:
        s = re.sub(pat, sub, s, flags=re.IGNORECASE)
    s = s.replace(";", ",")
    return s


def _patch_styles_xml(styles_bytes, style_requests):
    if not style_requests:
        return styles_bytes, {}

    root = ET.fromstring(styles_bytes)
    q = lambda tag: f"{{{NS_MAIN}}}{tag}"

    fonts_el = root.find(q("fonts"))
    fills_el = root.find(q("fills"))
    borders_el = root.find(q("borders"))
    cell_xfs_el = root.find(q("cellXfs"))

    if fonts_el is None or fills_el is None or borders_el is None or cell_xfs_el is None:
        return styles_bytes, {}

    fonts_list = list(fonts_el.findall(q("font")))
    fills_list = list(fills_el.findall(q("fill")))
    borders_list = list(borders_el.findall(q("border")))
    xfs_list = list(cell_xfs_el.findall(q("xf")))

    mapping = {}
    fmt_map = {"int": "3", "dec": "2", "money": "4", "pct": "10"}

    for orig_xf_idx, st_key in style_requests:
        if (orig_xf_idx, st_key) in mapping:
            continue
        st = json.loads(st_key)
        base_xf = xfs_list[orig_xf_idx] if 0 <= orig_xf_idx < len(xfs_list) else xfs_list[0]
        new_xf = ET.fromstring(ET.tostring(base_xf))

        if any(k in st for k in ("b", "i", "u", "color", "sz", "font")):
            orig_font_id = int(base_xf.get("fontId", "0"))
            base_font = fonts_list[orig_font_id] if 0 <= orig_font_id < len(fonts_list) else fonts_list[0]
            new_font = ET.fromstring(ET.tostring(base_font))

            for flag_key, tag_name in (("b", "b"), ("i", "i"), ("u", "u")):
                if flag_key in st:
                    existing = new_font.find(q(tag_name))
                    if st[flag_key]:
                        if existing is None:
                            ET.SubElement(new_font, q(tag_name))
                    else:
                        if existing is not None:
                            new_font.remove(existing)

            if "sz" in st and st["sz"]:
                sz_el = new_font.find(q("sz"))
                if sz_el is None:
                    sz_el = ET.SubElement(new_font, q("sz"))
                sz_el.set("val", str(st["sz"]))

            if "font" in st and st["font"]:
                nm_el = new_font.find(q("name"))
                if nm_el is None:
                    nm_el = ET.SubElement(new_font, q("name"))
                nm_el.set("val", str(st["font"]))

            if "color" in st:
                col_el = new_font.find(q("color"))
                argb = _clean_hex_argb(st["color"])
                if argb:
                    if col_el is None:
                        col_el = ET.SubElement(new_font, q("color"))
                    col_el.attrib.clear()
                    col_el.set("rgb", argb)
                elif col_el is not None and st["color"] in (None, "", "auto"):
                    new_font.remove(col_el)

            fonts_el.append(new_font)
            fonts_list.append(new_font)
            new_xf.set("fontId", str(len(fonts_list) - 1))
            new_xf.set("applyFont", "1")

        if "fill" in st:
            fill_val = st["fill"]
            new_fill = ET.Element(q("fill"))
            pf = ET.SubElement(new_fill, q("patternFill"))
            argb = _clean_hex_argb(fill_val)
            if argb:
                pf.set("patternType", "solid")
                fg = ET.SubElement(pf, q("fgColor"))
                fg.set("rgb", argb)
                bg = ET.SubElement(pf, q("bgColor"))
                bg.set("indexed", "64")
            else:
                pf.set("patternType", "none")
            fills_el.append(new_fill)
            fills_list.append(new_fill)
            new_xf.set("fillId", str(len(fills_list) - 1))
            new_xf.set("applyFill", "1")

        if "border" in st:
            b_val = st["border"]
            new_border = ET.Element(q("border"))
            for side in ("left", "right", "top", "bottom"):
                s_el = ET.SubElement(new_border, q(side))
                if b_val in ("thin", "medium", "thick", "double", "dashed"):
                    s_el.set("style", b_val)
                    c_el = ET.SubElement(s_el, q("color"))
                    c_el.set("indexed", "64")
            ET.SubElement(new_border, q("diagonal"))
            borders_el.append(new_border)
            borders_list.append(new_border)
            new_xf.set("borderId", str(len(borders_list) - 1))
            new_xf.set("applyBorder", "1")

        if any(k in st for k in ("align", "valign", "wrap")):
            al_el = new_xf.find(q("alignment"))
            if al_el is None:
                al_el = ET.SubElement(new_xf, q("alignment"))
            if "align" in st and st["align"]:
                al_el.set("horizontal", str(st["align"]))
            if "valign" in st and st["valign"]:
                al_el.set("vertical", str(st["valign"]))
            if "wrap" in st:
                if st["wrap"]:
                    al_el.set("wrapText", "1")
                elif "wrapText" in al_el.attrib:
                    del al_el.attrib["wrapText"]
            new_xf.set("applyAlignment", "1")

        if "fmt" in st and st["fmt"] in fmt_map:
            new_xf.set("numFmtId", fmt_map[st["fmt"]])
            new_xf.set("applyNumberFormat", "1")

        cell_xfs_el.append(new_xf)
        xfs_list.append(new_xf)
        mapping[(orig_xf_idx, st_key)] = len(xfs_list) - 1

    fonts_el.set("count", str(len(fonts_list)))
    fills_el.set("count", str(len(fills_list)))
    borders_el.set("count", str(len(borders_list)))
    cell_xfs_el.set("count", str(len(xfs_list)))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), mapping


def patch_existing_xlsx(source_bytes, edits_by_sheet, styles_by_sheet):
    in_zip = zipfile.ZipFile(io.BytesIO(source_bytes), "r")
    q = lambda tag: f"{{{NS_MAIN}}}{tag}"

    wb_root = ET.fromstring(in_zip.read("xl/workbook.xml"))
    rels_root = ET.fromstring(in_zip.read("xl/_rels/workbook.xml.rels"))
    rels_map = {r.get("Id"): r.get("Target") for r in rels_root}

    sheet_targets = {}
    for idx, sh in enumerate(wb_root.findall(f".//{q('sheet')}")):
        rid = sh.get(f"{{{NS_REL}}}id") or sh.get("r:id")
        tgt = rels_map.get(rid, "")
        tgt_path = tgt.lstrip("/") if tgt.startswith("/") else f"xl/{tgt}"
        sheet_key = f"imp{idx}"
        sheet_targets[sheet_key] = tgt_path
        sheet_targets[sh.get("name", "")] = tgt_path

    modified_sheets_xml = {}
    style_requests = set()

    for sh_key, tgt_path in list(sheet_targets.items()):
        if not sh_key.startswith("imp"):
            continue
        sh_edits = edits_by_sheet.get(sh_key) or {}
        sh_styles = styles_by_sheet.get(sh_key) or {}
        if not sh_edits and not sh_styles:
            continue
        if tgt_path not in in_zip.namelist():
            continue

        ws_root = ET.fromstring(in_zip.read(tgt_path))
        modified_sheets_xml[tgt_path] = (ws_root, sh_edits, sh_styles)

        sheet_data = ws_root.find(q("sheetData"))
        if sheet_data is None:
            continue
        row_map = {int(r.get("r", 0)): r for r in sheet_data.findall(q("row")) if r.get("r")}

        for rc_key, st_dict in sh_styles.items():
            if not st_dict:
                continue
            r_idx, c_idx = [int(x) for x in rc_key.split(":")]
            row_el = row_map.get(r_idx)
            orig_s = 0
            if row_el is not None:
                ref = f"{_col_to_letters(c_idx)}{r_idx}"
                for c_el in row_el.findall(q("c")):
                    if c_el.get("r") == ref:
                        orig_s = int(c_el.get("s", "0"))
                        break
            st_json = json.dumps(st_dict, sort_keys=True, ensure_ascii=False)
            style_requests.add((orig_s, st_json))

    styles_bytes = in_zip.read("xl/styles.xml")
    if style_requests:
        styles_bytes, xf_mapping = _patch_styles_xml(styles_bytes, style_requests)
    else:
        xf_mapping = {}

    patched_files = {"xl/styles.xml": styles_bytes}

    for tgt_path, (ws_root, sh_edits, sh_styles) in modified_sheets_xml.items():
        sheet_data = ws_root.find(q("sheetData"))
        if sheet_data is None:
            continue

        rows_by_num = {int(r.get("r", 0)): r for r in sheet_data.findall(q("row")) if r.get("r")}
        all_rc_keys = set(sh_edits.keys()) | set(sh_styles.keys())

        for rc_key in all_rc_keys:
            r_idx, c_idx = [int(x) for x in rc_key.split(":")]
            ref = f"{_col_to_letters(c_idx)}{r_idx}"

            row_el = rows_by_num.get(r_idx)
            if row_el is None:
                row_el = ET.Element(q("row"), {"r": str(r_idx)})
                inserted = False
                for i, existing_r in enumerate(list(sheet_data)):
                    if int(existing_r.get("r", 0)) > r_idx:
                        sheet_data.insert(i, row_el)
                        inserted = True
                        break
                if not inserted:
                    sheet_data.append(row_el)
                rows_by_num[r_idx] = row_el

            cell_el = None
            for c_candidate in row_el.findall(q("c")):
                if c_candidate.get("r") == ref:
                    cell_el = c_candidate
                    break

            if cell_el is None:
                cell_el = ET.Element(q("c"), {"r": ref})
                inserted_c = False
                for i, existing_c in enumerate(list(row_el)):
                    if _cell_col_index(existing_c) > c_idx:
                        row_el.insert(i, cell_el)
                        inserted_c = True
                        break
                if not inserted_c:
                    row_el.append(cell_el)

            if rc_key in sh_styles and sh_styles[rc_key]:
                orig_s = int(cell_el.get("s", "0"))
                st_json = json.dumps(sh_styles[rc_key], sort_keys=True, ensure_ascii=False)
                new_s = xf_mapping.get((orig_s, st_json))
                if new_s is not None:
                    cell_el.set("s", str(new_s))

            if rc_key in sh_edits:
                val = sh_edits[rc_key]
                for child in list(cell_el):
                    cell_el.remove(child)
                if "t" in cell_el.attrib:
                    del cell_el.attrib["t"]

                if val is None or val == "":
                    pass
                elif isinstance(val, (int, float)):
                    v_el = ET.SubElement(cell_el, q("v"))
                    v_el.text = str(val)
                else:
                    s_val = str(val)
                    if s_val.startswith("="):
                        f_el = ET.SubElement(cell_el, q("f"))
                        f_el.text = _ru_formula_to_en(s_val)
                    else:
                        clean_num = s_val.strip().replace("\u00a0", "").replace(" ", "")
                        if re.match(r"^-?\d+([.,]\d+)?$", clean_num):
                            v_el = ET.SubElement(cell_el, q("v"))
                            v_el.text = clean_num.replace(",", ".")
                        else:
                            cell_el.set("t", "inlineStr")
                            is_el = ET.SubElement(cell_el, q("is"))
                            t_el = ET.SubElement(is_el, q("t"))
                            t_el.text = s_val

        patched_files[tgt_path] = ET.tostring(ws_root, encoding="utf-8", xml_declaration=True)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for item in in_zip.infolist():
            if item.filename in patched_files:
                out_zip.writestr(item, patched_files[item.filename])
            else:
                out_zip.writestr(item, in_zip.read(item.filename))
    out_buf.seek(0)
    return out_buf


def build_generated_xlsx(sheets_payload):
    fonts = [
        {"name": "Calibri", "sz": "11", "b": False, "i": False, "u": False, "color": "FF242424"},
        {"name": "Calibri", "sz": "11", "b": True, "i": False, "u": False, "color": "FFFFFFFF"},
    ]
    fills = [
        {"type": "none", "rgb": None},
        {"type": "gray125", "rgb": None},
        {"type": "solid", "rgb": "FF107C41"},
    ]
    borders = [
        {"style": None, "color": None},
        {"style": "thin", "color": "FFD1D1D1"},
        {"style": "medium", "color": "FF000000"},
    ]
    xfs = [
        {"fontId": 0, "fillId": 0, "borderId": 0, "numFmtId": 0, "align": None, "valign": "center", "wrap": False},
    ]
    xf_cache = {}

    def get_xf_idx(st):
        if not st:
            return 0
        key_str = json.dumps(st, sort_keys=True, ensure_ascii=False)
        if key_str in xf_cache:
            return xf_cache[key_str]

        f_obj = {
            "name": st.get("font") or "Calibri",
            "sz": str(st.get("sz") or 11),
            "b": bool(st.get("b")),
            "i": bool(st.get("i")),
            "u": bool(st.get("u")),
            "color": _clean_hex_argb(st.get("color")) or "FF242424",
        }
        if f_obj not in fonts:
            fonts.append(f_obj)
        font_id = fonts.index(f_obj)

        fill_rgb = _clean_hex_argb(st.get("fill"))
        if fill_rgb:
            fl_obj = {"type": "solid", "rgb": fill_rgb}
            if fl_obj not in fills:
                fills.append(fl_obj)
            fill_id = fills.index(fl_obj)
        else:
            fill_id = 0

        b_style = st.get("border")
        if b_style in ("thin", "medium", "thick", "double"):
            b_obj = {"style": b_style, "color": "FF000000"}
            if b_obj not in borders:
                borders.append(b_obj)
            border_id = borders.index(b_obj)
        else:
            border_id = 1

        fmt_map = {"int": 3, "dec": 2, "money": 4, "pct": 10}
        num_fmt_id = fmt_map.get(st.get("fmt"), 0)

        xf_obj = {
            "fontId": font_id,
            "fillId": fill_id,
            "borderId": border_id,
            "numFmtId": num_fmt_id,
            "align": st.get("align"),
            "valign": st.get("valign") or "center",
            "wrap": bool(st.get("wrap")),
        }
        xfs.append(xf_obj)
        idx = len(xfs) - 1
        xf_cache[key_str] = idx
        return idx

    sheet_xmls = []
    sheet_names = []
    for s_idx, sh in enumerate(sheets_payload or [{"name": "Лист1", "cells": []}]):
        name = (sh.get("name") or f"Лист{s_idx + 1}")[:31]
        sheet_names.append(name)
        cols = sh.get("cols") or []
        row_h = sh.get("rowH") or {}
        merges = sh.get("merges") or []
        cells = sh.get("cells") or []

        rows_dict = {}
        for item in cells:
            r, c, raw_val, disp_val, st = item[0], item[1], item[2], item[3], (item[4] if len(item) > 4 else None)
            rows_dict.setdefault(int(r), []).append((int(c), raw_val, disp_val, st))

        ws_lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">',
        ]
        if cols:
            ws_lines.append("<cols>")
            for ci, w_px in enumerate(cols, start=1):
                w_ch = max(4.0, round((float(w_px) - 5.0) / 7.0, 2))
                ws_lines.append(f'<col min="{ci}" max="{ci}" width="{w_ch}" customWidth="1"/>')
            ws_lines.append("</cols>")

        ws_lines.append("<sheetData>")
        for r_num in sorted(rows_dict.keys()):
            ht = row_h.get(str(r_num)) or row_h.get(r_num)
            ht_attr = f' ht="{round(float(ht) * 0.75, 2)}" customHeight="1"' if ht else ""
            ws_lines.append(f'<row r="{r_num}"{ht_attr}>')
            for c_num, raw_val, disp_val, st in sorted(rows_dict[r_num], key=lambda x: x[0]):
                ref = f"{_col_to_letters(c_num)}{r_num}"
                s_idx = get_xf_idx(st)
                s_attr = f' s="{s_idx}"' if s_idx else ""
                if raw_val is None or raw_val == "":
                    ws_lines.append(f'<c r="{ref}"{s_attr}/>')
                elif isinstance(raw_val, (int, float)):
                    ws_lines.append(f'<c r="{ref}"{s_attr}><v>{raw_val}</v></c>')
                else:
                    s_raw = str(raw_val)
                    if s_raw.startswith("="):
                        f_en = _esc_xml(_ru_formula_to_en(s_raw))
                        v_tag = f"<v>{_esc_xml(str(disp_val))}</v>" if isinstance(disp_val, (int, float)) else ""
                        ws_lines.append(f'<c r="{ref}"{s_attr}><f>{f_en}</f>{v_tag}</c>')
                    else:
                        ws_lines.append(f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{_esc_xml(s_raw)}</t></is></c>')
            ws_lines.append("</row>")
        ws_lines.append("</sheetData>")

        if merges:
            ws_lines.append(f'<mergeCells count="{len(merges)}">')
            for r1, c1, r2, c2 in merges:
                ws_lines.append(f'<mergeCell ref="{_col_to_letters(c1)}{r1}:{_col_to_letters(c2)}{r2}"/>')
            ws_lines.append("</mergeCells>")

        ws_lines.append("</worksheet>")
        sheet_xmls.append("\n".join(ws_lines).encode("utf-8"))

    st_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<styleSheet xmlns="{NS_MAIN}">',
        f'<fonts count="{len(fonts)}">',
    ]
    for f in fonts:
        b_tag = "<b/>" if f["b"] else ""
        i_tag = "<i/>" if f["i"] else ""
        u_tag = "<u/>" if f["u"] else ""
        st_lines.append(
            f'<font>{b_tag}{i_tag}{u_tag}<sz val="{_esc_xml(f["sz"])}"/>'
            f'<color rgb="{f["color"]}"/><name val="{_esc_xml(f["name"])}"/></font>'
        )
    st_lines.append("</fonts>")

    st_lines.append(f'<fills count="{len(fills)}">')
    for fl in fills:
        if fl["type"] in ("none", "gray125"):
            st_lines.append(f'<fill><patternFill patternType="{fl["type"]}"/></fill>')
        else:
            st_lines.append(
                f'<fill><patternFill patternType="solid"><fgColor rgb="{fl["rgb"]}"/><bgColor indexed="64"/></patternFill></fill>'
            )
    st_lines.append("</fills>")

    st_lines.append(f'<borders count="{len(borders)}">')
    for bd in borders:
        if not bd["style"]:
            st_lines.append("<border><left/><right/><top/><bottom/><diagonal/></border>")
        else:
            side_xml = "".join(
                f'<{s} style="{bd["style"]}"><color rgb="{bd["color"]}"/></{s}>'
                for s in ("left", "right", "top", "bottom")
            )
            st_lines.append(f"<border>{side_xml}<diagonal/></border>")
    st_lines.append("</borders>")

    st_lines.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    st_lines.append(f'<cellXfs count="{len(xfs)}">')
    for x in xfs:
        al_attrs = []
        if x["align"]:
            al_attrs.append(f'horizontal="{x["align"]}"')
        if x["valign"]:
            al_attrs.append(f'vertical="{x["valign"]}"')
        if x["wrap"]:
            al_attrs.append('wrapText="1"')
        al_xml = f'<alignment {" ".join(al_attrs)}/>' if al_attrs else ""
        st_lines.append(
            f'<xf numFmtId="{x["numFmtId"]}" fontId="{x["fontId"]}" fillId="{x["fillId"]}" borderId="{x["borderId"]}" '
            f'applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">{al_xml}</xf>'
        )
    st_lines.append("</cellXfs></styleSheet>")

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        ct_overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(len(sheet_xmls))
        )
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            f"{ct_overrides}</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )
        wb_sheets = "".join(
            f'<sheet name="{_esc_xml(nm)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
            for i, nm in enumerate(sheet_names)
        )
        zf.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets>{wb_sheets}</sheets></workbook>',
        )
        wb_rels = "".join(
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
            for i in range(len(sheet_xmls))
        )
        wb_rels += f'<Relationship Id="rId{len(sheet_xmls) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{wb_rels}</Relationships>',
        )
        zf.writestr("xl/styles.xml", "\n".join(st_lines).encode("utf-8"))
        for i, sh_bytes in enumerate(sheet_xmls):
            zf.writestr(f"xl/worksheets/sheet{i + 1}.xml", sh_bytes)

    out_buf.seek(0)
    return out_buf


def _esc_xml(s):
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _clean(value, default, limit=60):
    value = (value or "").strip() if isinstance(value, str) else ""
    return value[:limit] or default


app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
)


@app.route("/")
def index():
    _ensure_runtime_files()
    try:
        return render_template("index.html")
    except Exception:
        raw_html = zlib.decompress(base64.b64decode(_EMBEDDED_FILES["templates/index.html"])).decode("utf-8")
        raw_html = raw_html.replace("{{ url_for('static', filename='css/style.css') }}", "/static/css/style.css")
        raw_html = raw_html.replace("{{ url_for('static', filename='js/app.js') }}", "/static/js/app.js")
        return Response(raw_html, mimetype="text/html; charset=utf-8")


@app.route("/api/workbook/default", methods=["GET"])
def get_default_workbook():
    _ensure_runtime_files()
    if os.path.exists(DEFAULT_XLSX_PATH):
        return send_file(
            DEFAULT_XLSX_PATH,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=False,
            download_name=DEFAULT_XLSX_NAME,
        )
    return jsonify({"error": "Default workbook not found"}), 404


@app.route("/api/export_xlsx", methods=["POST"])
def export_xlsx_route():
    _ensure_runtime_files()
    data = request.get_json(silent=True) or {}
    filename = _clean(data.get("filename"), DEFAULT_XLSX_NAME, 120)
    if not filename.lower().endswith(".xlsx"):
        filename = re.sub(r"\.(xls|csv|tsv|txt)$", "", filename, flags=re.IGNORECASE) + ".xlsx"

    mode = data.get("mode", "imported")
    edits = data.get("edits") or {}
    styles = data.get("styles") or {}
    source_b64 = data.get("source_b64")

    try:
        if mode == "imported":
            if source_b64:
                src_bytes = base64.b64decode(source_b64)
            elif os.path.exists(DEFAULT_XLSX_PATH):
                with open(DEFAULT_XLSX_PATH, "rb") as f:
                    src_bytes = f.read()
            else:
                src_bytes = None

            if src_bytes:
                out_buf = patch_existing_xlsx(src_bytes, edits, styles)
                return send_file(
                    out_buf,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    as_attachment=True,
                    download_name=filename,
                )

        sheets_payload = data.get("sheets") or []
        out_buf = build_generated_xlsx(sheets_payload)
        return send_file(
            out_buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:  # noqa: BLE001
        print(f"Ошибка экспорта XLSX: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/start", methods=["POST"])
def start_game():
    data = request.get_json(silent=True) or {}
    role = _clean(data.get("role"), "Начинающий искатель")
    boss = _clean(data.get("boss"), "Мастер Гильдии Морган")
    api_key = _clean(data.get("api_key"), "", 200)

    prompt = (
        f"Начни живое и лёгкое фэнтезийное приключение с классического пролога: солнечное утро, герой ({role}) едет в повозке "
        f"в уютный городок Оукхейвен на ярмарку и приём в Гильдию к мастеру «{boss}». "
        f"Стартовые параметры: ROI (Здоровье) 100, KPI (Бодрость/Мана) 100, Anxiety (Неприятности) 10. "
        f"Пиши просто и увлекательно (2-3 коротких предложения). Сделай 1-й вариант опасной глупостью, а 2-й и 3-й — спокойным разговором и выбором пути!"
    )
    res = normalize_llm(call_llm(prompt, api_key=api_key), {"ROI": 100, "KPI": 100, "Anxiety": 10})
    if res:
        res["scene_id"] = None
        return jsonify(res)

    s = SCENES["start"]
    perm = [0, 1, 2]
    return jsonify({
        "narrative": _fmt(s["narrative"], role, boss),
        "stats": {"ROI": 100, "KPI": 100, "Anxiety": 10},
        "options": [_fmt(s["options"][i][0], role, boss) for i in perm],
        "audit_log": s["audit_log"],
        "inventory_add": None,
        "game_over": False,
        "scene_id": _encode_scene_state("start", perm, -1, 0),
    })


@app.route("/api/action", methods=["POST"])
def take_action():
    data = request.get_json(silent=True) or {}
    action_text = _clean(data.get("action"), "Осмотреться вокруг", 200)
    raw = data.get("stats") or {}
    stats = {
        "ROI": clamp(raw.get("ROI", 100)),
        "KPI": clamp(raw.get("KPI", 100)),
        "Anxiety": clamp(raw.get("Anxiety", 10)),
    }
    role = _clean(data.get("role"), "Начинающий искатель")
    boss = _clean(data.get("boss"), "Мастер Гильдии Морган")
    api_key = _clean(data.get("api_key"), "", 200)
    clock = _clean(data.get("clock"), "09:00", 5)
    inventory = [str(i)[:40] for i in (data.get("inventory") or [])][:20]
    seen = [str(s) for s in (data.get("seen") or [])][:30]
    scene_id = data.get("scene_id")
    try:
        choice_idx = int(data.get("choice_idx", -1))
    except (TypeError, ValueError):
        choice_idx = -1

    prompt = (
        f"Герой: {role}, Мастер Гильдии: {boss}.\n"
        f"Время: {clock} (день заканчивается в 18:00).\n"
        f"Сумка: {', '.join(inventory) if inventory else 'пусто'}.\n"
        f"Действие игрока: «{action_text}»\n"
        f"Текущие параметры -> ROI (Здоровье): {stats['ROI']}, KPI (Бодрость/Мана): {stats['KPI']}, Anxiety (Неприятности): {stats['Anxiety']}.\n"
        f"Если игрок просто спросил что-то или выбрал направление куда идти — НЕ отнимай ROI и KPI (оставь их прежними)! "
        f"Если же игрок совершил явную глупость или полез в драку напролом — нанеси большой урон (-45..-55). "
        f"Пиши легко и интересно (2-3 простых предложения), а варианты 1-3 перемешай местами."
    )
    res = normalize_llm(call_llm(prompt, api_key=api_key), stats)
    if res:
        return jsonify(res)

    return jsonify(generate_fallback_turn(action_text, stats, role, boss, scene_id, choice_idx, seen))


_EMBEDDED_FILES = {
    "templates/index.html": "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",
    "static/css/style.css": "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",
    "static/data/\u041a\u043d\u0438\u0433\u04301.xlsx": "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",
    "static/guides/blue.jpeg": "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",
    "static/guides/excel.jpeg": "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",
    "static/guides/green.jpeg": "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",
    "static/js/app.js": "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"
}


def _ensure_runtime_files():
    """Автоматически восстанавливает templates/ и static/, если app.py запущен напрямую из ZIP-архива Windows без распаковки."""
    for rel_path, b64_data in _EMBEDDED_FILES.items():
        full_p = os.path.join(BASE_DIR, *rel_path.split("/"))
        if not os.path.exists(full_p):
            os.makedirs(os.path.dirname(full_p), exist_ok=True)
            with open(full_p, "wb") as f:
                f.write(zlib.decompress(base64.b64decode(b64_data)))


if __name__ == "__main__":
    _ensure_runtime_files()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
