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
    "templates/index.html": "eJzNW3tv29YV/7+f4pZFIAc1KZGSbCewNCRd+sCcNUizdts/BUVdSUwoUSNp2W5RwE7aJkWyeEu7rdj6WLZiGFAMUB07deLEBfoJqK+QL7CvsHPOJSmSIiXb7YYFkcXH5X2cx+/8zrnU8vM/ff2lK7+6dIF1vK5Vf24Zv5il99o1yVmV8ALXm/XnGPxb7nJPZ0ZHd1zu1aRfXHlZXpLit3p6l9ekgcnX+rbjScywex7vQdM1s+l1ak0+MA0u08k8M3umZ+qW7Bq6xWuqUgq78kzP4nX/z/5Tf99/4A9VZd1y19mzzU/YhXWDW8tF0UK0tszeNeZwqyaZMJzEOg5v1aSm7ulnza7e5kV30H5xvWvNnyq/BIcMDnturdDxvP7ZYnFtbU1ZKyu20y5qpVIJGxcYLuC8vV4rlFiJlTX4XzhVvgDPO9zwGFxfKLCNWqFcYGJdBU0rsA432x0PjuGms053W6Zl1QqntLKm6uWFhUJR9NLXvQ5r1goXF5ha6mjaYPHVhXdirdXSolFRw9bhmBqNuRiNqVbGY6pLYkxNqcb7Waoa5cWJUatKlamqUrUW2ZluST5jyXBQYK7n2Nc4PdhqtcJzOVyhokWXQOLc0Pu1gmOv9pph/yg6OJIm1OJ6GxZ3O5x7oXLefZetOtbbLduZK7ie7plGYR6nzcl+CobrFukhBY4Kp9l776EZFoUdLjfs5kYwxvOyzPx7/t5oc/SB/3S0Pdpm/nf+EOxmzz8Y3WH+t3DywD/0D+Cz4z/2h+z1VgtMkJUXqqworIm9xRtMloMucRDuMMPSXbcmkZ01dCdYFLVomoPodkO2eMuL3aUWjVXPs3thI7RKueH12JreallcYtRpTfLvjzZhnrujrdF1mNyhmO13oxujLZopHG6C/ePUH0I79IVHEtMdU5ctvYGC9b9MtxhtS8zuGZZpXKtJ5/p9xbPbbYuft1137nRqmjRVdIjQ2iW0dhVsckEKbEzCw8DG8Li+3BbWJRmrjgOO/ZJt2SCcZcN0DIszAzopg9dv0JcDzyhaVSom7i/NuK/O6iC4vzRjgNz76qwOgvvqrCXkN1Czuyi24QMST5lLUdhL6qrb1yMT4minsmW37ch6yHSPpFGtAv8jjWqlsUa1UkYH1EmAORKoH5ahZdoD9gSYg3eFTbwgUA5Wmt1pAEASwN5iR10YVAH1omcF5h3lWVWbfJiADky/rxumBzNWKtXZXakgmAGuIkC+aBJJ6JOUpXjP+R2HQtNIaNWx0LSYE1XSQgvmPnO2VUVjSwDaVQXEB7C9aMl0GJs+wHZ67qpyRkrDtkSwnTVipnWiIU6xzaZtyGSTEjObdHpFnCVwyCFof8t2rjVs+xpgUYSCnwPoPfSfAHZd9/fm2egmfB/634xuI5TvEUTuw22EN8Tw4eg6AvtjwQ1GN1jgCGm2MHPehmWvNmWMPquumLqrD/gbwXmE0f7h6AOYhIgpMIWTgGg1pv9qnstFiq4oaONKtbOoaLqmLDL8QK8lpQIaLy+eKytLDD94raKcYWgGZbAO/IhrooPIP3p2j8+0clUpH8FnFmCMM4pqqUqF4QfmJsPnmCNV8mwyfvmqbfaOZ6sTOhbalT2+jtTDbDZ5rz6p0kxTSV1bLkLYrz+XxwJcrjtGJ80DZtlFJYELiQiyoARBhg4ghlQAeDKFnIjEkzrFuDNGvBKoD/+sAPaplaN2kquuYmY8M3v9VY95G33oWIi+b+kG79gWMCtkLeDI+0hyGDjsU6Q/wN9uCB8n598BGNgi+jMk1rbP8P8utUI2Bx82d27lCnuR+V+fTnOisPfAqUkzoAHpyMp0UCVH43QplLP7vPcG9zyz13bjGPclwBYuBpZIiwVse0DYdoiMFdb2qf8pm7uiNy5ALtTn7JJjTy4r1cWPweYiw1gCEKkA2Cww/MBzpRLDgDO+ArkRoI/2DjQFOmMB7KgASxBeZKWC+QRdWMJTuiP+ysElui+LxivQQUVZsmS6gk3OALApVVk0WYLrwR3xxYJr1IBR6xWaQw68TTdlxOJcjDkZOdMHkGo6Y2X/Dnn86DoDoEEVQ0QChr4Pxr3D5uDm/dPTjaaObXLgRxyKFCWwX5ED3UebAMOnvEE40l0yrIPAqYa5vjZOfXp65AWO2WhgLo0eFBxnZ0Bd3ltN5kfpFuJx2dMbbry/K3hej61rSu+ybnim3XOzTD7pl91Ghle2bc9+A9PPuQJQ6ELcLyl/Qpe6SeFgP3DHXQoNO8QoHoeJ2GNIKe9CwzAVywmVs3ywHPPBctwHIWKzckdVdZWp6HAqfA8WozMZDl5dsOQyNIOmb6pa7JasvlkZPyerJ/IN7Zi+kfCGepYwsyJrvlfl6pMR3No9ayNfs3pzkNBsQB0fgZGDQrOVOLr7oytRRawrAU0HMiaXYwhagXPQL6ruJLo5AWPKI0jHk73w2a7d5OcnHCtRXJgQfSjqfWLsdOkhGggC1JBqHPuEWSLCz/m/h2i4NdqS8UEK/PujW4BbEBKzKhc/UE0QVEAVmgUKWqF6mCbciqngX5YqlyHXKsvw+d97UpSgdPhAqj/7YyY/PYkXuR3d4XLb4TytSL6O5dpfQtoUV+S//MdIw0Atm+hFgbKyEzA2R1nXaSRuh8zfga89TOrY6EN4Zg8jEJ3iFax37YNxbOOVH121Z4QiO5VB5aIagOXKInxD8gzpS5VVXoWrIViWZELZ2AWE2TEII7gO5PKJqHf1R/PZyDbqRG93sbwp/AqEmYuy/0WLimqable3LBbGeMKDRJQXCJGDHJepTdzk7pOxbAI3uSHMDKub+4yY77fpewG3ARPM4zYnNa54wUYLjSu9qktwUSKTKxPoryCJDpBELf0/gnyKbaVPJ0kbltrj6jyP5/UEGQXSmMtER+8jdvhPQCkHY5oZG6ZlO91VS5fzK+xYNYJprMtrjt5Ps0yRaKZaivniGSZ9bKBbqyD4c2oqn/oUprXNRtvEBR5hCBpb4eQ9CE87aH7gelSnYmCPuxjSRlvzDJBwGNALZD+b7Hx1nmjk03hpi13oedxJh7KEf/Yacq6LTlFWq5FDkKc67Xi5X8D0RIFtyOYo3Cad1dB7gPAXmqaH6cmzz/6QY11TIcK+Nh7wHm53+LswmBBJaji72zW9l4VpiBE/PsmIrXXBX6JRKfAIYkGM5H2wy6cQ4z4kJpichNlzueO9jOBUT+iotS7VW2F1cZb2Jqc9TY2BM5BVT7H3eGEFTT147jW8KWV2Bqa/6tkg177FPXjWpjpxn1sWTBhX3NItl0tZWJmo3EyzXTBCYBE6glUo8L9mgLZAacQHgO04PoDY7rw/Ne0d12wE0vwNuv+G4uBNNF302C1iGgg9Edp0dTOaY9sxm0KwJDg8fSsp5rg+qLVrODb6Stj+DXGeEgQkt1i3Gz83fgDhkm7nZ7m4n+7JduMqNwKV0pXXgwtB3TIfvJeLuMiEbD4HLHqYuSOK8toKMv+AjI3xeoso+HUsSwXZcCTHlm17433Rhg1m3c0RHG31ytOrAqINBI/ZOX0atgS8Dy454GtJpxXKwsrCnKxppVSijza3iWYY+D8lHphp7IoS4w4Rih0s8T/bfHRCKhRN7+fkn3nTO8nswnooze/xD5ofHFDa7OYxM3H3Iu+txtnZPXSwyISwBPnso/tHohrZ6h/XhOg8vyQ0Fef1JripeLsgnOgnuIMUg3qxdzS67T+KZp9aNnQiCgmA9/9+cjsXuLMNnrYZZhh85qsC1CqxY9GQ14JNMtqbctM21LHXwl003KtKZN73Mb6RBYkUi3x6k2DyUHCZKG/DxCuzfdggN0Wg2ZLOGpe5jvxwPHWHLtT9j2E8gpi8zZXpxtGQuyGEug1a5cxKoZtdts+SL5hPfsKRjE8TeJgZpY44EKmAGEhQaIzVOc5imYNecIrKIOOCyTzDegTeHH2AvAkn8c8TTmK8qTKcUYb+yYlG+AJ3bjEgk7PFiz7f0XbHFjDNQxLxEInSV7mD5Lg8TC9dBHzHtrvnN+ZkNQ6pfw/SwDujWyEGUNVpdIs08A0MfuveFAyNUy1H77WDrBbHuixOu5iGLUCe19XxjY8Feo+AQ8qGuV+QdailUirt+CwxiTC2T/EyHHEFH478DK+I/ib22F3u/RruxjEhMSCkMcmcBKLLAZgizPMUqBx3fU6xafWBY2slpZSdoHJxM0cpmdB7gtFFFHt51bIg6HKeqDHcowIVpm7o77/FuIs+CRbxl4fHzZuXi4IXJcjXPaqaHYAHYLXs7ugjWPRebFsGA9E3xHyHYz42zEyQKa65ct/ur/ZjkdK9JC4kqWE8Cyc2cRuDNA69D2r/jhDudqjyxCyAiA8h0T0IWSFRRlFhws3KqIaX2C6Kuwi+uxdkI3AUpCK6YfC+V5OoMjivGO5gXvHoT2w7frIvQxRHsLNGmyoll0Cv3Im86oWmwRda2tG6aGV30WqVSuiciS7i28C27gZkPDgE7sbDUC9liPwTkPKQRPgA929lyknCaLLnP87Ub9du6hZQZeNa07FDDQcgfBHvTUJE/PGmqVt2O5xbeEaAQz3D5J1VHocgizcbG9DUaou3dPIjKjSRcacxK6J2NPGuT9TL7B3uxyCFp2BnVDPGppR2b9PW93Kxox3Z3dPVAst2eXzPPQG3f4IphZ5wZ1rtYjq7QFlQLSxjlv14oyZ3jbyiY6PufwVwg6XJPciLycuwChJthxxCxKRi+lmYZJ2l9kHgeeAI4s4TEK54HQqrNnCEVIXWiZCDKpjIH1DupBY4RnI20XoIKHwjRhdBS3czyvziBVWsusLt0RZD5iQARZRqcdMHO0CuSW/fUk1WvJ4qSAENSlsDAeqwZcNu8nrw2hYdn1YySpr9DOmTolnLdsBAEAAu22SOooIn7O6A/HArWCNMBF8fGVvncpH6mMEBxuWW8TAY+S3ea2PxFsN/8i2Xz5NFwbOYlA8x8BEwYFTYR2Hsi7d9ySggQZ8PK6hDUdEGmL4xbvwEBZplhGkx4J7cWAyCA1ynFUPI+Zh2X+5gLVNYxkFQrIwKZE9FMnAi0dDQxxVNbIYT8/uMgtQDEaNnL/1c3/wZ36AXAmn/ka2sXGTnLr1G3GePXi3cFm9oCImj+5w+y15x7N8Ehth2r739bPMfgSmGAnqFA+MzgybnXntHHzc5mpz6gBKQ2zVjsgrmmhLPF6EegL0SSQtd7jbwaEoDQie9HrlvLNKEfgaX7qLjQavA8XYRTODSIb4cNVkSzKKhlGEHANcxe5BfAb25jZZ4GPS9R2kSSHSXTGcY7s0ENYIAN+JIMUeocgjwhwnn91+DPYiNxe1xejr8/iAU/Pdfv/j9wWnqEhedhJVwuAh0xE9FImjBcRFZn4AQ74A5BE6gAMPFpR07CiDdm12ymh2m6vGq+5EpL7LdvmN2dSdNeOnd1Xj30dsYlG2KvTTh5DeF0I5PdWOMB1i12feY6xg1qSh+wVG86hZBovYa7hYoV4kjiWb1yWdm/AAEu4JFXQ1//DHuCWZNP/8AykC/VvoP/r+CGQ==",
    "static/css/style.css": "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",
    "static/data/\u041a\u043d\u0438\u0433\u04301.xlsx": "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",
    "static/guides/blue.jpeg": "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",
    "static/guides/excel.jpeg": "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",
    "static/guides/green.jpeg": "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",
    "static/js/allowance.js": "eJyNkt9r2zAQx9/zV9z8UOTMUTb21pFBRxls0G3Qp1FKUa1zrCFLQTrXMyH/+yT/iJ22gepFPunuo6/vvuslfHW28ehW2uZCg8TKgtDaNsLkyOG3UBLQkCKNVdg9NEpreER4QqcKhRIeW6ASITDCEYfletEoI23Dr4XS7dXIgg3sFxCWVpWiS/iUdZEULZOC4rXBBq7DJ0vTITUuh1Q7A3cxiW+RvtVa/0HhWJrBeHZjDZUshffwcTrsUff8r1WGJask/dwxD/27DoVknqwTW+xrzirIrfEUhYYMKpXno+aBGBe5dlYxVT0JXUfyj9tfP/lOOI/jo1EiSwgfZGzTQzxV6JMMkv0hSWfouFQBrEPxTsZm08m5uIAr50TLle/2ISW31U4joTz5i2f93AdABsfUS7jjnMf/vw2ynnF4oTShY8EKmy9A7Q5tATEIOhJPTpltUHx/ONV8OEYHyAXlJexhvYTvJtBDsUeiUOjBk3AEomv+zHnL9YxwVvT46HGqlVAmYN842oF7I6jklfjHPmT9hDuPwqoPXlolnfVGo9lSeWqu8XaqUvJNJnOYWydHn73y8NTiaIl3ff5MjTK5riX6MKuAf3G9q30ZrybM6EZ/xo2dcfsZq6JlPXFuz6GFg+BX2z82ZxGG9R/sw12I",
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
