"""
Telegram viktorina boti (aiogram 3.x)

Ishga tushirish:
    pip install -U "aiogram>=3.7" python-docx
    $env:BOT_TOKEN = "123456:ABC..."     # PowerShell
    python bot.py
"""

import asyncio
import html
import json
import logging
import os
import random
import re
import socket
import sys
import tempfile
import uuid
from dataclasses import dataclass, field
from typing import Optional

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ChatMemberStatus, ChatType, ParseMode, PollType
from aiogram.exceptions import TelegramBadRequest, TelegramRetryAfter
from aiogram.filters import Command
from aiogram.types import BotCommand, CallbackQuery, InlineKeyboardButton, Message, PollAnswer
from aiogram.utils.keyboard import InlineKeyboardBuilder
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

# --------------------------------------------------------------------------
# Sozlamalar
# --------------------------------------------------------------------------
MAX_FILE_SIZE = 20 * 1024 * 1024  # Bot API yuklab olish limiti: 20 MB
LOCK_PORT = 48653           # bitta nusxa ishlashini ta'minlovchi lokal port

POLL_Q_MAX = 300            # Telegram: savol max 300 belgi
POLL_OPT_MAX = 100          # Telegram: variant max 100 belgi
POLL_OPT_COUNT = 10         # Telegram: variantlar max 10 ta
LETTERS = "ABCDEFGHIJ"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("quizbot")

router = Router()


# --------------------------------------------------------------------------
# Chat sozlamalari: botning ichidan /settings orqali o'zgartiriladi,
# settings.json faylida saqlanadi (bot qayta ishga tushsa ham yo'qolmaydi)
# --------------------------------------------------------------------------
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

DEFAULTS: dict = {
    "question_time": 30,   # har bir savol vaqti (soniya)
    "lobby_timeout": 60,   # kutish xonasi vaqti (soniya)
    "count": 25,           # savollar soni
    "pause": 3,            # to'g'ri javobdan keyingi pauza (soniya)
    "shuffle_q": True,     # savollarni random tanlash
    "shuffle_o": True,     # variantlarni aralashtirish
}
# kalit: (min, max, qadam, belgi, nomi, birlik)
NUM_SETTINGS = {
    "question_time": (5, 600, 5, "⏱", "Savol vaqti", "s"),
    "lobby_timeout": (10, 600, 10, "⏳", "Kutish xonasi", "s"),
    "count": (1, 200, 5, "🔢", "Savollar", "ta"),
    "pause": (0, 30, 1, "⏸", "Pauza", "s"),
}
BOOL_SETTINGS = ("shuffle_q", "shuffle_o")

CHAT_SETTINGS: dict[str, dict] = {}


def load_settings() -> None:
    global CHAT_SETTINGS
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            CHAT_SETTINGS = data
    except FileNotFoundError:
        pass
    except (OSError, ValueError):
        log.exception("settings.json o'qilmadi, standart sozlamalar ishlatiladi")


def save_settings() -> None:
    tmp = SETTINGS_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(CHAT_SETTINGS, f, ensure_ascii=False, indent=2)
        os.replace(tmp, SETTINGS_FILE)
    except OSError:
        log.exception("settings.json saqlanmadi")


def get_settings(chat_id: int) -> dict:
    """Chat sozlamalarini (standart qiymatlar bilan birlashtirib) qaytaradi."""
    cfg = dict(DEFAULTS)
    saved = CHAT_SETTINGS.get(str(chat_id), {})
    if not isinstance(saved, dict):
        saved = {}
    for key, spec in NUM_SETTINGS.items():
        v = saved.get(key)
        if isinstance(v, int) and not isinstance(v, bool):
            cfg[key] = max(spec[0], min(spec[1], v))
    for key in BOOL_SETTINGS:
        if isinstance(saved.get(key), bool):
            cfg[key] = saved[key]
    return cfg


def update_settings(chat_id: int, **changes) -> None:
    CHAT_SETTINGS.setdefault(str(chat_id), {}).update(changes)
    save_settings()


def reset_settings(chat_id: int) -> None:
    CHAT_SETTINGS.pop(str(chat_id), None)
    save_settings()


# --------------------------------------------------------------------------
# Ma'lumot tuzilmalari
# --------------------------------------------------------------------------
@dataclass
class Question:
    text: str
    options: list[str]
    correct: int  # options ichidagi indeks


@dataclass
class Session:
    sid: str
    chat_id: int
    host_id: int
    questions: list[Question]
    cfg: dict
    state: str = "lobby"  # lobby | running | stopped
    players: dict[int, str] = field(default_factory=dict)   # lobbidagilar
    names: dict[int, str] = field(default_factory=dict)     # barcha ismlar
    scores: dict[int, int] = field(default_factory=dict)
    go: asyncio.Event = field(default_factory=asyncio.Event)
    lobby_msg: Optional[Message] = None
    task: Optional[asyncio.Task] = None
    poll_ids: list[str] = field(default_factory=list)
    poll_answers: dict[str, dict[int, bool]] = field(default_factory=dict)
    open_poll: Optional[Message] = None


SESSIONS: dict[int, Session] = {}            # chat_id -> faol sessiya
POLLS: dict[str, tuple[Session, int]] = {}   # poll_id -> (sessiya, to'g'ri variant indeksi)


# --------------------------------------------------------------------------
# Fayl o'qish va parser
# --------------------------------------------------------------------------
def _walk(element, parent, lines: list[str]) -> None:
    """DOCX elementlarini (paragraf va jadvallar) tartib bilan o'qiydi."""
    for child in element.iterchildren():
        if child.tag == qn("w:p"):
            lines.append(Paragraph(child, parent).text)
        elif child.tag == qn("w:tbl"):
            table = Table(child, parent)
            for row in table.rows:
                seen = []
                for cell in row.cells:
                    if any(cell._tc is s for s in seen):  # birlashtirilgan katakni takrorlamaslik
                        continue
                    seen.append(cell._tc)
                    _walk(cell._tc, cell, lines)


def read_docx(path: str) -> str:
    doc = Document(path)
    lines: list[str] = []
    _walk(doc.element.body, doc, lines)
    return "\n".join(lines)


def read_txt(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16", errors="replace")
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


SEP_Q = re.compile(r"\+{5,}")
SEP_O = re.compile(r"={5,}")
NUM_PREFIX = re.compile(r"^\s*\d{1,4}\s*[.)]\s*(?!\d)")
ZERO_WIDTH = dict.fromkeys(map(ord, "\u200b\u200c\u200d\ufeff"), None)


def _clean(s: str) -> str:
    s = s.translate(ZERO_WIDTH).replace("\u00a0", " ")
    return " ".join(line.strip() for line in s.split("\n") if line.strip())


def parse_questions(text: str) -> list[Question]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    parts = SEP_Q.split(text)
    # birinchi bo'lak odatda sarlavha; agar unda savol bo'lsa (=====) uni ham olamiz
    blocks = parts if (parts and SEP_O.search(parts[0])) else parts[1:]

    result: list[Question] = []
    for block in blocks:
        chunks = SEP_O.split(block)
        if len(chunks) < 3:
            continue
        q_text = _clean(NUM_PREFIX.sub("", chunks[0].strip(), count=1))
        if not q_text:
            continue

        opts: list[tuple[str, bool]] = []
        index: dict[str, int] = {}
        found_correct = False
        for chunk in chunks[1:]:
            chunk = chunk.translate(ZERO_WIDTH).strip()
            is_correct = False
            if chunk.startswith("#"):
                chunk = chunk[1:]
                if not found_correct:
                    is_correct = True
                    found_correct = True
            opt = _clean(chunk)
            if not opt:
                continue
            if opt in index:  # Telegram bir xil variantlarni qabul qilmaydi
                if is_correct:
                    old, _ = opts[index[opt]]
                    opts[index[opt]] = (old, True)
                continue
            index[opt] = len(opts)
            opts.append((opt, is_correct))

        if len(opts) < 2 or not any(c for _, c in opts):
            continue
        correct = next(i for i, (_, c) in enumerate(opts) if c)
        result.append(Question(q_text, [o for o, _ in opts], correct))
    return result


def prepare_questions(
    pool: list[Question], count: int, shuffle_q: bool = True, shuffle_o: bool = True
) -> list[Question]:
    """Savollarni tanlaydi (random yoki fayl tartibida) va sozlamaga qarab variantlarni aralashtiradi."""
    chosen = random.sample(pool, min(count, len(pool))) if shuffle_q else pool[:count]
    prepared: list[Question] = []
    for q in chosen:
        items = [(o, i == q.correct) for i, o in enumerate(q.options)]
        if len(items) > POLL_OPT_COUNT:  # to'g'ri variantni saqlab, ortiqchasini tashlaymiz
            others = [i for i in range(len(items)) if i != q.correct]
            keep = {q.correct, *random.sample(others, POLL_OPT_COUNT - 1)}
            items = [it for i, it in enumerate(items) if i in keep]
        if shuffle_o:
            random.shuffle(items)
        correct = next(i for i, (_, c) in enumerate(items) if c)
        prepared.append(Question(q.text, [o for o, _ in items], correct))
    return prepared


# --------------------------------------------------------------------------
# Yordamchi funksiyalar
# --------------------------------------------------------------------------
esc = html.escape


async def tg(func, *args, **kwargs):
    """Telegram flood-limitiga (RetryAfter) chidamli chaqiruv."""
    for attempt in range(4):
        try:
            return await func(*args, **kwargs)
        except TelegramRetryAfter as e:
            if attempt == 3:
                raise
            await asyncio.sleep(e.retry_after + 1)


def lobby_keyboard(sid: str):
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Qo‘shilish", callback_data=f"join:{sid}")
    kb.button(text="🚀 Viktorinani boshlash", callback_data=f"go:{sid}")
    kb.adjust(1)
    return kb.as_markup()


def lobby_text(s: Session, started: bool = False) -> str:
    lines = [
        "📚 <b>Viktorina</b>",
        f"Savollar soni: <b>{len(s.questions)}</b>",
        f"Har bir savol uchun vaqt: <b>{s.cfg['question_time']} soniya</b>",
        "",
        f"👥 <b>Qo‘shilganlar ({len(s.players)}):</b>",
    ]
    if s.players:
        lines += [f"{i}. {esc(n)}" for i, n in enumerate(s.players.values(), 1)]
    else:
        lines.append("— hozircha hech kim yo‘q —")
    lines.append("")
    if started:
        lines.append("🚦 <b>Viktorina boshlandi!</b>")
    else:
        lines.append(f"⏳ {s.cfg['lobby_timeout']} soniyadan keyin avtomatik boshlanadi.")
    return "\n".join(lines)


async def refresh_lobby(s: Session, started: bool = False) -> None:
    if not s.lobby_msg:
        return
    try:
        await s.lobby_msg.edit_text(
            lobby_text(s, started),
            reply_markup=None if started else lobby_keyboard(s.sid),
        )
    except TelegramBadRequest as e:
        if "not modified" not in str(e):
            log.warning("Lobby yangilanmadi: %s", e)
    except TelegramRetryAfter:
        pass


def leaderboard_text(s: Session, title: str) -> str:
    ids = set(s.scores) | set(s.players)
    rows = sorted(ids, key=lambda u: (-s.scores.get(u, 0), s.names.get(u, "").lower()))
    lines = [title, ""]
    if not rows:
        lines.append("Hech kim qatnashmadi.")
    medals = ["🥇", "🥈", "🥉"]
    for i, uid in enumerate(rows[:50]):
        mark = medals[i] if i < 3 else f"{i + 1}."
        lines.append(f"{mark} {esc(s.names.get(uid, 'Noma‘lum'))} — <b>{s.scores.get(uid, 0)}</b> ball")
    return "\n".join(lines)


def cleanup(s: Session) -> None:
    """Sessiyani va unga bog'liq barcha ma'lumotlarni tozalaydi."""
    if SESSIONS.get(s.chat_id) is s:
        del SESSIONS[s.chat_id]
    for pid in s.poll_ids:
        POLLS.pop(pid, None)
    s.poll_ids.clear()
    s.poll_answers.clear()
    s.open_poll = None


# --------------------------------------------------------------------------
# Viktorina sikli
# --------------------------------------------------------------------------
async def ask_question(bot: Bot, s: Session, idx: int, q: Question) -> None:
    total = len(s.questions)
    prefix = f"{idx}/{total}. "
    long_mode = len(prefix + q.text) > POLL_Q_MAX or any(len(o) > POLL_OPT_MAX for o in q.options)

    if long_mode:
        # Uzun savol/variantlar: to'liq matn alohida xabarda, pollda esa harflar
        body = [f"<b>{idx}/{total}.</b> {esc(q.text)}", ""]
        body += [f"<b>{LETTERS[i]})</b> {esc(o)}" for i, o in enumerate(q.options)]
        explanation = "\n".join(body)
        # Telegram xabarining limiti 4096 belgi. HTML teglarini hisobga olib,
        # matnni xavfsizroq bo'laklarga ajratamiz.
        while explanation:
            if len(explanation) <= 3900:
                chunk, explanation = explanation, ""
            else:
                cut = explanation.rfind("\n", 0, 3900)
                if cut < 1:
                    cut = 3900
                chunk, explanation = explanation[:cut], explanation[cut:].lstrip("\n")
            await tg(bot.send_message, s.chat_id, chunk)
        poll_question = f"{prefix}Javobni tanlang"
        poll_options = list(LETTERS[: len(q.options)])
    else:
        poll_question = prefix + q.text
        poll_options = q.options

    poll_msg = await tg(
        bot.send_poll,
        chat_id=s.chat_id,
        question=poll_question,
        options=poll_options,
        type=PollType.QUIZ,
        is_anonymous=False,
        correct_option_id=q.correct,
        open_period=s.cfg["question_time"],
    )
    s.open_poll = poll_msg
    s.poll_ids.append(poll_msg.poll.id)
    POLLS[poll_msg.poll.id] = (s, q.correct)

    # Vaqt tugashini kutamiz (open_period o'zi yopadi; xavfsizlik uchun +1 soniya)
    await asyncio.sleep(s.cfg["question_time"] + 1)
    try:
        await bot.stop_poll(s.chat_id, poll_msg.message_id)
    except TelegramBadRequest:
        pass  # allaqachon yopilgan
    s.open_poll = None

    correct_text = q.options[q.correct]
    if long_mode:
        correct_text = f"{LETTERS[q.correct]}) {correct_text}"
    await tg(
        bot.send_message,
        s.chat_id,
        f"✅ To‘g‘ri javob: <b>{esc(correct_text[:1000])}</b>",
    )


async def session_main(bot: Bot, s: Session) -> None:
    try:
        # 1) Kutish xonasi
        try:
            await asyncio.wait_for(s.go.wait(), timeout=s.cfg["lobby_timeout"])
        except asyncio.TimeoutError:
            pass

        # 2) Boshlanish
        s.state = "running"
        await refresh_lobby(s, started=True)
        await tg(bot.send_message, s.chat_id, "🚦 Viktorina boshlanmoqda... Tayyor bo‘ling!")
        await asyncio.sleep(3)

        # 3) Savollar: bittasi to'liq yakunlanmaguncha keyingisi yuborilmaydi
        for idx, q in enumerate(s.questions, 1):
            await ask_question(bot, s, idx, q)
            if idx < len(s.questions):
                await asyncio.sleep(s.cfg["pause"])

        # 4) Reyting
        await tg(
            bot.send_message,
            s.chat_id,
            leaderboard_text(s, f"🏆 <b>Viktorina yakunlandi!</b> ({len(s.questions)} ta savol)"),
        )
    except asyncio.CancelledError:
        s.state = "stopped"
        if s.open_poll:
            try:
                await bot.stop_poll(s.chat_id, s.open_poll.message_id)
            except Exception:
                pass
        try:
            if s.scores:
                text = leaderboard_text(s, "🛑 <b>Viktorina to‘xtatildi.</b> Joriy reyting:")
            else:
                text = "🛑 Viktorina to‘xtatildi."
            await bot.send_message(s.chat_id, text)
        except Exception:
            pass
        raise
    except Exception:
        log.exception("Viktorinada xatolik")
        try:
            await bot.send_message(s.chat_id, "⚠️ Viktorinada kutilmagan xatolik yuz berdi. Viktorina to‘xtatildi.")
        except Exception:
            pass
    finally:
        cleanup(s)


# --------------------------------------------------------------------------
# Handlerlar
# --------------------------------------------------------------------------
HELP_TEXT = (
    "👋 Salom! Men viktorina botiman.\n\n"
    "📎 <b>.docx</b> yoki <b>.txt</b> test faylini yuboring.\n"
    f"Savollar soni, vaqt va boshqa parametrlar /settings orqali o‘zgartiriladi. Bitta viktorina uchun fayl izohiga (caption) raqam yozsangiz ham bo‘ladi, masalan: <code>30</code>.\n\n"
    "Format:\n"
    "<code>+++++\nSavol matni\n=====\n#To‘g‘ri javob\n=====\nVariant 2\n=====\nVariant 3</code>\n\n"
    "⚙️ /settings — sozlamalar\n"
    "🛑 /stop — faol viktorinani to‘xtatish."
)


@router.message(Command("start", "help"))
async def cmd_start(message: Message) -> None:
    await message.answer(HELP_TEXT)


@router.message(Command("stop"))
async def cmd_stop(message: Message, bot: Bot) -> None:
    s = SESSIONS.get(message.chat.id)
    if not s:
        await message.answer("Hozir faol viktorina yo‘q.")
        return
    user_id = message.from_user.id if message.from_user else 0
    allowed = user_id == s.host_id or message.chat.type == ChatType.PRIVATE
    if not allowed:
        try:
            member = await bot.get_chat_member(message.chat.id, user_id)
            allowed = member.status in (ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.CREATOR)
        except TelegramBadRequest:
            allowed = False
    if not allowed:
        await message.answer("Viktorinani faqat uni boshlagan foydalanuvchi yoki admin to‘xtata oladi.")
        return
    if s.task and not s.task.done():
        s.task.cancel()


def settings_text(cfg: dict) -> str:
    lines = ["⚙️ <b>Sozlamalar</b> (shu chat uchun)", ""]
    for key, (_lo, _hi, _step, icon, title, unit) in NUM_SETTINGS.items():
        lines.append(f"{icon} {title}: <b>{cfg[key]} {unit}</b>")
    lines.append(f"🔀 Savollar tartibi: <b>{'random' if cfg['shuffle_q'] else 'fayl tartibida'}</b>")
    lines.append(f"🔀 Variantlar: <b>{'aralashtiriladi' if cfg['shuffle_o'] else 'asl tartibda'}</b>")
    lines += [
        "",
        "ℹ️ O‘zgarishlar keyingi viktorinadan boshlab amal qiladi. "
        "Fayl izohidagi son esa savollar sonini faqat shu viktorina uchun almashtiradi.",
    ]
    return "\n".join(lines)


def settings_keyboard(cfg: dict):
    kb = InlineKeyboardBuilder()
    for key, (_lo, _hi, step, icon, title, unit) in NUM_SETTINGS.items():
        kb.row(
            InlineKeyboardButton(text=f"➖ {step}", callback_data=f"cfg:adj:{key}:dec"),
            InlineKeyboardButton(text=f"{icon} {title}: {cfg[key]} {unit}", callback_data="cfg:noop"),
            InlineKeyboardButton(text=f"➕ {step}", callback_data=f"cfg:adj:{key}:inc"),
        )
    kb.row(InlineKeyboardButton(
        text="🔀 Savollar: random" if cfg["shuffle_q"] else "📄 Savollar: fayl tartibida",
        callback_data="cfg:tog:shuffle_q",
    ))
    kb.row(InlineKeyboardButton(
        text="🔀 Variantlar: aralash" if cfg["shuffle_o"] else "📄 Variantlar: asl tartibda",
        callback_data="cfg:tog:shuffle_o",
    ))
    kb.row(
        InlineKeyboardButton(text="♻️ Standart", callback_data="cfg:reset"),
        InlineKeyboardButton(text="✖️ Yopish", callback_data="cfg:close"),
    )
    return kb.as_markup()


async def can_change_settings(bot: Bot, chat, user_id: int) -> bool:
    """Shaxsiy chatda hamma, guruhda faqat adminlar sozlamani o'zgartira oladi."""
    if chat.type == ChatType.PRIVATE:
        return True
    try:
        member = await bot.get_chat_member(chat.id, user_id)
    except TelegramBadRequest:
        return False
    return member.status in (ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.CREATOR)


@router.message(Command("settings"))
async def cmd_settings(message: Message) -> None:
    cfg = get_settings(message.chat.id)
    await message.answer(settings_text(cfg), reply_markup=settings_keyboard(cfg))


@router.callback_query(F.data.startswith("cfg:"))
async def cb_settings(cb: CallbackQuery, bot: Bot) -> None:
    if not cb.message:
        await cb.answer()
        return
    parts = cb.data.split(":")
    action = parts[1] if len(parts) > 1 else ""
    if action == "noop":
        await cb.answer()
        return

    chat = cb.message.chat
    if not await can_change_settings(bot, chat, cb.from_user.id):
        await cb.answer("Sozlamalarni faqat guruh adminlari o‘zgartira oladi.", show_alert=True)
        return

    if action == "close":
        try:
            await cb.message.delete()
        except TelegramBadRequest:
            pass
        await cb.answer()
        return

    cfg = get_settings(chat.id)
    if action == "reset":
        reset_settings(chat.id)
    elif action == "adj" and len(parts) == 4 and parts[2] in NUM_SETTINGS and parts[3] in ("inc", "dec"):
        key, direction = parts[2], parts[3]
        lo, hi, step = NUM_SETTINGS[key][:3]
        new = max(lo, min(hi, cfg[key] + (step if direction == "inc" else -step)))
        if new == cfg[key]:
            await cb.answer(f"Chegara: {lo}–{hi}", show_alert=False)
            return
        update_settings(chat.id, **{key: new})
    elif action == "tog" and len(parts) == 3 and parts[2] in BOOL_SETTINGS:
        update_settings(chat.id, **{parts[2]: not cfg[parts[2]]})
    else:
        await cb.answer()
        return

    cfg = get_settings(chat.id)
    try:
        await cb.message.edit_text(settings_text(cfg), reply_markup=settings_keyboard(cfg))
    except TelegramBadRequest as e:
        if "not modified" not in str(e):
            log.warning("Sozlamalar xabari yangilanmadi: %s", e)
    await cb.answer("Saqlandi ✅")


@router.message(F.document)
async def on_document(message: Message, bot: Bot) -> None:
    doc = message.document
    ext = os.path.splitext((doc.file_name or "").lower())[1]
    private = message.chat.type == ChatType.PRIVATE
    if ext not in (".docx", ".txt"):
        if private:
            await message.answer("Faqat <b>.docx</b> yoki <b>.txt</b> fayllar qabul qilinadi.")
        return

    chat_id = message.chat.id
    if chat_id in SESSIONS:
        await message.answer("Bu chatda viktorina allaqachon davom etmoqda. Avval uni tugating yoki /stop yozing.")
        return
    if doc.file_size and doc.file_size > MAX_FILE_SIZE:
        await message.answer("Fayl juda katta (maksimum 20 MB).")
        return

    # Faylni vaqtinchalik joyga yuklab olamiz va albatta o'chiramiz
    fd, tmp_path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    try:
        await bot.download(doc, destination=tmp_path)
        reader = read_docx if ext == ".docx" else read_txt
        text = await asyncio.to_thread(reader, tmp_path)
        pool = await asyncio.to_thread(parse_questions, text)
    except Exception:
        log.exception("Faylni o‘qishda xatolik")
        await message.answer("⚠️ Faylni o‘qib bo‘lmadi. Fayl buzilmaganini va formatga mosligini tekshiring.")
        return
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    if not pool:
        await message.answer(
            "❌ Faylda birorta ham to‘g‘ri savol topilmadi.\n"
            "<code>+++++</code>, <code>=====</code> va <code>#</code> belgilarini tekshiring."
        )
        return

    cfg = get_settings(chat_id)
    want = cfg["count"]
    m = re.search(r"\d+", message.caption or "")
    if m:
        want = max(1, int(m.group()))
    count = min(want, len(pool))

    # Parse paytida boshqa fayl kelgan bo'lishi mumkin — qayta tekshiramiz
    if chat_id in SESSIONS:
        await message.answer("Bu chatda viktorina allaqachon davom etmoqda.")
        return

    s = Session(
        sid=uuid.uuid4().hex[:8],
        chat_id=chat_id,
        host_id=message.from_user.id if message.from_user else 0,
        questions=prepare_questions(pool, count, cfg["shuffle_q"], cfg["shuffle_o"]),
        cfg=cfg,
    )
    SESSIONS[chat_id] = s

    note = f"Bazada {len(pool)} ta savol bor." if want > len(pool) else ""
    try:
        s.lobby_msg = await message.answer(lobby_text(s), reply_markup=lobby_keyboard(s.sid))
        if note:
            await message.answer(f"ℹ️ {note} Hammasi ishlatiladi.")
    except Exception:
        cleanup(s)
        raise
    s.task = asyncio.create_task(session_main(bot, s))


def _find_session(cb: CallbackQuery, sid: str) -> Optional[Session]:
    if not cb.message:
        return None
    s = SESSIONS.get(cb.message.chat.id)
    return s if s and s.sid == sid else None


@router.callback_query(F.data.startswith("join:"))
async def cb_join(cb: CallbackQuery) -> None:
    s = _find_session(cb, cb.data.split(":", 1)[1])
    if not s or s.state != "lobby":
        await cb.answer("Bu viktorina boshlangan yoki tugagan.", show_alert=True)
        return
    user = cb.from_user
    if user.id in s.players:
        await cb.answer("Siz allaqachon qo‘shilgansiz ✅")
        return
    s.players[user.id] = user.full_name
    s.names[user.id] = user.full_name
    s.scores.setdefault(user.id, 0)
    await cb.answer("Qo‘shildingiz! ✅")
    await refresh_lobby(s)


@router.callback_query(F.data.startswith("go:"))
async def cb_go(cb: CallbackQuery) -> None:
    s = _find_session(cb, cb.data.split(":", 1)[1])
    if not s or s.state != "lobby":
        await cb.answer("Bu viktorina boshlangan yoki tugagan.", show_alert=True)
        return
    if cb.from_user.id != s.host_id:
        await cb.answer("Viktorinani faqat fayl yuborgan foydalanuvchi boshlay oladi.", show_alert=True)
        return
    s.go.set()
    await cb.answer("Boshlanmoqda! 🚀")


@router.poll_answer()
async def on_poll_answer(pa: PollAnswer) -> None:
    entry = POLLS.get(pa.poll_id)
    if not entry or not pa.user:
        return
    s, correct = entry
    user = pa.user
    s.names[user.id] = user.full_name
    s.scores.setdefault(user.id, 0)

    answers = s.poll_answers.setdefault(pa.poll_id, {})
    was_correct = answers.get(user.id, False)
    is_correct = bool(pa.option_ids and pa.option_ids[0] == correct)
    if is_correct != was_correct:
        s.scores[user.id] += 1 if is_correct else -1
        s.scores[user.id] = max(0, s.scores[user.id])
    answers[user.id] = is_correct


# --------------------------------------------------------------------------
# Ishga tushirish
# --------------------------------------------------------------------------
def acquire_single_instance_lock() -> socket.socket:
    """Ikkinchi nusxa ishga tushmasligi uchun lokal portni band qiladi."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", LOCK_PORT))
    except OSError:
        print("❌ Bot allaqachon ishlayapti (boshqa terminal yoki jarayonda). Avval uni to‘xtating.")
        sys.exit(1)
    return sock


async def main() -> None:
    token = os.getenv("BOT_TOKEN")
    if not token:
        sys.exit("❌ BOT_TOKEN muhit o‘zgaruvchisi topilmadi. Masalan: $env:BOT_TOKEN = \"123456:ABC...\"")

    load_settings()

    bot = Bot(token=token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()
    dp.include_router(router)

    await bot.delete_webhook(drop_pending_updates=True)
    await bot.set_my_commands([
        BotCommand(command="start", description="Yordam"),
        BotCommand(command="settings", description="Sozlamalar"),
        BotCommand(command="stop", description="Viktorinani to'xtatish"),
    ])
    me = await bot.get_me()
    log.info("Bot ishga tushdi: @%s", me.username)
    try:
        await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
    finally:
        for s in list(SESSIONS.values()):
            if s.task and not s.task.done():
                s.task.cancel()
        await bot.session.close()


if __name__ == "__main__":
    _lock = acquire_single_instance_lock()
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        log.info("Bot to‘xtatildi.")