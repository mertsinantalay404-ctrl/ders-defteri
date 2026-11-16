#!/usr/bin/env python3
"""Join a Teams class at the scheduled start, then leave below 10 participants."""

import datetime
import json
import os
import re
import sys
import time
from urllib.parse import parse_qs, quote, urlparse
from zoneinfo import ZoneInfo

TZ = ZoneInfo("Europe/Istanbul")
THRESHOLD = 10


def log(message):
    now = datetime.datetime.now(TZ).isoformat(timespec="seconds")
    print(f"{now} {message}", flush=True)


def meeting_join_url(url):
    raw = (url or "").strip()
    if not raw:
        return ""
    href = re.sub(r"^msteams:", "https:", raw, flags=re.I)
    parsed = urlparse(href)
    match = re.search(r"/l/(?:team|channel)/([^/]+)", parsed.path, re.I)
    if not match:
        return raw
    thread = match.group(1)
    tenant = (parse_qs(parsed.query).get("tenantId") or [""])[0]
    context = quote(json.dumps({"Tid": tenant}, separators=(",", ":"))) if tenant else ""
    join = f"https://teams.microsoft.com/l/meetup-join/{thread}/0"
    if context:
        join += f"?context={context}"
    return join


def frames_of(page):
    frames = []
    for frame in page.frames:
        if frame not in frames:
            frames.append(frame)
    return frames


def click_first(page, names, timeout=8000):
    deadline = time.time() + timeout / 1000
    patterns = [re.compile(rf"^{re.escape(name)}$", re.I) for name in names]
    while time.time() < deadline:
        for frame in frames_of(page):
            for pattern in patterns:
                for role in ("button", "link"):
                    loc = frame.get_by_role(role, name=pattern)
                    try:
                        if loc.count() and loc.first.is_visible():
                            loc.first.click(timeout=2500)
                            return True
                    except Exception:
                        continue
        time.sleep(0.2)
    return False


def visible_button(page, names, timeout=1500):
    deadline = time.time() + timeout / 1000
    patterns = [re.compile(rf"^{re.escape(name)}$", re.I) for name in names]
    while time.time() < deadline:
        for frame in frames_of(page):
            for pattern in patterns:
                loc = frame.get_by_role("button", name=pattern)
                try:
                    if loc.count() and loc.first.is_visible():
                        return True
                except Exception:
                    continue
        time.sleep(0.2)
    return False


def locator_visible(page, selector):
    for frame in frames_of(page):
        loc = frame.locator(selector)
        try:
            if loc.count() and loc.first.is_visible():
                return loc.first
        except Exception:
            continue
    return None


def fill_first(page, selectors, value):
    for selector in selectors:
        loc = page.locator(selector)
        try:
            loc.first.wait_for(state="visible", timeout=8000)
            loc.first.fill(value)
            return True
        except Exception:
            continue
    return False


def parse_hm(value):
    parts = str(value or "18:00").split(":")
    return int(parts[0]), int(parts[1] if len(parts) > 1 else 0)


def class_plan(now):
    start_h, start_m = parse_hm(os.environ.get("TEAMS_CLASS_START", "18:00"))
    end_by_day = {
        0: os.environ.get("TEAMS_CLASS_END_HALK", "20:00"),
        3: os.environ.get("TEAMS_CLASS_END_FIZYO", "21:00"),
        4: os.environ.get("TEAMS_CLASS_END_ANATOMI", "21:00"),
    }
    end_h, end_m = parse_hm(end_by_day.get(now.weekday(), "21:00"))
    start = now.replace(hour=start_h, minute=start_m, second=0, microsecond=0)
    end = now.replace(hour=end_h, minute=end_m, second=0, microsecond=0)
    return start, end


def join_url_for_today(now):
    by_day = {
        0: os.environ.get("TEAMS_JOIN_HALK", "").strip(),
        3: os.environ.get("TEAMS_JOIN_FIZYO", "").strip(),
        4: os.environ.get("TEAMS_JOIN_ANATOMI", "").strip(),
    }
    url = meeting_join_url(by_day.get(now.weekday(), "") or os.environ.get("TEAMS_JOIN_URL", "").strip())
    if not url or "launchAgent=" in url:
        return url
    sep = "&" if "?" in url else "?"
    return url + sep + "launchAgent=join_only"


def wait_until(target):
    while True:
        now = datetime.datetime.now(TZ)
        remaining = (target - now).total_seconds()
        if remaining <= 0:
            return now
        time.sleep(0.05 if remaining < 2 else min(1.0, remaining))


LEAVE_NAMES = ["Ayrıl", "Leave", "Hang up", "Çağrıyı sonlandır", "Leave call", "Leave meeting"]
JOIN_SELECTORS = [
    "[data-tid='prejoin-join-button']",
    "[data-tid='join-button']",
    "[data-tid='calling-join-button']",
]
HANGUP_SELECTORS = [
    "[data-tid='call-hangup']",
    "[data-tid='hangup-button']",
    "[data-tid='calling-hangup-button']",
]
CONTINUE_NAMES = [
    "Tarayıcıda devam et",
    "Web üzerinde katıl",
    "Web'de katıl",
    "Bunun yerine web uygulamasını kullan",
    "Bunun yerine web’de katıl",
    "Continue on this browser",
    "Join on the web instead",
    "Use the web app instead",
]
AUDIO_NAMES = [
    "Ses kullanma",
    "Don't use audio",
    "Bilgisayar sesi",
    "Computer audio",
]
JOIN_NAMES = [
    "Şimdi katıl",
    "Toplantıya katıl",
    "Ses kapalı olarak katıl",
    "Katılma isteği gönder",
    "Katıl",
    "Join now",
    "Join meeting",
    "Ask to join",
]


def page_hint(page):
    parsed = urlparse(page.url)
    where = f"{parsed.netloc}{parsed.path}"[:160]
    labels = []
    for frame in frames_of(page)[:6]:
        try:
            found = frame.locator("button, a, [role='button']").evaluate_all(
                """els => els
                    .map(el => (el.innerText || el.getAttribute('aria-label') || '').replace(/\\s+/g, ' ').trim())
                    .filter(text => text && text.length < 60)
                    .slice(0, 12)"""
            )
        except Exception:
            continue
        for label in found or []:
            if "@" in label or label in labels:
                continue
            labels.append(label)
            if len(labels) >= 12:
                return where, labels
    return where, labels


def field_visible(page, selectors):
    for selector in selectors:
        if locator_visible(page, selector):
            return True
    return False


def mfa_wall(page):
    if "login.microsoftonline.com" not in page.url and "login.live.com" not in page.url:
        return False
    try:
        text = page.inner_text("body")
    except Exception:
        return False
    folded = text.lower()
    return any(
        token in folded
        for token in ("authenticator", "onaylayın", "doğrulama kodu", "approve a request")
    )


def advance(page, user, password, done):
    if mfa_wall(page):
        log("mfa_required")
        return "mfa"
    if field_visible(
        page,
        [
            'input[name="loginfmt"]',
            'input[type="email"]',
        ],
    ):
        filled = fill_first(
            page,
            ['input[name="loginfmt"]', 'input[type="email"]'],
            user,
        )
        if filled:
            click_first(page, ["İleri", "Next"], timeout=4000)
            log("login_email")
            return "email"
    if click_first(
        page,
        ["Parolanızı kullanın", "Parola kullan", "Use your password"],
        timeout=600,
    ):
        log("login_use_password")
        return "password_choice"
    if field_visible(page, ['input[name="passwd"]', 'input[type="password"]']):
        filled = fill_first(
            page,
            ['input[name="passwd"]', 'input[type="password"]'],
            password,
        )
        if filled:
            click_first(page, ["Oturum açın", "Oturum aç", "Sign in"], timeout=4000)
            log("login_password")
            return "password"
    if (
        "stay" not in done
        and "login.microsoftonline.com" in page.url
        and click_first(page, ["Evet", "Yes"], timeout=500)
    ):
        done.add("stay")
        log("login_stay")
        return "stay"
    web = locator_visible(page, "[data-tid='joinOnWeb'], #openTeamsClientInBrowser")
    if web is not None:
        try:
            web.click(timeout=2500)
            log("web_join")
            return "web"
        except Exception:
            pass
    if click_first(page, CONTINUE_NAMES, timeout=700):
        log("web_join")
        return "web"
    if "audio" not in done and click_first(page, AUDIO_NAMES, timeout=400):
        done.add("audio")
        log("audio_selected")
        return "audio"
    if "camera" not in done and click_first(page, ["Kamerayı kapat", "Turn camera off"], timeout=300):
        done.add("camera")
        log("camera_off")
        return "camera"
    if "mic" not in done and click_first(page, ["Mikrofonu kapat", "Mute microphone"], timeout=300):
        done.add("mic")
        log("mic_off")
        return "mic"
    return "wait"


def join_control(page):
    for selector in JOIN_SELECTORS:
        found = locator_visible(page, selector)
        if found is not None:
            return found
    deadline = time.time() + 0.4
    patterns = [re.compile(rf"^{re.escape(name)}$", re.I) for name in JOIN_NAMES]
    for frame in frames_of(page):
        for pattern in patterns:
            loc = frame.get_by_role("button", name=pattern)
            try:
                if loc.count() and loc.first.is_visible():
                    return loc.first
            except Exception:
                continue
        if time.time() > deadline:
            break
    return None


def click_join(page):
    control = join_control(page)
    if control is None:
        return False
    try:
        if control.is_disabled():
            now = time.time()
            if now - getattr(click_join, "disabled_logged_at", 0) > 15:
                log("join_button_disabled")
                click_join.disabled_logged_at = now
            return False
        control.click(timeout=3000)
        return True
    except Exception as error:
        log(f"join_click_error {type(error).__name__}")
        return False


def in_meeting(page):
    for selector in HANGUP_SELECTORS:
        if locator_visible(page, selector) is not None:
            return True
    return visible_button(page, LEAVE_NAMES, timeout=400)


def confirm_joined(page, timeout=45):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if in_meeting(page):
            return True
        time.sleep(1.5)
    return False


def participant_count(page):
    try:
        click_first(
            page,
            ["Katılımcılar", "Kişiler", "People", "Show participants", "Roster"],
            timeout=1500,
        )
        time.sleep(1)
        chunks = [page.inner_text("body")]
        labels = page.eval_on_selector_all(
            "[aria-label]",
            "els => els.map(el => el.getAttribute('aria-label') || '')",
        )
        chunks.append("\n".join(labels or []))
        text = "\n".join(chunks)
        found = [
            int(num)
            for num in re.findall(
                r"(\d+)\s*(?:katılımcı|kişiler|participants?|people)\b",
                text,
                flags=re.I,
            )
        ]
        if not found:
            return None
        return max(found)
    except Exception as error:
        log(f"participant_read_failed {error}")
        return None


def leave_meeting(page):
    for selector in HANGUP_SELECTORS:
        found = locator_visible(page, selector)
        if found is None:
            continue
        try:
            found.click(timeout=3000)
            return True
        except Exception:
            continue
    return click_first(page, LEAVE_NAMES, timeout=8000)


def watch_and_leave(page, max_seconds):
    seen_full = False
    left = False
    last_count = None
    deadline = time.time() + max_seconds
    log("watch_start leave_below=9 stay_at_or_above=10")
    while time.time() < deadline:
        if not in_meeting(page):
            log("meeting_ui_missing skip_leave")
            time.sleep(15)
            continue
        count = participant_count(page)
        if count is None:
            log("participants=unreadable skip_leave")
            time.sleep(15)
            continue
        last_count = count
        log(f"participants={count}")
        if count >= THRESHOLD:
            seen_full = True
        elif seen_full and count <= THRESHOLD - 1 and not left:
            log(f"leave_at={datetime.datetime.now(TZ).isoformat(timespec='seconds')} participants={count}")
            leave_meeting(page)
            left = True
            log("left_once")
            return last_count
        time.sleep(15)
    if not left and in_meeting(page):
        log(
            f"leave_at={datetime.datetime.now(TZ).isoformat(timespec='seconds')} "
            f"participants={last_count if last_count is not None else 'unknown'} reason=max_time"
        )
        leave_meeting(page)
    return last_count


def main():
    user = os.environ.get("TEAMS_USER", "").strip()
    password = os.environ.get("TEAMS_PASSWORD", "").strip()
    wait_for_start = os.environ.get("TEAMS_WAIT_FOR_START", "1") == "1"
    max_seconds = int(os.environ.get("TEAMS_MAX_SECONDS", "14400"))
    now = datetime.datetime.now(TZ)
    start, end = class_plan(now)
    url = join_url_for_today(now)
    log(f"weekday={now.weekday()} class_start={start.isoformat()} class_end={end.isoformat()} url_ok={bool(url)}")
    if not user or not password or not url:
        log("missing secrets")
        sys.exit(1)

    if now > end:
        log("too_late class_already_ended")
        return
    if now > start + datetime.timedelta(minutes=3):
        late = (now - start).total_seconds()
        log(f"too_late skip_join offset_s={late:.0f}")
        return

    from playwright.sync_api import sync_playwright

    login_at = start - datetime.timedelta(minutes=6)
    if wait_for_start:
        remaining_to_login = (login_at - datetime.datetime.now(TZ)).total_seconds()
        if remaining_to_login > 0:
            log(f"waiting_for_login remaining_s={remaining_to_login:.3f}")
            wait_until(login_at)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=True,
            args=[
                "--use-fake-ui-for-media-stream",
                "--use-fake-device-for-media-stream",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        context = browser.new_context(
            locale="tr-TR",
            viewport={"width": 1366, "height": 768},
            permissions=["microphone", "camera"],
        )
        page = context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=120000)
        done = set()
        ready_logged = False
        if wait_for_start:
            while datetime.datetime.now(TZ) < start:
                if join_control(page) is not None:
                    if not ready_logged:
                        log("prejoin_ready")
                        ready_logged = True
                else:
                    state = advance(page, user, password, done)
                    if state == "mfa":
                        context.close()
                        browser.close()
                        sys.exit(1)
                remaining = (start - datetime.datetime.now(TZ)).total_seconds()
                if remaining <= 0:
                    break
                time.sleep(0.25 if remaining < 3 else min(1.0, remaining))
        else:
            advance(page, user, password, done)

        joined = in_meeting(page)
        clicked = False
        join_deadline = max(start, datetime.datetime.now(TZ)) + datetime.timedelta(seconds=90)
        while not joined and datetime.datetime.now(TZ) < join_deadline:
            if click_join(page):
                click_at = datetime.datetime.now(TZ)
                offset = (click_at - start).total_seconds()
                log(f"join_click_at={click_at.isoformat(timespec='milliseconds')} offset_s={offset:.3f}")
                clicked = True
                joined = confirm_joined(page, timeout=20)
                if joined:
                    break
            state = advance(page, user, password, done)
            if state == "mfa":
                context.close()
                browser.close()
                sys.exit(1)
            if in_meeting(page):
                joined = True
                break
            time.sleep(0.4)

        log(f"joined={joined}")
        if not joined:
            where, labels = page_hint(page)
            log(f"join_not_confirmed skip_leave_watch clicked={clicked} page={where} controls={labels}")
            context.close()
            browser.close()
            sys.exit(1)

        watch_and_leave(page, max_seconds=max_seconds)
        context.close()
        browser.close()


if __name__ == "__main__":
    main()
