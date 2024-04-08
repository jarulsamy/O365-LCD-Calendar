#!/usr/bin/env python3
import sys
from datetime import date, timedelta, datetime
from zoneinfo import ZoneInfo

import board
import digitalio
from O365 import Account, FileSystemTokenBackend, MSGraphProtocol
from PIL import Image, ImageDraw, ImageFont
import socket
import time
import humanize

from collections.abc import Iterator
import concurrent.futures

from functools import partial
import threading
import queue

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.settimeout(0)
    try:
        # doesn't even have to be reachable
        s.connect(("10.254.254.254", 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = "127.0.0.1"
    finally:
        s.close()
    return IP


def get_now():
    return datetime.now(ZoneInfo("America/Denver"))

def o365_auth(
    client_id,
    client_secret,
    token_path="/home/joshua/cal/tokens",
    token_filename="token.txt",
):
    credentials = (client_id, client_secret)

    protocol = MSGraphProtocol()
    scopes = [
        "offline_access",
        "Calendars.Read",
        "Calendars.Read.Shared",
        "Calendars.ReadBasic",
    ]

    token_backend = FileSystemTokenBackend(
        token_path=token_path,
        token_filename=token_filename,
    )
    token = token_backend.load_token()
    account = Account(credentials, protocol=protocol, token_backend=token_backend)

    if not account.is_authenticated:
        account.authenticate(scopes=scopes)

    return account

def get_events(account):
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()

    today = date.today()
    tomorrow = today + timedelta(days=1)

    now = get_now()
    q = calendar.new_query("start").less_equal(now)
    q.chain("and").on_attribute("end").greater_equal(now)
    events = calendar.get_events(
        limit=1,
        query=q,
        include_recurring=True,
        order_by="start/dateTime",
    )

    return events


def scroll(seq, width=16):
    n = len(seq)
    if n < width:
        yield seq
        return

    for i in range(n - width + 1):
        yield seq[i : i + width]
    end = n + 1
    for j in range(n - width + 1, end):
        l = width - (end - j) + 1
        res = f"{seq[j:end]} {seq[0:l]}"
        yield res


def lcd_setup():
    from adafruit_character_lcd.character_lcd import Character_LCD_Mono

    lcd_rs = digitalio.DigitalInOut(board.D25)
    lcd_en = digitalio.DigitalInOut(board.D24)
    lcd_d4 = digitalio.DigitalInOut(board.D23)
    lcd_d5 = digitalio.DigitalInOut(board.D17)
    lcd_d6 = digitalio.DigitalInOut(board.D18)
    lcd_d7 = digitalio.DigitalInOut(board.D22)
    # lcd_backlight = digitalio.DigitalInOut(board.
    lcd_columns = 16
    lcd_rows = 2

    lcd = Character_LCD_Mono(
        lcd_rs,
        lcd_en,
        lcd_d4,
        lcd_d5,
        lcd_d6,
        lcd_d7,
        lcd_columns,
        lcd_rows,
        None,
    )

    ip = socket.gethostbyname(socket.gethostname())
    lcd.message = str(get_ip())

    return lcd


current_event = None

def event_thread(account):
    global current_event
    while True:
        events = list(get_events(account))
        if events:
            current_event = events[0]
        else:
            current_event = None
        time.sleep(30)


def lcd_thread(lcd):
    global current_event
    lcd.clear()
    while True:

        lcd.clear()
        lcd.message = "Please Knock"

        now = get_now()
        while current_event is not None and current_event.start <= now <= current_event.end:
            lcd.cursor_position(0, 0)
            msg = "In a Meeting"
            lcd.message = msg.ljust(15)

            now = get_now()
            try:
                remaining = current_event.end - now
            except AttributeError:
                break
            delta = humanize.precisedelta(
                        remaining,
                        suppress=["seconds"],
                        format="%0.0f",
                    )
            second_line = f"{delta} left".ljust(15)
            for line in scroll(second_line):
                lcd.cursor_position(0, 1)
                lcd.message = line
                time.sleep(0.3)
            time.sleep(3)
        time.sleep(3)


def main():
    CLIENT_ID = "<REDACTED>"
    CLIENT_SECRET = "<REDACTED>"

    account = o365_auth(CLIENT_ID, CLIENT_SECRET)
    lcd = lcd_setup()

    event_queue = queue.Queue(maxsize=1)
    interrupt = threading.Event()

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        executor.submit(partial(event_thread, account))
        lcd_thread(lcd)

    return 0


def lcd_loop(lcd, client_id, client_secret):
    while True:
        lcd.clear()
        events = list(get_events(client_id, client_secret))
        if not len(events):
            lcd.clear()
            lcd.message = "Please knock."
            time.sleep(60)
            continue

        current = events[0]
        second_line = current.subject

        now = get_now()
        event_poll = 0
        while current.start <= now <= current.end:
            if event_poll > 32:
                event_poll = 0
                events = list(get_events(client_id, client_secret))
                if not events:
                    break
                current = events[0]
            event_poll += 1

            lcd.cursor_position(0, 0)
            lcd.message = "In a Meeting"

            now = get_now()
            remaining = current.end - now

            delta = humanize.precisedelta(
                remaining, suppress=["seconds"], format="%0.0f"
            )

            second_line = f"{delta} left"
            if len(second_line) > 16:
                for line in Scroll(second_line):
                    lcd.cursor_position(0, 1)
                    lcd.message = line
                    time.sleep(0.25)
                time.sleep(1)
            else:
                lcd.cursor_position(0, 1)
                lcd.message = second_line
                time.sleep(1)
        time.sleep(30)




if __name__ == "__main__":
    sys.exit(main())
