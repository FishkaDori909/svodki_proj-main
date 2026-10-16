#!/usr/bin/env python3
"""Диагностика почтового ящика (mbox): показывает структуру каждого письма и вложений.
Нужна, если вложения не находятся. Только чтение, ничего не меняет.

Запуск: python tools/diagnose_mbox.py "/путь/к/почтовому/ящику"
"""
import mailbox
import sys

if len(sys.argv) < 2:
    sys.exit('Укажите путь к почтовому ящику: python tools/diagnose_mbox.py "/путь/к/ящику"')

box = mailbox.mbox(sys.argv[1], create=False)


def dump(part, depth=0):
    indent = "  " * depth
    size = len(part.get_payload(decode=True) or b"") if not part.is_multipart() else "-"
    print(f"{indent}- тип={part.get_content_type()} имя_файла={part.get_filename()!r} "
          f"вложение={part.get('Content-Disposition')!r} размер={size}")
    if part.is_multipart():
        for sub in part.get_payload():
            dump(sub, depth + 1)


for key in box.iterkeys():
    msg = box.get_message(key)
    print(f"=== письмо #{key} ===")
    print("Тема:", msg.get("Subject"))
    print("От:", msg.get("From"))
    print("Состав письма:")
    dump(msg)
    print()
