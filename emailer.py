# -*- coding: utf-8 -*-
"""
Mock-deck Generator — 이메일 전송 (생성된 .pptx 를 팀원에게 발송)

SMTP(기본 Gmail)로 .pptx 를 첨부해 보낸다. 보내는 사람 Gmail + 앱 비밀번호가 필요하다.
(Gmail은 일반 비밀번호가 아니라 '앱 비밀번호'를 발급받아 써야 함 — 2단계 인증 활성화 후
 https://myaccount.google.com/apppasswords 에서 발급.)
"""

import os
import smtplib
from email.message import EmailMessage

DEFAULT_SMTP_HOST = "smtp.gmail.com"
DEFAULT_SMTP_PORT = 587
PPTX_MIME = ("application", "vnd.openxmlformats-officedocument.presentationml.presentation")


def build_message(sender, recipients, subject, body, attachment_path):
    """MIME 메시지 구성(첨부 포함). 전송과 분리해 테스트 가능."""
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            data = f.read()
        msg.add_attachment(data, maintype=PPTX_MIME[0], subtype=PPTX_MIME[1],
                           filename=os.path.basename(attachment_path))
    return msg


def parse_recipients(text):
    """쉼표/세미콜론/공백/줄바꿈으로 구분된 수신자 문자열을 리스트로."""
    raw = text.replace(";", ",").replace("\n", ",").replace(" ", ",")
    return [r.strip() for r in raw.split(",") if r.strip()]


def send_email(sender, password, recipients, subject, body, attachment_path,
               smtp_host=DEFAULT_SMTP_HOST, smtp_port=DEFAULT_SMTP_PORT):
    """SMTP(STARTTLS)로 발송. 실패 시 예외 발생."""
    if not sender or not password:
        raise ValueError("보내는 사람 이메일과 앱 비밀번호가 필요합니다.")
    if not recipients:
        raise ValueError("받는 사람이 비어 있습니다.")
    msg = build_message(sender, recipients, subject, body, attachment_path)
    with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as s:
        s.starttls()
        s.login(sender, password)
        s.send_message(msg)
    return True
