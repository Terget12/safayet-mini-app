#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Paid Number Telegram Bot — single-file edition

Features
--------
• Telegram user panel + Telegram admin panel
• SQLite database (created automatically)
• Live API price and stock
• Global markup (default $0.005) and per-item markup override
• Country and service catalog controlled by admin
• Country → service → live provider-price selection for users
• Balance deposits credited only after automatic payment-provider confirmation
• Number purchase, SMS polling, OTP delivery, cancellation and refund
• Cancellation is blocked after an OTP is received
• Expired orders are not refunded (as requested)
• Order history, transaction ledger, statistics and broadcast
• getPricesV3 provider prices/stock and providerIds purchasing
• Admin-configurable Support ID/link with direct Send Message button
• Styled Telegram buttons with safe compatibility fallback
• Admin-set minimum deposit and per-user balance management
• Cheapest-first provider grid: 2 columns, 10 items per page
• Two OTP forwarding groups with private admin balance details

Secret setup
------------
Set BOT_TOKEN, ADMIN_IDS, and IPSHOP_API_KEY in the hosting Environment panel
or in a .env file beside this script. SMSBower and Local USA API keys can be
managed from Admin → Panel Management; SMS_API_KEY and LOCAL_API_KEY are
optional fallback values. Hosting Environment values take priority. Never put
real credentials in this Python file or share .env publicly.

Install:
    python -m pip install "python-telegram-bot[rate-limiter]>=22.8,<23" "aiohttp>=3.9,<4" "openpyxl>=3.1,<4"

Run:
    python paid_number_bot.py

Non-secret URLs and defaults are configured in the CONFIG section below.

API compatibility
-----------------
This adapter supports the common action-based API flow:
getBalance, getCountries, getServicesList, getPricesV3, getNumberV2, getStatus,
setStatus. It accepts both plain-text SMS-Activate-style responses and common
JSON response shapes. Set SMS_API_URL to the exact handler URL shown in your
SMS Power documentation.
"""

from __future__ import annotations

import json
import csv
import aiohttp
from aiohttp import web
import asyncio
import html
import inspect
import logging
from logging.handlers import RotatingFileHandler
import math
import os
import sys
import re

# Load .env from the same directory as this script. Variables already supplied
# by the hosting panel take priority over values in .env.
def _load_env_file(env_path: str | None = None) -> None:
    env_path = env_path or os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.isfile(env_path):
        return
    try:
        with open(env_path, "r", encoding="utf-8") as env_file:
            for raw_line in env_file:
                line = raw_line.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("export "):
                    line = line[7:].lstrip()
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()
                if not key or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key):
                    continue
                if len(value) >= 2 and value[0] == value[-1] and value[0] in ("\"", "'"):
                    value = value[1:-1]
                else:
                    # Treat a whitespace-prefixed # as an unquoted comment.
                    value = re.split(r"\s+#", value, maxsplit=1)[0].rstrip()
                os.environ.setdefault(key, value)
    except OSError as exc:
        print(f"WARNING: Could not read .env file: {exc}", file=sys.stderr)


_load_env_file()
import re
import sqlite3
import importlib.util
import threading
import time
import tempfile
import shutil
import zipfile
import unicodedata
from datetime import datetime
from urllib.parse import quote
from contextlib import suppress
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Iterable, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from telegram import (
    BotCommand,
    CopyTextButton,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
)
from telegram.constants import ParseMode
from telegram.error import BadRequest, Forbidden, NetworkError, RetryAfter, TelegramError
from telegram.ext import (
    Application,
    ApplicationBuilder,
    AIORateLimiter,
    BaseUpdateProcessor,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)  






_APP_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
_SMSBOWER_COUNTRY_DATA = (
    ('74', 'Afghanistan', '🇦🇫'),
    ('155', 'Albania', '🇦🇱'),
    ('58', 'Algeria', '🇩🇿'),
    ('76', 'Angola', '🇦🇴'),
    ('181', 'Anguilla', '🇦🇮'),
    ('169', 'Antigua and Barbuda', '🇦🇬'),
    ('39', 'Argentina', '🇦🇷'),
    ('148', 'Armenia', '🇦🇲'),
    ('179', 'Aruba', '🇦🇼'),
    ('175', 'Australia', '🇦🇺'),
    ('50', 'Austria', '🇦🇹'),
    ('35', 'Azerbaijan', '🇦🇿'),
    ('122', 'Bahamas', '🇧🇸'),
    ('145', 'Bahrain', '🇧🇭'),
    ('60', 'Bangladesh', '🇧🇩'),
    ('118', 'Barbados', '🇧🇧'),
    ('51', 'Belarus', '🇧🇾'),
    ('82', 'Belgium', '🇧🇪'),
    ('124', 'Belize', '🇧🇿'),
    ('120', 'Benin', '🇧🇯'),
    ('1003', 'Bermuda', '🇧🇲'),
    ('158', 'Bhutan', '🇧🇹'),
    ('92', 'Bolivia', '🇧🇴'),
    ('108', 'Bosnia and Herzegovina', '🇧🇦'),
    ('123', 'Botswana', '🇧🇼'),
    ('73', 'Brazil', '🇧🇷'),
    ('121', 'Brunei Darussalam', '🇧🇳'),
    ('83', 'Bulgaria', '🇧🇬'),
    ('152', 'Burkina Faso', '🇧🇫'),
    ('119', 'Burundi', '🇧🇮'),
    ('24', 'Cambodia', '🇰🇭'),
    ('41', 'Cameroon', '🇨🇲'),
    ('36', 'Canada', '🇨🇦'),
    ('186', 'Cape Verde', '🇨🇻'),
    ('170', 'Cayman Islands', '🇰🇾'),
    ('125', 'Central African Republic', '🇨🇫'),
    ('42', 'Chad', '🇹🇩'),
    ('151', 'Chile', '🇨🇱'),
    ('3', 'China', '🇨🇳'),
    ('33', 'Colombia', '🇨🇴'),
    ('133', 'Comoros', '🇰🇲'),
    ('150', 'Congo', '🇨🇬'),
    ('18', 'Congo (Democratic Republic)', '🇨🇩'),
    ('93', 'Costa Rica', '🇨🇷'),
    ('45', 'Croatia', '🇭🇷'),
    ('113', 'Cuba', '🇨🇺'),
    ('77', 'Cyprus', '🇨🇾'),
    ('63', 'Czech Republic', '🇨🇿'),
    ('27', 'Côte d’Ivoire', '🇨🇮'),
    ('172', 'Denmark', '🇩🇰'),
    ('168', 'Djibouti', '🇩🇯'),
    ('126', 'Dominica', '🇩🇲'),
    ('109', 'Dominican Republic', '🇩🇴'),
    ('105', 'Ecuador', '🇪🇨'),
    ('21', 'Egypt', '🇪🇬'),
    ('101', 'El Salvador', '🇸🇻'),
    ('167', 'Equatorial Guinea', '🇬🇶'),
    ('176', 'Eritrea', '🇪🇷'),
    ('34', 'Estonia', '🇪🇪'),
    ('71', 'Ethiopia', '🇪🇹'),
    ('189', 'Fiji', '🇫🇯'),
    ('163', 'Finland', '🇫🇮'),
    ('78', 'France', '🇫🇷'),
    ('162', 'French Guiana', '🇬🇫'),
    ('154', 'Gabon', '🇬🇦'),
    ('28', 'Gambia', '🇬🇲'),
    ('128', 'Georgia', '🇬🇪'),
    ('43', 'Germany', '🇩🇪'),
    ('38', 'Ghana', '🇬🇭'),
    ('201', 'Gibraltar', '🇬🇮'),
    ('129', 'Greece', '🇬🇷'),
    ('1008', 'Greenland', '🇬🇱'),
    ('127', 'Grenada', '🇬🇩'),
    ('160', 'Guadeloupe', '🇬🇵'),
    ('94', 'Guatemala', '🇬🇹'),
    ('68', 'Guinea', '🇬🇳'),
    ('130', 'Guinea-Bissau', '🇬🇼'),
    ('131', 'Guyana', '🇬🇾'),
    ('26', 'Haiti', '🇭🇹'),
    ('88', 'Honduras', '🇭🇳'),
    ('14', 'Hong Kong', '🇭🇰'),
    ('84', 'Hungary', '🇭🇺'),
    ('132', 'Iceland', '🇮🇸'),
    ('22', 'India', '🇮🇳'),
    ('6', 'Indonesia', '🇮🇩'),
    ('57', 'Iran', '🇮🇷'),
    ('47', 'Iraq', '🇮🇶'),
    ('23', 'Ireland', '🇮🇪'),
    ('13', 'Israel', '🇮🇱'),
    ('86', 'Italy', '🇮🇹'),
    ('103', 'Jamaica', '🇯🇲'),
    ('1001', 'Japan', '🇯🇵'),
    ('116', 'Jordan', '🇯🇴'),
    ('2', 'Kazakhstan', '🇰🇿'),
    ('8', 'Kenya', '🇰🇪'),
    ('1002', 'Korea', '🇰🇷'),
    ('1004', 'Kosovo', '🇽🇰'),
    ('100', 'Kuwait', '🇰🇼'),
    ('11', 'Kyrgyzstan', '🇰🇬'),
    ('25', 'Laos', '🇱🇦'),
    ('49', 'Latvia', '🇱🇻'),
    ('153', 'Lebanon', '🇱🇧'),
    ('136', 'Lesotho', '🇱🇸'),
    ('135', 'Liberia', '🇱🇷'),
    ('102', 'Libya', '🇱🇾'),
    ('1005', 'Liechtenstein', '🇱🇮'),
    ('44', 'Lithuania', '🇱🇹'),
    ('165', 'Luxembourg', '🇱🇺'),
    ('20', 'Macau', '🇲🇴'),
    ('183', 'Macedonia', '🇲🇰'),
    ('17', 'Madagascar', '🇲🇬'),
    ('137', 'Malawi', '🇲🇼'),
    ('7', 'Malaysia', '🇲🇾'),
    ('159', 'Maldives', '🇲🇻'),
    ('69', 'Mali', '🇲🇱'),
    ('199', 'Malta', '🇲🇹'),
    ('1011', 'Martinique', '🇲🇶'),
    ('114', 'Mauritania', '🇲🇷'),
    ('157', 'Mauritius', '🇲🇺'),
    ('54', 'Mexico', '🇲🇽'),
    ('85', 'Moldova', '🇲🇩'),
    ('144', 'Monaco', '🇲🇨'),
    ('72', 'Mongolia', '🇲🇳'),
    ('171', 'Montenegro', '🇲🇪'),
    ('180', 'Montserrat', '🇲🇸'),
    ('37', 'Morocco', '🇲🇦'),
    ('80', 'Mozambique', '🇲🇿'),
    ('5', 'Myanmar', '🇲🇲'),
    ('138', 'Namibia', '🇳🇦'),
    ('81', 'Nepal', '🇳🇵'),
    ('48', 'Netherlands', '🇳🇱'),
    ('185', 'New Caledonia', '🇳🇨'),
    ('67', 'New Zealand', '🇳🇿'),
    ('90', 'Nicaragua', '🇳🇮'),
    ('139', 'Niger', '🇳🇪'),
    ('19', 'Nigeria', '🇳🇬'),
    ('204', 'Niue', '🇳🇺'),
    ('174', 'Norway', '🇳🇴'),
    ('107', 'Oman', '🇴🇲'),
    ('66', 'Pakistan', '🇵🇰'),
    ('188', 'Palestine', '🇵🇸'),
    ('112', 'Panama', '🇵🇦'),
    ('79', 'Papua New Guinea', '🇵🇬'),
    ('87', 'Paraguay', '🇵🇾'),
    ('65', 'Peru', '🇵🇪'),
    ('4', 'Philippines', '🇵🇭'),
    ('15', 'Poland', '🇵🇱'),
    ('117', 'Portugal', '🇵🇹'),
    ('97', 'Puerto Rico', '🇵🇷'),
    ('111', 'Qatar', '🇶🇦'),
    ('32', 'Romania', '🇷🇴'),
    ('0', 'Russian Federation', '🇷🇺'),
    ('140', 'Rwanda', '🇷🇼'),
    ('146', 'Réunion', '🇷🇪'),
    ('134', 'Saint Kitts and Nevis', '🇰🇳'),
    ('164', 'Saint Lucia', '🇱🇨'),
    ('166', 'Saint Vincent', '🇻🇨'),
    ('178', 'Sao Tome and Principe', '🇸🇹'),
    ('53', 'Saudi Arabia', '🇸🇦'),
    ('61', 'Senegal', '🇸🇳'),
    ('29', 'Serbia', '🇷🇸'),
    ('184', 'Seychelles', '🇸🇨'),
    ('115', 'Sierra Leone', '🇸🇱'),
    ('196', 'Singapore', '🇸🇬'),
    ('1006', 'Sint Maarten', '🇸🇽'),
    ('141', 'Slovakia', '🇸🇰'),
    ('59', 'Slovenia', '🇸🇮'),
    ('149', 'Somalia', '🇸🇴'),
    ('31', 'South Africa', '🇿🇦'),
    ('177', 'South Sudan', '🇸🇸'),
    ('56', 'Spain', '🇪🇸'),
    ('64', 'Sri Lanka', '🇱🇰'),
    ('1010', 'Sudan', '🇸🇩'),
    ('142', 'Suriname', '🇸🇷'),
    ('106', 'Swaziland', '🇸🇿'),
    ('46', 'Sweden', '🇸🇪'),
    ('173', 'Switzerland', '🇨🇭'),
    ('1333', 'Syrian Arab Republic', '🇸🇾'),
    ('55', 'Taiwan', '🇹🇼'),
    ('143', 'Tajikistan', '🇹🇯'),
    ('9', 'Tanzania', '🇹🇿'),
    ('52', 'Thailand', '🇹🇭'),
    ('91', 'Timor-Leste', '🇹🇱'),
    ('99', 'Togo', '🇹🇬'),
    ('104', 'Trinidad and Tobago', '🇹🇹'),
    ('89', 'Tunisia', '🇹🇳'),
    ('62', 'Turkey', '🇹🇷'),
    ('161', 'Turkmenistan', '🇹🇲'),
    ('75', 'Uganda', '🇺🇬'),
    ('1', 'Ukraine', '🇺🇦'),
    ('95', 'United Arab Emirates', '🇦🇪'),
    ('16', 'United Kingdom', '🇬🇧'),
    ('187', 'United States', '🇺🇸'),
    ('12', 'United States (virtual)', '🇺🇸'),
    ('156', 'Uruguay', '🇺🇾'),
    ('40', 'Uzbekistan', '🇺🇿'),
    ('1007', 'Vanuatu', '🇻🇺'),
    ('70', 'Venezuela', '🇻🇪'),
    ('10', 'Viet Nam', '🇻🇳'),
    ('30', 'Yemen', '🇾🇪'),
    ('147', 'Zambia', '🇿🇲'),
    ('96', 'Zimbabwe', '🇿🇼'),
)
_SMSBOWER_COUNTRY_FLAGS_BY_NAME = {
    " ".join(name.split()).casefold(): flag
    for _country_id, name, flag in _SMSBOWER_COUNTRY_DATA
}
_SMSBOWER_COUNTRY_FLAGS_BY_ID = {
    country_id: flag for country_id, _name, flag in _SMSBOWER_COUNTRY_DATA
}
def _country_name_key(value: Any) -> str:
    ascii_name = (
        unicodedata.normalize("NFKD", str(value or ""))
        .encode("ascii", "ignore")
        .decode("ascii")
        .casefold()
    )
    return re.sub(r"[^a-z0-9]+", "", ascii_name)


_SMSBOWER_COUNTRY_ID_BY_NAME = {
    _country_name_key(name): str(country_id)
    for country_id, name, _flag in _SMSBOWER_COUNTRY_DATA
}
_SMSBOWER_COUNTRY_NAME_ALIASES = {
    "us": "united states",
    "usa": "united states",
    "united states of america": "united states",
    "uk": "united kingdom",
    "uae": "united arab emirates",
    "ivory coast": "côte d’ivoire",
    "czech republic": "czechia",
    "turkey": "türkiye",
    "republic of korea": "south korea",
    "korea republic of": "south korea",
    "democratic people s republic of korea": "north korea",
    "cape verde": "cabo verde",
    "macau": "macao",
}
for _country_alias, _canonical_country_name in _SMSBOWER_COUNTRY_NAME_ALIASES.items():
    _canonical_country_id = _SMSBOWER_COUNTRY_ID_BY_NAME.get(
        _country_name_key(_canonical_country_name)
    )
    if _canonical_country_id:
        _SMSBOWER_COUNTRY_ID_BY_NAME[_country_name_key(_country_alias)] = (
            _canonical_country_id
        )


def normalize_live_stock_country_ids(
    stock_by_country: dict[str, int],
    country_names: Optional[dict[str, str]] = None,
) -> dict[str, int]:
    """Translate getPrices country labels to the numeric IDs used by catalog rows."""
    id_by_name = dict(_SMSBOWER_COUNTRY_ID_BY_NAME)
    if country_names:
        for country_id, country_name in country_names.items():
            country_id = str(country_id or "").strip()
            if country_id.isdigit() and country_name:
                id_by_name[_country_name_key(country_name)] = country_id
    for alias, canonical in _SMSBOWER_COUNTRY_NAME_ALIASES.items():
        country_id = id_by_name.get(_country_name_key(canonical))
        if country_id:
            id_by_name[_country_name_key(alias)] = country_id

    normalized: dict[str, int] = {}
    unmapped = 0
    for raw_country, raw_stock in stock_by_country.items():
        key = str(raw_country or "").strip()
        country_id = (
            key
            if key.isdigit()
            else id_by_name.get(_country_name_key(key))
        )
        if not country_id:
            unmapped += 1
            continue
        try:
            stock = max(0, int(decimal_value(raw_stock)))
        except (ValueError, TypeError, InvalidOperation):
            continue
        normalized[country_id] = max(normalized.get(country_id, 0), stock)
    if unmapped:
        logger.warning(
            "SMSBower returned %s stock country label(s) that could not be mapped to IDs.",
            unmapped,
        )
    return normalized

# Premium emoji data is embedded from the supplied TXT files. Admin uploads are
# merged and persisted in premium_emoji_data.json beside this script.
_PREMIUM_EMOJI_DATA_PATH = os.path.join(_APP_DIRECTORY, "premium_emoji_data.json")
_BASE_FLAG_ROWS = [{"name":"Sint Maarten","flag":"🇸🇽","iso":"SX","calling":"1","id":"5461113820955027461"},{"name":"Bangladesh","flag":"🇧🇩","iso":"BD","calling":"880","id":"5911365056594973179"},{"name":"India","flag":"🇮🇳","iso":"IN","calling":"91","id":"5913754823643107921"},{"name":"Pakistan","flag":"🇵🇰","iso":"PK","calling":"92","id":"5913705895375672082"},{"name":"Isle of Man","flag":"🇮🇲","iso":"IM","calling":"44","id":"5226538255029121667"},{"name":"Ukraine","flag":"🇺🇦","iso":"UA","calling":"380","id":"5911406692007941050"},{"name":"Poland","flag":"🇵🇱","iso":"PL","calling":"48","id":"5913550391789752571"},{"name":"Kazakhstan","flag":"🇰🇿","iso":"KZ","calling":"7","id":"5913724621433082323"},{"name":"China","flag":"🇨🇳","iso":"CN","calling":"86","id":"5913779335021466780"},{"name":"Azerbaijan","flag":"🇦🇿","iso":"AZ","calling":"994","id":"5911197578640233518"},{"name":"Armenia","flag":"🇦🇲","iso":"AM","calling":"374","id":"5913272455866093666"},{"name":"Russia","flag":"🇷🇺","iso":"RU","calling":"79","id":"5913274246867456342"},{"name":"Uzbekistan","flag":"🇺🇿","iso":"UZ","calling":"998","id":"5911051846104912282"},{"name":"Germany","flag":"🇩🇪","iso":"DE","calling":"49","id":"5911096835887337583"},{"name":"Japan","flag":"🇯🇵","iso":"JP","calling":"81","id":"5913293711659241040"},{"name":"Turkey","flag":"🇹🇷","iso":"TR","calling":"90","id":"5910995113881901195"},{"name":"Belarus","flag":"🇧🇾","iso":"BY","calling":"375","id":"5911011185649521599"},{"name":"Brazil","flag":"🇧🇷","iso":"BR","calling":"55","id":"5911148568768418614"},{"name":"Zambia","flag":"🇿🇲","iso":"ZM","calling":"260","id":"5913564754160389778"},{"name":"Yemen","flag":"🇾🇪","iso":"YE","calling":"967","id":"5913346492512341993"},{"name":"Vietnam","flag":"🇻🇳","iso":"VN","calling":"84","id":"5913428887164949581"},{"name":"Holy See (Vatican City State)","flag":"🇻🇦","iso":"VA","calling":"379","id":"5911211932420938860"},{"name":"Vanuatu","flag":"🇻🇺","iso":"VU","calling":"678","id":"5913511535220625585"},{"name":"Uruguay","flag":"🇺🇾","iso":"UY","calling":"598","id":"5913623088406204470"},{"name":"United Arab Emirates","flag":"🇦🇪","iso":"AE","calling":"971","id":"5913726554168365343"},{"name":"Uganda","flag":"🇺🇬","iso":"UG","calling":"256","id":"5913488939397681980"},{"name":"Turkmenistan","flag":"🇹🇲","iso":"TM","calling":"993","id":"5913315521503170180"},{"name":"Tunisia","flag":"🇹🇳","iso":"TN","calling":"216","id":"5911332947419468671"},{"name":"Togo","flag":"🇹🇬","iso":"TG","calling":"228","id":"5913423260757790970"},{"name":"Thailand","flag":"🇹🇭","iso":"TH","calling":"66","id":"5913617968805187987"},{"name":"Tanzania","flag":"🇹🇿","iso":"TZ","calling":"255","id":"5911418949844603556"},{"name":"Tajikistan","flag":"🇹🇯","iso":"TJ","calling":"992","id":"5911287639809463107"},{"name":"Switzerland","flag":"🇨🇭","iso":"CH","calling":"41","id":"5913271227505448072"},{"name":"Sweden","flag":"🇸🇪","iso":"SE","calling":"46","id":"5911156510162949403"},{"name":"Eswatini (Swaziland)","flag":"🇸🇿","iso":"SZ","calling":"268","id":"5913374525763883286"},{"name":"Suriname","flag":"🇸🇷","iso":"SR","calling":"597","id":"5913275539652611719"},{"name":"Sudan","flag":"🇸🇩","iso":"SD","calling":"249","id":"5911387497799094470"},{"name":"Spain","flag":"🇪🇸","iso":"ES","calling":"34","id":"5911193287967904547"},{"name":"Sri Lanka","flag":"🇱🇰","iso":"LK","calling":"94","id":"5911293163137406640"},{"name":"South Sudan","flag":"🇸🇸","iso":"SS","calling":"211","id":"5911406262511211744"},{"name":"South Africa","flag":"🇿🇦","iso":"ZA","calling":"27","id":"5911203119148044594"},{"name":"Somalia","flag":"🇸🇴","iso":"SO","calling":"252","id":"5911397852965244436"},{"name":"Solomon Islands","flag":"🇸🇧","iso":"SB","calling":"677","id":"5911482712929080608"},{"name":"Slovenia","flag":"🇸🇮","iso":"SI","calling":"386","id":"5913431983836368644"},{"name":"Slovakia","flag":"🇸🇰","iso":"SK","calling":"421","id":"5913751666842145020"},{"name":"Singapore","flag":"🇸🇬","iso":"SG","calling":"65","id":"5911531460808051849"},{"name":"Sierra Leone","flag":"🇸🇱","iso":"SL","calling":"232","id":"5911210450657218661"},{"name":"Seychelles","flag":"🇸🇨","iso":"SC","calling":"248","id":"5911185183364616913"},{"name":"Serbia","flag":"🇷🇸","iso":"RS","calling":"381","id":"5913592598433369871"},{"name":"Senegal","flag":"🇸🇳","iso":"SN","calling":"221","id":"5910995302860461643"},{"name":"Sao Tome and Principe","flag":"🇸🇹","iso":"ST","calling":"239","id":"5913574331937462345"},{"name":"San Marino","flag":"🇸🇲","iso":"SM","calling":"378","id":"5913587968458625465"},{"name":"Samoa","flag":"🇼🇸","iso":"WS","calling":"685","id":"5913325971158602854"},{"name":"Palestine","flag":"🇵🇸","iso":"PS","calling":"970","id":"5913684768431541668"},{"name":"Rwanda","flag":"🇷🇼","iso":"RW","calling":"250","id":"5911455229433352234"},{"name":"Romania","flag":"🇷🇴","iso":"RO","calling":"40","id":"5913460373570195273"},{"name":"Qatar","flag":"🇶🇦","iso":"QA","calling":"974","id":"5911260864983339619"},{"name":"Portugal","flag":"🇵🇹","iso":"PT","calling":"351","id":"5911023653939581472"},{"name":"Philippines","flag":"🇵🇭","iso":"PH","calling":"63","id":"5911268638874145162"},{"name":"Peru","flag":"🇵🇪","iso":"PE","calling":"51","id":"5911207993935925780"},{"name":"Paraguay","flag":"🇵🇾","iso":"PY","calling":"595","id":"5911014265141072316"},{"name":"Papua","flag":"🇵🇬","iso":"PG","calling":"675","id":"5911107251183030903"},{"name":"Panama","flag":"🇵🇦","iso":"PA","calling":"507","id":"5913428968769327174"},{"name":"Palau","flag":"🇵🇼","iso":"PW","calling":"680","id":"5911283903187915549"},{"name":"Oman","flag":"🇴🇲","iso":"OM","calling":"968","id":"5913570801474343473"},{"name":"Svalbard and Jan Mayen","flag":"🇸🇯","iso":"SJ","calling":"47","id":"5780471598922337683"},{"name":"Nigeria","flag":"🇳🇬","iso":"NG","calling":"234","id":"5911143844304393105"},{"name":"Niger","flag":"🇳🇪","iso":"NE","calling":"227","id":"5911270086278124251"},{"name":"Pitcairn Islands","flag":"🇵🇳","iso":"PN","calling":"64","id":"5780471598922337683"},{"name":"Netherlands","flag":"🇳🇱","iso":"NL","calling":"31","id":"5913367645226275100"},{"name":"Nepal","flag":"🇳🇵","iso":"NP","calling":"977","id":"5913496520014958723"},{"name":"Namibia","flag":"🇳🇦","iso":"NA","calling":"264","id":"5911108535378252443"},{"name":"Mozambique","flag":"🇲🇿","iso":"MZ","calling":"258","id":"5911333419865871464"},{"name":"Morocco","flag":"🇲🇦","iso":"MA","calling":"212","id":"5911482111633658301"},{"name":"Montenegro","flag":"🇲🇪","iso":"ME","calling":"382","id":"5913239436157522151"},{"name":"Mongolia","flag":"🇲🇳","iso":"MN","calling":"976","id":"5911041383564580038"},{"name":"Monaco","flag":"🇲🇨","iso":"MC","calling":"377","id":"5911245347266500057"},{"name":"Moldova, Republic of","flag":"🇲🇩","iso":"MD","calling":"373","id":"5913456847402045950"},{"name":"Maldives","flag":"🇲🇻","iso":"MV","calling":"960","id":"5913501399097806832"},{"name":"Mali","flag":"🇲🇱","iso":"ML","calling":"223","id":"5911305266355245916"},{"name":"Malta","flag":"🇲🇹","iso":"MT","calling":"356","id":"5911023714069123567"},{"name":"Martinique","flag":"🇲🇶","iso":"MQ","calling":"596","id":"5911378005921370347"},{"name":"Marshall Islands","flag":"🇲🇭","iso":"MH","calling":"692","id":"5913235935759175692"},{"name":"Mauritius","flag":"🇲🇺","iso":"MU","calling":"230","id":"5913291113204027321"},{"name":"Mexico","flag":"🇲🇽","iso":"MX","calling":"52","id":"5913687302462246518"},{"name":"Micronesia, Federated States of","flag":"🇫🇲","iso":"FM","calling":"691","id":"5911271104185373336"},{"name":"Malaysia","flag":"🇲🇾","iso":"MY","calling":"60","id":"5913654360063087453"},{"name":"Kenya","flag":"🇰🇪","iso":"KE","calling":"254","id":"5911154710571651231"},{"name":"Madagascar","flag":"🇲🇬","iso":"MG","calling":"261","id":"5913766918271012920"},{"name":"Republic of North Macedonia","flag":"🇲🇰","iso":"MK","calling":"389","id":"5913394029210374721"},{"name":"Luxembourg","flag":"🇱🇺","iso":"LU","calling":"352","id":"5913390842344640293"},{"name":"Lithuania","flag":"🇱🇹","iso":"LT","calling":"370","id":"5911172315642597775"},{"name":"Liechtenstein","flag":"🇱🇮","iso":"LI","calling":"423","id":"5911166650580734660"},{"name":"Libya","flag":"🇱🇾","iso":"LY","calling":"218","id":"5911236989260140996"},{"name":"Liberia","flag":"🇱🇷","iso":"LR","calling":"231","id":"5913324167272337727"},{"name":"Kiribati","flag":"🇰🇮","iso":"KI","calling":"686","id":"5911294443037660118"},{"name":"Kosovo","flag":"🇽🇰","iso":"XK","calling":"383","id":"5911433681582429010"},{"name":"Kuwait","flag":"🇰🇼","iso":"KW","calling":"965","id":"5913290705182134003"},{"name":"Kyrgyzstan","flag":"🇰🇬","iso":"KG","calling":"996","id":"5911202161370337549"},{"name":"Lao People's Democratic Republic","flag":"🇱🇦","iso":"LA","calling":"856","id":"5913718526874489279"},{"name":"Latvia","flag":"🇱🇻","iso":"LV","calling":"371","id":"5913738489882480243"},{"name":"Lebanon","flag":"🇱🇧","iso":"LB","calling":"961","id":"5911504273664905447"},{"name":"Lesotho","flag":"🇱🇸","iso":"LS","calling":"266","id":"5911059881988723711"},{"name":"Indonesia","flag":"🇮🇩","iso":"ID","calling":"62","id":"5913479361620611038"},{"name":"Iran","flag":"🇮🇷","iso":"IR","calling":"98","id":"5911308891307643032"},{"name":"Iraq","flag":"🇮🇶","iso":"IQ","calling":"964","id":"5911382442622587735"},{"name":"Ireland","flag":"🇮🇪","iso":"IE","calling":"353","id":"5913440715504881532"},{"name":"Israel","flag":"🇮🇱","iso":"IL","calling":"972","id":"5911471936856134692"},{"name":"Italy","flag":"🇮🇹","iso":"IT","calling":"39","id":"5913688444923547525"},{"name":"Jordan","flag":"🇯🇴","iso":"JO","calling":"962","id":"5913234136167878475"},{"name":"Iceland","flag":"🇮🇸","iso":"IS","calling":"354","id":"5911047899029967246"},{"name":"Hungary","flag":"🇭🇺","iso":"HU","calling":"36","id":"5913767635530551104"},{"name":"Honduras","flag":"🇭🇳","iso":"HN","calling":"504","id":"5911406889576436289"},{"name":"Haiti","flag":"🇭🇹","iso":"HT","calling":"509","id":"5913459789454643194"},{"name":"Guyana","flag":"🇬🇾","iso":"GY","calling":"592","id":"5913579412883771480"},{"name":"Bissau","flag":"🇬🇼","iso":"GW","calling":"245","id":"5911398694778836149"},{"name":"Guinea","flag":"🇬🇳","iso":"GN","calling":"224","id":"5913471858312744319"},{"name":"Guatemala","flag":"🇬🇹","iso":"GT","calling":"502","id":"5913324858762072330"},{"name":"Greece","flag":"🇬🇷","iso":"GR","calling":"30","id":"5911210399117611448"},{"name":"Ghana","flag":"🇬🇭","iso":"GH","calling":"233","id":"5913391155877252952"},{"name":"Georgia","flag":"🇬🇪","iso":"GE","calling":"995","id":"5913434771270144023"},{"name":"Gambia","flag":"🇬🇲","iso":"GM","calling":"220","id":"5913657267755945883"},{"name":"Gabon","flag":"🇬🇦","iso":"GA","calling":"241","id":"5911037896051137264"},{"name":"France","flag":"🇫🇷","iso":"FR","calling":"33","id":"5913605586414473124"},{"name":"Aland Islands","flag":"🇦🇽","iso":"AX","calling":"358","id":"5780471598922337683"},{"name":"Fiji","flag":"🇫🇯","iso":"FJ","calling":"679","id":"5911393832875856716"},{"name":"Ethiopia","flag":"🇪🇹","iso":"ET","calling":"251","id":"5911078333168227043"},{"name":"Timor-Leste","flag":"🇹🇱","iso":"TL","calling":"670","id":"5911141915864076479"},{"name":"Ecuador","flag":"🇪🇨","iso":"EC","calling":"593","id":"5911273865849347408"},{"name":"Egypt","flag":"🇪🇬","iso":"EG","calling":"20","id":"5913694831539916769"},{"name":"El Salvador","flag":"🇸🇻","iso":"SV","calling":"503","id":"5913238624408703010"},{"name":"Estonia","flag":"🇪🇪","iso":"EE","calling":"372","id":"5910986042910969906"},{"name":"Djibouti","flag":"🇩🇯","iso":"DJ","calling":"253","id":"5911407709915190157"},{"name":"Denmark","flag":"🇩🇰","iso":"DK","calling":"45","id":"5911206009661034712"},{"name":"Cyprus","flag":"🇨🇾","iso":"CY","calling":"357","id":"5911023550860366409"},{"name":"Croatia","flag":"🇭🇷","iso":"HR","calling":"385","id":"5913692684056269311"},{"name":"Costa Rica","flag":"🇨🇷","iso":"CR","calling":"506","id":"5911261745451635030"},{"name":"Congo","flag":"🇨🇬","iso":"CG","calling":"242","id":"5911338788574990168"},{"name":"Congo, The Democratic Republic of the","flag":"🇨🇩","iso":"CD","calling":"243","id":"5913770362834783827"},{"name":"Comoros","flag":"🇰🇲","iso":"KM","calling":"269","id":"5911338582416560604"},{"name":"Cambodia","flag":"🇰🇭","iso":"KH","calling":"855","id":"5913699998385573485"},{"name":"Cameroon","flag":"🇨🇲","iso":"CM","calling":"237","id":"5911172109484167745"},{"name":"Cape Verde","flag":"🇨🇻","iso":"CV","calling":"238","id":"5913571501554012193"},{"name":"Central African Republic","flag":"🇨🇫","iso":"CF","calling":"236","id":"5913443245240619222"},{"name":"Chad","flag":"🇹🇩","iso":"TD","calling":"235","id":"5913299849167507310"},{"name":"Czechia","flag":"🇨🇿","iso":"CZ","calling":"420","id":"5911198691036764307"},{"name":"Chile","flag":"🇨🇱","iso":"CL","calling":"56","id":"5911470957603592832"},{"name":"Colombia","flag":"🇨🇴","iso":"CO","calling":"57","id":"5913773060074246009"},{"name":"Burundi","flag":"🇧🇮","iso":"BI","calling":"257","id":"5913766441529642752"},{"name":"Botswana","flag":"🇧🇼","iso":"BW","calling":"267","id":"5911513782722499475"},{"name":"Bosnia and Herzegovina","flag":"🇧🇦","iso":"BA","calling":"387","id":"5913700002680541032"},{"name":"Bolivia","flag":"🇧🇴","iso":"BO","calling":"591","id":"5913638795101606133"},{"name":"Bhutan","flag":"🇧🇹","iso":"BT","calling":"975","id":"5913236734623093021"},{"name":"Benin","flag":"🇧🇯","iso":"BJ","calling":"229","id":"5913735869952430547"},{"name":"Argentina","flag":"🇦🇷","iso":"AR","calling":"54","id":"5913573356979884082"},{"name":"Australia","flag":"🇦🇺","iso":"AU","calling":"61","id":"5913632326880858455"},{"name":"Austria","flag":"🇦🇹","iso":"AT","calling":"43","id":"5911338831524664592"},{"name":"Bahrain","flag":"🇧🇭","iso":"BH","calling":"973","id":"5913581663446634403"},{"name":"Belgium","flag":"🇧🇪","iso":"BE","calling":"32","id":"5913529642802745141"},{"name":"Belize","flag":"🇧🇿","iso":"BZ","calling":"501","id":"5913355005137522807"},{"name":"Angola","flag":"🇦🇴","iso":"AO","calling":"244","id":"5913753316109586411"},{"name":"Andorra","flag":"🇦🇩","iso":"AD","calling":"376","id":"5911314702398396902"},{"name":"Algeria","flag":"🇩🇿","iso":"DZ","calling":"213","id":"5913782968563800236"},{"name":"Albania","flag":"🇦🇱","iso":"AL","calling":"355","id":"5911357458797826163"},{"name":"Afghanistan","flag":"🇦🇫","iso":"AF","calling":"93","id":"5913492040364068694"},{"name":"Zimbabwe","flag":"🇿🇼","iso":"ZW","calling":"263","id":"5911092502265336396"},{"name":"Cuba","flag":"🇨🇺","iso":"CU","calling":"53","id":"5431551436502611633"},{"name":"Korea","flag":"🇰🇵","iso":"KP","calling":"850","id":"5434142701941437163"},{"name":"Venezuela","flag":"🇻🇪","iso":"VE","calling":"58","id":"5434009132753499322"},{"name":"Syria","flag":"🇸🇾","iso":"SY","calling":"963","id":"5433910876786670092"},{"name":"Myanmar","flag":"🇲🇲","iso":"MM","calling":"95","id":"5433666360003540231"},{"name":"Nicaragua","flag":"🇳🇮","iso":"NI","calling":"505","id":"5334807849418003620"},{"name":"South Korea","flag":"🇰🇷","iso":"KR","calling":"82","id":"5913371673905598425"},{"name":"Equatorial Guinea","flag":"🇬🇶","iso":"GQ","calling":"240","id":"5911306279967529251"},{"name":"Greenland","flag":"🇬🇱","iso":"GL","calling":"299","id":"5292014752283774878"},{"name":"Faroe Islands","flag":"🇫🇴","iso":"FO","calling":"298","id":"5296469342039327674"},{"name":"Côte d'Ivoire (Ivory Coast)","flag":"🇨🇮","iso":"CI","calling":"225","id":"5222233374948602940"},{"name":"Brunei","flag":"🇧🇳","iso":"BN","calling":"673","id":"5911336409163109113"},{"name":"Bulgaria","flag":"🇧🇬","iso":"BG","calling":"359","id":"5294329219965272288"},{"name":"Burkina Faso","flag":"🇧🇫","iso":"BF","calling":"226","id":"5913407764515786948"},{"name":"Eritrea","flag":"🇪🇷","iso":"ER","calling":"291","id":"5433723401464198287"},{"name":"Malawi","flag":"🇲🇼","iso":"MW","calling":"265","id":"5433968339154122439"},{"name":"Mauritania","flag":"🇲🇷","iso":"MR","calling":"222","id":"5433859405898594234"},{"name":"Nauru","flag":"🇳🇷","iso":"NR","calling":"674","id":"5434131139889478358"},{"name":"Saudi Arabia","flag":"🇸🇦","iso":"SA","calling":"966","id":"4985897134424328239"},{"name":"Tonga","flag":"🇹🇴","iso":"TO","calling":"676","id":"5433640100573491806"},{"name":"Tuvalu","flag":"🇹🇻","iso":"TV","calling":"688","id":"5433684690923961019"},{"name":"Taiwan","flag":"🇹🇼","iso":"TW","calling":"886","id":"5366187256937726720"},{"name":"Hong Kong","flag":"🇭🇰","iso":"HK","calling":"852","id":"5292166459118606932"},{"name":"Macau","flag":"🇲🇴","iso":"MO","calling":"853","id":"6323557758096377611"},{"name":"Aruba","flag":"🇦🇼","iso":"AW","calling":"297","id":"5780471598922337683"},{"name":"Bonaire","flag":"🇧🇶","iso":"BQ","calling":"599","id":"5780471598922337683"},{"name":"Falkland Islands","flag":"🇫🇰","iso":"FK","calling":"500","id":"5780471598922337683"},{"name":"French Guiana","flag":"🇬🇫","iso":"GF","calling":"594","id":"5780471598922337683"},{"name":"Guadeloupe","flag":"🇬🇵","iso":"GP","calling":"590","id":"5780471598922337683"},{"name":"Reunion","flag":"🇷🇪","iso":"RE","calling":"262","id":"5780471598922337683"},{"name":"New Caledonia","flag":"🇳🇨","iso":"NC","calling":"687","id":"5780471598922337683"},{"name":"Niue","flag":"🇳🇺","iso":"NU","calling":"683","id":"5780471598922337683"},{"name":"Norfolk Island","flag":"🇳🇫","iso":"NF","calling":"672","id":"5780471598922337683"},{"name":"Saint Helena","flag":"🇸🇭","iso":"SH","calling":"290","id":"5780471598922337683"},{"name":"Tokelau","flag":"🇹🇰","iso":"TK","calling":"690","id":"5780471598922337683"},{"name":"Wallis and Futuna","flag":"🇼🇫","iso":"WF","calling":"681","id":"5780471598922337683"},{"name":"Cook Islands","flag":"🇨🇰","iso":"CK","calling":"682","id":"5780471598922337683"},{"name":"French Polynesia","flag":"🇵🇫","iso":"PF","calling":"689","id":"5780471598922337683"},{"name":"Gibraltar","flag":"🇬🇮","iso":"GI","calling":"350","id":"5780471598922337683"},{"name":"Saint Pierre and Miquelon","flag":"🇵🇲","iso":"PM","calling":"508","id":"5780471598922337683"}]
_BASE_SERVICE_ROWS = [{"name":"FACEBOOK","emoji":"📘","id":"5296523110734930940"},{"name":"WHATSAPP","emoji":"💬","id":"5296439324512918594"},{"name":"TELEGRAM","emoji":"✈️","id":"5294193228415801857"},{"name":"WHATSAPP BUSINESSES","emoji":"💼","id":"5336814486701514414"},{"name":"IMO","emoji":"💭","id":"5296393407017556498"},{"name":"INSTAGRAM","emoji":"📸","id":"5296719399330292238"},{"name":"APPLE","emoji":"🍎","id":"5294101883051353361"},{"name":"GOOGLE","emoji":"🔍","id":"5335010201005231986"},{"name":"MICROSOFT","emoji":"🪟","id":"5334880948259427772"},{"name":"OTHER","emoji":"📶","id":"5305517382138112561"},{"name":"TEAMS","emoji":"🧑‍🤝‍🧑","id":"5334590977837403844"},{"name":"TIKTOK","emoji":"🎵","id":"5339213256001102461"},{"name":"BKASH","emoji":"🏦","id":"5348469219761626211"},{"name":"ROCKET","emoji":"🚀","id":"5346042941196507141"},{"name":"BYBIT","emoji":"📈","id":"5348372939479751825"},{"name":"BINANCE","emoji":"💱","id":"5348212415077064131"},{"name":"MELBET","emoji":"🌟","id":"5337102391244263212"},{"name":"SNAPCHAT","emoji":"👻","id":"5359441366554255082"},{"name":"UBER","emoji":"🚗","id":"5298715455316303708"},{"name":"PAYPAL","emoji":"💵","id":"5776103539872896061"},{"name":"DISCORD","emoji":"🎬","id":"5116246243646898866"},{"name":"AMAZON","emoji":"🌟","id":"4995019580536524226"},{"name":"VIBER","emoji":"💜","id":"5463060437572528782"},{"name":"LINKEDIN","emoji":"💼","id":"6224222994265279792"},{"name":"LINE","emoji":"🔒","id":"5399818044866327279"},{"name":"💜 WECHAT","emoji":"🌟","id":"5782757599560602950"},{"name":"TWITTER","emoji":"🐦","id":"5215726959056662534"},{"name":"REDDIT","emoji":"👽","id":"4992421103847604984"},{"name":"PINTEREST","emoji":"📌","id":"5346103513120258857"},{"name":"TWITCH","emoji":"🎮","id":"5233333563306301418"},{"name":"ZOOM","emoji":"📹","id":"5881799193219043268"},{"name":"SIGNAL","emoji":"💬","id":"5293998404404272267"},{"name":"SLACK","emoji":"💻","id":"4994972469040251302"},{"name":"SKYPE","emoji":"☎️","id":"4992613535562334989"},{"name":"NETFLIX","emoji":"🎥","id":"6255738712664050133"},{"name":"SPOTIFY","emoji":"🎵","id":"5411392711146095115"},{"name":"AMAZON PRIME","emoji":"📺","id":"6111801057061374810"},{"name":"HOICHOI","emoji":"🍿","id":"6104822598493801746"},{"name":"DARAZ","emoji":"📦","id":"5336879280578138635"},{"name":"FOODPANDA","emoji":"🐼","id":"5336879280578138635"},{"name":"PATHAO","emoji":"🛵","id":"5336879280578138635"},{"name":"ALIEXPRESS","emoji":"🛒","id":"5336879280578138635"},{"name":"SHOPEE","emoji":"🛍️","id":"5336879280578138635"},{"name":"PAYONEER","emoji":"💳","id":"5336879280578138635"},{"name":"WISE","emoji":"🦉","id":"5336879280578138635"},{"name":"CHATGPT","emoji":"🤖","id":"5296516998996445955"},{"name":"NOTION","emoji":"📓","id":"5336879280578138635"},{"name":"GITHUB","emoji":"🐙","id":"5417836094098007862"},{"name":"CANVA","emoji":"🖌️","id":"5111661409008092227"},{"name":"FIGMA","emoji":"🎨","id":"5336879280578138635"},{"name":"UPWORK","emoji":"💼","id":"5336879280578138635"},{"name":"FIVERR","emoji":"🟢","id":"5336879280578138635"},{"name":"YAHOO","emoji":"🌐","id":"5336879280578138635"},{"name":"DROPBOX","emoji":"☁️","id":"5336879280578138635"},{"name":"COURSERA","emoji":"📚","id":"5336879280578138635"},{"name":"DUOLINGO","emoji":"🗣️","id":"5336879280578138635"}]

def _emoji_name_key(value: Any) -> str:
    return " ".join(str(value or "").split()).casefold()

def _build_emoji_maps(flag_rows: list[dict[str, str]], service_rows: list[dict[str, str]]) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    country_codes: dict[str, str] = {}
    country_ids: dict[str, str] = {}
    service_ids: dict[str, str] = {}
    for row in flag_rows:
        name = str(row.get("name", "")).strip()
        flag = str(row.get("flag", "")).strip()
        eid = str(row.get("id", "")).strip()
        iso = str(row.get("iso", "")).strip().upper()
        if name and eid.isdigit(): country_ids[_emoji_name_key(name)] = eid
        if flag and eid.isdigit(): country_ids[_emoji_name_key(flag)] = eid
        if name and re.fullmatch(r"[A-Z]{2}", iso): country_codes[name] = iso
    for row in service_rows:
        name = str(row.get("name", "")).strip()
        eid = str(row.get("id", "")).strip()
        if name and eid.isdigit(): service_ids[_emoji_name_key(name)] = eid
    return country_codes, country_ids, service_ids

def _load_premium_emoji_data() -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    if os.path.isfile(_PREMIUM_EMOJI_DATA_PATH):
        try:
            with open(_PREMIUM_EMOJI_DATA_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, dict) and isinstance(saved.get("flags"), list) and isinstance(saved.get("services"), list):
                if saved.get("managed") is True:
                    return saved["flags"], saved["services"]
                # Older data files are additive overrides; merge them into the
                # TXT-derived defaults rather than accidentally hiding defaults.
                merged_flags = { _emoji_name_key(r.get("name")): dict(r) for r in _BASE_FLAG_ROWS }
                merged_services = { _emoji_name_key(r.get("name")): dict(r) for r in _BASE_SERVICE_ROWS }
                for row in saved["flags"]:
                    if isinstance(row, dict) and row.get("name"):
                        merged_flags[_emoji_name_key(row.get("name"))] = row
                for row in saved["services"]:
                    if isinstance(row, dict) and row.get("name"):
                        merged_services[_emoji_name_key(row.get("name"))] = row
                return list(merged_flags.values()), list(merged_services.values())
        except Exception:
            pass  # Keep the embedded TXT-derived mappings if the saved file is unreadable.
    return list(_BASE_FLAG_ROWS), list(_BASE_SERVICE_ROWS)

def _save_premium_emoji_data() -> None:
    temp_path = _PREMIUM_EMOJI_DATA_PATH + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as f:
        json.dump({"managed": True, "flags": PREMIUM_FLAG_ROWS, "services": PREMIUM_SERVICE_ROWS}, f, ensure_ascii=False, indent=2)
    os.replace(temp_path, _PREMIUM_EMOJI_DATA_PATH)

PREMIUM_FLAG_ROWS, PREMIUM_SERVICE_ROWS = _load_premium_emoji_data()
_COUNTRY_EMOJI_CODES, _PREMIUM_COUNTRY_EMOJI_IDS_NORMALIZED, _PREMIUM_SERVICE_EMOJI_IDS_NORMALIZED = _build_emoji_maps(PREMIUM_FLAG_ROWS, PREMIUM_SERVICE_ROWS)
_COUNTRY_EMOJI_CODES_NORMALIZED = {_emoji_name_key(k): str(v).upper() for k, v in _COUNTRY_EMOJI_CODES.items()}


async def create_ipshop_invoice(user_id, amount, customer_name):
    async with aiohttp.ClientSession() as session:

        data = {
            "api_key": IPSHOP_API_KEY,
            "user_id": user_id,
            "amount": amount,
            "customer_name": customer_name,
            "customer_phone": "",
            "customer_note": "Balance Deposit",
            "callback_url": CALLBACK_URL,
            "return_url": RETURN_URL,
        }

        async with session.post(
            "https://payvora.online/api/create_invoice.php",
            data=data,
        ) as resp:

            text = await resp.text()

            print("CREATE INVOICE =", text)

            try:
                return json.loads(text)
            except Exception:
                return {
                    "status": False,
                    "message": text,
                }


async def check_ipshop_invoice(invoice_id):

    async with aiohttp.ClientSession() as session:

        data = {
            "api_key": IPSHOP_API_KEY,
            "invoice_id": invoice_id,
        }

        async with session.post(
            "https://payvora.online/api/check_invoice.php",
            data=data,
        ) as resp:

            text = await resp.text()

            print("CHECK INVOICE =", text)

            try:
                return json.loads(text)
            except Exception:
                return {
                    "status": False,
                    "message": text,
                }

        

# ================= IPSHOP PAYMENT =================

IPSHOP_API_URL = "https://payvora.online/api/create_invoice.php"

CALLBACK_URL = "https://t.me/renoxa_otp_bot"
RETURN_URL = "https://t.me/renoxa_otp_bot?start=payment_success"

# IPShop invoice monitors run as background tasks, independent of user input.
_IPSHOP_PAYMENT_TASKS: dict[str, asyncio.Task[Any]] = {}


# =============================================================================
# TELEGRAM BUTTON STYLE COMPATIBILITY
# =============================================================================
# Telegram currently exposes primary, success and danger styles. A neutral
# ("secondary") action is represented by omitting the style. The wrappers below
# keep the bot compatible with older python-telegram-bot versions while enabling
# colors automatically on python-telegram-bot 22.7+.
_TelegramKeyboardButton = KeyboardButton
_TelegramInlineKeyboardButton = InlineKeyboardButton
_REPLY_STYLE_SUPPORTED = "style" in inspect.signature(_TelegramKeyboardButton).parameters
_INLINE_STYLE_SUPPORTED = "style" in inspect.signature(_TelegramInlineKeyboardButton).parameters
_INLINE_CUSTOM_EMOJI_SUPPORTED = (
    "icon_custom_emoji_id"
    in inspect.signature(_TelegramInlineKeyboardButton).parameters
)
_VALID_BUTTON_STYLES = {"primary", "success", "danger", "secondary"}


def _normalise_button_style(style: Optional[str]) -> Optional[str]:
    value = str(style or "").strip().lower()
    if value not in _VALID_BUTTON_STYLES:
        return None
    # Telegram has no literal secondary color; no style is the neutral option.
    return None if value == "secondary" else value


def _infer_inline_style(text: str) -> str:
    value = str(text or "").casefold()
    navigation_words = (
        "previous", "prev", "next", "back", "refresh",
        "পূর্ববর্তী", "পরবর্তী", "পেছনে", "ফিরে", "রিফ্রেশ",
    )
    navigation_icons = ("⬅", "➡", "←", "→", "🔙", "◀", "▶", "⏪", "⏮")
    if any(word in value for word in navigation_words) or value.lstrip().startswith(
        navigation_icons
    ):
        return "secondary"

    danger_words = (
        "cancel", "reject", "delete", "remove", "disable", "stop",
        "ক্যানসেল", "বাতিল", "রিমুভ", "ডিলিট", "❌", "🗑",
    )
    success_words = (
        "buy", "approve", "accept", "confirm", "save", "add", "enable",
        "কিনুন", "অ্যাপ্রুভ", "সেভ", "যোগ", "✅",
    )
    if any(word in value for word in danger_words):
        return "danger"
    if any(word in value for word in success_words):
        return "primary"
    return "primary"


def KeyboardButton(text: str, *args: Any, style: Optional[str] = None, **kwargs: Any) -> Any:
    # Main-menu and admin-menu buttons are green by default, as requested.
    selected = _normalise_button_style(style or "primary")
    if _REPLY_STYLE_SUPPORTED and selected:
        kwargs["style"] = selected
    return _TelegramKeyboardButton(text, *args, **kwargs)


def InlineKeyboardButton(
    text: str, *args: Any, style: Optional[str] = None, **kwargs: Any
) -> Any:
    inferred_style = _infer_inline_style(text)
    # Navigation actions stay neutral even if a caller supplies a color style.
    requested_style = (
        "secondary" if inferred_style == "secondary" else (style or inferred_style)
    )
    selected = _normalise_button_style(requested_style)
    if _INLINE_STYLE_SUPPORTED and selected:
        kwargs["style"] = selected
    return _TelegramInlineKeyboardButton(text, *args, **kwargs)


_COUNTRY_EMOJI_CODES_NORMALIZED = {
    _emoji_name_key(name): str(code).strip().upper()
    for name, code in _COUNTRY_EMOJI_CODES.items()
}
_COUNTRY_EMOJI_CODE_ALIASES = {
    "anguilla": "AI",
    "aruba": "AW",
    "antigua and barbuda": "AG",
    "bahamas": "BS",
    "barbados": "BB",
    "bermuda": "BM",
    "bhutan": "BT",
    "brunei darussalam": "BN",
    "cambodia": "KH",
    "canada": "CA",
    "cayman islands": "KY",
    "cape verde": "CV",
    "comoros": "KM",
    "congo (democratic republic)": "CD",
    "congo (republic of the)": "CG",
    "czech republic": "CZ",
    "côte d'ivoire": "CI",
    "cote d'ivoire (ivory coast)": "CI",
    "côte d’ivoire": "CI",
    "cote d'ivoire": "CI",
    "cuba": "CU",
    "dominica": "DM",
    "dominican republic": "DO",
    "french guiana": "GF",
    "gibraltar": "GI",
    "guadeloupe": "GP",
    "greece": "GR",
    "greenland": "GL",
    "grenada": "GD",
    "guyana": "GY",
    "hong kong": "HK",
    "india": "IN",
    "jamaica": "JM",
    "kazakhstan": "KZ",
    "korea": "KR",
    "liechtenstein": "LI",
    "macau": "MO",
    "macedonia": "MK",
    "maldives": "MV",
    "martinique": "MQ",
    "montserrat": "MS",
    "new caledonia": "NC",
    "niue": "NU",
    "pakistan": "PK",
    "palestine": "PS",
    "puerto rico": "PR",
    "réunion": "RE",
    "reunion": "RE",
    "russian federation": "RU",
    "saint kitts and nevis": "KN",
    "saint lucia": "LC",
    "saint vincent": "VC",
    "seychelles": "SC",
    "sint maarten": "SX",
    "south sudan": "SS",
    "suriname": "SR",
    "swaziland": "SZ",
    "syrian arab republic (syria)": "SY",
    "taiwan": "TW",
    "timor-leste": "TL",
    "trinidad and tobago": "TT",
    "united states": "US",
    "united states (virtual)": "US",
    "united states / canada": "US",
    "usa (virtual)": "US",
    "vanuatu": "VU",
    "vietnam": "VN",
}


def _country_emoji_code(country_name: Any) -> Optional[str]:
    key = " ".join(str(country_name or "").split()).casefold()
    code = _COUNTRY_EMOJI_CODE_ALIASES.get(key)
    if not code:
        code = _COUNTRY_EMOJI_CODES_NORMALIZED.get(key)
    if not code:
        mapped = _PREMIUM_COUNTRY_EMOJI_IDS_NORMALIZED.get(key, "")
        if re.fullmatch(r"[A-Z]{2}", mapped.upper()):
            code = mapped.upper()
    return code if code and re.fullmatch(r"[A-Z]{2}", code) else None


def _country_flag_emoji(country_name: Any, country_id: Any = None) -> str:
    name_key = " ".join(str(country_name or "").split()).casefold()
    embedded_flag = _SMSBOWER_COUNTRY_FLAGS_BY_NAME.get(name_key)
    if embedded_flag:
        return embedded_flag

    code = _country_emoji_code(country_name)
    if code:
        return "".join(chr(0x1F1E6 + ord(char) - ord("A")) for char in code)

    country_id_key = str(country_id or "").strip()
    if country_id_key in _SMSBOWER_COUNTRY_FLAGS_BY_ID:
        return _SMSBOWER_COUNTRY_FLAGS_BY_ID[country_id_key]
    return "🌍"


def _country_premium_emoji_id(country_name: Any, country_id: Any = None) -> Optional[str]:
    key = _emoji_name_key(country_name)
    direct = _PREMIUM_COUNTRY_EMOJI_IDS_NORMALIZED.get(key, "")
    if direct.isdigit():
        return direct
    flag = _country_flag_emoji(country_name, country_id)
    emoji_id = _PREMIUM_COUNTRY_EMOJI_IDS_NORMALIZED.get(_emoji_name_key(flag), "")
    return emoji_id if emoji_id.isdigit() else None


def _country_tg_emoji(country_name: Any, country_id: Any = None) -> str:
    flag = _country_flag_emoji(country_name, country_id)
    emoji_id = _country_premium_emoji_id(country_name, country_id)
    if emoji_id:
        return f'<tg-emoji emoji-id="{emoji_id}">{html.escape(flag)}</tg-emoji>'
    return html.escape(flag)


def _service_premium_emoji_id(service_name: Any) -> Optional[str]:
    value = _PREMIUM_SERVICE_EMOJI_IDS_NORMALIZED.get(_emoji_name_key(service_name), "")
    return value if value.isdigit() else None


def _service_fallback_emoji(service_name: Any) -> str:
    key = _emoji_name_key(service_name)
    for row in PREMIUM_SERVICE_ROWS:
        if _emoji_name_key(row.get("name")) == key:
            return str(row.get("emoji") or "📱")
    return "📱"


def _service_tg_emoji(service_name: Any) -> str:
    fallback = html.escape(_service_fallback_emoji(service_name))
    emoji_id = _service_premium_emoji_id(service_name)
    if emoji_id:
        return f'<tg-emoji emoji-id="{emoji_id}">{fallback}</tg-emoji>'
    return fallback


def _inline_custom_emoji_kwargs(emoji_id: Optional[str]) -> dict[str, Any]:
    if not emoji_id or not str(emoji_id).isdigit():
        return {}
    if _INLINE_CUSTOM_EMOJI_SUPPORTED:
        return {"icon_custom_emoji_id": str(emoji_id)}
    return {"api_kwargs": {"icon_custom_emoji_id": str(emoji_id)}}


def _country_button_pair(
    country_name: str,
    callback_data: str,
    *,
    country_id: Any = None,
    text_prefix: str = "",
    text_limit: int = 25,
) -> tuple[InlineKeyboardButton, InlineKeyboardButton, bool]:
    """Use the supplied country Premium Emoji ID as the button icon."""
    label = f"{text_prefix}{short_text(country_name, text_limit)}"
    flag = _country_flag_emoji(country_name, country_id)
    emoji_id = _country_premium_emoji_id(country_name, country_id)
    fallback_button = InlineKeyboardButton(f"{flag} {label}", callback_data=callback_data)
    if not emoji_id:
        return fallback_button, fallback_button, False
    premium_button = InlineKeyboardButton(label, callback_data=callback_data, **_inline_custom_emoji_kwargs(emoji_id))
    return premium_button, fallback_button, True


# =============================================================================
# CONFIG — non-secret URLs and runtime settings
# =============================================================================

SMS_API_URL = os.getenv("SMS_API_URL", "https://smsbower.page/stubs/handler_api.php").strip()

# Separate GetSMSWeb API used only for USA local numbers.
LOCAL_API_URL = os.getenv(
    "LOCAL_API_URL", "https://getsmsweb.com/developer_api"
).strip()
LOCAL_OTP_NUMBER_PARAM = os.getenv("LOCAL_OTP_NUMBER_PARAM", "phone").strip() or "phone"
DEFAULT_LOCAL_PRICE_BDT = os.getenv("DEFAULT_LOCAL_PRICE_BDT", "10")
DEFAULT_LOCAL_SESSION_MINUTES = int(
    os.getenv("DEFAULT_LOCAL_SESSION_MINUTES", "720")
)
LOCAL_CANCEL_LOCK_SECONDS = 10 * 60
LOCAL_CANCEL_WINDOW_SECONDS = 60 * 60
# Telegram message refresh intervals for local-number screens.
# Keep the OTP API poll fast, but edit Telegram messages less often to avoid flood limits.
LOCAL_DETAIL_REFRESH_SECONDS = max(
    30, int(os.getenv("LOCAL_DETAIL_REFRESH_SECONDS", "30"))
)
LOCAL_ACTIVE_REFRESH_SECONDS = max(
    300, int(os.getenv("LOCAL_ACTIVE_REFRESH_SECONDS", "300"))
)
LOCAL_OTP_POLL_SECONDS = 5
LOCAL_ACTIVE_PAGE_SIZE = 10

# Comma-separated Telegram numeric IDs. Use /id in the bot to see your ID.
ADMIN_IDS_RAW = os.getenv("ADMIN_IDS", "").strip()

DATABASE_PATH = os.getenv("DATABASE_PATH", "paid_number_bot.sqlite3")
BACKUP_DIRECTORY = os.path.abspath(
    os.getenv(
        "BACKUP_DIRECTORY",
        os.path.join(os.path.dirname(os.path.abspath(DATABASE_PATH)) or ".", "backups"),
    )
)
BACKUP_KEEP_COUNT = max(30, int(os.getenv("BACKUP_KEEP_COUNT", "30")))
DATABASE_AUTO_BACKUP_INTERVAL_SECONDS = max(
    60 * 60, int(os.getenv("DATABASE_AUTO_BACKUP_HOURS", "6")) * 60 * 60
)
DATABASE_RECOVERY_ANCHOR = "paid-number-bot-recovery-anchor.db"
DATABASE_REQUIRED_TABLES = {
    "users",
    "settings",
    "catalog",
    "orders",
    "deposits",
    "transactions",
}
DATABASE_STARTUP_NOTICE: Optional[str] = None
DATABASE_BACKUP_MAX_BYTES = max(
    1 * 1024 * 1024,
    int(os.getenv("DATABASE_BACKUP_MAX_BYTES", str(50 * 1024 * 1024))),
)
DATABASE_BACKUP_PREFIX = "paid_number_bot_backup"
DEFAULT_MARKUP_USD = os.getenv("DEFAULT_MARKUP_USD", "0.005")
DEFAULT_EXPIRY_MINUTES = int(os.getenv("DEFAULT_EXPIRY_MINUTES", "20"))
POLL_INTERVAL_SECONDS = max(10, int(os.getenv("POLL_INTERVAL_SECONDS", "15")))
HTTP_TIMEOUT_SECONDS = max(5, int(os.getenv("HTTP_TIMEOUT_SECONDS", "25")))

# High-load tuning. Defaults are conservative enough for a small VPS while
# allowing many different users to be served concurrently.
BOT_MAX_CONCURRENT_UPDATES = max(8, int(os.getenv("BOT_MAX_CONCURRENT_UPDATES", "64")))
BOT_CONNECTION_POOL_SIZE = max(
    BOT_MAX_CONCURRENT_UPDATES,
    int(os.getenv("BOT_CONNECTION_POOL_SIZE", "128")),
)
BOT_POOL_TIMEOUT_SECONDS = max(5, int(os.getenv("BOT_POOL_TIMEOUT_SECONDS", "30")))
BOT_MEDIA_WRITE_TIMEOUT_SECONDS = max(
    30, int(os.getenv("BOT_MEDIA_WRITE_TIMEOUT_SECONDS", "180"))
)
# Keep normal Bot API traffic below Telegram's free flood-control ceiling.
# These values can be overridden with environment variables.
TELEGRAM_OVERALL_RATE = max(1, min(15, int(os.getenv("TELEGRAM_OVERALL_RATE", "10"))))
TELEGRAM_GROUP_RATE = max(1, min(20, int(os.getenv("TELEGRAM_GROUP_RATE", "15"))))
PROVIDER_HTTP_CONNECTIONS = max(4, int(os.getenv("PROVIDER_HTTP_CONNECTIONS", "32")))
PROVIDER_MAX_CONCURRENCY = max(2, int(os.getenv("PROVIDER_MAX_CONCURRENCY", "24")))
LOCAL_PROVIDER_MAX_CONCURRENCY = max(
    2, int(os.getenv("LOCAL_PROVIDER_MAX_CONCURRENCY", "16"))
)
SMS_POLL_CONCURRENCY = max(2, int(os.getenv("SMS_POLL_CONCURRENCY", "8")))
SMS_POLL_BATCH_SIZE = max(100, int(os.getenv("SMS_POLL_BATCH_SIZE", "2000")))
COUNTDOWN_REFRESH_SECONDS = max(
    60, int(os.getenv("COUNTDOWN_REFRESH_SECONDS", "120"))
)
FILE_TASK_CONCURRENCY = max(1, int(os.getenv("FILE_TASK_CONCURRENCY", "2")))

# If the provider's numeric prices need conversion, change this. Normally 1.
PROVIDER_PRICE_MULTIPLIER = Decimal(os.getenv("PROVIDER_PRICE_MULTIPLIER", "1"))

SUPPORT_TEXT = os.getenv(
    "SUPPORT_TEXT",
    "Support: contact the administrator of this bot.",
)
CURRENCY_SYMBOL = "$"
UNITS_PER_USD = 1_000_000  # 1 unit = 1/1,000,000 of a currency unit — matches legacy DB format
MAX_CATALOG_ITEMS_SHOWN = 30
CATALOG_COUNTRIES_PAGE_SIZE = 5
USER_COUNTRIES_PAGE_SIZE = 20
MAX_ORDER_HISTORY = 15
MAX_BROADCAST_LENGTH = 3900


# Conversation states
(
    CAT_COUNTRY_ID,
    CAT_COUNTRY_NAME,
    CAT_SERVICE_CODE,
    CAT_SERVICE_NAME,
    CAT_EXTRA_PARAMS,
    CAT_MARKUP,
) = range(10, 16)
SET_MARKUP = 20

# VPN/Proxy shop workflows
VSHOP_ADD_CATEGORY = 130
VSHOP_ADD_SERVICE = 131
VSHOP_SET_PRICE = 132
VSHOP_UPLOAD_FILE = 133
VSHOP_QUANTITY = 134
VSHOP_RENAME_CATEGORY = 135
VSHOP_RENAME_SERVICE = 136
VSHOP_SET_DURATION = 137
VSHOP_PROVIDE_CONTENT = 138
VSHOP_PROVIDE_CONFIRM = 139
VSHOP_ORDER_LOOKUP = 140
VSHOP_SET_NOTICE_GROUP = 141
VSHOP_CUSTOM_QUANTITY = 142
VSHOP_SET_QTY = 143
BUY_COUNTRY_SEARCH = 144

# Legacy defaults for services created before per-service limits were added.
VSHOP_MIN_QTY = 1
VSHOP_MAX_QTY = 15


# =============================================================================
# Logging and helpers
# =============================================================================

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
# Persist errors on the hosting server without changing existing console logs.
# The file rotates at 5 MB and keeps five backups to prevent unlimited growth.
_ERROR_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.log")
_error_file_handler = RotatingFileHandler(
    _ERROR_LOG_PATH,
    maxBytes=5 * 1024 * 1024,
    backupCount=5,
    encoding="utf-8",
)
_error_file_handler.setLevel(logging.ERROR)
_error_file_handler.setFormatter(logging.Formatter(
    "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
))
logging.getLogger().addHandler(_error_file_handler)

# HTTPX INFO records include full Telegram Bot API URLs, which contain the bot
# token. Keep request URLs out of logs while preserving warnings and errors.
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger("paid-number-bot")


def _log_uncaught_exception(exc_type, exc_value, exc_traceback):
    """Write otherwise-unhandled Python exceptions to error.log as well."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logger.critical(
        "Uncaught exception",
        exc_info=(exc_type, exc_value, exc_traceback),
    )


sys.excepthook = _log_uncaught_exception


@dataclass(slots=True)
class _UpdateLockEntry:
    lock: asyncio.Lock
    users: int = 0


class PerUserUpdateProcessor(BaseUpdateProcessor):
    """Process different users concurrently while preserving per-user order.

    ConversationHandler needs sequential updates for a conversation. A global
    sequential queue, however, lets one slow API/file operation delay everyone.
    This processor keeps updates from the same Telegram user serialized while
    allowing unrelated users to run in parallel.
    """

    def __init__(self, max_concurrent_updates: int):
        super().__init__(max_concurrent_updates=max_concurrent_updates)
        self._entries: dict[tuple[str, int], _UpdateLockEntry] = {}

    @staticmethod
    def _update_key(update: object) -> tuple[str, int]:
        if isinstance(update, Update):
            if update.effective_user:
                return ("user", int(update.effective_user.id))
            if update.effective_chat:
                return ("chat", int(update.effective_chat.id))
        return ("global", 0)

    async def do_process_update(self, update: object, coroutine: Any) -> None:
        key = self._update_key(update)
        entry = self._entries.get(key)
        if entry is None:
            entry = _UpdateLockEntry(asyncio.Lock())
            self._entries[key] = entry
        entry.users += 1
        try:
            async with entry.lock:
                await coroutine
        finally:
            entry.users -= 1
            if entry.users <= 0 and not entry.lock.locked():
                self._entries.pop(key, None)

    async def initialize(self) -> None:
        return None

    async def shutdown(self) -> None:
        self._entries.clear()
 





FILE_JOB_SEMAPHORE = asyncio.Semaphore(FILE_TASK_CONCURRENCY)


async def run_file_job(function: Any, *args: Any) -> Any:
    """Run CPU/disk-heavy export/backup work without blocking bot updates."""
    async with FILE_JOB_SEMAPHORE:
        return await asyncio.to_thread(function, *args)


def now_ts() -> int:
    return int(time.time())


def parse_admin_ids(raw: str) -> set[int]:
    result: set[int] = set()
    for item in re.split(r"[,\s]+", raw.strip()):
        if item.isdigit():
            result.add(int(item))
    return result


ADMIN_IDS = parse_admin_ids(ADMIN_IDS_RAW)


def is_admin(user_id: Optional[int]) -> bool:
    return user_id is not None and user_id in ADMIN_IDS


def decimal_value(value: Any) -> Decimal:
    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise ValueError(f"Invalid decimal value: {value!r}") from exc


def usd_to_units(value: Any) -> int:
    amount = decimal_value(value)
    return int((amount * Decimal(str(UNITS_PER_USD))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def units_to_decimal(units: int) -> Decimal:
    return Decimal(int(units)) / Decimal(str(UNITS_PER_USD))


def short_text(value: Any, limit: int = 80) -> str:
    text = str(value or "").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def parse_extra_params(raw: str) -> dict[str, str]:
    """Accept '-', JSON object, or URL-query style key=value&key2=value2."""
    raw = raw.strip()
    if not raw or raw == "-":
        return {}
    if raw.startswith("{"):
        data = json.loads(raw)
        if not isinstance(data, dict):
            raise ValueError("JSON must be an object")
        return {str(k): str(v) for k, v in data.items() if str(k).strip()}

    output: dict[str, str] = {}
    for pair in raw.split("&"):
        if "=" not in pair:
            raise ValueError("Use key=value&key2=value2")
        key, value = pair.split("=", 1)
        key = key.strip()
        if not key:
            raise ValueError("Parameter key cannot be empty")
        output[key] = value.strip()
    return output


def smsbower_countries_from_csv() -> list[tuple[str, str]]:
    """Return the built-in SMSBower country table; no external CSV is required.

    The function name is retained for compatibility with existing call sites.
    Country IDs, names, and flags are embedded in this Python file so deployment
    only needs this script (plus the bot's normal runtime configuration).
    """
    return [
        (str(country_id), str(country_name))
        for country_id, country_name, _flag in _SMSBOWER_COUNTRY_DATA
    ]


def safe_json_loads(raw: str, default: Any) -> Any:
    try:
        return json.loads(raw)
    except Exception:
        return default


def status_label(status: str) -> str:
    return {
        "waiting": "⏳ Waiting for SMS",
        "otp_received": "✅ OTP received",
        "cancelled": "↩️ Cancelled / refunded",
        "expired": "⌛ Expired",
        "failed": "❌ Failed",
    }.get(status, status.replace("_", " ").title())


def admin_menu_markup() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("📊 Statistics", callback_data="admin:stats"),
                InlineKeyboardButton("🌍 Countries & services", callback_data="admin:catalog"),
            ],
            [
                InlineKeyboardButton("➕ Add service", callback_data="admin:addcatalog"),
                InlineKeyboardButton("🧾 Order logs", callback_data="admin:orderlogs"),
            ],
            [
                InlineKeyboardButton("⚙️ Profit markup", callback_data="admin:setmarkup"),
                InlineKeyboardButton("🔌 Provider balance", callback_data="admin:providerbalance"),
            ],
            [
                InlineKeyboardButton("📧 Mail Services", callback_data="mail:admin:open"),
                InlineKeyboardButton("🧩 Panel Management", callback_data="panel:home"),
            ],
            [InlineKeyboardButton("⭐ Premium Emoji", callback_data="premiumemoji:menu")],
        ]
    )


# =============================================================================
# Database
# =============================================================================


class Database:
    def __init__(self, path: str):
        self.conn = sqlite3.connect(
            path,
            check_same_thread=False,
            timeout=30.0,
        )
        self.conn.row_factory = sqlite3.Row
        self.lock = threading.RLock()
        with self.lock:
            self.conn.execute("PRAGMA journal_mode=WAL")
            self.conn.execute("PRAGMA synchronous=NORMAL")
            self.conn.execute("PRAGMA foreign_keys=ON")
            self.conn.execute("PRAGMA busy_timeout=30000")
            self.conn.execute("PRAGMA temp_store=MEMORY")
            self.conn.execute("PRAGMA cache_size=-32768")
            self.conn.execute("PRAGMA wal_autocheckpoint=1000")
            with suppress(sqlite3.Error):
                self.conn.execute("PRAGMA mmap_size=67108864")
        self.create_schema()

    def create_schema(self) -> None:
        schema = """
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            first_name TEXT,
            balance_units INTEGER NOT NULL DEFAULT 0,
            accepted_terms INTEGER NOT NULL DEFAULT 0,
            is_blocked INTEGER NOT NULL DEFAULT 0,
            created_at INTEGER NOT NULL,
            updated_at INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS catalog (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            country_id TEXT NOT NULL,
            country_name TEXT NOT NULL,
            service_code TEXT NOT NULL,
            service_name TEXT NOT NULL,
            extra_params_json TEXT NOT NULL DEFAULT '{}',
            markup_override_units INTEGER,
            enabled INTEGER NOT NULL DEFAULT 1,
            created_at INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS country_favorites (
            user_id INTEGER NOT NULL,
            service_code TEXT NOT NULL,
            country_id TEXT NOT NULL,
            created_at INTEGER NOT NULL,
            PRIMARY KEY(user_id, service_code, country_id)
        );
        CREATE INDEX IF NOT EXISTS idx_country_favorites_service
            ON country_favorites(service_code, country_id);

        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            catalog_id INTEGER,
            activation_id TEXT NOT NULL UNIQUE,
            phone TEXT NOT NULL,
            provider_id TEXT,
            country_name TEXT NOT NULL,
            service_name TEXT NOT NULL,
            provider_cost_units INTEGER NOT NULL,
            charged_units INTEGER NOT NULL,
            status TEXT NOT NULL,
            otp_code TEXT,
            last_api_status TEXT,
            created_at INTEGER NOT NULL,
            expires_at INTEGER NOT NULL,
            updated_at INTEGER NOT NULL,
            refunded INTEGER NOT NULL DEFAULT 0,
            message_id INTEGER,
            FOREIGN KEY(user_id) REFERENCES users(user_id)
        );

        CREATE INDEX IF NOT EXISTS idx_orders_user ON orders(user_id, id DESC);
        CREATE INDEX IF NOT EXISTS idx_orders_active ON orders(status, expires_at);

        CREATE TABLE IF NOT EXISTS deposits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            amount_units INTEGER NOT NULL,
            txid TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'approved',
            created_at INTEGER NOT NULL,
            reviewed_at INTEGER,
            admin_id INTEGER,
            FOREIGN KEY(user_id) REFERENCES users(user_id)
        );

        CREATE INDEX IF NOT EXISTS idx_deposits_status ON deposits(status, id DESC);

        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            amount_units INTEGER NOT NULL,
            kind TEXT NOT NULL,
            reference TEXT,
            note TEXT,
            created_at INTEGER NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(user_id)
        );

        CREATE TABLE IF NOT EXISTS referral_rewards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            referrer_id INTEGER NOT NULL,
            referred_user_id INTEGER NOT NULL UNIQUE,
            deposit_id INTEGER NOT NULL UNIQUE,
            amount_units INTEGER NOT NULL,
            created_at INTEGER NOT NULL,
            FOREIGN KEY(referrer_id) REFERENCES users(user_id),
            FOREIGN KEY(referred_user_id) REFERENCES users(user_id),
            FOREIGN KEY(deposit_id) REFERENCES deposits(id)
        );

        CREATE TABLE IF NOT EXISTS how_to_use_buttons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            video_file_id TEXT,
            created_at INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS vshop_categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            enabled INTEGER NOT NULL DEFAULT 1,
            created_at INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS vshop_services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            price_units INTEGER NOT NULL DEFAULT 0,
            enabled INTEGER NOT NULL DEFAULT 1,
            automation INTEGER NOT NULL DEFAULT 0,
            duration_days INTEGER NOT NULL DEFAULT 0,
            duration_label TEXT NOT NULL DEFAULT '',
            min_quantity INTEGER NOT NULL DEFAULT 1,
            max_quantity INTEGER NOT NULL DEFAULT 15,
            created_at INTEGER NOT NULL,
            UNIQUE(category_id, name),
            FOREIGN KEY(category_id) REFERENCES vshop_categories(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS vshop_inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_id INTEGER NOT NULL,
            kind TEXT NOT NULL,
            email TEXT,
            password TEXT,
            proxy_line TEXT,
            sold INTEGER NOT NULL DEFAULT 0,
            sold_to INTEGER,
            sold_order_id INTEGER,
            sold_at INTEGER,
            created_at INTEGER NOT NULL,
            FOREIGN KEY(service_id) REFERENCES vshop_services(id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_vshop_inventory_stock
            ON vshop_inventory(service_id, sold, id);

        CREATE TABLE IF NOT EXISTS vshop_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            service_id INTEGER NOT NULL,
            category_id INTEGER,
            quantity INTEGER NOT NULL,
            auto_delivered INTEGER NOT NULL DEFAULT 0,
            pending_quantity INTEGER NOT NULL DEFAULT 0,
            total_units INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'success',
            delivery_info TEXT,
            delivered_at INTEGER,
            notice_chat_id INTEGER,
            notice_message_id INTEGER,
            created_at INTEGER NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(user_id),
            FOREIGN KEY(service_id) REFERENCES vshop_services(id)
        );

        CREATE TABLE IF NOT EXISTS vshop_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            inventory_id INTEGER NOT NULL UNIQUE,
            FOREIGN KEY(order_id) REFERENCES vshop_orders(id) ON DELETE CASCADE,
            FOREIGN KEY(inventory_id) REFERENCES vshop_inventory(id)
        );

        CREATE TABLE IF NOT EXISTS purchase_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            source TEXT NOT NULL,
            source_order_id INTEGER NOT NULL,
            created_at INTEGER NOT NULL,
            product_name TEXT NOT NULL,
            item_detail TEXT NOT NULL DEFAULT '',
            amount_units INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT '',
            refunded INTEGER NOT NULL DEFAULT 0,
            UNIQUE(source, source_order_id)
        );
        CREATE INDEX IF NOT EXISTS idx_purchase_history_user_recent
            ON purchase_history(user_id, created_at DESC, id DESC);
        """
        with self.lock:
            self.conn.executescript(schema)
            # Safe migration for databases created by an older bot version.
            order_columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(orders)")}
            if "provider_id" not in order_columns:
                self.conn.execute("ALTER TABLE orders ADD COLUMN provider_id TEXT")
            user_columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(users)")}
            if "referred_by" not in user_columns:
                self.conn.execute("ALTER TABLE users ADD COLUMN referred_by INTEGER")
            # vshop_services migrations
            svc_columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(vshop_services)")}
            if "automation" not in svc_columns:
                self.conn.execute("ALTER TABLE vshop_services ADD COLUMN automation INTEGER NOT NULL DEFAULT 0")
            if "duration_days" not in svc_columns:
                self.conn.execute("ALTER TABLE vshop_services ADD COLUMN duration_days INTEGER NOT NULL DEFAULT 0")
            if "duration_label" not in svc_columns:
                self.conn.execute("ALTER TABLE vshop_services ADD COLUMN duration_label TEXT NOT NULL DEFAULT ''")
            if "min_quantity" not in svc_columns:
                self.conn.execute("ALTER TABLE vshop_services ADD COLUMN min_quantity INTEGER NOT NULL DEFAULT 1")
            if "max_quantity" not in svc_columns:
                self.conn.execute("ALTER TABLE vshop_services ADD COLUMN max_quantity INTEGER NOT NULL DEFAULT 15")
            # vshop_orders migrations
            vorder_cols = {row["name"] for row in self.conn.execute("PRAGMA table_info(vshop_orders)")}
            if "category_id" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN category_id INTEGER")
            if "auto_delivered" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN auto_delivered INTEGER NOT NULL DEFAULT 0")
            if "pending_quantity" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN pending_quantity INTEGER NOT NULL DEFAULT 0")
            if "status" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN status TEXT NOT NULL DEFAULT 'success'")
            if "delivery_info" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN delivery_info TEXT")
            if "delivered_at" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN delivered_at INTEGER")
            if "notice_chat_id" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN notice_chat_id INTEGER")
            if "notice_message_id" not in vorder_cols:
                self.conn.execute("ALTER TABLE vshop_orders ADD COLUMN notice_message_id INTEGER")
            # Create these only after the migrations above. Older databases may
            # have vshop_orders without the status column yet.
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_vshop_orders_status "
                "ON vshop_orders(status, id DESC)"
            )
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_vshop_orders_user "
                "ON vshop_orders(user_id, id DESC)"
            )
            self.conn.execute(
                "INSERT OR IGNORE INTO settings(key,value) VALUES('vshop_notice_group_id','')"
            )
            self.conn.execute(
                "INSERT OR IGNORE INTO settings(key,value) VALUES('vshop_manage_inline_enabled','1')"
            )
            self.conn.execute(
                "INSERT OR IGNORE INTO settings(key,value) VALUES('markup_usd',?)",
                (DEFAULT_MARKUP_USD,),
            )
            self.conn.execute(
                "INSERT OR IGNORE INTO settings(key,value) VALUES('expiry_minutes',?)",
                (str(DEFAULT_EXPIRY_MINUTES),),
            )
            self.conn.commit()

    def execute(self, sql: str, params: Iterable[Any] = ()) -> sqlite3.Cursor:
        with self.lock:
            cur = self.conn.execute(sql, tuple(params))
            self.conn.commit()
            return cur

    def fetchone(self, sql: str, params: Iterable[Any] = ()) -> Optional[sqlite3.Row]:
        with self.lock:
            return self.conn.execute(sql, tuple(params)).fetchone()

    def fetchall(self, sql: str, params: Iterable[Any] = ()) -> list[sqlite3.Row]:
        with self.lock:
            return self.conn.execute(sql, tuple(params)).fetchall()

    def register_user(self, user: Any) -> sqlite3.Row:
        ts = now_ts()
        with self.lock:
            self.conn.execute(
                """
                INSERT INTO users(user_id,username,first_name,created_at,updated_at)
                VALUES(?,?,?,?,?)
                ON CONFLICT(user_id) DO UPDATE SET
                    username=excluded.username,
                    first_name=excluded.first_name,
                    updated_at=excluded.updated_at
                """,
                (user.id, user.username, user.first_name, ts, ts),
            )
            self.conn.commit()
            return self.conn.execute("SELECT * FROM users WHERE user_id=?", (user.id,)).fetchone()

    def get_setting(self, key: str, default: str = "") -> str:
        row = self.fetchone("SELECT value FROM settings WHERE key=?", (key,))
        return str(row["value"]) if row else default

    def set_setting(self, key: str, value: str) -> None:
        self.execute(
            "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )

    def set_referrer_from_start(self, user_id: int, referrer_id: int) -> bool:
        if user_id <= 0 or referrer_id <= 0 or user_id == referrer_id:
            return False
        with self.lock:
            cur = self.conn.execute(
                """
                UPDATE users
                SET referred_by=?, updated_at=?
                WHERE user_id=? AND (referred_by IS NULL OR referred_by=0)
                """,
                (referrer_id, now_ts(), user_id),
            )
            self.conn.commit()
            return cur.rowcount == 1

    def award_referral_for_deposit(self, deposit_id: int) -> Optional[sqlite3.Row]:
        """Credit the referrer once when a qualifying deposit is approved."""
        with self.lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                deposit = self.conn.execute(
                    "SELECT * FROM deposits WHERE id=? AND status='approved'",
                    (deposit_id,),
                ).fetchone()
                if not deposit:
                    self.conn.rollback()
                    return None
                user = self.conn.execute(
                    "SELECT referred_by FROM users WHERE user_id=?",
                    (deposit["user_id"],),
                ).fetchone()
                referrer_id = int(user["referred_by"] or 0) if user else 0
                threshold = bdt_to_units(self.get_setting("referral_min_deposit_bdt", "0"))
                reward = bdt_to_units(self.get_setting("referral_reward_bdt", "0"))
                referrer = self.conn.execute(
                    "SELECT 1 FROM users WHERE user_id=?", (referrer_id,)
                ).fetchone()
                if not referrer or reward <= 0 or threshold <= 0:
                    self.conn.rollback()
                    return None
                approved_total = self.conn.execute(
                    """
                    SELECT COALESCE(SUM(amount_units), 0) AS total
                    FROM deposits
                    WHERE user_id=? AND status='approved'
                    """,
                    (deposit["user_id"],),
                ).fetchone()
                if int(approved_total["total"] or 0) < threshold:
                    self.conn.rollback()
                    return None
                inserted = self.conn.execute(
                    """
                    INSERT OR IGNORE INTO referral_rewards(
                        referrer_id,referred_user_id,deposit_id,amount_units,created_at
                    ) VALUES(?,?,?,?,?)
                    """,
                    (referrer_id, deposit["user_id"], deposit_id, reward, now_ts()),
                )
                if inserted.rowcount != 1:
                    self.conn.rollback()
                    return None
                self.conn.execute(
                    "UPDATE users SET balance_units=balance_units+?,updated_at=? WHERE user_id=?",
                    (reward, now_ts(), referrer_id),
                )
                self.conn.execute(
                    """
                    INSERT INTO transactions(
                        user_id,amount_units,kind,reference,note,created_at
                    ) VALUES(?,?,?,?,?,?)
                    """,
                    (
                        referrer_id,
                        reward,
                        "referral_reward",
                        str(deposit_id),
                        f"Referral reward for user {deposit['user_id']}",
                        now_ts(),
                    ),
                )
                self.conn.commit()
                return self.conn.execute(
                    "SELECT * FROM referral_rewards WHERE deposit_id=?",
                    (deposit_id,),
                ).fetchone()
            except Exception:
                self.conn.rollback()
                raise

    def global_markup_units(self) -> int:
        return usd_to_units(self.get_setting("markup_usd", DEFAULT_MARKUP_USD))

    def expiry_seconds(self) -> int:
        raw = self.get_setting("expiry_minutes", str(DEFAULT_EXPIRY_MINUTES))
        try:
            minutes = max(1, int(raw))
        except ValueError:
            minutes = DEFAULT_EXPIRY_MINUTES
        return minutes * 60

    def hold_balance(self, user_id: int, amount_units: int, reference: str) -> bool:
        ts = now_ts()
        with self.lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                cur = self.conn.execute(
                    """
                    UPDATE users
                    SET balance_units=balance_units-?, updated_at=?
                    WHERE user_id=? AND is_blocked=0 AND balance_units>=?
                    """,
                    (amount_units, ts, user_id, amount_units),
                )
                if cur.rowcount != 1:
                    self.conn.rollback()
                    return False
                self.conn.execute(
                    "INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at) VALUES(?,?,?,?,?,?)",
                    (user_id, -amount_units, "purchase_hold", reference, "Number purchase", ts),
                )
                self.conn.commit()
                return True
            except Exception:
                self.conn.rollback()
                raise

    def credit_balance(
        self,
        user_id: int,
        amount_units: int,
        kind: str,
        reference: str,
        note: str,
    ) -> None:
        ts = now_ts()
        with self.lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                self.conn.execute(
                    "UPDATE users SET balance_units=balance_units+?, updated_at=? WHERE user_id=?",
                    (amount_units, ts, user_id),
                )
                self.conn.execute(
                    "INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at) VALUES(?,?,?,?,?,?)",
                    (user_id, amount_units, kind, reference, note, ts),
                )
                self.conn.commit()
            except Exception:
                self.conn.rollback()
                raise

    def refund_order(self, order_id: int, api_status: str) -> bool:
        ts = now_ts()
        with self.lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                order = self.conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
                if not order or order["refunded"] or order["status"] != "waiting":
                    self.conn.rollback()
                    return False
                self.conn.execute(
                    "UPDATE users SET balance_units=balance_units+?, updated_at=? WHERE user_id=?",
                    (order["charged_units"], ts, order["user_id"]),
                )
                self.conn.execute(
                    """
                    UPDATE orders SET status='cancelled', refunded=1,
                    last_api_status=?, updated_at=? WHERE id=?
                    """,
                    (api_status, ts, order_id),
                )
                self.conn.execute(
                    "INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at) VALUES(?,?,?,?,?,?)",
                    (
                        order["user_id"],
                        order["charged_units"],
                        "cancel_refund",
                        str(order_id),
                        "Number cancelled before OTP",
                        ts,
                    ),
                )
                self.conn.commit()
                return True
            except Exception:
                self.conn.rollback()
                raise

    def complete_auto_deposit(self, invoice_id: str) -> Optional[sqlite3.Row]:
        """Credit and record an invoice exactly once after the provider says paid."""
        ts = now_ts()
        with self.lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                invoice = self.conn.execute(
                    "SELECT * FROM payment_invoices WHERE invoice_id=? AND status='awaiting_payment'",
                    (str(invoice_id),),
                ).fetchone()
                if not invoice:
                    self.conn.rollback()
                    return None

                updated = self.conn.execute(
                    """
                    UPDATE payment_invoices
                    SET status='paid', updated_at=?
                    WHERE invoice_id=? AND status='awaiting_payment'
                    """,
                    (ts, str(invoice_id)),
                )
                if updated.rowcount != 1:
                    self.conn.rollback()
                    return None

                cur = self.conn.execute(
                    """
                    INSERT INTO deposits(user_id,amount_units,txid,status,created_at,reviewed_at)
                    VALUES(?,?,?,'approved',?,?)
                    """,
                    (
                        invoice["user_id"],
                        invoice["amount_units"],
                        str(invoice_id),
                        ts,
                        ts,
                    ),
                )
                self.conn.execute(
                    "UPDATE users SET balance_units=balance_units+?,updated_at=? WHERE user_id=?",
                    (invoice["amount_units"], ts, invoice["user_id"]),
                )
                self.conn.execute(
                    """
                    INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at)
                    VALUES(?,?,?,?,?,?)
                    """,
                    (
                        invoice["user_id"],
                        invoice["amount_units"],
                        "deposit",
                        str(invoice_id),
                        "IPShop Auto Payment",
                        ts,
                    ),
                )
                deposit = self.conn.execute(
                    "SELECT * FROM deposits WHERE id=?",
                    (cur.lastrowid,),
                ).fetchone()
                self.conn.commit()
                return deposit
            except Exception:
                self.conn.rollback()
                raise


def _sqlite_read_uri(path: str, *, immutable: bool) -> str:
    encoded_path = quote(os.path.abspath(path), safe="/")
    uri = f"file:{encoded_path}?mode=ro"
    return uri + ("&immutable=1" if immutable else "")


def _database_file_summary(
    path: str, *, immutable: bool = True
) -> dict[str, int]:
    """Read and validate a bot database without changing its contents."""
    with open(path, "rb") as handle:
        if handle.read(16) != b"SQLite format 3\x00":
            raise ValueError("Not a SQLite database.")

    connection = sqlite3.connect(
        _sqlite_read_uri(path, immutable=immutable),
        uri=True,
        timeout=10,
    )
    try:
        integrity = connection.execute("PRAGMA integrity_check").fetchone()
        if not integrity or str(integrity[0]).strip().lower() != "ok":
            raise ValueError(
                f"SQLite integrity check failed: {integrity[0] if integrity else 'unknown'}"
            )

        tables = {
            str(row[0])
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }
        missing = sorted(DATABASE_REQUIRED_TABLES - tables)
        if missing:
            raise ValueError("Missing required tables: " + ", ".join(missing))

        user_columns = {
            str(row[1]) for row in connection.execute("PRAGMA table_info(users)")
        }
        if "balance_units" not in user_columns:
            raise ValueError("The users table has no balance column.")

        summary = {
            "users": int(connection.execute("SELECT COUNT(*) FROM users").fetchone()[0]),
            "balance_units": int(
                connection.execute(
                    "SELECT COALESCE(SUM(balance_units),0) FROM users"
                ).fetchone()[0]
                or 0
            ),
        }
        for table in (
            "orders",
            "deposits",
            "transactions",
            "local_orders",
            "vshop_orders",
            "purchase_history",
        ):
            summary[table] = (
                int(connection.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
                if table in tables
                else 0
            )
        return summary
    finally:
        connection.close()


def _database_activity_count(summary: dict[str, int]) -> int:
    return sum(
        int(summary.get(table, 0))
        for table in (
            "orders",
            "deposits",
            "transactions",
            "local_orders",
            "vshop_orders",
            "purchase_history",
        )
    )


def _database_backup_paths() -> list[str]:
    if not os.path.isdir(BACKUP_DIRECTORY):
        return []
    candidates = []
    for name in os.listdir(BACKUP_DIRECTORY):
        if not name.lower().endswith(".db"):
            continue
        if name.startswith(("database_export_", "pre_restore_")) or name == DATABASE_RECOVERY_ANCHOR:
            candidates.append(os.path.join(BACKUP_DIRECTORY, name))
    return candidates


def _valid_database_backups() -> list[tuple[str, dict[str, int]]]:
    valid: list[tuple[str, dict[str, int]]] = []
    for path in _database_backup_paths():
        try:
            valid.append((path, _database_file_summary(path, immutable=True)))
        except Exception as exc:
            logger.warning("Ignoring invalid database backup %s: %s", path, exc)
    return valid


def _select_recovery_backup(
    candidates: list[tuple[str, dict[str, int]]],
) -> Optional[tuple[str, dict[str, int]]]:
    if not candidates:
        return None
    greatest_user_count = max(summary["users"] for _, summary in candidates)
    # Ignore a recently-created, nearly empty database when an older full copy
    # is available; among similarly complete copies, prefer the newest one.
    minimum_users = math.ceil(greatest_user_count * 0.9)
    user_eligible = [
        candidate
        for candidate in candidates
        if candidate[1]["users"] >= minimum_users
    ]
    greatest_activity = max(
        _database_activity_count(summary) for _, summary in user_eligible
    )
    if greatest_activity:
        minimum_activity = max(1, math.ceil(greatest_activity * 0.9))
        eligible = [
            candidate
            for candidate in user_eligible
            if _database_activity_count(candidate[1]) >= minimum_activity
        ]
    else:
        eligible = user_eligible
    return max(eligible, key=lambda item: os.path.getmtime(item[0]))


def _preserve_database_file_and_sidecars(
    database_path: str, *, label: str
) -> None:
    existing = [
        database_path,
        database_path + "-wal",
        database_path + "-shm",
    ]
    if not any(os.path.exists(path) for path in existing):
        return
    os.makedirs(BACKUP_DIRECTORY, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    for suffix, source in zip(("", "-wal", "-shm"), existing):
        if os.path.isfile(source):
            destination = os.path.join(
                BACKUP_DIRECTORY,
                f"{label}_{stamp}.db{suffix}",
            )
            shutil.copy2(source, destination)


def prepare_database_before_open() -> None:
    """Recover a missing/corrupt DB before Database() can create a blank one."""
    global DATABASE_STARTUP_NOTICE
    database_path = os.path.abspath(os.path.expanduser(DATABASE_PATH))
    exists = os.path.isfile(database_path)
    has_sidecars = any(
        os.path.exists(database_path + suffix) for suffix in ("-wal", "-shm")
    )
    current_summary: Optional[dict[str, int]] = None
    current_error: Optional[Exception] = None
    if exists:
        try:
            current_summary = _database_file_summary(database_path, immutable=False)
        except Exception as exc:
            current_error = exc

    candidates = _valid_database_backups()
    if current_summary is not None:
        selected = _select_recovery_backup(candidates)
        if selected:
            backup_path, backup_summary = selected
            current_users = current_summary["users"]
            if (
                backup_summary["users"] >= 20
                and backup_summary["users"] >= max(20, current_users * 5)
                and current_summary["balance_units"] == 0
                and _database_activity_count(current_summary) == 0
                and _database_activity_count(backup_summary) > 0
            ):
                DATABASE_STARTUP_NOTICE = (
                    "⚠️ The current database looks newly initialized: "
                    f"{current_users} users and no balances or transactions. "
                    f"A valid fuller backup is available: {os.path.basename(backup_path)} "
                    f"({backup_summary['users']} users). It was not restored "
                    "automatically to avoid overwriting newer activity. Check "
                    "Admin Panel → Database Settings → Upload DB."
                )
        return

    backup_files = _database_backup_paths()
    if not backup_files and not exists and not has_sidecars:
        return  # Brand-new installation; there is nothing to recover.
    if not candidates:
        reason = (
            f" The existing database is invalid: {current_error}"
            if current_error
            else " SQLite sidecar files exist without the main database."
            if has_sidecars
            else ""
        )
        raise RuntimeError(
            "The bot database is missing or invalid, and no valid recovery backup "
            f"was found.{reason} Refusing to initialize an empty database."
        )

    selected = _select_recovery_backup(candidates)
    assert selected is not None
    source_path, summary = selected
    os.makedirs(os.path.dirname(database_path) or ".", exist_ok=True)
    _preserve_database_file_and_sidecars(database_path, label="pre_recovery")
    staging_path = database_path + ".recovery.tmp"
    for suffix in ("-wal", "-shm"):
        with suppress(FileNotFoundError):
            os.remove(staging_path + suffix)
    with suppress(FileNotFoundError):
        os.remove(staging_path)
    shutil.copy2(source_path, staging_path)
    _database_file_summary(staging_path, immutable=True)
    for suffix in ("-wal", "-shm"):
        with suppress(FileNotFoundError):
            os.remove(database_path + suffix)
    os.replace(staging_path, database_path)
    DATABASE_STARTUP_NOTICE = (
        "✅ The bot automatically recovered its missing or invalid database from "
        f"{os.path.basename(source_path)}. "
        f"Recovered {summary['users']} users and "
        f"{summary['orders'] + summary['local_orders'] + summary['vshop_orders']} "
        "number/VPN orders. Please verify balances and recent activity."
    )
    logger.warning("%s", DATABASE_STARTUP_NOTICE)


prepare_database_before_open()
DB = Database(DATABASE_PATH)


# =============================================================================
# SMS provider API adapter
# =============================================================================


@dataclass
class LiveOffer:
    provider_cost_units: int
    available: int
    raw: Any


@dataclass
class ProviderOffer:
    provider_id: str
    provider_cost_units: int
    available: int
    raw: Any


@dataclass
class PurchasedNumber:
    activation_id: str
    phone: str
    provider_cost_units: Optional[int]
    expires_in: Optional[int]
    raw: Any


@dataclass
class ActivationStatus:
    state: str  # waiting, code, cancelled, expired, unknown, error
    code: Optional[str]
    raw_text: str
    raw: Any


class SmsPowerAPI:
    def __init__(self, base_url: str, api_key: str):
        self.base_url = base_url
        self.api_key = api_key
        timeout = aiohttp.ClientTimeout(total=HTTP_TIMEOUT_SECONDS)
        connector = aiohttp.TCPConnector(
            limit=PROVIDER_HTTP_CONNECTIONS,
            limit_per_host=PROVIDER_HTTP_CONNECTIONS,
            ttl_dns_cache=300,
            enable_cleanup_closed=True,
        )
        self.session = aiohttp.ClientSession(timeout=timeout, connector=connector)
        self.request_semaphore = asyncio.Semaphore(PROVIDER_MAX_CONCURRENCY)
        # Live lowest prices from the most recent getPrices response, keyed by service
        # then country ID. Used to sort country pages cheapest-first without making
        # one extra API request per country.
        self.country_price_cache: dict[str, dict[str, int]] = {}

    async def close(self) -> None:
        await self.session.close()

    async def request(self, action: str, **params: Any) -> tuple[Any, str]:
        query: dict[str, Any] = {"api_key": self.api_key, "action": action}
        for key, value in params.items():
            if value is not None and value != "":
                query[key] = value

        try:
            async with self.request_semaphore:
                async with self.session.get(self.base_url, params=query) as response:
                    text = (await response.text()).strip()
                    if response.status >= 400:
                        raise RuntimeError(f"Provider HTTP {response.status}: {short_text(text, 300)}")
        except asyncio.TimeoutError as exc:
            raise RuntimeError("Provider request timed out") from exc
        except aiohttp.ClientError as exc:
            raise RuntimeError(f"Provider connection error: {exc}") from exc

        data: Any = text
        if text.startswith("{") or text.startswith("["):
            try:
                data = json.loads(text)
            except json.JSONDecodeError:
                data = text
        return data, text

    async def get_balance(self) -> Decimal:
        data, text = await self.request("getBalance")
        if isinstance(data, dict):
            for key in ("balance", "data", "amount"):
                if key in data and not isinstance(data[key], (dict, list)):
                    return decimal_value(data[key])
            if isinstance(data.get("data"), dict):
                for key in ("balance", "amount"):
                    if key in data["data"]:
                        return decimal_value(data["data"][key])
        match = re.search(r"ACCESS_BALANCE:([-+]?\d+(?:\.\d+)?)", text, re.I)
        if match:
            return decimal_value(match.group(1))
        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
            return decimal_value(text)
        raise RuntimeError(f"Could not parse provider balance: {short_text(text, 200)}")

    @staticmethod
    def _extract_cost_count(node: Any) -> Optional[tuple[Decimal, int]]:
        if isinstance(node, (int, float, str, Decimal)):
            return decimal_value(node), 0
        if not isinstance(node, dict):
            return None

        cost = None
        for key in ("cost", "price", "amount", "rate"):
            if key in node:
                try:
                    cost = decimal_value(node[key])
                    break
                except ValueError:
                    pass
        if cost is None:
            return None

        count = -1
        for key in ("count", "available", "stock", "quantity", "qty"):
            if key in node:
                try:
                    count = int(decimal_value(node[key]))
                except Exception:
                    count = 0
                break
        return cost, max(-1, count)

    @classmethod
    def _search_offer(cls, data: Any, country: str, service: str) -> Optional[tuple[Decimal, int]]:
        country_s = str(country)
        service_s = str(service)

        # Standard shape: {"country_id": {"service": {"cost": ..., "count": ...}}}
        if isinstance(data, dict):
            for country_key in (country_s, int(country_s) if country_s.isdigit() else None):
                if country_key is not None and country_key in data:
                    country_node = data[country_key]
                    if isinstance(country_node, dict):
                        for service_key in (service_s, service_s.lower(), service_s.upper()):
                            if service_key in country_node:
                                found = cls._extract_cost_count(country_node[service_key])
                                if found:
                                    return found

            # Wrapped response shapes.
            for wrapper in ("data", "prices", "result", "offers"):
                if wrapper in data:
                    found = cls._search_offer(data[wrapper], country_s, service_s)
                    if found:
                        return found

            # Single offer object or a dict keyed another way.
            direct = cls._extract_cost_count(data)
            if direct:
                node_country = str(data.get("country", data.get("country_id", country_s)))
                node_service = str(data.get("service", data.get("service_code", service_s)))
                if node_country == country_s and node_service.lower() == service_s.lower():
                    return direct

            for value in data.values():
                found = cls._search_offer(value, country_s, service_s)
                if found:
                    return found

        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    item_country = str(item.get("country", item.get("country_id", country_s)))
                    item_service = str(item.get("service", item.get("service_code", item.get("code", service_s))))
                    if item_country == country_s and item_service.lower() == service_s.lower():
                        direct = cls._extract_cost_count(item)
                        if direct:
                            return direct
                found = cls._search_offer(item, country_s, service_s)
                if found:
                    return found
        return None

    @staticmethod
    def _provider_id_from(value: Any) -> Optional[str]:
        if value is None:
            return None
        candidate = str(value).strip()
        # SMSBower provider IDs are numeric. Restricting the value also keeps
        # Telegram callback_data compact and safe.
        return candidate if candidate.isdigit() else None

    @classmethod
    def _collect_provider_offers(
        cls,
        node: Any,
        output: dict[str, ProviderOffer],
        parent_key: Optional[str] = None,
    ) -> None:
        if isinstance(node, list):
            for item in node:
                cls._collect_provider_offers(item, output, parent_key)
            return
        if not isinstance(node, dict):
            return

        price_value = None
        for key in ("price", "cost", "amount", "rate"):
            if key in node:
                price_value = node[key]
                break
        if price_value is not None:
            provider_id = cls._provider_id_from(
                node.get("provider_id", node.get("providerId", node.get("provider", parent_key)))
            )
            if provider_id:
                try:
                    price = decimal_value(price_value) * PROVIDER_PRICE_MULTIPLIER
                except ValueError:
                    price = None
                if price is not None and price >= 0:
                    count = -1
                    for key in ("count", "available", "stock", "quantity", "qty"):
                        if key in node:
                            with suppress(Exception):
                                count = int(decimal_value(node[key]))
                            break
                    offer = ProviderOffer(
                        provider_id=provider_id,
                        provider_cost_units=usd_to_units(price),
                        available=max(-1, count),
                        raw=node,
                    )
                    current = output.get(provider_id)
                    if current is None:
                        output[provider_id] = offer
                    else:
                        # API responses can contain the same provider in more
                        # than one nested wrapper. Keep the best observed price
                        # and the greatest explicitly reported stock instead
                        # of accidentally discarding stock from another copy.
                        best_price_offer = (
                            offer
                            if offer.provider_cost_units < current.provider_cost_units
                            else current
                        )
                        known_counts = [
                            count for count in (current.available, offer.available)
                            if count >= 0
                        ]
                        merged_count = max(known_counts) if known_counts else -1
                        output[provider_id] = ProviderOffer(
                            provider_id=provider_id,
                            provider_cost_units=best_price_offer.provider_cost_units,
                            available=merged_count,
                            raw=best_price_offer.raw,
                        )

        for key, value in node.items():
            cls._collect_provider_offers(value, output, str(key))

    async def get_provider_offers(
        self,
        country_id: str,
        service_code: str,
        extra_params: Optional[dict[str, str]] = None,
    ) -> list[ProviderOffer]:
        # SMSBower documents only country and service for getPricesV3. Ignore
        # legacy server/hash/provider parameters saved by older bot versions.
        data, text = await self.request(
            "getPricesV3",
            country=country_id,
            service=service_code,
        )
        if isinstance(data, str):
            raise RuntimeError(self.human_api_error(text))
        offers: dict[str, ProviderOffer] = {}
        self._collect_provider_offers(data, offers)
        result = sorted(
            offers.values(),
            key=lambda offer: (offer.provider_cost_units, -offer.available, int(offer.provider_id)),
        )
        if not result:
            raise RuntimeError(f"No provider price/stock found in getPricesV3 response: {short_text(text, 250)}")
        return result

    async def get_service_country_stock(self, service_code: str) -> dict[str, int]:
        """Return getPrices stock keyed by numeric country ID.

        SMSBower's documented getPrices response is country-name -> service ->
        {cost, count}; older compatible responses may instead use numeric IDs.
        Country labels are normalized against the corrected 202-entry SMSBower
        country catalog embedded in this file. Prefer live getCountries data when
        callers supply it to normalize_live_stock_country_ids().
        """
        data, text = await self.request("getPrices", service=service_code)
        if isinstance(data, str):
            raise RuntimeError(self.human_api_error(text))

        stock_by_country: dict[str, int] = {}
        price_by_country: dict[str, int] = {}
        service_key = str(service_code).strip().casefold()
        wrapper_keys = {
            "data",
            "result",
            "prices",
            "offers",
            "response",
            "status",
            "success",
            "message",
        }

        def country_key(raw_key: Any) -> Optional[str]:
            raw = str(raw_key or "").strip()
            if raw.isdigit():
                return raw
            normalized = _country_name_key(raw)
            mapped = _SMSBOWER_COUNTRY_ID_BY_NAME.get(normalized)
            if mapped:
                return mapped
            return None if normalized in wrapper_keys else raw or None

        def stock_count(node: Any) -> Optional[int]:
            if isinstance(node, dict):
                for key in ("count", "available", "stock", "quantity", "qty"):
                    if key in node and node[key] is not None:
                        try:
                            return max(0, int(decimal_value(node[key])))
                        except (ValueError, TypeError, InvalidOperation):
                            continue
                nested_counts = [
                    count
                    for value in node.values()
                    if (count := stock_count(value)) is not None
                ]
                return sum(nested_counts) if nested_counts else None
            if isinstance(node, list):
                nested_counts = [
                    count
                    for value in node
                    if (count := stock_count(value)) is not None
                ]
                return sum(nested_counts) if nested_counts else None
            return None

        def lowest_price(node: Any) -> Optional[int]:
            """Extract the lowest cost/price in a country-service offer node."""
            values: list[int] = []
            if isinstance(node, dict):
                for key, value in node.items():
                    if str(key).strip().casefold() in {"price", "cost", "amount", "rate"}:
                        try:
                            parsed = decimal_value(value)
                            if parsed >= 0:
                                values.append(usd_to_units(parsed * PROVIDER_PRICE_MULTIPLIER))
                        except (ValueError, TypeError, InvalidOperation):
                            pass
                    elif isinstance(value, (dict, list)):
                        nested = lowest_price(value)
                        if nested is not None:
                            values.append(nested)
            elif isinstance(node, list):
                for value in node:
                    nested = lowest_price(value)
                    if nested is not None:
                        values.append(nested)
            return min(values) if values else None

        def add_price(raw_country: Any, node: Any) -> None:
            resolved_id = country_key(raw_country)
            price = lowest_price(node)
            if resolved_id is not None and price is not None:
                price_by_country[resolved_id] = min(
                    price_by_country.get(resolved_id, price), price
                )

        def add_stock(raw_country: Any, count: Optional[int]) -> None:
            resolved_id = country_key(raw_country)
            if resolved_id is None or count is None:
                return
            stock_by_country[resolved_id] = max(
                stock_by_country.get(resolved_id, 0), max(0, int(count))
            )

        def visit(node: Any, current_country_id: Optional[str] = None) -> None:
            if isinstance(node, list):
                for value in node:
                    visit(value, current_country_id)
                return
            if not isinstance(node, dict):
                return

            explicit_country = (
                node.get("country_id")
                or node.get("countryId")
                or node.get("country")
                or node.get("country_name")
                or node.get("countryName")
            )
            explicit_service = (
                node.get("service")
                or node.get("service_code")
                or node.get("serviceCode")
            )
            if (
                explicit_country is not None
                and str(explicit_service or "").strip().casefold() == service_key
            ):
                add_stock(explicit_country, stock_count(node))
                add_price(explicit_country, node)

            for key, value in node.items():
                key_text = str(key).strip()
                if key_text.casefold() == service_key and current_country_id:
                    count = stock_count(value)
                    if count is not None:
                        add_stock(current_country_id, count)
                    add_price(current_country_id, value)
                    if count is None:
                        visit(value, current_country_id)
                    continue

                nested_country_id = country_key(key_text)
                visit(value, nested_country_id or current_country_id)

        visit(data)
        if not stock_by_country:
            # Do not silently interpret an unrecognized response as zero
            # stock. This distinction prevents every country from vanishing
            # when the upstream API changes its JSON shape or returns an error.
            raise RuntimeError(
                "SMSBower getPrices response contained no recognizable country/service stock: "
                + short_text(text, 350)
            )
        self.country_price_cache[service_key] = price_by_country
        return stock_by_country

    async def get_countries(self) -> list[tuple[str, str]]:
        """Return the current SMSBower country ID/name list."""
        data, text = await self.request("getCountries")
        if isinstance(data, str):
            raise RuntimeError(self.human_api_error(text))

        countries: dict[str, str] = {}

        def add_country(country_id: Any, name: Any) -> None:
            country_id = str(country_id or "").strip()
            name = str(name or "").strip()
            if country_id.isdigit() and name:
                countries[country_id] = name

        def visit(node: Any) -> None:
            if isinstance(node, list):
                for value in node:
                    visit(value)
                return
            if not isinstance(node, dict):
                return

            country_id = (
                node.get("id")
                or node.get("country_id")
                or node.get("countryId")
            )
            country_name = (
                node.get("eng")
                or node.get("name")
                or node.get("country_name")
                or node.get("countryName")
            )
            add_country(country_id, country_name)

            for key, value in node.items():
                if str(key).strip().isdigit() and isinstance(value, dict):
                    add_country(
                        value.get("id") or key,
                        value.get("eng")
                        or value.get("name")
                        or value.get("country_name")
                        or value.get("countryName"),
                    )
                if isinstance(value, (dict, list)):
                    visit(value)

        visit(data)
        if not countries:
            raise RuntimeError("SMSBower country list could not be parsed.")
        return sorted(countries.items(), key=lambda item: item[1].casefold())

    async def get_offer(
        self,
        country_id: str,
        service_code: str,
        extra_params: dict[str, str],
    ) -> LiveOffer:
        offers = await self.get_provider_offers(country_id, service_code, extra_params)
        # Unknown stock (-1) is not the same as available stock. Only mark
        # an offer as purchasable when the API explicitly reports count > 0.
        available_offers = [offer for offer in offers if offer.available > 0]
        selected = available_offers[0] if available_offers else offers[0]
        return LiveOffer(selected.provider_cost_units, selected.available, selected.raw)

    async def get_number(
        self,
        country_id: str,
        service_code: str,
        max_price_units: int,
        extra_params: dict[str, str],
        provider_id: Optional[str] = None,
    ) -> PurchasedNumber:
        max_price = units_to_decimal(max_price_units)
        allowed_optional = {
            "exceptProviderIds", "phoneException", "ref", "userID", "minPrice"
        }
        params: dict[str, Any] = {
            key: value for key, value in extra_params.items() if key in allowed_optional
        }
        if provider_id:
            params["providerIds"] = provider_id
        params.update(
            country=country_id,
            service=service_code,
            maxPrice=f"{max_price:.6f}",
        )

        data, text = await self.request("getNumberV2", **params)
        # Some compatible providers disable V2. Fall back to the standard method
        # while preserving providerIds and maxPrice.
        if isinstance(data, str) and "BAD_ACTION" in text.upper():
            data, text = await self.request("getNumber", **params)

        if isinstance(data, dict):
            if str(data.get("status", "")).lower() in {"error", "failed", "fail"}:
                raise RuntimeError(str(data.get("message") or data.get("error") or data))

            source = data.get("data") if isinstance(data.get("data"), dict) else data
            activation_id = (
                source.get("activationId")
                or source.get("activation_id")
                or source.get("order_id")
                or source.get("id")
            )
            phone = (
                source.get("phoneNumber")
                or source.get("number")
                or source.get("phone_number")
                or source.get("phone")
            )
            if activation_id and phone:
                cost_units = None
                for key in ("activationCost", "price", "cost", "amount"):
                    if key in source:
                        with suppress(Exception):
                            cost_units = usd_to_units(decimal_value(source[key]) * PROVIDER_PRICE_MULTIPLIER)
                        break
                expires_in = None
                for key in ("expires_in", "expires", "ttl"):
                    if key in source:
                        with suppress(Exception):
                            expires_in = int(source[key])
                        break
                return PurchasedNumber(str(activation_id), str(phone), cost_units, expires_in, data)

        match = re.search(r"ACCESS_NUMBER:([^:]+):(.+)$", text, re.I)
        if match:
            return PurchasedNumber(match.group(1).strip(), match.group(2).strip(), None, None, text)

        raise RuntimeError(self.human_api_error(text))

    async def get_status(self, activation_id: str, extra_params: dict[str, str]) -> ActivationStatus:
        # getStatus accepts only the activation ID; ignore legacy catalog extras.
        data, text = await self.request("getStatus", id=activation_id)
        upper = text.upper()

        if isinstance(data, dict):
            source = data.get("data") if isinstance(data.get("data"), dict) else data
            code = source.get("code") or source.get("sms_code") or source.get("otp")
            state = str(source.get("status") or source.get("state") or "").lower()
            if code:
                return ActivationStatus("code", str(code), text, data)
            if any(word in state for word in ("wait", "pending", "active")):
                return ActivationStatus("waiting", None, text, data)
            if "cancel" in state:
                return ActivationStatus("cancelled", None, text, data)
            if any(word in state for word in ("expire", "timeout")):
                return ActivationStatus("expired", None, text, data)

        if upper.startswith("STATUS_OK:"):
            return ActivationStatus("code", text.split(":", 1)[1].strip(), text, data)
        if upper in {
            "STATUS_WAIT_CODE",
            "STATUS_WAIT_RETRY",
            "STATUS_WAIT_RESEND",
            "STATUS_WAITING",
        } or upper.startswith(("STATUS_WAIT_RETRY:", "STATUS_WAIT_RESEND:")):
            # SMSBower can return STATUS_WAIT_RETRY:LastCode when the
            # previous SMS is known and the activation is waiting for another.
            # Treat the whole documented response as waiting, not unknown.
            return ActivationStatus("waiting", None, text, data)
        if "CANCEL" in upper:
            return ActivationStatus("cancelled", None, text, data)
        if "EXPIRE" in upper or "TIMEOUT" in upper:
            return ActivationStatus("expired", None, text, data)
        if upper.startswith(("BAD_", "NO_", "ERROR", "WRONG_")):
            return ActivationStatus("error", None, text, data)
        return ActivationStatus("unknown", None, text, data)

    async def set_status(
        self,
        activation_id: str,
        status: int,
        extra_params: dict[str, str],
    ) -> tuple[Any, str]:
        # setStatus accepts only ID and status; ignore legacy catalog extras.
        return await self.request("setStatus", id=activation_id, status=status)

    @staticmethod
    def cancellation_succeeded(data: Any, text: str) -> bool:
        upper = text.upper()
        if any(token in upper for token in ("ACCESS_CANCEL", "CANCEL_SUCCESS", "CANCELLED", "CANCELED")):
            return True
        if isinstance(data, dict):
            status = str(data.get("status", "")).lower()
            message = str(data.get("message", "")).lower()
            result = str(data.get("result", "")).lower()
            return status in {"success", "ok", "cancelled", "canceled"} or "cancel" in message or "cancel" in result
        return False

    @staticmethod
    def human_api_error(text: str) -> str:
        upper = text.upper().strip()
        mapping = {
            "NO_NUMBERS": "No numbers are available right now.",
            "NO_BALANCE": "The provider account has insufficient balance.",
            "BAD_KEY": "The provider API key is invalid.",
            "NO_KEY": "The provider API key is missing.",
            "BAD_SERVICE": "The country/service configuration is invalid.",
            "WRONG_MAX_PRICE": "The live price changed. Please try again.",
            "CHANNELS_LIMIT": "The provider has too many pending activations.",
            "EARLY_CANCEL_DENIED": "The provider does not allow cancellation yet.",
        }
        for token, message in mapping.items():
            if token in upper:
                return message
        return f"Provider error: {short_text(text, 250)}"


API: Optional[SmsPowerAPI] = None
PURCHASE_LOCKS: dict[int, asyncio.Lock] = {}
POLL_TASK: Optional[asyncio.Task[Any]] = None
COUNTDOWN_MINUTE_CACHE: dict[int, int] = {}


# =============================================================================
# Rendering and access checks
# =============================================================================


async def ensure_registered(update: Update) -> Optional[sqlite3.Row]:
    user = update.effective_user
    if not user:
        return None
    return DB.register_user(user)


def referral_start_id(context: ContextTypes.DEFAULT_TYPE) -> Optional[int]:
    args = list(getattr(context, "args", []) or [])
    if not args:
        return None
    raw = str(args[0]).strip()
    if raw.startswith("ref_"):
        raw = raw[4:]
    return int(raw) if raw.isdigit() else None


async def require_terms(update: Update) -> bool:
    row = await ensure_registered(update)
    if not row:
        return False
    if row["is_blocked"]:
        if update.effective_message:
            await update.effective_message.reply_text("Your account is blocked.")
        return False
    if row["accepted_terms"]:
        return True

    markup = InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ I agree", callback_data="terms:accept")]]
    )
    text = (
        "<b>Usage agreement</b>\n\n"
        "Use this bot only for lawful purposes, for accounts and services you are authorized to access, "
        "and only where virtual/temporary numbers are permitted. Fraud, spam, account theft, and bypassing "
        "a platform's rules are prohibited.\n\nTap <b>I agree</b> to continue."
    )
    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
    elif update.effective_message:
        await update.effective_message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
    return False


async def require_admin(update: Update) -> bool:
    if is_admin(update.effective_user.id if update.effective_user else None):
        return True
    if update.callback_query:
        await update.callback_query.answer("Admin only.", show_alert=True)
    elif update.effective_message:
        await update.effective_message.reply_text("Admin only.")
    return False


def get_catalog_item(item_id: int) -> Optional[sqlite3.Row]:
    return DB.fetchone("SELECT * FROM catalog WHERE id=?", (item_id,))


def catalog_extra(item: sqlite3.Row) -> dict[str, str]:
    data = safe_json_loads(item["extra_params_json"], {})
    return {str(k): str(v) for k, v in data.items()} if isinstance(data, dict) else {}


def item_markup_units(item: sqlite3.Row) -> int:
    override = item["markup_override_units"]
    return int(override) if override is not None else DB.global_markup_units()


def sync_live_service_countries(
    service_ref: int,
    stock_by_country: dict[str, int],
    country_names: dict[str, str],
) -> int:
    """Add newly supported, currently stocked provider countries to this service."""
    template = DB.fetchone(
        "SELECT * FROM catalog WHERE id=? AND enabled=1", (service_ref,)
    )
    if not template:
        return 0

    candidates = [
        (str(country_id), str(country_names.get(str(country_id), "")).strip())
        for country_id, stock in stock_by_country.items()
        if int(stock) > 0 and str(country_names.get(str(country_id), "")).strip()
    ]
    if not candidates:
        return 0

    inserted = 0
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            for country_id, country_name in candidates:
                existing = DB.conn.execute(
                    "SELECT 1 FROM catalog WHERE country_id=? AND service_code=? LIMIT 1",
                    (country_id, template["service_code"]),
                ).fetchone()
                if existing:
                    # Existing catalog rows may have been created with an older,
                    # mismatched static country-ID table. Reconcile their display
                    # names against the live SMSBower country catalog whenever
                    # that country has stock, rather than keeping stale labels.
                    DB.conn.execute(
                        "UPDATE catalog SET country_name=? WHERE country_id=? AND service_code=?",
                        (country_name, country_id, template["service_code"]),
                    )
                    continue
                DB.conn.execute(
                    """
                    INSERT INTO catalog(
                        country_id,country_name,service_code,service_name,
                        extra_params_json,markup_override_units,enabled,created_at
                    ) VALUES(?,?,?,?,?,?,1,?)
                    """,
                    (
                        country_id,
                        country_name,
                        template["service_code"],
                        template["service_name"],
                        template["extra_params_json"] or "{}",
                        template["markup_override_units"],
                        now_ts(),
                    ),
                )
                inserted += 1
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            raise
    return inserted


async def show_id(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await ensure_registered(update)
    if update.effective_message and update.effective_user:
        await update.effective_message.reply_text(
            f"Your Telegram ID: <code>{update.effective_user.id}</code>",
            parse_mode=ParseMode.HTML,
        )


async def show_buy_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Show enabled services first. The user selects a country afterwards."""
    if not await require_terms(update):
        return
    query = update.callback_query
    message = update.effective_message or (query.message if query else None)
    if not message:
        return
    if query:
        await query.answer()

    services = DB.fetchall(
        """
        SELECT MIN(id) AS service_ref,
               MIN(service_code) AS service_code,
               MAX(service_name) AS service_name
        FROM catalog
        WHERE enabled=1
        GROUP BY service_code COLLATE NOCASE
        ORDER BY MAX(service_name) COLLATE NOCASE
        """
    )
    if not services:
        value = "No service available."
        if query:
            await replace_callback_message(query, context.bot, value)
        else:
            await message.reply_text(value)
        return

    buttons: list[list[InlineKeyboardButton]] = []
    for row in services:
        service_name = str(row["service_name"])
        service_emoji_id = _service_premium_emoji_id(service_name)
        fallback_emoji = _service_fallback_emoji(service_name)
        label = short_text(service_name, 34)
        service_button_kwargs = _inline_custom_emoji_kwargs(service_emoji_id) if bool(getattr(update.effective_user, "is_premium", False)) else {}
        service_button = InlineKeyboardButton(
            label if service_button_kwargs else f"{fallback_emoji} {label}",
            callback_data=f"buy:svc:{row['service_ref']}",
            **service_button_kwargs,
        )
        buttons.append([service_button])

    # Always keep the user's still-valid Buy Number orders reachable.
    buttons.append([InlineKeyboardButton("📋 Active Number List", callback_data="order:active:list")])
    value = "<b>Select service</b>"
    markup = InlineKeyboardMarkup(buttons)
    if query:
        await replace_callback_message(query, context.bot, 
            value,
            parse_mode=ParseMode.HTML,
            reply_markup=markup,
        )
    else:
        await message.reply_text(
            value,
            parse_mode=ParseMode.HTML,
            reply_markup=markup,
        )


async def show_service_countries(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_terms(update):
        return
    await query.answer("Checking live stock…")

    parts = str(query.data).split(":")
    try:
        if len(parts) == 4 and parts[1] == "svcpage":
            service_ref = int(parts[2])
            page = int(parts[3])
        else:
            service_ref = int(parts[-1])
            page = 0
    except Exception:
        await query.edit_message_text("Invalid service.")
        return

    selected = get_catalog_item(service_ref)
    if not selected or not selected["enabled"]:
        await query.edit_message_text(
            "Service unavailable.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")]]
            ),
        )
        return

    if API is None:
        await query.edit_message_text(
            "Live stock check is not ready right now.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")]]
            ),
        )
        return

    stock_result, countries_result = await asyncio.gather(
        API.get_service_country_stock(str(selected["service_code"])),
        API.get_countries(),
        return_exceptions=True,
    )
    if isinstance(stock_result, Exception):
        logger.warning(
            "Live country stock check failed for service %s: %s",
            selected["service_code"],
            short_text(str(stock_result), 180),
        )
        await query.edit_message_text(
            "Could not check live stock right now.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")]]
            ),
        )
        return

    stock_by_country = {
        str(country_id): max(0, int(stock))
        for country_id, stock in stock_result.items()
    }
    country_names = (
        {str(country_id): str(name) for country_id, name in countries_result}
        if not isinstance(countries_result, Exception)
        else {}
    )
    stock_by_country = normalize_live_stock_country_ids(
        stock_by_country, country_names
    )
    if isinstance(countries_result, Exception):
        logger.warning(
            "SMSBower country-name refresh failed; using saved names: %s",
            short_text(str(countries_result), 180),
        )
    try:
        sync_live_service_countries(service_ref, stock_by_country, country_names)
    except Exception:
        logger.exception("Could not add newly stocked country rows for service %s", selected["service_code"])

    configured_rows = DB.fetchall(
        """
        SELECT c.* FROM catalog AS c
        INNER JOIN (
            SELECT MIN(id) AS id
            FROM catalog
            WHERE enabled=1 AND service_code=?
            GROUP BY country_id
        ) AS country_rows ON country_rows.id=c.id
        ORDER BY c.country_name COLLATE NOCASE
        """,
        (selected["service_code"],),
    )
    available_rows = [
        row
        for row in configured_rows
        if stock_by_country.get(str(row["country_id"]), 0) > 0
    ]
    favorite_rows = DB.fetchall(
        "SELECT country_id FROM country_favorites WHERE user_id=? AND service_code=?",
        (int(query.from_user.id), str(selected["service_code"])),
    )
    favorite_country_ids = {str(row["country_id"]) for row in favorite_rows}
    live_country_prices = getattr(API, "country_price_cache", {}).get(
        str(selected["service_code"]).strip().casefold(), {}
    )
    # Favorites always stay at the top; within each group sort by cheapest live rate.
    available_rows.sort(
        key=lambda row: (
            0 if str(row["country_id"]) in favorite_country_ids else 1,
            live_country_prices.get(str(row["country_id"]), 10**18),
            str(row["country_name"]).casefold(),
        )
    )
    total = len(available_rows)
    if not total:
        await query.edit_message_text(
            "No countries have numbers in stock right now.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")]]
            ),
        )
        return

    pages = max(1, math.ceil(total / USER_COUNTRIES_PAGE_SIZE))
    page = max(0, min(page, pages - 1))
    start = page * USER_COUNTRIES_PAGE_SIZE
    rows = available_rows[start : start + USER_COUNTRIES_PAGE_SIZE]

    country_buttons: list[InlineKeyboardButton] = []
    fallback_country_buttons: list[InlineKeyboardButton] = []
    has_premium_icons = False
    for row in rows:
        button, fallback_button, uses_premium = _country_button_pair(
            str(row["country_name"]),
            f"buy:service:{row['id']}",
            country_id=str(row["country_id"]),
            text_limit=18,
        )
        country_buttons.append(button)
        fallback_country_buttons.append(fallback_button)
        has_premium_icons = has_premium_icons or uses_premium

    buttons: list[list[InlineKeyboardButton]] = []
    fallback_buttons: list[list[InlineKeyboardButton]] = []
    for index in range(0, len(country_buttons), 2):
        country_row = country_buttons[index : index + 2]
        fallback_country_row = fallback_country_buttons[index : index + 2]
        if len(country_row) == 1:
            spacer = InlineKeyboardButton(
                "\u2800" * 20, callback_data="noop", style="secondary"
            )
            country_row.append(spacer)
            fallback_country_row.append(spacer)
        buttons.append(country_row)
        fallback_buttons.append(fallback_country_row)

    previous_button = InlineKeyboardButton(
        "⬅️",
        callback_data=f"buy:svcpage:{service_ref}:{page - 1}" if page > 0 else "noop",
    )
    page_button = InlineKeyboardButton(
        f"📄 {page + 1}/{pages}", callback_data="noop"
    )
    next_button = InlineKeyboardButton(
        "➡️",
        callback_data=(
            f"buy:svcpage:{service_ref}:{page + 1}"
            if page < pages - 1
            else "noop"
        ),
    )
    page_nav = [previous_button, page_button, next_button]
    buttons.append(page_nav)
    fallback_buttons.append(page_nav)
    country_search_button = InlineKeyboardButton(
        "🔎 Search country",
        callback_data=f"buy:countrysearch:{service_ref}:{page}",
        style="danger",
    )
    buttons.append([country_search_button])
    fallback_buttons.append([country_search_button])
    back_button = InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")
    buttons.append([back_button])
    fallback_buttons.append([back_button])

    text = (
        f"🌍 <b>Select a Country for "
        f"{html.escape(str(selected['service_name']).upper())}</b>\n"
        f"{' ' * 4}"
    )
    await edit_callback_message_with_country_fallback(
        query,
        text,
        InlineKeyboardMarkup(buttons),
        InlineKeyboardMarkup(fallback_buttons),
        parse_mode=ParseMode.HTML,
        has_premium_icons=has_premium_icons,
    )


async def country_search_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    query = update.callback_query
    if not query or not await require_terms(update):
        return ConversationHandler.END
    try:
        parts = str(query.data).split(":")
        service_ref = int(parts[2])
        page = max(0, int(parts[3]))
    except (IndexError, ValueError):
        await query.answer("Invalid search request.", show_alert=True)
        return ConversationHandler.END

    service = get_catalog_item(service_ref)
    if not service or not service["enabled"]:
        await query.answer("This service is no longer available.", show_alert=True)
        return ConversationHandler.END

    await query.answer()
    context.user_data["country_search_service_ref"] = service_ref
    context.user_data["country_search_page"] = page
    try:
        await query.edit_message_text(
            "🔎 Send the country name to search (for example, Canada).",
            reply_markup=InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            "❌ Cancel",
                            callback_data="workflow:cancel",
                            style="danger",
                        )
                    ]
                ]
            ),
        )
        if query.message:
            context.user_data["country_search_prompt_id"] = query.message.message_id
    except TelegramError:
        if query.message:
            prompt = await query.message.reply_text(
                "🔎 Send the country name to search (for example, Canada).",
                reply_markup=InlineKeyboardMarkup(
                    [[InlineKeyboardButton("❌ Cancel", callback_data="workflow:cancel", style="danger")]]
                ),
            )
            context.user_data["country_search_prompt_id"] = prompt.message_id
    return BUY_COUNTRY_SEARCH


async def country_search_receive(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    message = update.effective_message
    user = update.effective_user
    if not message or not user or not await require_terms(update):
        return ConversationHandler.END

    search_text = " ".join(str(message.text or "").split()).casefold()
    if not search_text:
        await message.reply_text("Please enter a country name.")
        return BUY_COUNTRY_SEARCH

    service_ref = context.user_data.get("country_search_service_ref")
    page = int(context.user_data.get("country_search_page", 0))
    prompt_id = context.user_data.get("country_search_prompt_id")
    try:
        service_ref = int(service_ref)
    except (TypeError, ValueError):
        await message.reply_text("Search expired. Open the service and try again.")
        return ConversationHandler.END

    service = get_catalog_item(service_ref)
    if not service or not service["enabled"] or API is None:
        result_text = "This service is no longer available for live search."
        result_markup = InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")]]
        )
    else:
        stock_result, countries_result = await asyncio.gather(
            API.get_service_country_stock(str(service["service_code"])),
            API.get_countries(),
            return_exceptions=True,
        )
        if isinstance(stock_result, Exception):
            logger.warning(
                "Country search stock check failed for %s: %s",
                service["service_code"],
                short_text(str(stock_result), 180),
            )
            result_text = "Could not check live stock. Try searching again."
            result_markup = InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            "🔎 Search again",
                            callback_data=f"buy:countrysearch:{service_ref}:{page}",
                            style="danger",
                        )
                    ],
                    [
                        InlineKeyboardButton(
                            "⬅️ Back to countries",
                            callback_data=f"buy:svcpage:{service_ref}:{page}",
                        )
                    ],
                ]
            )
        else:
            stock_by_country = {
                str(country_id): max(0, int(stock))
                for country_id, stock in stock_result.items()
            }
            if isinstance(countries_result, Exception):
                logger.warning(
                    "Country names unavailable during search for %s; using catalog names.",
                    service["service_code"],
                )
                country_names = {
                    str(row["country_id"]): str(row["country_name"])
                    for row in DB.fetchall(
                        "SELECT country_id,country_name FROM catalog WHERE service_code=?",
                        (str(service["service_code"]),),
                    )
                }
            else:
                country_names = {
                    str(country_id): str(name)
                    for country_id, name in countries_result
                }
            stock_by_country = normalize_live_stock_country_ids(
                stock_by_country, country_names
            )
            try:
                sync_live_service_countries(
                    service_ref, stock_by_country, country_names
                )
            except Exception:
                logger.exception(
                    "Could not sync countries during search for service %s",
                    service["service_code"],
                )

            configured_rows = DB.fetchall(
                """
                SELECT c.* FROM catalog AS c
                INNER JOIN (
                    SELECT MIN(id) AS id
                    FROM catalog
                    WHERE enabled=1 AND service_code=?
                    GROUP BY country_id
                ) AS country_rows ON country_rows.id=c.id
                ORDER BY c.country_name COLLATE NOCASE
                """,
                (str(service["service_code"]),),
            )
            available = [
                row
                for row in configured_rows
                if stock_by_country.get(str(row["country_id"]), 0) > 0
                and search_text
                in " ".join(str(row["country_name"]).split()).casefold()
            ]
            favorite_rows = DB.fetchall(
                "SELECT country_id FROM country_favorites WHERE user_id=? AND service_code=?",
                (int(user.id), str(service["service_code"])),
            )
            favorite_ids = {str(row["country_id"]) for row in favorite_rows}
            live_country_prices = getattr(API, "country_price_cache", {}).get(
                str(service["service_code"]).strip().casefold(), {}
            )
            available.sort(
                key=lambda row: (
                    0 if str(row["country_id"]) in favorite_ids else 1,
                    live_country_prices.get(str(row["country_id"]), 10**18),
                    0
                    if " ".join(str(row["country_name"]).split()).casefold()
                    == search_text
                    else 1,
                    str(row["country_name"]).casefold(),
                )
            )

            if not available:
                result_text = (
                    f"No in-stock country matched “{html.escape(search_text)}”."
                )
                result_markup = InlineKeyboardMarkup(
                    [
                        [
                            InlineKeyboardButton(
                                "🔎 Search again",
                                callback_data=f"buy:countrysearch:{service_ref}:{page}",
                                style="danger",
                            )
                        ],
                        [
                            InlineKeyboardButton(
                                "⬅️ Back to countries",
                                callback_data=f"buy:svcpage:{service_ref}:{page}",
                            )
                        ],
                    ]
                )
            else:
                buttons: list[list[InlineKeyboardButton]] = []
                for row in available[:10]:
                    country_id = str(row["country_id"])
                    country_name = str(row["country_name"])
                    stock = stock_by_country.get(country_id, 0)
                    favorite_prefix = "⭐ " if country_id in favorite_ids else ""
                    buttons.append(
                        [
                            InlineKeyboardButton(
                                (
                                    f"{favorite_prefix}"
                                    f"{short_text(country_name, 31)} · {stock} in stock"
                                    if bool(getattr(query.from_user, "is_premium", False)) and _country_premium_emoji_id(country_name, country_id)
                                    else f"{favorite_prefix}{_country_flag_emoji(country_name, country_id)} {short_text(country_name, 31)} · {stock} in stock"
                                ),
                                callback_data=f"buy:service:{row['id']}",
                                **(_inline_custom_emoji_kwargs(_country_premium_emoji_id(country_name, country_id)) if bool(getattr(query.from_user, "is_premium", False)) else {}),
                            )
                        ]
                    )
                buttons.append(
                    [
                        InlineKeyboardButton(
                            "🔎 Search another country",
                            callback_data=f"buy:countrysearch:{service_ref}:{page}",
                            style="danger",
                        )
                    ]
                )
                buttons.append(
                    [
                        InlineKeyboardButton(
                            "⬅️ Back to countries",
                            callback_data=f"buy:svcpage:{service_ref}:{page}",
                        )
                    ]
                )
                result_text = (
                    f"<b>Available matches for “{html.escape(search_text)}”</b>\n"
                    "Tap a country to view its number offers."
                )
                result_markup = InlineKeyboardMarkup(buttons)

    for key in (
        "country_search_service_ref",
        "country_search_page",
        "country_search_prompt_id",
    ):
        context.user_data.pop(key, None)
    try:
        if prompt_id:
            await context.bot.edit_message_text(
                chat_id=message.chat_id,
                message_id=int(prompt_id),
                text=result_text,
                reply_markup=result_markup,
                parse_mode=ParseMode.HTML,
            )
        else:
            await message.reply_text(
                result_text,
                reply_markup=result_markup,
                parse_mode=ParseMode.HTML,
            )
    except TelegramError:
        await message.reply_text(
            result_text,
            reply_markup=result_markup,
            parse_mode=ParseMode.HTML,
        )
    return ConversationHandler.END


async def deposit_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_terms(update):
        return ConversationHandler.END

    minimum = minimum_deposit_units()
    minimum_text = (
        f"\nMinimum deposit: <b>{money(minimum)}</b>\n" if minimum > 0 else ""
    )
    await update.effective_message.reply_text(
        "<b>💳 Add Money</b>\n\n"
        "Enter the amount you want to deposit.\n\n"
        f"{minimum_text}"
        "Example:\n"
        "<code>100</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(
            [[
                InlineKeyboardButton(
                    "❌ Cancel",
                    callback_data="workflow:cancel",
                    style="danger",
                )
            ]]
        ),
    )

    return DEP_AMOUNT_NEW

async def deposit_amount(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    message = update.effective_message
    user = update.effective_user

    if not message or not user:
        return ConversationHandler.END

    txt = (message.text or "").strip()

    try:
        amount_decimal = bdt_input_value(txt)
        if amount_decimal <= 0:
            raise ValueError
    except Exception:
        await message.reply_text(
            "❌ Please enter a valid BDT amount.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton(
                    "❌ Cancel",
                    callback_data="workflow:cancel",
                    style="danger",
                )]]
            ),
        )
        return DEP_AMOUNT_NEW

    amount_units = bdt_to_units(amount_decimal)
    minimum = minimum_deposit_units()
    if minimum > 0 and amount_units < minimum:
        await message.reply_text(
            f"❌ Minimum deposit is <b>{money(minimum)}</b>.\n"
            f"Please enter at least <b>{money(minimum)}</b>.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton(
                    "❌ Cancel",
                    callback_data="workflow:cancel",
                    style="danger",
                )]]
            ),
        )
        return DEP_AMOUNT_NEW

    amount = float(amount_decimal)
    await message.reply_text("⏳ Creating payment invoice...")

    result = await create_ipshop_invoice(
        user.id,
        amount,
        user.full_name,
    )

    if not result.get("status"):
        await message.reply_text(
            f"❌ {result.get('message','Failed to create invoice.')}"
        )
        return ConversationHandler.END

    invoice_id = result["invoice_id"]
    payment_url = result["payment_url"]

    try:
        DB.execute(
            """
            INSERT INTO payment_invoices
                (invoice_id,user_id,chat_id,amount_units,status,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                str(invoice_id),
                int(user.id),
                int(message.chat_id),
                amount_units,
                "awaiting_payment",
                now_ts(),
                now_ts(),
            ),
        )
    except sqlite3.IntegrityError:
        logger.error("Payment provider returned a duplicate invoice ID: %s", invoice_id)
        await message.reply_text(
            "❌ Could not track this invoice safely. Please try again or contact support."
        )
        return ConversationHandler.END

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "💳 Pay Now",
                    url=payment_url,
                ),
            ]
        ]
    )

    await message.reply_text(
        f"✅ Invoice Created\n\n"
        f"Invoice: `{invoice_id}`\n"
        f"Amount: {amount:.2f} BDT",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=keyboard,
    )
    old_task = _IPSHOP_PAYMENT_TASKS.pop(str(invoice_id), None)
    if old_task and not old_task.done():
        old_task.cancel()
    _IPSHOP_PAYMENT_TASKS[str(invoice_id)] = asyncio.create_task(
        monitor_ipshop_payment(
            context.application,
            invoice_id=str(invoice_id),
            user_id=int(user.id),
            chat_id=int(message.chat_id),
            amount_units=amount_units,
        ),
        name=f"ipshop-payment-{invoice_id}",
    )

    return ConversationHandler.END


async def notify_successful_deposit_group(
    application: Application, *, user_id: int, amount_units: int, transaction_id: str
) -> None:
    """Send only successful automatic deposits to the configured deposit group."""
    group_raw = DB.get_setting("deposit_group_id", "").strip()
    if not re.fullmatch(r"-?\d+", group_raw):
        return
    user = DB.fetchone("SELECT first_name,username,balance_units FROM users WHERE user_id=?", (user_id,))
    if not user:
        return
    text = _deposit_notice_text(
        deposit_id=None,
        user_id=user_id,
        name=str(user["first_name"] or "Unknown"),
        username=f"@{user['username']}" if user["username"] else "none",
        method="Automatic Payment",
        account="—",
        amount=amount_units,
        transaction_id=str(transaction_id),
        balance=int(user["balance_units"] or 0),
        status="SUCCESS",
        status_icon="✅",
    )
    with suppress(TelegramError):
        await application.bot.send_message(chat_id=int(group_raw), text=text, parse_mode=ParseMode.HTML)


def _deposit_notice_text(
    *,
    deposit_id: Optional[int],
    user_id: int,
    name: str,
    username: str,
    method: str,
    account: str,
    amount: int,
    transaction_id: str,
    balance: Optional[int],
    status: str,
    status_icon: str,
) -> str:
    """Build deposit-group notices with the same quote layout as VPN orders."""
    deposit_label = f"#{deposit_id}" if deposit_id is not None else "Automatic"
    details = [
        f"🆔 Deposit: <code>{deposit_label}</code>",
        f"👤 Name: {html.escape(name)}",
        f"👤 User: {html.escape(username)}",
        f"🆔 Id : (<code>{user_id}</code>)",
        f"💳 Method: {html.escape(method)}",
        f"📱 Account: <code>{html.escape(account)}</code>",
        f"💵 Deposit Amount: <b>{money(amount)}</b>",
        f"🧾 Transaction ID: <code>{html.escape(transaction_id)}</code>",
    ]
    if balance is not None:
        details.append(f"💳 Current Balance: <b>{money(balance)}</b>")
    return "\n".join([
        f"💳 <b>Deposit — {html.escape(status)}</b>",
        "━━━━━━━━━━━━━━━━",
        "<blockquote>" + "\n".join(details) + "</blockquote>",
        "━━━━━━━━━━━━━━━━",
        f"<blockquote>📌 Status: {status_icon} <b>{html.escape(status)}</b></blockquote>",
    ])


async def monitor_ipshop_payment(
    application: Application,
    *,
    invoice_id: str,
    user_id: int,
    chat_id: int,
    amount_units: int,
) -> None:
    """Check an IPShop invoice without blocking the user's conversation."""
    task_key = str(invoice_id)
    try:
        for _ in range(60):
            await asyncio.sleep(5)
            check = await check_ipshop_invoice(task_key)
            if not check.get("status"):
                continue

            payment_status = str(check.get("payment_status") or "").strip().lower()
            if payment_status == "paid":
                dep = DB.complete_auto_deposit(task_key)
                if not dep:
                    return
                await notify_successful_deposit_group(
                    application, user_id=user_id, amount_units=amount_units,
                    transaction_id=str(check.get("trxid") or task_key)
                )
                referral_reward = DB.award_referral_for_deposit(int(dep["id"]))
                with suppress(TelegramError):
                    await application.bot.send_message(
                        chat_id,
                        f"✅ Payment Successful!\n\n"
                        f"Invoice: {task_key}\n"
                        f"TRX ID: {check.get('trxid', 'N/A')}\n"
                        f"Balance Added: {money(amount_units)}",
                    )
                if referral_reward:
                    with suppress(TelegramError):
                        await application.bot.send_message(
                            int(referral_reward["referrer_id"]),
                            f"🎉 Referral reward added: "
                            f"{money(referral_reward['amount_units'])}",
                        )
                return

            # Provider terminal states never create a deposit or change balance.
            if payment_status in {"failed", "expired", "cancelled", "canceled", "rejected"}:
                DB.execute(
                    """
                    UPDATE payment_invoices SET status=?,updated_at=?
                    WHERE invoice_id=? AND status='awaiting_payment'
                    """,
                    (payment_status, now_ts(), task_key),
                )
                with suppress(TelegramError):
                    await application.bot.send_message(
                        chat_id,
                        f"❌ Payment invoice {payment_status}. No balance was added.",
                    )
                return

        with suppress(TelegramError):
            await application.bot.send_message(
                chat_id,
                "⌛ Payment was not detected within 5 minutes.",
            )
        DB.execute(
            """
            UPDATE payment_invoices SET status='unpaid',updated_at=?
            WHERE invoice_id=? AND status='awaiting_payment'
            """,
            (now_ts(), task_key),
        )
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.exception("IPShop payment monitor failed for %s", task_key)
    finally:
        current = _IPSHOP_PAYMENT_TASKS.get(task_key)
        if current is asyncio.current_task():
            _IPSHOP_PAYMENT_TASKS.pop(task_key, None)


SUPPORT_MESSAGE = (
    "👋 𝙷𝚎𝚕𝚕𝚘! আপনার যেকোনো প্রয়োজনে সরাসরি আমাদের সাপোর্ট এডমিনের সাথে কথা বলুন।\n\n"
    "👤 𝙰𝚍𝚖𝚒𝚗: SAFAYET KHONDOKAR\n"
    "⚡ 𝚂𝚝𝚊𝚝𝚞𝚜: 𝙰𝚌𝚝𝚒𝚟𝚎 ✅\n\n"
    "👇 মেসেজ দিতে নিচের বাটনে ক্লিক করুন:"
)


def normalise_support_link(raw: str) -> str:
    value = str(raw or "").strip()
    if not value:
        raise ValueError("Support ID or link cannot be empty.")

    if re.fullmatch(r"\d{5,20}", value):
        return f"tg://user?id={value}"

    if value.startswith("@"): 
        username = value[1:]
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{4,31}", username):
            raise ValueError("Invalid Telegram username.")
        return f"https://t.me/{username}"

    if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{4,31}", value):
        return f"https://t.me/{value}"

    if value.startswith("http://t.me/"):
        value = "https://" + value[len("http://"):]

    if value.startswith("https://t.me/"):
        username = value.split("https://t.me/", 1)[1].split("?", 1)[0].strip("/")
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{4,31}", username):
            raise ValueError("Invalid t.me support link.")
        return f"https://t.me/{username}"

    if value.startswith("tg://user?id="):
        user_id = value.split("=", 1)[1]
        if not re.fullmatch(r"\d{5,20}", user_id):
            raise ValueError("Invalid Telegram numeric ID.")
        return value

    raise ValueError("Send a numeric Telegram ID, @username, username, or t.me link.")


async def show_support(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_terms(update):
        return
    link = DB.get_setting("support_link", "").strip()
    if not link:
        await update.effective_message.reply_text(
            SUPPORT_MESSAGE + "\n\n⚠️ Support link has not been configured yet."
        )
        return
    await update.effective_message.reply_text(
        SUPPORT_MESSAGE,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("💬 Send Message", url=link, style="primary")]
        ]),
    )



async def set_support_link_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    current = DB.get_setting("support_link", "").strip() or "Not set"
    await update.effective_message.reply_text(
        "<b>Set Support ID/Link</b>\n\n"
        f"Current: <code>{html.escape(current)}</code>\n\n"
        "Send one of these:\n"
        "• Numeric Telegram ID\n"
        "• @username\n"
        "• username\n"
        "• https://t.me/username\n\n"
        "Send <code>-</code> to remove it.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_SUPPORT_LINK


async def set_support_link_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    raw = (update.effective_message.text or "").strip()
    if raw == "-":
        DB.set_setting("support_link", "")
        await update.effective_message.reply_text(
            "✅ Support link removed.", reply_markup=admin_menu()
        )
        return ConversationHandler.END
    try:
        link = normalise_support_link(raw)
    except ValueError as exc:
        await update.effective_message.reply_text(
            f"❌ {html.escape(str(exc))}\nPlease try again or send /cancel.",
            parse_mode=ParseMode.HTML,
        )
        return SET_SUPPORT_LINK
    DB.set_setting("support_link", link)
    await update.effective_message.reply_text(
        f"✅ Support link saved:\n<code>{html.escape(link)}</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=admin_menu(),
    )
    return ConversationHandler.END


async def admin_stats(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    await query.answer()
    start_ts, end_ts, day_label = _dhaka_day_bounds()
    users = DB.fetchone("SELECT COUNT(*) AS n,COALESCE(SUM(balance_units),0) AS bal FROM users")
    today = DB.fetchone(
        """SELECT COUNT(*) AS bought,
                  SUM(CASE WHEN status='otp_received' THEN 1 ELSE 0 END) AS otp_count,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status IN ('otp_received','expired') THEN charged_units ELSE 0 END),0) AS revenue,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status IN ('otp_received','expired') THEN provider_cost_units ELSE 0 END),0) AS cost
           FROM orders WHERE created_at>=? AND created_at<?""", (start_ts, end_ts))
    all_orders = DB.fetchone(
        """SELECT COUNT(*) AS n,
                  SUM(CASE WHEN status='waiting' THEN 1 ELSE 0 END) AS waiting,
                  SUM(CASE WHEN status='otp_received' THEN 1 ELSE 0 END) AS success,
                  SUM(CASE WHEN status='cancelled' THEN 1 ELSE 0 END) AS cancelled,
                  SUM(CASE WHEN status='expired' THEN 1 ELSE 0 END) AS expired,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status IN ('otp_received','expired') THEN charged_units ELSE 0 END),0) AS revenue,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status IN ('otp_received','expired') THEN provider_cost_units ELSE 0 END),0) AS cost
           FROM orders""")
    local_today = DB.fetchone(
        """SELECT COUNT(*) AS bought,
                  COALESCE(SUM(CASE WHEN has_otp=1 THEN 1 ELSE 0 END),0) AS otp_count,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status<>'cancelled' THEN charged_units ELSE 0 END),0) AS revenue
           FROM local_orders WHERE created_at>=? AND created_at<?""", (start_ts, end_ts))
    today_revenue = int(today['revenue'] or 0) + int(local_today['revenue'] or 0)
    # Provider costs are stored in USD units; convert them using the bot's
    # current USD-to-BDT exchange rate before displaying or calculating profit.
    today_cost = usd_units_to_bdt_units(int(today['cost'] or 0))
    total_revenue = int(all_orders['revenue'] or 0)
    total_cost = usd_units_to_bdt_units(int(all_orders['cost'] or 0))
    text = (
        "<b>📊 Admin Statistics</b>\n\n"
        f"📅 <b>Today ({html.escape(day_label)})</b>\n"
        f"👥 Numbers bought: <b>{int(today['bought'] or 0) + int(local_today['bought'] or 0)}</b>\n"
        f"📩 OTPs received: <b>{int(today['otp_count'] or 0) + int(local_today['otp_count'] or 0)}</b>\n"
        f"💵 User charges: <b>{money(today_revenue)}</b>\n"
        f"🏷️ Provider cost (standard orders): <b>{money(today_cost)}</b>\n"
        f"📈 Estimated profit (standard orders): <b>{money(int(today['revenue'] or 0) - today_cost)}</b>\n\n"
        "<b>📦 All-time (standard orders)</b>\n"
        f"Total orders: <b>{int(all_orders['n'] or 0)}</b>\n"
        f"Waiting: <b>{int(all_orders['waiting'] or 0)}</b>\n"
        f"OTP received: <b>{int(all_orders['success'] or 0)}</b>\n"
        f"Cancelled: <b>{int(all_orders['cancelled'] or 0)}</b>\n"
        f"Expired: <b>{int(all_orders['expired'] or 0)}</b>\n"
        f"Total user charges: <b>{money(total_revenue)}</b>\n"
        f"Total provider cost: <b>{money(total_cost)}</b>\n"
        f"Estimated total profit: <b>{money(total_revenue - total_cost)}</b>\n\n"
        f"👤 Total users: <b>{int(users['n'] or 0)}</b>\n"
        f"💰 Current user balances: <b>{money(int(users['bal'] or 0))}</b>"
    )
    await replace_callback_message(query, context.bot, text, parse_mode=ParseMode.HTML, reply_markup=admin_menu_markup())


async def admin_catalog(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Inline-button entry — show service-first catalog view."""
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    await query.answer()
    text, markup = admin_catalog_services_content()
    await replace_callback_message(query, context.bot, text, parse_mode=ParseMode.HTML, reply_markup=markup)


async def add_catalog_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query or not await require_admin(update):
        return ConversationHandler.END
    await query.answer()
    context.user_data["new_catalog"] = {}
    await query.message.reply_text(
        "Enter the service short code, for example Instagram is <code>ig</code>.\n"
        "Send /cancel to stop.",
        parse_mode=ParseMode.HTML,
        reply_markup=ReplyKeyboardRemove(),
    )
    return CAT_SERVICE_CODE


async def add_catalog_country_id(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = (update.effective_message.text or "").strip()
    if not value.isdigit() or len(value) > 12:
        await update.effective_message.reply_text("Enter a numeric country ID, for example 33.")
        return CAT_COUNTRY_ID
    context.user_data["new_catalog"]["country_id"] = value
    await update.effective_message.reply_text("Enter the display country name, for example Cambodia.")
    return CAT_COUNTRY_NAME


async def add_catalog_country_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = (update.effective_message.text or "").strip()
    if not value or len(value) > 80:
        await update.effective_message.reply_text("Enter a valid country name.")
        return CAT_COUNTRY_NAME
    context.user_data["new_catalog"]["country_name"] = value
    await update.effective_message.reply_text("Enter the service short code, for example Instagram is ig.")
    return CAT_SERVICE_CODE


async def add_catalog_service_code(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = (update.effective_message.text or "").strip().lower()
    if not re.fullmatch(r"[a-z0-9_-]{1,50}", value):
        await update.effective_message.reply_text("Enter a valid service short code, for example ig.")
        return CAT_SERVICE_CODE
    context.user_data["new_catalog"]["service_code"] = value
    await update.effective_message.reply_text("Enter the display service name, for example Telegram.")
    return CAT_SERVICE_NAME


async def add_catalog_extra(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw = (update.effective_message.text or "").strip()
    try:
        extra = parse_extra_params(raw)
    except Exception as exc:
        await update.effective_message.reply_text(f"Invalid parameters: {exc}")
        return CAT_EXTRA_PARAMS
    context.user_data["new_catalog"]["extra_params_json"] = json.dumps(extra, ensure_ascii=False)
    await update.effective_message.reply_text(
        f"Enter markup in USD, for example <code>0.005</code>.\n"
        f"Send <code>-</code> to use global markup ({money(DB.global_markup_units())}).",
        parse_mode=ParseMode.HTML,
    )
    return CAT_MARKUP


async def set_markup_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query or not await require_admin(update):
        return ConversationHandler.END
    await query.answer()
    await query.message.reply_text(
        f"Current global markup: <b>{money(DB.global_markup_units())}</b>\n"
        "Enter the new markup in USD, for example <code>0.005</code>.",
        parse_mode=ParseMode.HTML,
    )
    return SET_MARKUP


async def admin_home_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    await query.answer()
    await replace_callback_message(query, context.bot, "<b>Admin panel</b>", parse_mode=ParseMode.HTML, reply_markup=admin_menu_markup())


async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    text = " ".join(context.args).strip()
    if not text:
        await update.effective_message.reply_text("Usage: /broadcast YOUR MESSAGE")
        return
    text = text[:MAX_BROADCAST_LENGTH]
    users = DB.fetchall("SELECT user_id FROM users WHERE is_blocked=0")
    sent = failed = 0
    progress = await update.effective_message.reply_text(f"Broadcasting to {len(users)} users…")
    for row in users:
        try:
            await context.bot.send_message(row["user_id"], text)
            sent += 1
        except (Forbidden, TelegramError):
            failed += 1
        await asyncio.sleep(0.075)
    await progress.edit_text(f"Broadcast finished. Sent: {sent}, failed: {failed}.")


# =============================================================================
# Error handler and lifecycle
# =============================================================================


def database_absolute_path() -> str:
    return os.path.abspath(os.path.expanduser(DATABASE_PATH))


def database_remove_sidecars(path: str) -> None:
    for suffix in ("-wal", "-shm"):
        with suppress(FileNotFoundError):
            os.remove(path + suffix)


def database_snapshot_to(path: str) -> None:
    """Create a transactionally consistent SQLite snapshot while the bot runs."""
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    database = DB
    with database.lock:
        destination = sqlite3.connect(path)
        try:
            database.conn.backup(destination)
            destination.commit()
        finally:
            destination.close()


def database_validate_file(path: str) -> dict[str, int]:
    """Validate SQLite integrity and return a small restore summary."""
    return _database_file_summary(path, immutable=False)


def database_zip_snapshot(snapshot_path: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    fd, zip_path = tempfile.mkstemp(
        prefix=f"{DATABASE_BACKUP_PREFIX}_{timestamp}_",
        suffix=".zip",
    )
    os.close(fd)

    summary = database_validate_file(snapshot_path)
    manifest = {
        "format": "paid-number-bot-database-backup",
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "database_filename": os.path.basename(DATABASE_PATH),
        "summary": summary,
    }

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(snapshot_path, arcname="database.sqlite3")
        archive.writestr(
            "manifest.json",
            json.dumps(manifest, ensure_ascii=False, indent=2),
        )
    return zip_path


def create_database_backup_zip() -> str:
    fd, snapshot_path = tempfile.mkstemp(
        prefix="bot_db_snapshot_",
        suffix=".sqlite3",
    )
    os.close(fd)
    try:
        database_snapshot_to(snapshot_path)
        return database_zip_snapshot(snapshot_path)
    finally:
        with suppress(FileNotFoundError):
            os.remove(snapshot_path)


def database_extract_candidate(upload_path: str, original_name: str) -> str:
    """
    Return a validated local SQLite file extracted from a ZIP or copied from
    a raw .db/.sqlite/.sqlite3 upload.
    """
    suffix = os.path.splitext(original_name or upload_path)[1].lower()
    fd, candidate_path = tempfile.mkstemp(
        prefix="db_restore_candidate_",
        suffix=".sqlite3",
    )
    os.close(fd)

    try:
        if suffix == ".zip":
            with zipfile.ZipFile(upload_path, "r") as archive:
                entries = [
                    item
                    for item in archive.infolist()
                    if not item.is_dir()
                    and os.path.splitext(item.filename)[1].lower()
                    in {".db", ".sqlite", ".sqlite3"}
                ]
                if len(entries) != 1:
                    raise ValueError(
                        "Backup ZIP must contain exactly one SQLite database file."
                    )
                entry = entries[0]
                safe_name = entry.filename.replace("\\", "/")
                if safe_name.startswith("/") or ".." in safe_name.split("/"):
                    raise ValueError("Unsafe path found inside ZIP.")
                if entry.file_size > DATABASE_BACKUP_MAX_BYTES:
                    raise ValueError("Database inside ZIP is too large.")
                if entry.compress_size <= 0 and entry.file_size > 0:
                    raise ValueError("Invalid compressed database entry.")
                if entry.compress_size and entry.file_size / entry.compress_size > 250:
                    raise ValueError("Suspicious ZIP compression ratio.")
                with archive.open(entry, "r") as source, open(
                    candidate_path, "wb"
                ) as target:
                    copied = 0
                    while True:
                        chunk = source.read(1024 * 1024)
                        if not chunk:
                            break
                        copied += len(chunk)
                        if copied > DATABASE_BACKUP_MAX_BYTES:
                            raise ValueError("Database inside ZIP is too large.")
                        target.write(chunk)
        elif suffix in {".db", ".sqlite", ".sqlite3"}:
            shutil.copy2(upload_path, candidate_path)
        else:
            raise ValueError("Upload a .zip, .db, .sqlite or .sqlite3 file.")

        database_validate_file(candidate_path)
        return candidate_path
    except Exception:
        with suppress(FileNotFoundError):
            os.remove(candidate_path)
        raise


async def database_stop_background_tasks() -> None:
    """Pause all tasks that can read/write the database during replacement."""
    global POLL_TASK, DAILY_REPORT_TASK, DAILY_BACKUP_TASK
    global PERSISTENT_DATABASE_BACKUP_TASK

    tasks = [
        POLL_TASK,
        DAILY_REPORT_TASK,
        DAILY_BACKUP_TASK,
        PERSISTENT_DATABASE_BACKUP_TASK,
        *list(LOCAL_ORDER_POLL_TASKS.values()),
        *list(LOCAL_DETAIL_TASKS.values()),
        *list(LOCAL_ACTIVE_TASKS.values()),
    ]
    for task in tasks:
        if task and not task.done():
            task.cancel()
    for task in tasks:
        if task:
            with suppress(asyncio.CancelledError):
                await task

    POLL_TASK = None
    DAILY_REPORT_TASK = None
    DAILY_BACKUP_TASK = None
    PERSISTENT_DATABASE_BACKUP_TASK = None
    LOCAL_ORDER_POLL_TASKS.clear()
    LOCAL_DETAIL_TASKS.clear()
    LOCAL_ACTIVE_TASKS.clear()
    LOCAL_ORDER_VIEWS.clear()


def database_reopen_from(path: str) -> None:
    """Atomically replace the live database and run all safe migrations."""
    global DB

    live_path = database_absolute_path()
    os.makedirs(os.path.dirname(live_path) or ".", exist_ok=True)
    staging_path = live_path + ".restore.tmp"

    with suppress(FileNotFoundError):
        os.remove(staging_path)
    shutil.copy2(path, staging_path)
    database_validate_file(staging_path)

    old_db = DB
    with old_db.lock:
        with suppress(sqlite3.Error):
            old_db.conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        old_db.conn.close()

    database_remove_sidecars(live_path)
    os.replace(staging_path, live_path)

    DB = Database(DATABASE_PATH)
    ensure_advanced_schema()
    ensure_management_schema()
    ensure_local_usa_schema()


def retry_after_seconds(exc: RetryAfter) -> int:
    value = exc.retry_after
    if hasattr(value, "total_seconds"):
        return max(1, int(value.total_seconds()))
    return max(1, int(value))


async def database_restore_receive(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not await require_admin(update):
        return ConversationHandler.END

    document = update.effective_message.document
    if not document:
        await update.effective_message.reply_text(
            "Send a .zip, .db, .sqlite or .sqlite3 document.",
            reply_markup=current_workflow_keyboard(update, context),
        )
        return DB_RESTORE_FILE

    if document.file_size and document.file_size > DATABASE_BACKUP_MAX_BYTES:
        await update.effective_message.reply_text(
            "❌ File is too large.",
            reply_markup=current_workflow_keyboard(update, context),
        )
        return DB_RESTORE_FILE

    filename = document.file_name or "database_upload"
    suffix = os.path.splitext(filename)[1].lower()
    if suffix not in {".zip", ".db", ".sqlite", ".sqlite3"}:
        await update.effective_message.reply_text(
            "❌ Unsupported file. Send ZIP or SQLite.",
            reply_markup=current_workflow_keyboard(update, context),
        )
        return DB_RESTORE_FILE

    fd, upload_path = tempfile.mkstemp(prefix="db_upload_", suffix=suffix)
    os.close(fd)
    candidate_path = ""

    try:
        telegram_file = await document.get_file()
        await telegram_file.download_to_drive(custom_path=upload_path)
        candidate_path = await run_file_job(
            database_extract_candidate,
            upload_path,
            filename,
        )
        summary = await run_file_job(
            database_validate_file,
            candidate_path,
        )
    except Exception as exc:
        logger.exception("Uploaded database validation failed")
        if candidate_path:
            with suppress(FileNotFoundError):
                os.remove(candidate_path)
        await update.effective_message.reply_text(
            f"❌ Invalid backup: {short_text(str(exc), 700)}",
            reply_markup=current_workflow_keyboard(update, context),
        )
        return DB_RESTORE_FILE
    finally:
        with suppress(FileNotFoundError):
            os.remove(upload_path)

    old_candidate = context.user_data.pop("db_restore_candidate", None)
    if old_candidate:
        with suppress(FileNotFoundError):
            os.remove(old_candidate)

    context.user_data["db_restore_candidate"] = candidate_path
    context.user_data["db_restore_filename"] = filename

    await update.effective_message.reply_text(
        "<b>Restore this database?</b>\n"
        f"Users: <b>{summary['users']}</b>\n"
        f"Total balance: <b>{money(summary['balance_units'])}</b>\n"
        f"Global orders: <b>{summary['orders']}</b>\n"
        f"Local orders: <b>{summary['local_orders']}</b>\n"
        f"Deposits: <b>{summary['deposits']}</b>\n\n"
        "The current database will be sent to you as a safety backup first.",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(
                        "✅ Confirm Restore",
                        callback_data="dbrestore:confirm",
                        style="danger",
                    ),
                    InlineKeyboardButton(
                        "❌ Cancel",
                        callback_data="dbrestore:cancel",
                        style="secondary",
                    ),
                ]
            ]
        ),
    )
    return DB_RESTORE_CONFIRM


async def database_restore_cancel(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    query = update.callback_query
    if query:
        await query.answer("Restore cancelled.")
        with suppress(BadRequest, TelegramError):
            await query.edit_message_reply_markup(reply_markup=None)

    candidate = context.user_data.pop("db_restore_candidate", None)
    context.user_data.pop("db_restore_filename", None)
    if candidate:
        with suppress(FileNotFoundError):
            os.remove(candidate)

    await update.effective_message.reply_text(
        "Restore cancelled.",
        reply_markup=admin_database_settings_menu(),
    )
    return ConversationHandler.END


def validate_config() -> None:
    problems = []
    if (
        not BOT_TOKEN
        or BOT_TOKEN.upper() == "XXX"
        or BOT_TOKEN.upper().startswith("PASTE_")
        or "DEMO" in BOT_TOKEN.upper()
        or "REPLACE" in BOT_TOKEN.upper()
    ):
        problems.append("Set a real BOT_TOKEN in Zero-Bot Environment or the private .env file; demo placeholders are not accepted.")
    if not DB.get_setting("panel_smsbower_api_key", SMS_API_KEY):
        logger.warning("SMSBower API key is not configured; add it from Admin → Panel Management.")
    if not SMS_API_URL or SMS_API_URL.upper().startswith("PASTE_") or "REPLACE" in SMS_API_URL.upper():
        problems.append("Set SMS_API_URL to the exact handler URL from the SMS Power API page.")
    if not SMS_API_URL.lower().startswith("https://"):
        problems.append("SMS_API_URL must use HTTPS.")
    if not DB.get_setting("panel_localusa_api_key", LOCAL_API_KEY):
        logger.warning("Local USA API key is not configured; add it from Admin → Panel Management.")
    if not LOCAL_API_URL.lower().startswith("https://"):
        problems.append("LOCAL_API_URL must use HTTPS.")
    if (
        not IPSHOP_API_KEY
        or IPSHOP_API_KEY.upper() == "XXX"
        or IPSHOP_API_KEY.upper().startswith("PASTE_")
        or "DEMO" in IPSHOP_API_KEY.upper()
        or "REPLACE" in IPSHOP_API_KEY.upper()
    ):
        problems.append("Set a real IPSHOP_API_KEY in Zero-Bot Environment or the private .env file; demo placeholders are not accepted.")
    if problems:
        raise SystemExit("\nCONFIG ERROR:\n- " + "\n- ".join(problems))
    if not ADMIN_IDS:
        logger.warning("ADMIN_IDS is empty. The bot can run and /id will work, but admin features stay disabled until you add your Telegram ID and restart.")


# =============================================================================
# ADVANCED BDT / ADMIN / EXPORT OVERRIDES
# =============================================================================

# User-facing balances and prices are stored in micro-BDT. Provider prices and
# markup remain in micro-USD until they are converted using the admin rate.
USD_SYMBOL = "$"
DEFAULT_USD_BDT_RATE = os.getenv("DEFAULT_USD_BDT_RATE", "125")
DAILY_REPORT_TASK: Optional[asyncio.Task[Any]] = None
DAILY_BACKUP_TASK: Optional[asyncio.Task[Any]] = None
PERSISTENT_DATABASE_BACKUP_TASK: Optional[asyncio.Task[Any]] = None

# Additional conversation states.
DEP_AMOUNT_NEW = 101
SET_DEPOSIT_GROUP = 106
SET_EXCHANGE_RATE = 107
SET_SUPPORT_LINK = 108


def ensure_advanced_schema() -> None:
    """Apply migrations for automatic deposits and the current service schema."""
    with DB.lock:
        DB.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS payment_invoices (
                invoice_id TEXT PRIMARY KEY,
                user_id INTEGER NOT NULL,
                chat_id INTEGER NOT NULL,
                amount_units INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'awaiting_payment',
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(user_id)
            );
            CREATE INDEX IF NOT EXISTS idx_payment_invoices_status
                ON payment_invoices(status, created_at);
            """
        )
        order_columns = {row["name"] for row in DB.conn.execute("PRAGMA table_info(orders)")}
        for column, ddl in (
            ("balance_after_units", "ALTER TABLE orders ADD COLUMN balance_after_units INTEGER"),
            ("otp_message_id", "ALTER TABLE orders ADD COLUMN otp_message_id INTEGER"),
            ("cc_controls_enabled", "ALTER TABLE orders ADD COLUMN cc_controls_enabled INTEGER NOT NULL DEFAULT 1"),
            ("ever_received_otp", "ALTER TABLE orders ADD COLUMN ever_received_otp INTEGER NOT NULL DEFAULT 0"),
            ("otp_delivery_blocked", "ALTER TABLE orders ADD COLUMN otp_delivery_blocked INTEGER NOT NULL DEFAULT 0"),
        ):
            if column not in order_columns:
                DB.conn.execute(ddl)

        DB.conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES('usd_bdt_rate',?)",
            (DEFAULT_USD_BDT_RATE,),
        )
        DB.conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES('deposit_group_id','')"
        )
        DB.conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES('support_link','')"
        )

        # One-time migration from the older USD-based user ledger to BDT. The
        # provider_cost_units column deliberately stays in USD for API maxPrice.
        mode_row = DB.conn.execute("SELECT value FROM settings WHERE key='currency_mode'").fetchone()
        if not mode_row:
            rate = decimal_value(DB.conn.execute(
                "SELECT value FROM settings WHERE key='usd_bdt_rate'"
            ).fetchone()[0])
            rate_text = str(rate)
            DB.conn.execute("UPDATE users SET balance_units=CAST(ROUND(balance_units * ?) AS INTEGER)", (rate_text,))
            DB.conn.execute("UPDATE deposits SET amount_units=CAST(ROUND(amount_units * ?) AS INTEGER)", (rate_text,))
            DB.conn.execute("UPDATE transactions SET amount_units=CAST(ROUND(amount_units * ?) AS INTEGER)", (rate_text,))
            DB.conn.execute("UPDATE orders SET charged_units=CAST(ROUND(charged_units * ?) AS INTEGER)", (rate_text,))
            DB.conn.execute("INSERT INTO settings(key,value) VALUES('currency_mode','BDT')")

        # Old manual requests and unpaid automatic invoices were stored as
        # pending deposits. They are not valid balance credits in the new flow.
        DB.conn.execute("DELETE FROM deposits WHERE status='pending'")
        DB.conn.commit()


ensure_advanced_schema()


def bdt_to_units(value: Any) -> int:
    amount = decimal_value(value)
    return int((amount * UNITS_PER_USD).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def bdt_input_value(value: Any) -> Decimal:
    """Parse an admin-entered BDT amount without requiring plain digits."""
    raw = str(value or "").strip()
    raw = raw.replace("৳", "").replace("BDT", "").replace(",", "").strip()
    return decimal_value(raw)


def money(units: int) -> str:
    value = units_to_decimal(int(units or 0))
    text = f"{value:.2f}"
    return f"৳{text}"


def user_account_summary(user_id: int) -> dict[str, int]:
    """Return current balance and combined Global + USA Local order totals."""
    user = DB.fetchone(
        "SELECT balance_units FROM users WHERE user_id=?",
        (user_id,),
    )
    current_balance = int(user["balance_units"] or 0) if user else 0

    global_summary = DB.fetchone(
        """
        SELECT COUNT(*) AS order_count,
               COALESCE(SUM(charged_units),0) AS purchase_units,
               COALESCE(SUM(CASE WHEN refunded=1 THEN 1 ELSE 0 END),0)
                   AS refund_count,
               COALESCE(SUM(CASE WHEN refunded=1 THEN charged_units ELSE 0 END),0)
                   AS refund_units
        FROM orders
        WHERE user_id=?
        """,
        (user_id,),
    )
    local_summary = DB.fetchone(
        """
        SELECT COUNT(*) AS order_count,
               COALESCE(SUM(charged_units),0) AS purchase_units,
               COALESCE(SUM(CASE WHEN refunded=1 THEN 1 ELSE 0 END),0)
                   AS refund_count,
               COALESCE(SUM(CASE WHEN refunded=1 THEN charged_units ELSE 0 END),0)
                   AS refund_units
        FROM local_orders
        WHERE user_id=?
        """,
        (user_id,),
    )

    return {
        "current_balance": current_balance,
        "order_count": int(global_summary["order_count"] or 0)
        + int(local_summary["order_count"] or 0),
        "purchase_units": int(global_summary["purchase_units"] or 0)
        + int(local_summary["purchase_units"] or 0),
        "refund_count": int(global_summary["refund_count"] or 0)
        + int(local_summary["refund_count"] or 0),
        "refund_units": int(global_summary["refund_units"] or 0)
        + int(local_summary["refund_units"] or 0),
    }


def sync_recent_purchase_history(user_id: Optional[int] = None) -> None:
    """Snapshot recent purchases before daily cleanup removes source order rows."""
    user_filter = " WHERE user_id=?" if user_id is not None else ""
    user_params: tuple[Any, ...] = (int(user_id),) if user_id is not None else ()

    with DB.lock:
        conn = DB.conn
        conn.execute("BEGIN IMMEDIATE")
        try:
            snapshots: list[tuple[Any, ...]] = []

            global_orders = conn.execute(
                """
                SELECT user_id,id,created_at,country_name,service_name,phone,
                       charged_units,status,refunded
                FROM orders
                """ + user_filter,
                user_params,
            ).fetchall()
            for row in global_orders:
                product = "Global number · " + short_text(
                    f"{row['country_name']} / {row['service_name']}", 100
                )
                snapshots.append(
                    (
                        int(row["user_id"]), "global", int(row["id"]),
                        int(row["created_at"]), product, str(row["phone"] or ""),
                        int(row["charged_units"] or 0), str(row["status"] or ""),
                        int(row["refunded"] or 0),
                    )
                )

            local_orders = conn.execute(
                """
                SELECT user_id,id,created_at,phone,charged_units,status,refunded
                FROM local_orders
                """ + user_filter,
                user_params,
            ).fetchall()
            for row in local_orders:
                snapshots.append(
                    (
                        int(row["user_id"]), "local", int(row["id"]),
                        int(row["created_at"]), "USA Local number",
                        str(row["phone"] or ""), int(row["charged_units"] or 0),
                        str(row["status"] or ""), int(row["refunded"] or 0),
                    )
                )

            vshop_filter = " WHERE vo.user_id=?" if user_id is not None else ""
            vshop_orders = conn.execute(
                """
                SELECT vo.user_id,vo.id,vo.created_at,vo.quantity,vo.total_units,
                       vo.status,COALESCE(vc.name,'Product') AS category_name,
                       COALESCE(vs.name,'Item') AS service_name
                FROM vshop_orders AS vo
                LEFT JOIN vshop_categories AS vc ON vc.id=vo.category_id
                LEFT JOIN vshop_services AS vs ON vs.id=vo.service_id
                """ + vshop_filter,
                user_params,
            ).fetchall()
            for row in vshop_orders:
                product = short_text(
                    f"{row['category_name']} / {row['service_name']}", 100
                )
                quantity = max(0, int(row["quantity"] or 0))
                snapshots.append(
                    (
                        int(row["user_id"]), "vshop", int(row["id"]),
                        int(row["created_at"]), product,
                        f"{quantity} item{'s' if quantity != 1 else ''}",
                        int(row["total_units"] or 0), str(row["status"] or ""), 0,
                    )
                )

            conn.executemany(
                """
                INSERT INTO purchase_history(
                    user_id,source,source_order_id,created_at,product_name,
                    item_detail,amount_units,status,refunded
                ) VALUES(?,?,?,?,?,?,?,?,?)
                ON CONFLICT(source,source_order_id) DO UPDATE SET
                    user_id=excluded.user_id,
                    created_at=excluded.created_at,
                    product_name=excluded.product_name,
                    item_detail=excluded.item_detail,
                    amount_units=excluded.amount_units,
                    status=excluded.status,
                    refunded=excluded.refunded
                """,
                snapshots,
            )

            if user_id is None:
                conn.execute(
                    """
                    DELETE FROM purchase_history
                    WHERE id IN (
                        SELECT id FROM (
                            SELECT id,ROW_NUMBER() OVER(
                                PARTITION BY user_id
                                ORDER BY created_at DESC,id DESC
                            ) AS row_number
                            FROM purchase_history
                        )
                        WHERE row_number>10
                    )
                    """
                )
            else:
                conn.execute(
                    """
                    DELETE FROM purchase_history
                    WHERE user_id=? AND id NOT IN (
                        SELECT id FROM purchase_history
                        WHERE user_id=?
                        ORDER BY created_at DESC,id DESC
                        LIMIT 10
                    )
                    """,
                    (int(user_id), int(user_id)),
                )
            conn.commit()
        except Exception:
            conn.rollback()
            raise


def money_usd(units: int) -> str:
    value = units_to_decimal(int(units or 0))
    text = f"{value:.6f}".rstrip("0").rstrip(".")
    if "." not in text:
        text += ".00"
    return f"{USD_SYMBOL}{text}"


def exchange_rate() -> Decimal:
    try:
        rate = decimal_value(DB.get_setting("usd_bdt_rate", DEFAULT_USD_BDT_RATE))
        return rate if rate > 0 else Decimal(DEFAULT_USD_BDT_RATE)
    except Exception:
        return Decimal(DEFAULT_USD_BDT_RATE)


def usd_units_to_bdt_units(usd_units: int) -> int:
    usd = units_to_decimal(int(usd_units or 0))
    return bdt_to_units(usd * exchange_rate())


def format_timestamp(ts: Any) -> str:
    if not ts:
        return ""
    try:
        return datetime.fromtimestamp(int(ts)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def copy_button(label: str, value: str) -> InlineKeyboardButton:
    # Copy buttons (including OTP/code copy buttons) use Telegram's success style.
    return InlineKeyboardButton(label, copy_text=CopyTextButton(str(value)), style="success")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    row = await ensure_registered(update)
    if not row or not update.effective_message:
        return
    referrer_id = referral_start_id(context)
    if referrer_id:
        DB.set_referrer_from_start(update.effective_user.id, referrer_id)
    if not row["accepted_terms"]:
        await require_terms(update)
        return
    await update.effective_message.reply_text(
        "Welcome. Choose an option below.",
        reply_markup=main_menu(update.effective_user.id if update.effective_user else None),
    )


async def accept_terms(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not update.effective_user:
        return
    await query.answer()
    DB.execute(
        "UPDATE users SET accepted_terms=1, updated_at=? WHERE user_id=?",
        (now_ts(), update.effective_user.id),
    )
    accepted_message = await replace_callback_message(
        query, context.bot, "✅ Agreement accepted."
    )
    await accepted_message.reply_text(
        "Choose an option below.", reply_markup=main_menu(update.effective_user.id)
    )


def render_account_summary_text(user_id: int) -> str:
    summary = user_account_summary(user_id)
    return (
        "<b>💰 Account Summary</b>\n\n"
        f"Current balance: <b>{money(summary['current_balance'])}</b>\n"
        f"Numbers purchased: <b>{summary['order_count']}</b>\n"
        f"Total purchase amount: <b>{money(summary['purchase_units'])}</b>\n"
        f"Refunded numbers: <b>{summary['refund_count']}</b>\n"
        f"Total refund amount: <b>{money(summary['refund_units'])}</b>"
    )


def user_account_summary_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [[
            InlineKeyboardButton("🔗 Refer", callback_data="referral:show"),
            InlineKeyboardButton(
                "📜 My History",
                callback_data="purchasehistory:mine",
                style="danger",
            ),
        ]]
    )


def render_purchase_history_text(user_id: int, *, admin_view: bool = False) -> str:
    sync_recent_purchase_history(user_id)
    rows = DB.fetchall(
        """
        SELECT created_at,product_name,item_detail,amount_units,status,refunded
        FROM purchase_history
        WHERE user_id=?
        ORDER BY created_at DESC,id DESC
        LIMIT 10
        """,
        (user_id,),
    )
    title = (
        f"🧾 Purchase History — User {int(user_id)}"
        if admin_view
        else "🧾 Last 10 purchases"
    )
    lines = [f"<b>{html.escape(title)}</b>"]
    if not rows:
        lines.append("\nNo purchase history yet.")
        return "".join(lines)

    for index, row in enumerate(rows, start=1):
        product = html.escape(short_text(row["product_name"], 100))
        detail = html.escape(short_text(row["item_detail"], 80) or "-")
        status = (
            "↩️ Refunded"
            if int(row["refunded"] or 0)
            else status_label(str(row["status"] or "unknown"))
        )
        lines.append(
            f"\n\n<b>{index}. {product}</b>\n"
            f"📦 <code>{detail}</code>\n"
            f"💵 {money(int(row['amount_units'] or 0))} • "
            f"{html.escape(status)}\n"
            f"🕒 {html.escape(_dhaka_format_timestamp(row['created_at']))}"
        )
    return "".join(lines)


async def show_balance(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    if not await require_terms(update):
        return

    user_id = int(update.effective_user.id)
    await update.effective_message.reply_text(
        render_account_summary_text(user_id),
        parse_mode=ParseMode.HTML,
        reply_markup=user_account_summary_keyboard(),
    )


async def show_my_purchase_history(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_terms(update):
        return
    await query.answer()
    user_id = int(update.effective_user.id)
    await replace_callback_message(
        query,
        context.bot,
        render_purchase_history_text(user_id),
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Back to My Balance", callback_data="purchasehistory:back")]]
        ),
    )


async def show_my_purchase_history_back(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_terms(update):
        return
    await query.answer()
    user_id = int(update.effective_user.id)
    await replace_callback_message(
        query,
        context.bot,
        render_account_summary_text(user_id),
        parse_mode=ParseMode.HTML,
        reply_markup=user_account_summary_keyboard(),
    )


async def admin_user_purchase_history(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    try:
        user_id = int(query.data.rsplit(":", 1)[1])
        row = user_profile_row(user_id)
    except (ValueError, OverflowError, IndexError):
        row = None
        user_id = 0
    if not row:
        await query.answer("User not found.", show_alert=True)
        return
    await query.answer()
    await replace_callback_message(
        query,
        context.bot,
        render_purchase_history_text(user_id, admin_view=True),
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Back to User", callback_data=f"umprofile:{user_id}")]]
        ),
    )


async def admin_user_profile_back(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    try:
        user_id = int(query.data.rsplit(":", 1)[1])
        row = user_profile_row(user_id)
    except (ValueError, OverflowError, IndexError):
        row = None
        user_id = 0
    if not row:
        await query.answer("User not found.", show_alert=True)
        return
    await query.answer()
    await replace_callback_message(
        query,
        context.bot,
        render_user_profile(row),
        parse_mode=ParseMode.HTML,
        reply_markup=user_profile_keyboard(user_id),
    )


async def show_referral(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not await require_terms(update):
        return
    await query.answer()
    user_id = int(update.effective_user.id)
    username = DB.get_setting("bot_username", "").strip().lstrip("@")
    link = (
        f"https://t.me/{username}?start=ref_{user_id}"
        if username
        else f"/start ref_{user_id}"
    )
    reward = bdt_to_units(DB.get_setting("referral_reward_bdt", "0"))
    threshold = bdt_to_units(DB.get_setting("referral_min_deposit_bdt", "0"))
    successful = DB.fetchone(
        "SELECT COUNT(*) AS n FROM referral_rewards WHERE referrer_id=?",
        (user_id,),
    )
    pending = DB.fetchone(
        """
        SELECT COUNT(*) AS n
        FROM users u
        WHERE u.referred_by=?
          AND NOT EXISTS (
              SELECT 1
              FROM referral_rewards rr
              WHERE rr.referred_user_id=u.user_id
          )
        """,
        (user_id,),
    )
    successful_count = int(successful["n"] or 0) if successful else 0
    pending_count = int(pending["n"] or 0) if pending else 0
    await query.message.reply_text(
        "<b>🔗 Referral</b>\n\n"
        f"Your referral link:\n<code>{html.escape(link)}</code>\n\n"
        f"Per referral income: <b>{money(reward)}</b>\n"
        f"Required deposit by referred user: <b>{money(threshold)}</b>\n"
        f"Pending referrals: <b>{pending_count}</b>\n"
        f"Successful referrals: <b>{successful_count}</b>\n\n"
        "A referral is counted only once after the referred user deposits at least "
        f"<b>{money(threshold)}</b>.",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(
            [[copy_button("📋 Copy referral link", link)]]
        ),
    )


async def show_how_to_use(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_terms(update):
        return
    rows = DB.fetchall(
        "SELECT id,name,video_file_id FROM how_to_use_buttons ORDER BY id ASC"
    )
    if not rows:
        await update.effective_message.reply_text(
            "How to use videos have not been added yet."
        )
        return
    buttons = [
        [InlineKeyboardButton(str(row["name"]), callback_data=f"howuse:view:{row['id']}")]
        for row in rows
    ]
    await update.effective_message.reply_text(
        "📚 <b>How to use</b>\n\nChoose a topic:",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(buttons),
    )


async def show_how_to_use_video(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_terms(update):
        return
    await query.answer()
    try:
        button_id = int(str(query.data).rsplit(":", 1)[1])
    except (TypeError, ValueError):
        return
    row = DB.fetchone(
        "SELECT name,video_file_id FROM how_to_use_buttons WHERE id=?",
        (button_id,),
    )
    if not row:
        await query.message.reply_text("This How to use button is no longer available.")
        return
    if not row["video_file_id"]:
        await query.message.reply_text(
            f"🎬 {html.escape(str(row['name']))}\n\nThe video has not been set yet.",
            parse_mode=ParseMode.HTML,
        )
        return
    await context.bot.send_video(
        chat_id=query.message.chat_id,
        video=str(row["video_file_id"]),
        caption=f"📚 {str(row['name'])}",
    )


async def view_provider_offer(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not await require_terms(update):
        return
    await query.answer("Checking price…")
    try:
        _, _, raw_item_id, provider_id = query.data.split(":", 3)
        item_id = int(raw_item_id)
    except Exception:
        return
    item = get_catalog_item(item_id)
    if not item or not item["enabled"]:
        await query.edit_message_text("This item is no longer available.")
        return
    try:
        offers = await API.get_provider_offers(item["country_id"], item["service_code"], catalog_extra(item))
        offer = next(value for value in offers if value.provider_id == provider_id)
    except StopIteration:
        await query.edit_message_text(
            "That provider is no longer available. Refresh the live price list.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("🔄 Refresh prices", callback_data=f"buy:service:{item_id}")]]
            ),
        )
        return
    except Exception as exc:
        await query.edit_message_text(
            f"Live price error: <code>{html.escape(str(exc))}</code>",
            parse_mode=ParseMode.HTML,
        )
        return
    if offer.available <= 0:
        await query.edit_message_text(
            "This provider just ran out of stock. Refresh the live price list.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton(
                    "🔄 Refresh prices", callback_data=f"buy:service:{item_id}"
                )]]
            ),
        )
        return
    total_bdt = usd_units_to_bdt_units(offer.provider_cost_units + item_markup_units(item))
    stock = offer.available
    text_value = (
        f"{_country_tg_emoji(item['country_name'], item['country_id'])} {_service_tg_emoji(item['service_name'])} <b>{html.escape(item['country_name'])} — {html.escape(item['service_name'])}</b>\n\n"
        f"Price: <b>{money(total_bdt)}</b>\n"
        f"Available: <b>{stock}</b>\n"
        f"Validity: <b>{DB.expiry_seconds() // 60} minutes</b>\n\n"
        "The live price and stock will be checked again before purchase."
    )
    await query.edit_message_text(
        text_value,
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Buy this number", callback_data=f"buy:confirm:{item_id}:{provider_id}")],
            [InlineKeyboardButton("⬅️ Price list", callback_data=f"buy:service:{item_id}")],
        ]),
    )


# =============================================================================
# PREMIUM EMOJI MANAGEMENT — TXT upload/download and single custom emoji
# =============================================================================

PREMIUM_EMOJI_UPLOAD_MAX_BYTES = 2 * 1024 * 1024


def premium_emoji_management_markup() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⬆️ Upload Flags (TXT)", callback_data="premiumemoji:upload:flags"), InlineKeyboardButton("⬇️ Download Flags", callback_data="premiumemoji:download:flags")],
        [InlineKeyboardButton("⬆️ Upload Services (TXT)", callback_data="premiumemoji:upload:services"), InlineKeyboardButton("⬇️ Download Services", callback_data="premiumemoji:download:services")],
        [InlineKeyboardButton("🗑 Delete All Flags & Services", callback_data="premiumemoji:deleteall")],
        [InlineKeyboardButton("➕ Add Single Emoji", callback_data="premiumemoji:addsingle")],
        [InlineKeyboardButton("⬅️ Back", callback_data="premiumemoji:back")],
    ])


async def premium_emoji_menu_reply(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    context.user_data.pop("premium_emoji_mode", None)
    await update.effective_message.reply_text(
        "<b>⭐ Premium Emoji Management</b>\nUpload TXT files or manage individual emojis below.",
        parse_mode=ParseMode.HTML,
        reply_markup=premium_emoji_management_markup(),
    )


async def premium_emoji_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    data = str(query.data or "")
    await query.answer()
    if data == "premiumemoji:menu":
        await query.edit_message_text(
            "<b>⭐ Premium Emoji Management</b>\nUpload TXT files or manage individual emojis below.",
            parse_mode=ParseMode.HTML,
            reply_markup=premium_emoji_management_markup(),
        )
    elif data in ("premiumemoji:back",):
        context.user_data.pop("premium_emoji_mode", None)
        await query.edit_message_text("Admin panel", reply_markup=admin_menu_markup())
    elif data.startswith("premiumemoji:upload:"):
        kind = data.rsplit(":", 1)[-1]
        context.user_data["premium_emoji_mode"] = "upload_" + kind
        await query.message.reply_text(
            f"Send the {kind.upper()} TXT file now. Existing entries will be kept; new entries will be added and matching names updated.",
            reply_markup=ReplyKeyboardMarkup([[KeyboardButton("❌ Cancel")]], resize_keyboard=True),
        )
    elif data.startswith("premiumemoji:download:"):
        kind = data.rsplit(":", 1)[-1]
        rows = PREMIUM_FLAG_ROWS if kind == "flags" else PREMIUM_SERVICE_ROWS
        lines = []
        if kind == "flags":
            for row in rows:
                lines.append(f"{row.get('flag','')} ({row.get('calling','')}) ({row.get('iso','')}) {row.get('name','')} {{\"emoji\": \"{row.get('flag','')}\", \"id\": \"{row.get('id','')}\"}}")
        else:
            for row in rows:
                lines.append(f"{row.get('emoji','')} {row.get('name','')} {{\"emoji\": \"{row.get('emoji','')}\", \"id\": \"{row.get('id','')}\"}}")
        payload = ("\n".join(lines) + "\n").encode("utf-8")
        from io import BytesIO
        handle = BytesIO(payload)
        handle.name = f"{kind.title()}_Emojis.txt"
        await query.message.reply_document(document=handle, filename=handle.name, caption=f"Current {kind} Premium Emoji list")
    elif data == "premiumemoji:deleteall":
        PREMIUM_FLAG_ROWS.clear()
        PREMIUM_SERVICE_ROWS.clear()
        _refresh_premium_emoji_maps()
        _save_premium_emoji_data()
        await query.edit_message_text(
            "✅ All Country Flags and Service Emoji mappings have been deleted.",
            reply_markup=premium_emoji_management_markup(),
        )
    elif data == "premiumemoji:addsingle":
        context.user_data["premium_emoji_mode"] = "add_single"
        await query.message.reply_text(
            "Send one message in this format with a Premium Custom Emoji attached: "
            "<code>flag: Bangladesh</code> or <code>service: Instagram</code>.\n"
            "The bot will read the Custom Emoji ID from the message entity.",
            parse_mode=ParseMode.HTML,
        )


def _refresh_premium_emoji_maps() -> None:
    global _COUNTRY_EMOJI_CODES, _COUNTRY_EMOJI_CODES_NORMALIZED
    global _PREMIUM_COUNTRY_EMOJI_IDS_NORMALIZED, _PREMIUM_SERVICE_EMOJI_IDS_NORMALIZED
    _COUNTRY_EMOJI_CODES, _PREMIUM_COUNTRY_EMOJI_IDS_NORMALIZED, _PREMIUM_SERVICE_EMOJI_IDS_NORMALIZED = _build_emoji_maps(PREMIUM_FLAG_ROWS, PREMIUM_SERVICE_ROWS)
    _COUNTRY_EMOJI_CODES_NORMALIZED = {_emoji_name_key(k): str(v).upper() for k, v in _COUNTRY_EMOJI_CODES.items()}


def _parse_premium_emoji_txt(raw: str, kind: str) -> list[dict[str, str]]:
    parsed: list[dict[str, str]] = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        payload_match = re.search(r"(\{.*\})\s*$", line)
        if not payload_match:
            continue
        try:
            payload = json.loads(payload_match.group(1))
        except Exception:
            continue
        emoji = str(payload.get("emoji", "")).strip()
        emoji_id = str(payload.get("id", "")).strip()
        if not emoji or not emoji_id.isdigit():
            continue
        prefix = line[:payload_match.start()].strip()
        if kind == "flags":
            match = re.match(r"^(\S+)\s+\((\d+)\)\s+\(([A-Z]{2})\)\s+(.+)$", prefix)
            if not match:
                continue
            flag, calling, iso, name = match.groups()
            parsed.append({"name": name.strip(), "flag": emoji, "iso": iso, "calling": calling, "id": emoji_id})
        else:
            name = prefix
            if name.startswith(emoji):
                name = name[len(emoji):].strip()
            if name:
                parsed.append({"name": name, "emoji": emoji, "id": emoji_id})
    return parsed


async def premium_emoji_document_receive(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.user_data.get("premium_emoji_mode", "").startswith("upload_"):
        return
    if not await require_admin(update):
        return
    message = update.effective_message
    document = message.document if message else None
    if not document or not (document.file_name or "").lower().endswith(".txt"):
        await message.reply_text("❌ Please upload a .txt file.")
        return
    if document.file_size and document.file_size > PREMIUM_EMOJI_UPLOAD_MAX_BYTES:
        await message.reply_text("❌ TXT file is too large (maximum 2 MB).")
        return
    kind = context.user_data["premium_emoji_mode"].replace("upload_", "", 1)
    fd, path = tempfile.mkstemp(prefix="premium_emoji_", suffix=".txt")
    os.close(fd)
    try:
        tg_file = await document.get_file()
        await tg_file.download_to_drive(custom_path=path)
        with open(path, "r", encoding="utf-8-sig") as f:
            raw = f.read()
        incoming = _parse_premium_emoji_txt(raw, kind)
        if not incoming:
            await message.reply_text("❌ No valid emoji entries found. Check the TXT format.")
            return
        target = PREMIUM_FLAG_ROWS if kind == "flags" else PREMIUM_SERVICE_ROWS
        key_field = "name"
        existing = {_emoji_name_key(row.get(key_field)): i for i, row in enumerate(target)}
        added = updated = 0
        for row in incoming:
            key = _emoji_name_key(row.get(key_field))
            if key in existing:
                target[existing[key]] = row
                updated += 1
            else:
                existing[key] = len(target)
                target.append(row)
                added += 1
        _refresh_premium_emoji_maps()
        _save_premium_emoji_data()
        context.user_data.pop("premium_emoji_mode", None)
        await message.reply_text(
            f"✅ {kind.title()} uploaded.\nAdded: <b>{added}</b>\nUpdated: <b>{updated}</b>\nTotal: <b>{len(target)}</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=admin_menu(),
        )
    except Exception as exc:
        logger.exception("Premium emoji TXT upload failed")
        await message.reply_text(f"❌ Upload failed: {short_text(str(exc), 300)}")
    finally:
        with suppress(FileNotFoundError):
            os.remove(path)


async def premium_emoji_single_receive(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if context.user_data.get("premium_emoji_mode") != "add_single":
        return
    if not await require_admin(update):
        return
    message = update.effective_message
    text_value = str(message.text or "").strip() if message else ""
    entities = list(message.entities or []) if message else []
    custom_entity = next((e for e in entities if str(getattr(e, "type", "")) in ("custom_emoji", "MessageEntityType.CUSTOM_EMOJI")), None)
    emoji_id = str(getattr(custom_entity, "custom_emoji_id", "") or "")
    match = re.match(r"^(flag|service)\s*:\s*(.+)$", text_value, re.IGNORECASE)
    if not match or not emoji_id.isdigit():
        await message.reply_text("❌ Send `flag: Country Name` or `service: App Name` and include one Premium Custom Emoji.")
        return
    kind, name = match.group(1).lower(), match.group(2).strip()
    # Remove the visible custom emoji character from the label if it is appended.
    if custom_entity:
        visible = text_value
        # Telegram entity offsets are UTF-16; avoid slicing by offset and just trim trailing emoji-like token.
        name = re.sub(r"\s+\S+$", "", name).strip() or name
    if kind == "flag":
        flag = next((r.get("flag") for r in PREMIUM_FLAG_ROWS if _emoji_name_key(r.get("name")) == _emoji_name_key(name)), "🏳️")
        row = {"name": name, "flag": flag, "iso": "", "calling": "", "id": emoji_id}
        target = PREMIUM_FLAG_ROWS
    else:
        fallback = next((r.get("emoji") for r in PREMIUM_SERVICE_ROWS if _emoji_name_key(r.get("name")) == _emoji_name_key(name)), "📱")
        row = {"name": name, "emoji": fallback, "id": emoji_id}
        target = PREMIUM_SERVICE_ROWS
    key = _emoji_name_key(name)
    index = next((i for i,r in enumerate(target) if _emoji_name_key(r.get("name")) == key), None)
    if index is None: target.append(row)
    else: target[index] = row
    _refresh_premium_emoji_maps()
    _save_premium_emoji_data()
    context.user_data.pop("premium_emoji_mode", None)
    await message.reply_text(f"✅ Saved Premium Emoji for <b>{html.escape(name)}</b>.", parse_mode=ParseMode.HTML, reply_markup=admin_menu())


async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    context.user_data["admin_mode"] = True
    await update.effective_message.reply_text("<b>Admin panel</b>", parse_mode=ParseMode.HTML, reply_markup=admin_menu())


async def admin_back_main(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.pop("admin_mode", None)
    await update.effective_message.reply_text(
        "Main menu", reply_markup=main_menu(update.effective_user.id if update.effective_user else None)
    )


def admin_catalog_services_content() -> tuple[str, Optional[InlineKeyboardMarkup]]:
    """Service-first view: one button per distinct service with enabled/total count."""
    services = DB.fetchall(
        """
        SELECT MIN(id) AS service_ref,
               MIN(service_code) AS service_code,
               MAX(service_name) AS service_name,
               COUNT(DISTINCT country_id) AS total,
               COUNT(DISTINCT CASE WHEN enabled=1 THEN country_id END) AS enabled_count
        FROM catalog
        GROUP BY service_code COLLATE NOCASE
        ORDER BY MAX(service_name) COLLATE NOCASE
        """
    )
    if not services:
        return "Catalog is empty.\nAdd a service with ➕ Add Service.", None

    buttons: list[list[InlineKeyboardButton]] = []
    for row in services:
        total = int(row["total"])
        enabled = int(row["enabled_count"] or 0)
        buttons.append(
            [
                InlineKeyboardButton(
                    f"📱 {short_text(row['service_name'], 28)}  ({enabled}/{total} countries)",
                    callback_data=f"admin:catserviceid:{int(row['service_ref'])}:0",
                )
            ]
        )
    return (
        "<b>🌍 Catalog</b>\nService ক্লিক করুন → country-wise live stock দেখতে।",
        InlineKeyboardMarkup(buttons),
    )


def admin_catalog_service_detail(service_ref: int) -> tuple[str, InlineKeyboardMarkup]:
    """Service-specific admin screen with current markup and country controls."""
    row = DB.fetchone("SELECT * FROM catalog WHERE id=?", (service_ref,))
    if not row:
        return "Service no longer exists.", InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Services", callback_data="admin:catmenu")]]
        )
    service_code = str(row["service_code"])
    service_name = str(row["service_name"])
    count = DB.fetchone(
        "SELECT COUNT(DISTINCT country_id) AS total FROM catalog WHERE service_code=?",
        (service_code,),
    )
    override = row["markup_override_units"]
    margin_text = money_usd(int(override)) if override is not None else f"Default ({money_usd(DB.global_markup_units())})"
    text = (
        f"<b>📱 {html.escape(service_name)}</b>\n"
        f"Service code: <code>{html.escape(service_code)}</code>\n"
        f"Countries configured: <b>{int(count['total'] or 0) if count else 0}</b>\n"
        f"Current margin per number: <b>{html.escape(margin_text)}</b>\n\n"
        "Choose an action below. Margin changes apply to this service across all countries."
    )
    markup = InlineKeyboardMarkup([
        [InlineKeyboardButton("⚙️ Set / Change Margin", callback_data=f"admin:catmargin:{service_ref}")],
        [InlineKeyboardButton("🌍 View Countries", callback_data=f"admin:catcountries:{service_ref}:0")],
        [InlineKeyboardButton("⬅️ Services", callback_data="admin:catmenu")],
    ])
    return text, markup


async def admin_catalog_countries_content(
    service_code: str, page: int = 0
) -> tuple[str, InlineKeyboardMarkup, InlineKeyboardMarkup, bool]:
    """Render five serial-numbered country controls per admin catalog page."""
    total_row = DB.fetchone(
        "SELECT COUNT(DISTINCT country_id) AS total FROM catalog WHERE service_code=?",
        (service_code,),
    )
    total = int(total_row["total"] or 0) if total_row else 0
    if not total:
        empty_markup = InlineKeyboardMarkup(
            [[InlineKeyboardButton("⬅️ Services", callback_data="admin:catmenu")]]
        )
        return ".", empty_markup, empty_markup, False

    pages = max(1, math.ceil(total / CATALOG_COUNTRIES_PAGE_SIZE))
    page = max(0, min(int(page), pages - 1))
    rows = DB.fetchall(
        """
        SELECT c.* FROM catalog AS c
        INNER JOIN (
            SELECT MIN(id) AS id
            FROM catalog
            WHERE service_code=?
            GROUP BY country_id
        ) AS country_rows ON country_rows.id=c.id
        ORDER BY c.country_name COLLATE NOCASE
        LIMIT ? OFFSET ?
        """,
        (service_code, CATALOG_COUNTRIES_PAGE_SIZE, page * CATALOG_COUNTRIES_PAGE_SIZE),
    )
    service_ref_row = DB.fetchone(
        "SELECT MIN(id) AS service_ref FROM catalog WHERE service_code=?",
        (service_code,),
    )
    service_ref = int(service_ref_row["service_ref"]) if service_ref_row else 0
    buttons: list[list[InlineKeyboardButton]] = []
    fallback_buttons: list[list[InlineKeyboardButton]] = []
    has_premium_icons = False
    for serial, row in enumerate(
        rows, start=page * CATALOG_COUNTRIES_PAGE_SIZE + 1
    ):
        icon = "🟢" if row["enabled"] else "⚪"
        country_name = str(row["country_name"])
        country_id = str(row["country_id"])
        country_button, fallback_country_button, uses_premium = _country_button_pair(
            country_name,
            f"admin:cattoggle:{row['id']}:{page}",
            country_id=country_id,
            text_prefix=f"{serial}. {icon} ",
            text_limit=20,
        )
        buttons.append(
            [
                country_button,
                InlineKeyboardButton(
                    "🗑",
                    callback_data=f"admin:catdelete:{row['id']}:{page}",
                    style="danger",
                ),
            ]
        )
        fallback_buttons.append(
            [
                fallback_country_button,
                InlineKeyboardButton(
                    "🗑",
                    callback_data=f"admin:catdelete:{row['id']}:{page}",
                    style="danger",
                ),
            ]
        )
        has_premium_icons = has_premium_icons or uses_premium
    previous_button = InlineKeyboardButton(
        "⬅️",
        callback_data=f"admin:catpage:{service_ref}:{page - 1}" if page > 0 else "noop",
    )
    page_button = InlineKeyboardButton(
        f"📄 {page + 1}/{pages}", callback_data="noop"
    )
    next_button = InlineKeyboardButton(
        "➡️",
        callback_data=(
            f"admin:catpage:{service_ref}:{page + 1}"
            if page < pages - 1
            else "noop"
        ),
    )
    page_buttons = [previous_button, page_button, next_button]
    buttons.append(page_buttons)
    fallback_buttons.append(page_buttons)
    back_and_delete = [
        InlineKeyboardButton("⬅️ Service", callback_data=f"admin:catserviceid:{service_ref}:0"),
        InlineKeyboardButton(
            "🗑 Delete service",
            callback_data=f"admin:catservicedelete:{service_ref}",
            style="danger",
        ),
    ]
    buttons.append(back_and_delete)
    fallback_buttons.append(back_and_delete)
    return (
        ".",
        InlineKeyboardMarkup(buttons),
        InlineKeyboardMarkup(fallback_buttons),
        has_premium_icons,
    )


async def admin_catalog_reply(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    text, markup = admin_catalog_services_content()
    await update.effective_message.reply_text(
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=markup,
    )


async def admin_catalog_action(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return

    parts = query.data.split(":")
    action = parts[1] if len(parts) > 1 else ""

    if action in (
        "catservicedelete",
        "catservicedeleteconfirm",
        "catservicedeletecancel",
    ):
        try:
            service_ref = int(parts[2])
        except (ValueError, IndexError):
            await query.answer("Invalid service.", show_alert=True)
            return
        service = DB.fetchone(
            "SELECT service_code,service_name FROM catalog WHERE id=?",
            (service_ref,),
        )
        if not service:
            await query.answer("Service no longer exists.", show_alert=True)
            text, markup = admin_catalog_services_content()
            with suppress(TelegramError):
                await query.edit_message_text(
                    text, parse_mode=ParseMode.HTML, reply_markup=markup
                )
            return

        service_code = str(service["service_code"])
        if action == "catservicedelete":
            count_row = DB.fetchone(
                "SELECT COUNT(DISTINCT country_id) AS total FROM catalog WHERE service_code=?",
                (service_code,),
            )
            country_count = int(count_row["total"] or 0) if count_row else 0
            await query.answer()
            await query.edit_message_text(
                (
                    f"Delete <b>{html.escape(str(service['service_name']))}</b> "
                    f"and all {country_count} country entries?\n\n"
                    "Existing order history will be kept. This removes the service "
                    "from the catalog until it is added again."
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup(
                    [
                        [
                            InlineKeyboardButton(
                                "🗑 Delete service",
                                callback_data=f"admin:catservicedeleteconfirm:{service_ref}",
                                style="danger",
                            ),
                            InlineKeyboardButton(
                                "Keep service",
                                callback_data=f"admin:catservicedeletecancel:{service_ref}",
                            ),
                        ]
                    ]
                ),
            )
            return

        if action == "catservicedeletecancel":
            await query.answer("Service kept.")
            text, markup, fallback_markup, has_premium_icons = (
                await admin_catalog_countries_content(service_code, 0)
            )
            with suppress(TelegramError):
                await edit_callback_message_with_country_fallback(
                    query,
                    text,
                    markup,
                    fallback_markup,
                    parse_mode=ParseMode.HTML,
                    has_premium_icons=has_premium_icons,
                )
            return

        try:
            with DB.lock:
                DB.conn.execute("BEGIN IMMEDIATE")
                DB.conn.execute(
                    "DELETE FROM country_favorites WHERE service_code=?",
                    (service_code,),
                )
                DB.conn.execute(
                    "DELETE FROM catalog WHERE service_code=?",
                    (service_code,),
                )
                DB.conn.commit()
        except Exception:
            with suppress(Exception):
                DB.conn.rollback()
            logger.exception("Could not delete catalog service %s", service_code)
            await query.answer("Service could not be deleted.", show_alert=True)
            return
        await query.answer("Service and all country entries deleted.")
        text, markup = admin_catalog_services_content()
        with suppress(TelegramError):
            await query.edit_message_text(
                text, parse_mode=ParseMode.HTML, reply_markup=markup
            )
        return

    # ── Navigate back to service list ──────────────────────────────────────
    if action == "catmenu":
        await query.answer()
        text, markup = admin_catalog_services_content()
        with suppress(TelegramError):
            await query.edit_message_text(
                text, parse_mode=ParseMode.HTML, reply_markup=markup
            )
        return

    # ── Open service detail / margin controls ───────────────────────────────
    if action == "catmargin":
        try:
            service_ref = int(parts[2])
        except (ValueError, IndexError):
            await query.answer("Invalid service.", show_alert=True)
            return
        service = DB.fetchone("SELECT service_code,service_name FROM catalog WHERE id=?", (service_ref,))
        if not service:
            await query.answer("Service no longer exists.", show_alert=True)
            return
        context.user_data["admin_margin_service"] = {
            "service_ref": service_ref,
            "service_code": str(service["service_code"]),
            "service_name": str(service["service_name"]),
        }
        current = DB.fetchone("SELECT markup_override_units FROM catalog WHERE service_code=? ORDER BY id LIMIT 1", (str(service["service_code"]),))
        current_text = money_usd(int(current["markup_override_units"])) if current and current["markup_override_units"] is not None else f"Default ({money_usd(DB.global_markup_units())})"
        await query.answer()
        await query.message.reply_text(
            f"⚙️ <b>Set margin: {html.escape(str(service['service_name']))}</b>\n"
            f"Current margin: <b>{html.escape(current_text)}</b>\n\n"
            "Send the new USD margin, e.g. <code>0.02</code>. Send <code>-</code> to use the global default.",
            parse_mode=ParseMode.HTML,
        )
        return

    if action == "catcountries":
        try:
            service_ref = int(parts[2]); page = int(parts[3]) if len(parts) > 3 else 0
        except (ValueError, IndexError):
            await query.answer("Invalid service.", show_alert=True)
            return
        service = DB.fetchone("SELECT service_code FROM catalog WHERE id=?", (service_ref,))
        if not service:
            await query.answer("Service no longer exists.", show_alert=True)
            return
        await query.answer()
        text, markup, fallback_markup, has_premium_icons = await admin_catalog_countries_content(str(service["service_code"]), page)
        with suppress(TelegramError):
            await edit_callback_message_with_country_fallback(query, text, markup, fallback_markup, parse_mode=ParseMode.HTML, has_premium_icons=has_premium_icons)
        return

    # ── Open a service's country list ──────────────────────────────────────
    if action == "catservice":
        await query.answer()
        service_code = parts[2] if len(parts) > 2 else ""
        text, markup, fallback_markup, has_premium_icons = (
            await admin_catalog_countries_content(service_code)
        )
        with suppress(TelegramError):
            await edit_callback_message_with_country_fallback(
                query,
                text,
                markup,
                fallback_markup,
                parse_mode=ParseMode.HTML,
                has_premium_icons=has_premium_icons,
            )
        return

    if action in ("catserviceid", "catpage"):
        try:
            service_ref = int(parts[2])
            page = int(parts[3]) if len(parts) > 3 else 0
        except (ValueError, IndexError):
            await query.answer("Invalid service.", show_alert=True)
            return
        service = DB.fetchone(
            "SELECT service_code FROM catalog WHERE id=?",
            (service_ref,),
        )
        if not service:
            await query.answer("Service no longer exists.", show_alert=True)
            return
        await query.answer()
        if action == "catserviceid":
            text, markup = admin_catalog_service_detail(service_ref)
            with suppress(TelegramError):
                await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
            return
        text, markup, fallback_markup, has_premium_icons = (
            await admin_catalog_countries_content(str(service["service_code"]), page)
        )
        with suppress(TelegramError):
            await edit_callback_message_with_country_fallback(
                query, text, markup, fallback_markup,
                parse_mode=ParseMode.HTML, has_premium_icons=has_premium_icons,
            )
        return

    # ── Toggle / Delete a catalog row ──────────────────────────────────────
    try:
        item_id = int(parts[2])
        page = int(parts[3]) if len(parts) > 3 else 0
    except (ValueError, IndexError):
        await query.answer("Invalid item.", show_alert=True)
        return

    row = DB.fetchone("SELECT * FROM catalog WHERE id=?", (item_id,))
    service_code = row["service_code"] if row else None

    if not row:
        await query.answer("Already deleted.", show_alert=True)
    elif action == "cattoggle":
        DB.execute(
            "UPDATE catalog SET enabled=CASE enabled WHEN 1 THEN 0 ELSE 1 END WHERE id=?",
            (item_id,),
        )
        await query.answer("Updated.")
    elif action == "catdelete":
        DB.execute("DELETE FROM catalog WHERE id=?", (item_id,))
        await query.answer("Deleted.")
    else:
        await query.answer("Unknown action.", show_alert=True)
        return

    # Refresh the country list for this service
    if service_code:
        text, markup, fallback_markup, has_premium_icons = (
            await admin_catalog_countries_content(service_code, page)
        )
    else:
        text, markup = admin_catalog_services_content()
    if service_code:
        with suppress(TelegramError):
            await edit_callback_message_with_country_fallback(
                query,
                text,
                markup,
                fallback_markup,
                parse_mode=ParseMode.HTML,
                has_premium_icons=has_premium_icons,
            )
    else:
        with suppress(TelegramError):
            await query.edit_message_text(
                text, parse_mode=ParseMode.HTML, reply_markup=markup
            )


async def add_catalog_start_reply(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    context.user_data["new_catalog"] = {}
    await update.effective_message.reply_text(
        "Enter the service short code, for example Instagram is ig.\n"
        "Press ❌ Cancel to stop this form and return to the main menu.",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return CAT_SERVICE_CODE


def _markup_prompt_keyboard() -> InlineKeyboardMarkup:
    """Inline keyboard shown with the markup question — one-tap Default option."""
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton(
            f"⚙️ Default  ({money_usd(DB.global_markup_units())})",
            callback_data="cat:markup:default",
        )]]
    )


async def _finalize_catalog_add(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    override: Optional[int],
) -> bool:
    """Add or complete a service across every country in the embedded table."""
    data = context.user_data.get("new_catalog", {})
    # Country IDs/names/flags are embedded in this Python file; no CSV sidecar
    # is required, so adding a service works from a single-file deployment.
    countries = smsbower_countries_from_csv()

    service_code = data.get("service_code", "")
    service_name = data.get("service_name", "")
    if not service_code or not service_name:
        return False

    inserted = 0
    updated = 0
    timestamp = now_ts()
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            for country_id, country_name in countries:
                existing = DB.conn.execute(
                    "SELECT id FROM catalog WHERE country_id=? AND service_code=?",
                    (country_id, service_code),
                ).fetchall()
                if existing:
                    DB.conn.execute(
                        """
                        UPDATE catalog
                        SET country_name=?, service_name=?, markup_override_units=?
                        WHERE country_id=? AND service_code=?
                        """,
                        (country_name, service_name, override, country_id, service_code),
                    )
                    updated += 1
                else:
                    DB.conn.execute(
                        """
                        INSERT INTO catalog(
                            country_id,country_name,service_code,service_name,
                            extra_params_json,markup_override_units,enabled,created_at
                        ) VALUES(?,?,?,?,?,?,1,?)
                        """,
                        (
                            country_id,
                            country_name,
                            service_code,
                            service_name,
                            "{}",
                            override,
                            timestamp,
                        ),
                    )
                    inserted += 1
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            raise

    context.user_data.pop("new_catalog", None)
    service_ref_row = DB.fetchone(
        "SELECT MIN(id) AS service_ref FROM catalog WHERE service_code=?",
        (service_code,),
    )
    service_ref = int(service_ref_row["service_ref"])
    live_stock_count: Optional[int] = None
    if API is not None:
        stock_result, countries_result = await asyncio.gather(
            API.get_service_country_stock(str(service_code)),
            API.get_countries(),
            return_exceptions=True,
        )
        if not isinstance(stock_result, Exception):
            stock_by_country = {
                str(country_id): max(0, int(stock))
                for country_id, stock in stock_result.items()
            }
            live_country_names = (
                {
                    str(country_id): str(name)
                    for country_id, name in countries_result
                }
                if not isinstance(countries_result, Exception)
                else {}
            )
            stock_by_country = normalize_live_stock_country_ids(
                stock_by_country, live_country_names
            )
            live_stock_count = sum(
                1 for stock in stock_by_country.values() if stock > 0
            )
            if not isinstance(countries_result, Exception):
                try:
                    sync_live_service_countries(
                        service_ref,
                        stock_by_country,
                        live_country_names,
                    )
                except Exception:
                    logger.exception(
                        "Could not sync live SMSBower countries for service %s",
                        service_code,
                    )
        else:
            logger.warning(
                "Initial live stock check failed for service %s: %s",
                service_code,
                short_text(str(stock_result), 180),
            )

    msg = update.effective_message or (
        update.callback_query.message if update.callback_query else None
    )
    if msg:
        availability_note = (
            f"📦 In stock right now: <b>{live_stock_count}</b> countries."
            if live_stock_count is not None
            else "📦 Live availability will be checked when the country list opens."
        )
        await msg.reply_text(
            f"✅ Service <b>{html.escape(service_name)}</b> saved.\n"
            f"🌍 Countries configured: <b>{len(countries)}</b>\n"
            f"➕ New country entries: <b>{inserted}</b> • Updated existing: <b>{updated}</b>\n"
            f"{availability_note}\n"
            "The country list refreshes live stock whenever it is opened or refreshed.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                [[
                    InlineKeyboardButton(
                        "🌍 View country list",
                        callback_data=f"admin:catserviceid:{service_ref}:0",
                    )
                ]]
            ),
        )
    return True


async def add_catalog_service_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = (update.effective_message.text or "").strip()
    if not value or len(value) > 80:
        await update.effective_message.reply_text("Enter a valid service name (max 80 chars).")
        return CAT_SERVICE_NAME
    context.user_data["new_catalog"]["service_name"] = value
    context.user_data["new_catalog"]["extra_params_json"] = "{}"
    await update.effective_message.reply_text(
        f"⚙️ <b>Service markup / margin</b>\n\n"
        f"প্রতি নাম্বারে USD markup লিখুন, যেমন <code>0.005</code> বা <code>0.02</code>।\n"
        f"অথবা <b>Default</b> বাটনে ক্লিক করলে global markup "
        f"(<b>{money_usd(DB.global_markup_units())}</b>) সেট হবে।",
        parse_mode=ParseMode.HTML,
        reply_markup=_markup_prompt_keyboard(),
    )
    return CAT_MARKUP


async def add_catalog_markup_default(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    """Callback: admin tapped the Default markup button."""
    query = update.callback_query
    if not query:
        return CAT_MARKUP
    await query.answer(f"Default markup: {money_usd(DB.global_markup_units())}")
    return (
        ConversationHandler.END
        if await _finalize_catalog_add(update, context, override=None)
        else CAT_MARKUP
    )


async def add_catalog_markup(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle typed markup value. Also guards against accidental non-numeric input."""
    raw = (update.effective_message.text or "").strip()

    # Treat "-" or empty as global default
    if raw in ("-", ""):
        return (
            ConversationHandler.END
            if await _finalize_catalog_add(update, context, override=None)
            else CAT_MARKUP
        )

    try:
        override = usd_to_units(raw)
        if override < 0:
            raise ValueError("negative")
    except Exception:
        await update.effective_message.reply_text(
            "❌ Invalid markup. USD amount লিখুন (যেমন <code>0.005</code>), "
            "অথবা Default বাটনে ক্লিক করুন।",
            parse_mode=ParseMode.HTML,
            reply_markup=_markup_prompt_keyboard(),
        )
        return CAT_MARKUP

    return (
        ConversationHandler.END
        if await _finalize_catalog_add(update, context, override=override)
        else CAT_MARKUP
    )



async def set_markup_start_reply(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    await update.effective_message.reply_text(
        f"Current global markup: {money_usd(DB.global_markup_units())}\n"
        "Enter the new markup in USD, for example 0.005.",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_MARKUP


async def set_markup_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        units = usd_to_units((update.effective_message.text or "").strip())
        if units < 0:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text("Enter a valid non-negative USD amount.")
        return SET_MARKUP
    DB.set_setting("markup_usd", str(units_to_decimal(units)))
    await update.effective_message.reply_text(
        f"✅ Global markup changed to {money_usd(units)}.", reply_markup=admin_menu()
    )
    return ConversationHandler.END



async def set_exchange_rate_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    await update.effective_message.reply_text(
        f"Current exchange rate: 1 USD = ৳{exchange_rate():.2f}\n"
        "Enter the new BDT rate.",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_EXCHANGE_RATE


async def set_exchange_rate_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        value = decimal_value((update.effective_message.text or "").strip())
        if value <= 0 or value > 10000:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text("Enter a valid positive rate, for example 125.")
        return SET_EXCHANGE_RATE
    DB.set_setting("usd_bdt_rate", str(value))
    await update.effective_message.reply_text(
        f"✅ Exchange rate saved: 1 USD = ৳{value:.2f}", reply_markup=admin_menu()
    )
    return ConversationHandler.END



async def set_deposit_group_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    current = DB.get_setting("deposit_group_id", "") or "Not set"
    await update.effective_message.reply_text(
        f"Current deposit group ID: {current}\n"
        "Send the group chat ID, for example -1001234567890.",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_DEPOSIT_GROUP


async def set_deposit_group_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw = (update.effective_message.text or "").strip()
    if not re.fullmatch(r"-?\d{5,20}", raw):
        await update.effective_message.reply_text("Enter a valid numeric Telegram group chat ID.")
        return SET_DEPOSIT_GROUP
    DB.set_setting("deposit_group_id", raw)
    await update.effective_message.reply_text(
        f"✅ Deposit group ID saved: {raw}", reply_markup=admin_group_settings_menu()
    )
    return ConversationHandler.END


def style_workbook(ws: Any, headers: list[str]) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, header in enumerate(headers, start=1):
        width = min(max(len(header) + 4, 14), 34)
        ws.column_dimensions[get_column_letter(idx)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def create_payment_list_xlsx() -> str:
    rows = DB.fetchall(
        """
        SELECT u.user_id,u.username,u.first_name,u.balance_units,u.created_at,
               COALESCE((SELECT SUM(d.amount_units) FROM deposits d
                         WHERE d.user_id=u.user_id AND d.status='approved'),0) AS total_deposit,
               (
                   COALESCE((SELECT COUNT(*) FROM orders o WHERE o.user_id=u.user_id),0)
                   + COALESCE((SELECT COUNT(*) FROM local_orders lo WHERE lo.user_id=u.user_id),0)
               ) AS total_orders
        FROM users u
        ORDER BY u.created_at ASC
        """
    )
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.append(["Metric", "Value"])
    summary.append(["Total Users", len(rows)])
    summary.append(["Total Current Balance (BDT)", float(sum(units_to_decimal(int(row["balance_units"] or 0)) for row in rows))])
    summary.append(["Total Approved Deposits (BDT)", float(sum(units_to_decimal(int(row["total_deposit"] or 0)) for row in rows))])
    summary["B3"].number_format = '৳#,##0.00'
    summary["B4"].number_format = '৳#,##0.00'
    style_workbook(summary, ["Metric", "Value"])

    ws = wb.create_sheet("Users")
    headers = ["Username", "Chat ID", "Current Balance (BDT)", "Total Deposit (BDT)", "Total Orders", "Joined At"]
    ws.append(headers)
    for row in rows:
        username = f"@{row['username']}" if row["username"] else (row["first_name"] or "")
        ws.append([
            username,
            str(row["user_id"]),
            float(units_to_decimal(int(row["balance_units"] or 0))),
            float(units_to_decimal(int(row["total_deposit"] or 0))),
            int(row["total_orders"] or 0),
            format_timestamp(row["created_at"]),
        ])
    for cell in ws["C"][1:] + ws["D"][1:]:
        cell.number_format = '৳#,##0.00'
    style_workbook(ws, headers)
    path = os.path.join(tempfile.gettempdir(), f"payment_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(path)
    return path


def create_user_order_history_xlsx(user_id: int) -> str:
    summary_data = user_account_summary(user_id)

    global_rows = DB.fetchall(
        """
        SELECT *
        FROM orders
        WHERE user_id=?
        ORDER BY id DESC
        """,
        (user_id,),
    )
    local_rows = DB.fetchall(
        """
        SELECT lo.*,
               COALESCE(
                   (SELECT COUNT(*)
                    FROM local_otp_history h
                    WHERE h.order_id=lo.id),
                   0
               ) AS otp_count
        FROM local_orders lo
        WHERE lo.user_id=?
        ORDER BY lo.id DESC
        """,
        (user_id,),
    )

    wb = Workbook()

    summary_ws = wb.active
    summary_ws.title = "Summary"
    summary_headers = ["Metric", "Value"]
    summary_ws.append(summary_headers)
    summary_ws.append(
        ["Current Balance (BDT)", float(units_to_decimal(summary_data["current_balance"]))]
    )
    summary_ws.append(["Total Numbers Purchased", summary_data["order_count"]])
    summary_ws.append(
        ["Total Purchase Amount (BDT)", float(units_to_decimal(summary_data["purchase_units"]))]
    )
    summary_ws.append(["Refunded Numbers", summary_data["refund_count"]])
    summary_ws.append(
        ["Total Refund Amount (BDT)", float(units_to_decimal(summary_data["refund_units"]))]
    )
    for cell in ("B2", "B4", "B6"):
        summary_ws[cell].number_format = '৳#,##0.00'
    style_workbook(summary_ws, summary_headers)

    global_ws = wb.create_sheet("Global Orders")
    global_headers = [
        "Order ID",
        "Country",
        "Service",
        "Phone Number",
        "Price (BDT)",
        "Status",
        "OTP",
        "Purchased At",
        "Updated At",
        "Expires At",
        "Refunded",
    ]
    global_ws.append(global_headers)
    for row in global_rows:
        global_ws.append(
            [
                row["id"],
                row["country_name"],
                row["service_name"],
                row["phone"],
                float(units_to_decimal(int(row["charged_units"] or 0))),
                status_label(str(row["status"])),
                row["otp_code"] or "",
                format_timestamp(row["created_at"]),
                format_timestamp(row["updated_at"]),
                format_timestamp(row["expires_at"]),
                "Yes" if int(row["refunded"] or 0) else "No",
            ]
        )
    for cell in global_ws["E"][1:]:
        cell.number_format = '৳#,##0.00'
    style_workbook(global_ws, global_headers)

    local_ws = wb.create_sheet("USA Local")
    local_headers = [
        "Local ID",
        "Phone Number",
        "Price (BDT)",
        "Status",
        "Latest OTP",
        "OTP Count",
        "Purchased At",
        "Updated At",
        "Expires At",
        "Refunded",
    ]
    local_ws.append(local_headers)
    for row in local_rows:
        local_ws.append(
            [
                row["id"],
                row["phone"],
                float(units_to_decimal(int(row["charged_units"] or 0))),
                status_label(str(row["status"])),
                row["latest_otp"] or "",
                int(row["otp_count"] or 0),
                format_timestamp(row["created_at"]),
                format_timestamp(row["updated_at"]),
                format_timestamp(row["expires_at"]),
                "Yes" if int(row["refunded"] or 0) else "No",
            ]
        )
    for cell in local_ws["C"][1:]:
        cell.number_format = '৳#,##0.00'
    style_workbook(local_ws, local_headers)

    path = os.path.join(
        tempfile.gettempdir(),
        f"my_orders_{user_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )
    wb.save(path)
    return path


def create_order_logs_xlsx() -> str:
    rows = DB.fetchall(
        """
        SELECT o.*,u.username,u.first_name
        FROM orders o LEFT JOIN users u ON u.user_id=o.user_id
        ORDER BY o.id DESC
        """
    )
    wb = Workbook()
    ws = wb.active
    ws.title = "Order Logs"
    headers = [
        "Order ID", "Username", "Chat ID", "Country", "Service", "Phone Number",
        "Provider ID", "Price (BDT)", "Provider Cost (USD)", "Status", "OTP",
        "Purchased At", "Updated At", "Expires At", "Refunded"
    ]
    ws.append(headers)
    for row in rows:
        username = f"@{row['username']}" if row["username"] else (row["first_name"] or "")
        ws.append([
            row["id"], username, str(row["user_id"]), row["country_name"], row["service_name"],
            row["phone"], row["provider_id"] or "",
            float(units_to_decimal(int(row["charged_units"] or 0))),
            float(units_to_decimal(int(row["provider_cost_units"] or 0))),
            row["status"], row["otp_code"] or "", format_timestamp(row["created_at"]),
            format_timestamp(row["updated_at"]), format_timestamp(row["expires_at"]),
            "Yes" if row["refunded"] else "No",
        ])
    for cell in ws["H"][1:]:
        cell.number_format = '৳#,##0.00'
    for cell in ws["I"][1:]:
        cell.number_format = '$0.000000'
    style_workbook(ws, headers)

    local_ws = wb.create_sheet("USA Local")
    local_headers = [
        "Local ID", "Username", "Chat ID", "Phone Number", "Provider Order ID",
        "Price (BDT)", "Status", "Latest OTP", "OTP Count", "Purchased At",
        "Updated At", "Expires At", "Refunded"
    ]
    local_ws.append(local_headers)
    local_rows = DB.fetchall(
        """
        SELECT lo.*,u.username,u.first_name,
               COALESCE((SELECT COUNT(*) FROM local_otp_history h
                         WHERE h.order_id=lo.id),0) AS otp_count
        FROM local_orders lo
        LEFT JOIN users u ON u.user_id=lo.user_id
        ORDER BY lo.id DESC
        """
    )
    for row in local_rows:
        username = f"@{row['username']}" if row["username"] else (row["first_name"] or "")
        local_ws.append([
            row["id"], username, str(row["user_id"]), row["phone"],
            row["provider_order_id"],
            float(units_to_decimal(int(row["charged_units"] or 0))),
            row["status"], row["latest_otp"] or "", int(row["otp_count"] or 0),
            format_timestamp(row["created_at"]), format_timestamp(row["updated_at"]),
            format_timestamp(row["expires_at"]),
            "Yes" if row["refunded"] else "No",
        ])
    for cell in local_ws["F"][1:]:
        cell.number_format = '৳#,##0.00'
    style_workbook(local_ws, local_headers)

    path = os.path.join(tempfile.gettempdir(), f"order_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(path)
    return path


def collect_countries(node: Any, output: dict[str, str], parent_key: Optional[str] = None) -> None:
    if isinstance(node, list):
        for item in node:
            collect_countries(item, output, parent_key)
        return
    if not isinstance(node, dict):
        if parent_key and str(parent_key).isdigit() and isinstance(node, str):
            output[str(parent_key)] = node
        return
    country_id = node.get("id", node.get("country_id", node.get("countryId")))
    country_name = node.get("name", node.get("country", node.get("country_name", node.get("title"))))
    if country_id is not None and country_name:
        output[str(country_id)] = str(country_name)
    for key, value in node.items():
        if str(key).isdigit() and isinstance(value, str):
            output[str(key)] = value
        else:
            collect_countries(value, output, str(key))


async def create_country_list_xlsx() -> str:
    data, text = await API.request("getCountries")
    if isinstance(data, str):
        raise RuntimeError(API.human_api_error(text))
    countries: dict[str, str] = {}
    collect_countries(data, countries)
    if not countries:
        raise RuntimeError(f"Could not parse country list: {short_text(text, 250)}")
    wb = Workbook()
    ws = wb.active
    ws.title = "Country List"
    headers = ["Country Name", "Country ID"]
    ws.append(headers)
    for country_id, name in sorted(countries.items(), key=lambda item: item[1].lower()):
        ws.append([name, country_id])
    style_workbook(ws, headers)
    path = os.path.join(tempfile.gettempdir(), f"country_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(path)
    return path


async def send_xlsx(
    chat_id: int,
    path: str,
    caption: str,
    bot: Any,
    filename: Optional[str] = None,
) -> None:
    """Send an XLSX using a clean user-facing name while keeping temp files unique."""
    try:
        with open(path, "rb") as handle:
            await bot.send_document(
                chat_id=chat_id,
                document=handle,
                filename=filename or os.path.basename(path),
                caption=caption,
            )
    finally:
        with suppress(OSError):
            os.remove(path)


async def admin_country_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    status = await update.effective_message.reply_text("Loading the latest country list…")
    try:
        path = await create_country_list_xlsx()
        await send_xlsx(
            update.effective_chat.id,
            path,
            "Latest API country list",
            context.bot,
            filename="country_list.xlsx",
        )
        await status.delete()
    except Exception as exc:
        await status.edit_text(f"Could not create country list: {exc}")


async def add_balance_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    if len(context.args) != 2:
        await update.effective_message.reply_text("Usage: /addbalance USER_ID BDT_AMOUNT")
        return
    try:
        user_id = int(context.args[0])
        units = bdt_to_units(context.args[1])
        if units <= 0:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text("Invalid user ID or BDT amount.")
        return
    user = DB.fetchone("SELECT * FROM users WHERE user_id=?", (user_id,))
    if not user:
        await update.effective_message.reply_text("User not found. They must start the bot first.")
        return
    DB.credit_balance(user_id, units, "admin_credit", str(update.effective_user.id), "Manual admin credit")
    await update.effective_message.reply_text(f"✅ Added {money(units)} to user {user_id}.")
    with suppress(TelegramError):
        await context.bot.send_message(user_id, f"✅ Admin added {money(units)} to your balance.")


async def daily_order_report_loop(application: Application) -> None:
    while True:
        await asyncio.sleep(24 * 60 * 60)
        for admin_id in sorted(ADMIN_IDS):
            try:
                path = await run_file_job(create_order_logs_xlsx)
                await send_xlsx(
                    admin_id,
                    path,
                    "Automatic 24-hour order report",
                    application.bot,
                    filename="daily_order_report.xlsx",
                )
            except Exception:
                logger.exception("Could not send daily order report to %s", admin_id)



async def replace_callback_message(
    query: Any, bot: Any, text: str, *, reply_markup: Any = None, parse_mode: Any = None
) -> Any:
    """Delete the old inline message and send a fresh message instead of editing it."""
    chat_id = query.message.chat_id if query and query.message else None
    if query and query.message:
        with suppress(TelegramError):
            await query.message.delete()
    if chat_id is None:
        return None
    kwargs = {"chat_id": chat_id, "text": text, "reply_markup": reply_markup}
    if parse_mode is not None:
        kwargs["parse_mode"] = parse_mode
    return await bot.send_message(**kwargs)


async def replace_callback_message_with_country_fallback(
    query: Any,
    bot: Any,
    text: str,
    premium_markup: Any,
    fallback_markup: Any,
    *,
    parse_mode: Any = None,
    has_premium_icons: bool = False,
) -> Any:
    """Use standard flags for non-Premium users and fall back if icons are rejected."""
    query_user = getattr(query, "from_user", None)
    user_has_premium = bool(getattr(query_user, "is_premium", False))
    if not user_has_premium:
        # Telegram exposes the requesting user's Premium status on CallbackQuery.
        # Keep country selection fully usable without Premium by sending the
        # Unicode-flag keyboard directly instead of attempting custom icons.
        premium_markup = fallback_markup
        has_premium_icons = False

    try:
        return await replace_callback_message(
            query, bot, text, reply_markup=premium_markup, parse_mode=parse_mode
        )
    except BadRequest:
        if not has_premium_icons:
            raise
        logger.info("Telegram rejected custom country button icons; using Unicode flags.")
        return await replace_callback_message(
            query, bot, text, reply_markup=fallback_markup, parse_mode=parse_mode
        )


async def edit_callback_message_with_country_fallback(
    query: Any,
    text: str,
    premium_markup: Any,
    fallback_markup: Any,
    *,
    parse_mode: Any = None,
    has_premium_icons: bool = False,
) -> Any:
    """Edit the current catalog message while retaining Unicode flag fallback."""
    query_user = getattr(query, "from_user", None)
    if not bool(getattr(query_user, "is_premium", False)):
        premium_markup = fallback_markup
        has_premium_icons = False

    kwargs = {"text": text, "reply_markup": premium_markup}
    if parse_mode is not None:
        kwargs["parse_mode"] = parse_mode
    try:
        return await query.edit_message_text(**kwargs)
    except BadRequest:
        if not has_premium_icons:
            raise
        logger.info("Telegram rejected custom catalog icons; editing with Unicode flags.")
        kwargs["reply_markup"] = fallback_markup
        return await query.edit_message_text(**kwargs)


async def cancel_conversation(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancel any partial form, erase its temporary state, and restore Main Menu."""
    user_id = update.effective_user.id if update.effective_user else None
    clear_transient_workflow(context)
    # Cancel always exits the admin/user sub-flow completely. Admins still see
    # the Admin Panel key inside their normal main menu.
    context.user_data.pop("admin_mode", None)

    message = update.effective_message
    if update.callback_query:
        with suppress(TelegramError):
            await update.callback_query.answer("Cancelled.")
        message = update.callback_query.message
        await replace_callback_message(
            update.callback_query, context.bot,
            "✅ Current action cancelled. Main menu restored.",
            reply_markup=main_menu(user_id),
        )
        return ConversationHandler.END

    if message:
        await message.reply_text(
            "✅ Current action cancelled. Main menu restored.",
            reply_markup=main_menu(user_id),
        )
    return ConversationHandler.END


async def unknown_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Panel API key entry must be consumed before other admin text workflows.
    if await panel_api_key_input(update, context):
        return
    # Admin margin input is handled here so no other conversation or menu behavior changes.
    pending_margin = context.user_data.get("admin_margin_service")
    if pending_margin and update.effective_user and is_admin(update.effective_user.id):
        raw = (update.effective_message.text or "").strip()
        try:
            if raw == "-":
                units = None
            else:
                units = usd_to_units(raw)
                if units < 0:
                    raise ValueError("negative margin")
        except Exception:
            await update.effective_message.reply_text(
                "❌ Invalid margin. Send a USD amount such as 0.02, or send - for the global default."
            )
            return
        service_code = str(pending_margin["service_code"])
        try:
            with DB.lock:
                DB.conn.execute(
                    "UPDATE catalog SET markup_override_units=? WHERE service_code=?",
                    (units, service_code),
                )
                DB.conn.commit()
            context.user_data.pop("admin_margin_service", None)
            service_ref = int(pending_margin["service_ref"])
            text, markup = admin_catalog_service_detail(service_ref)
            await update.effective_message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=markup)
        except Exception:
            logger.exception("Could not update catalog margin for service %s", service_code)
            await update.effective_message.reply_text("❌ Could not save margin. Please check the error log and try again.")
        return
    if not await require_terms(update):
        return
    await update.effective_message.reply_text(
        "Use the menu buttons below.", reply_markup=main_menu(update.effective_user.id if update.effective_user else None)
    )



# =============================================================================
# MINIMUM DEPOSIT / USER MANAGER / OTP GROUPS / PROVIDER PAGINATION
# =============================================================================

PROVIDER_PAGE_SIZE = 10
SET_MINIMUM_DEPOSIT = 109
ADMIN_USER_LOOKUP_ID = 110
ADMIN_USER_BALANCE_AMOUNT = 111
SET_OTP_FORWARD_GROUP = 112
SET_LOCAL_PRICE = 113
SET_LOCAL_SESSION = 114
DB_RESTORE_FILE = 115
DB_RESTORE_CONFIRM = 116
# Keep these feature states outside the older database/broadcast state range
# below. Reusing those numbers sends the admin's amount/name/video to the
# wrong ConversationHandler state, so settings appear unsaved.
SET_REFERRAL_REWARD = 124
SET_REFERRAL_THRESHOLD = 125
SET_HOW_BUTTON_NAME = 126
SET_HOW_VIDEO_FILE = 127


def ensure_management_schema() -> None:
    """Add settings/flags required by the latest admin features."""
    with DB.lock:
        order_columns = {row["name"] for row in DB.conn.execute("PRAGMA table_info(orders)")}
        for column, ddl in (
            (
                "otp_public_forwarded",
                "ALTER TABLE orders ADD COLUMN otp_public_forwarded INTEGER NOT NULL DEFAULT 0",
            ),
            (
                "otp_private_forwarded",
                "ALTER TABLE orders ADD COLUMN otp_private_forwarded INTEGER NOT NULL DEFAULT 0",
            ),
            (
                "ever_received_otp",
                "ALTER TABLE orders ADD COLUMN ever_received_otp INTEGER NOT NULL DEFAULT 0",
            ),
        ):
            if column not in order_columns:
                DB.conn.execute(ddl)

        for key, value in (
            ("minimum_deposit_bdt", "0"),
            ("otp_forward_group_1", ""),
            ("otp_admin_group_2", ""),
        ):
            DB.conn.execute(
                "INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (key, value)
            )
        DB.conn.commit()


ensure_management_schema()


def minimum_deposit_units() -> int:
    try:
        value = bdt_to_units(DB.get_setting("minimum_deposit_bdt", "0"))
        return max(0, value)
    except Exception:
        return 0


CANCEL_BUTTON_TEXT = "❌ Cancel"

# Only temporary workflow values are removed here. cancel_conversation also
# clears admin_mode because every cancellation must return to the main menu.
TRANSIENT_WORKFLOW_KEYS = {
    "new_catalog",
    "balance_action",
    "balance_user_id",
    "otp_group_setting",
    "country_search_service_ref",
    "country_search_page",
    "country_search_prompt_id",
}


def clear_transient_workflow(context: ContextTypes.DEFAULT_TYPE) -> None:
    pending_database = context.user_data.pop("db_restore_candidate", None)
    context.user_data.pop("db_restore_filename", None)
    if pending_database:
        with suppress(FileNotFoundError):
            os.remove(pending_database)

    for key in TRANSIENT_WORKFLOW_KEYS:
        context.user_data.pop(key, None)


def begin_workflow(
    context: ContextTypes.DEFAULT_TYPE, *, admin_mode: Optional[bool] = None
) -> None:
    """Start a clean workflow without leaving stale partial form data behind."""
    clear_transient_workflow(context)
    if admin_mode is True:
        context.user_data["admin_mode"] = True
    elif admin_mode is False:
        context.user_data.pop("admin_mode", None)


def workflow_keyboard(
    user_id: Optional[int], *, admin: bool = False
) -> ReplyKeyboardMarkup:
    """During a multi-step form, hide every menu key and show only Cancel."""
    # user_id/admin are intentionally accepted for compatibility with all
    # existing workflow calls. A workflow must have one clear escape route and
    # no unrelated menu buttons that can leave stale state behind.
    return ReplyKeyboardMarkup(
        [[KeyboardButton(CANCEL_BUTTON_TEXT, style="danger")]],
        resize_keyboard=True,
        one_time_keyboard=False,
        input_field_placeholder="Complete the current step or press Cancel",
    )


def current_workflow_keyboard(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> ReplyKeyboardMarkup:
    user_id = update.effective_user.id if update.effective_user else None
    admin_view = bool(context.user_data.get("admin_mode")) and is_admin(user_id)
    return workflow_keyboard(user_id, admin=admin_view)


async def noop_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.callback_query:
        await update.callback_query.answer()


async def show_provider_prices(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Show provider prices in a two-column, cheapest-first, paginated grid."""
    query = update.callback_query
    if not query or not await require_terms(update):
        return
    await query.answer("Loading live prices…")

    try:
        parts = query.data.split(":")
        item_id = int(parts[2])
        requested_page = int(parts[3]) if len(parts) > 3 else 0
    except Exception:
        return

    item = get_catalog_item(item_id)
    if not item or not item["enabled"]:
        await query.edit_message_text("This country/service is no longer available.")
        return

    try:
        offers = await API.get_provider_offers(
            item["country_id"], item["service_code"], catalog_extra(item)
        )
    except Exception as exc:
        await query.edit_message_text(
            f"Could not load live prices: <code>{html.escape(str(exc))}</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Back", callback_data="buy:menu")]]
            ),
        )
        return

    active = sorted(
        (offer for offer in offers if offer.available > 0),
        key=lambda offer: (
            int(offer.provider_cost_units),
            -int(offer.available),
            int(offer.provider_id) if str(offer.provider_id).isdigit() else str(offer.provider_id),
        ),
    )
    if not active:
        await query.edit_message_text(
            "No provider currently has numbers in stock. Refresh to check again.",
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("🔄 Refresh", callback_data=f"buy:service:{item_id}")],
                    [InlineKeyboardButton("⬅️ Back", callback_data="buy:menu")],
                ]
            ),
        )
        return

    total_pages = max(1, (len(active) + PROVIDER_PAGE_SIZE - 1) // PROVIDER_PAGE_SIZE)
    page = min(max(0, requested_page), total_pages - 1)
    start = page * PROVIDER_PAGE_SIZE
    page_offers = active[start : start + PROVIDER_PAGE_SIZE]
    markup_units = item_markup_units(item)

    provider_buttons: list[InlineKeyboardButton] = []
    for offer in page_offers:
        total_bdt = usd_units_to_bdt_units(offer.provider_cost_units + markup_units)
        stock = str(offer.available)
        label = f"{money(total_bdt)} • {stock}"
        provider_buttons.append(
            InlineKeyboardButton(
                short_text(label, 34),
                callback_data=f"buy:provider:{item_id}:{offer.provider_id}",
            )
        )

    rows: list[list[InlineKeyboardButton]] = [
        provider_buttons[index : index + 2]
        for index in range(0, len(provider_buttons), 2)
    ]

    nav = [
        InlineKeyboardButton(
            "⬅️",
            callback_data=f"buy:service:{item_id}:{page - 1}" if page > 0 else "noop",
        ),
        InlineKeyboardButton(
            f"📄 {page + 1}/{total_pages}", callback_data="noop"
        ),
        InlineKeyboardButton(
            "➡️",
            callback_data=(
                f"buy:service:{item_id}:{page + 1}"
                if page + 1 < total_pages
                else "noop"
            ),
        ),
    ]
    rows.append(nav)
    is_favorite = bool(
        DB.fetchone(
            """
            SELECT 1 FROM country_favorites
            WHERE user_id=? AND service_code=? AND country_id=?
            """,
            (int(query.from_user.id), str(item["service_code"]), str(item["country_id"])),
        )
    )
    rows.append(
        [
            InlineKeyboardButton(
                "⭐ Remove favorite" if is_favorite else "☆ Add to favorites",
                callback_data=f"buy:favorite:{item_id}",
            )
        ]
    )
    rows.append(
        [
            InlineKeyboardButton(
                "🔄 Refresh", callback_data=f"buy:service:{item_id}:{page}"
            ),
            InlineKeyboardButton(
                "⬅️ Services", callback_data="buy:menu"
            ),
        ]
    )

    await query.edit_message_text(
        f"{_country_tg_emoji(item['country_name'], item['country_id'])} {_service_tg_emoji(item['service_name'])} <b>{html.escape(item['country_name'])} — {html.escape(item['service_name'])}</b>\n"
        "Cheapest in-stock providers appear first. Each button shows price and stock.",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(rows),
    )


async def toggle_country_favorite(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user or not await require_terms(update):
        return
    try:
        item_id = int(str(query.data).split(":")[2])
    except (IndexError, ValueError):
        await query.answer("Invalid country.", show_alert=True)
        return

    item = get_catalog_item(item_id)
    if not item or not item["enabled"]:
        await query.answer("This country/service is no longer available.", show_alert=True)
        return

    favorite_key = (
        int(user.id),
        str(item["service_code"]),
        str(item["country_id"]),
    )
    existing = DB.fetchone(
        """
        SELECT 1 FROM country_favorites
        WHERE user_id=? AND service_code=? AND country_id=?
        """,
        favorite_key,
    )
    if existing:
        DB.execute(
            """
            DELETE FROM country_favorites
            WHERE user_id=? AND service_code=? AND country_id=?
            """,
            favorite_key,
        )
        is_favorite = False
        answer = "Removed from favorites."
    else:
        DB.execute(
            """
            INSERT OR IGNORE INTO country_favorites(
                user_id,service_code,country_id,created_at
            ) VALUES(?,?,?,?)
            """,
            (*favorite_key, now_ts()),
        )
        is_favorite = True
        answer = "Added to favorites."

    await query.answer(answer)
    if not query.message or not query.message.reply_markup:
        return
    target_callback = f"buy:favorite:{item_id}"
    new_rows: list[list[InlineKeyboardButton]] = []
    for old_row in query.message.reply_markup.inline_keyboard:
        new_row: list[InlineKeyboardButton] = []
        for button in old_row:
            if button.callback_data == target_callback:
                new_row.append(
                    InlineKeyboardButton(
                        "⭐ Remove favorite" if is_favorite else "☆ Add to favorites",
                        callback_data=target_callback,
                    )
                )
            else:
                new_row.append(button)
        new_rows.append(new_row)
    try:
        await query.edit_message_reply_markup(
            reply_markup=InlineKeyboardMarkup(new_rows)
        )
    except TelegramError:
        logger.debug("Favorite state saved; provider page markup refresh was skipped.")



async def set_minimum_deposit_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    current = minimum_deposit_units()
    await update.effective_message.reply_text(
        f"Current minimum deposit: <b>{money(current)}</b>\n"
        "Enter the new minimum amount in BDT. Enter 0 to disable the minimum.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_MINIMUM_DEPOSIT


async def set_minimum_deposit_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    raw = (update.effective_message.text or "").strip()
    try:
        value = decimal_value(raw)
        if value < 0:
            raise ValueError
        units = bdt_to_units(value)
    except Exception:
        await update.effective_message.reply_text(
            "Enter a valid amount, for example 20 or 100. Enter 0 to disable it."
        )
        return SET_MINIMUM_DEPOSIT

    DB.set_setting("minimum_deposit_bdt", str(value))
    await update.effective_message.reply_text(
        f"✅ Minimum deposit set to {money(units)}.", reply_markup=admin_menu()
    )
    return ConversationHandler.END






def user_profile_row(user_id: int) -> Optional[sqlite3.Row]:
    start_ts, end_ts, _ = _dhaka_day_bounds()
    return DB.fetchone(
        """
        SELECT u.*,
               COALESCE((
                   SELECT SUM(d.amount_units) FROM deposits d
                   WHERE d.user_id=u.user_id AND d.status='approved'
               ),0) AS total_deposit_units,
               COALESCE((
                   SELECT COUNT(*) FROM orders o WHERE o.user_id=u.user_id
               ),0) AS total_orders,
               COALESCE((
                   SELECT COUNT(*) FROM orders o
                   WHERE o.user_id=u.user_id AND o.status='otp_received'
               ),0) AS successful_orders,
               COALESCE((
                   SELECT COUNT(*) FROM orders o
                   WHERE o.user_id=u.user_id AND o.status='otp_received'
                     AND o.updated_at>=? AND o.updated_at<?
               ),0) AS today_otp_count,
               COALESCE((
                   SELECT SUM(o.charged_units) FROM orders o
                   WHERE o.user_id=u.user_id AND o.status IN ('otp_received','expired')
                     AND o.refunded=0 AND o.created_at>=? AND o.created_at<?
               ),0) AS today_spent_units
        FROM users u WHERE u.user_id=?
        """,
        (start_ts, end_ts, start_ts, end_ts, user_id),
    )


async def admin_user_balance_action_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    query = update.callback_query
    if not query or not await require_admin(update):
        return ConversationHandler.END
    await query.answer()
    try:
        _, action, raw_user_id = query.data.split(":", 2)
        user_id = int(raw_user_id)
    except Exception:
        return ConversationHandler.END
    if not user_profile_row(user_id):
        await query.answer("User not found.", show_alert=True)
        return ConversationHandler.END
    context.user_data["balance_action"] = action
    context.user_data["balance_user_id"] = user_id
    verb = "add to" if action == "add" else "subtract from"
    await query.message.reply_text(
        f"Enter the BDT amount to {verb} user <code>{user_id}</code>.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return ADMIN_USER_BALANCE_AMOUNT


def adjust_user_balance(
    user_id: int, delta_units: int, admin_id: int
) -> tuple[bool, int, str]:
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            row = DB.conn.execute(
                "SELECT balance_units FROM users WHERE user_id=?", (user_id,)
            ).fetchone()
            if not row:
                DB.conn.rollback()
                return False, 0, "User not found."
            current = int(row["balance_units"] or 0)
            new_balance = current + int(delta_units)
            if new_balance < 0:
                DB.conn.rollback()
                return False, current, "The user does not have enough balance."
            ts = now_ts()
            DB.conn.execute(
                "UPDATE users SET balance_units=?,updated_at=? WHERE user_id=?",
                (new_balance, ts, user_id),
            )
            kind = "admin_credit" if delta_units >= 0 else "admin_debit"
            note = f"Balance adjusted by admin {admin_id}"
            DB.conn.execute(
                """
                INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at)
                VALUES(?,?,?,?,?,?)
                """,
                (user_id, delta_units, kind, str(admin_id), note, ts),
            )
            DB.conn.commit()
            return True, new_balance, ""
        except Exception:
            DB.conn.rollback()
            raise


async def admin_user_balance_amount(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    try:
        amount_units = bdt_to_units((update.effective_message.text or "").strip())
        if amount_units <= 0:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text("Enter a valid amount greater than 0.")
        return ADMIN_USER_BALANCE_AMOUNT

    action = context.user_data.pop("balance_action", "")
    user_id = int(context.user_data.pop("balance_user_id", 0))
    delta = amount_units if action == "add" else -amount_units
    success, new_balance, error = adjust_user_balance(
        user_id, delta, update.effective_user.id
    )
    if not success:
        await update.effective_message.reply_text(error, reply_markup=admin_menu())
        return ConversationHandler.END

    sign_text = "added to" if delta > 0 else "subtracted from"
    row = user_profile_row(user_id)
    await update.effective_message.reply_text(
        f"✅ {money(amount_units)} was {sign_text} the user's balance.\n"
        f"New balance: <b>{money(new_balance)}</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=admin_menu(),
    )
    if row:
        await update.effective_message.reply_text(
            render_user_profile(row),
            parse_mode=ParseMode.HTML,
            reply_markup=user_profile_keyboard(user_id),
        )
    with suppress(TelegramError):
        await context.bot.send_message(
            user_id,
            f"Your balance was adjusted by an administrator.\n"
            f"Change: {'+' if delta > 0 else '-'}{money(amount_units)}\n"
            f"Current balance: {money(new_balance)}",
        )
    return ConversationHandler.END



async def set_otp_group_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    label = (update.effective_message.text or "").strip()
    if label == "📨 OTP Group 1":
        setting_key = "otp_forward_group_1"
        description = "public OTP forwarding group"
    else:
        setting_key = "otp_admin_group_2"
        description = "private admin OTP group"
    context.user_data["otp_group_setting"] = setting_key
    current = DB.get_setting(setting_key, "").strip() or "Not set"
    await update.effective_message.reply_text(
        f"Current {description}: <code>{html.escape(current)}</code>\n"
        "Enter the numeric group ID (usually starts with -100). Enter - to remove it.\n"
        "The bot must be a member and allowed to send messages.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_OTP_FORWARD_GROUP


async def set_otp_group_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    setting_key = context.user_data.pop("otp_group_setting", "")
    if setting_key not in {"otp_forward_group_1", "otp_admin_group_2"}:
        await update.effective_message.reply_text("Setting session expired.", reply_markup=admin_menu())
        return ConversationHandler.END
    raw = (update.effective_message.text or "").strip()
    if raw == "-":
        DB.set_setting(setting_key, "")
        await update.effective_message.reply_text("✅ OTP group removed.", reply_markup=admin_group_settings_menu())
        return ConversationHandler.END
    if not re.fullmatch(r"-?\d+", raw):
        await update.effective_message.reply_text(
            "Enter a valid numeric group ID, for example -1001234567890, or - to remove it."
        )
        context.user_data["otp_group_setting"] = setting_key
        return SET_OTP_FORWARD_GROUP
    DB.set_setting(setting_key, raw)
    await update.effective_message.reply_text(
        f"✅ OTP group saved: <code>{html.escape(raw)}</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=admin_group_settings_menu(),
    )
    return ConversationHandler.END


def _setting_group_id(key: str) -> Optional[int]:
    raw = DB.get_setting(key, "").strip()
    if re.fullmatch(r"-?\d+", raw):
        return int(raw)
    return None


@dataclass(slots=True)
class LocalApiResult:
    ok: bool
    raw: str
    data: Any = None
    error: str = ""


@dataclass(slots=True)
class LocalPurchaseResult:
    ok: bool
    provider_order_id: str = ""
    provider_order_id_verified: bool = False
    phone: str = ""
    provider_price: Optional[str] = None
    message: str = ""
    raw: str = ""


@dataclass(slots=True)
class LocalOtpResult:
    state: str  # received, waiting, cancelled, error
    otp: str = ""
    message: str = ""
    raw: str = ""


class LocalUsaAPI:
    def __init__(self, base_url: str, api_key: str):
        self.base_url = base_url
        self.api_key = api_key
        timeout = aiohttp.ClientTimeout(total=HTTP_TIMEOUT_SECONDS)
        connector = aiohttp.TCPConnector(
            limit=PROVIDER_HTTP_CONNECTIONS,
            limit_per_host=PROVIDER_HTTP_CONNECTIONS,
            ttl_dns_cache=300,
            enable_cleanup_closed=True,
        )
        self.session = aiohttp.ClientSession(
            timeout=timeout,
            connector=connector,
            headers={
                "Accept": "application/json, text/plain, */*",
                "User-Agent": "Paid-Number-Bot-USA-Local/1.0",
                # Always ask the provider/proxy for a fresh OTP response.
                "Cache-Control": "no-cache, no-store, max-age=0",
                "Pragma": "no-cache",
            },
        )
        self.request_semaphore = asyncio.Semaphore(LOCAL_PROVIDER_MAX_CONCURRENCY)

    async def close(self) -> None:
        await self.session.close()

    async def request(self, action: str, **params: Any) -> LocalApiResult:
        query: dict[str, Any] = {"api_key": self.api_key, "action": action}
        query.update(
            {
                key: value
                for key, value in params.items()
                if value is not None and value != ""
            }
        )
        try:
            async with self.request_semaphore:
                async with self.session.get(
                    self.base_url, params=query, ssl=True
                ) as response:
                    raw = (await response.text(errors="replace")).strip()
                    if response.status >= 400:
                        return LocalApiResult(
                            False,
                            raw,
                            error=f"Provider HTTP {response.status}",
                        )
                    try:
                        data: Any = json.loads(raw)
                    except (json.JSONDecodeError, TypeError):
                        data = None
                    return LocalApiResult(True, raw, data)
        except asyncio.TimeoutError:
            return LocalApiResult(False, "", error="Provider request timed out.")
        except aiohttp.ClientError as exc:
            logger.warning("Local API connection error (%s): %s", action, exc)
            return LocalApiResult(False, "", error=f"Connection error: {exc}")
        except Exception as exc:
            logger.exception("Unexpected local API error (%s)", action)
            return LocalApiResult(False, "", error=f"Unexpected error: {exc}")

    async def buy_local(self) -> LocalPurchaseResult:
        return local_parse_purchase(await self.request("buy_local"))

    async def get_local_otp(
        self,
        phone: str,
        seen_otps: Optional[Iterable[str]] = None,
    ) -> LocalOtpResult:
        return local_parse_otp(
            await self.request(
                "get_local_otp",
                **{LOCAL_OTP_NUMBER_PARAM: phone},
            ),
            seen_otps=seen_otps,
        )

    async def cancel_local(self, phone: str) -> tuple[bool, str, str]:
        result = await self.request(
            "cancel_local",
            phone=phone,
        )
        success, message = local_parse_cancel(result)
        return success, message, result.raw

    async def get_balance(self) -> tuple[bool, str]:
        return local_parse_balance(await self.request("get_balance"))


LOCAL_API: Optional[LocalUsaAPI] = None
LOCAL_BUY_LOCKS: set[int] = set()
LOCAL_ORDER_POLL_TASKS: dict[int, asyncio.Task[Any]] = {}
LOCAL_DETAIL_TASKS: dict[tuple[int, int], asyncio.Task[Any]] = {}
LOCAL_ACTIVE_TASKS: dict[tuple[int, int], asyncio.Task[Any]] = {}
LOCAL_ORDER_VIEWS: dict[int, dict[tuple[int, int], int]] = {}


def ensure_local_usa_schema() -> None:
    with DB.lock:
        DB.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS local_orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                provider_order_id TEXT NOT NULL UNIQUE,
                phone TEXT NOT NULL,
                provider_price TEXT,
                charged_units INTEGER NOT NULL,
                balance_after_units INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'active',
                latest_otp TEXT,
                has_otp INTEGER NOT NULL DEFAULT 0,
                created_at INTEGER NOT NULL,
                cancel_after INTEGER NOT NULL,
                expires_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                refunded INTEGER NOT NULL DEFAULT 0,
                otp_message_id INTEGER,
                raw_buy_response TEXT,
                raw_last_response TEXT,
                FOREIGN KEY(user_id) REFERENCES users(user_id)
            );

            CREATE INDEX IF NOT EXISTS idx_local_orders_user
            ON local_orders(user_id, id DESC);

            CREATE INDEX IF NOT EXISTS idx_local_orders_active
            ON local_orders(status, expires_at);

            CREATE TABLE IF NOT EXISTS local_otp_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                otp TEXT NOT NULL,
                received_at INTEGER NOT NULL,
                raw_response TEXT,
                public_forwarded INTEGER NOT NULL DEFAULT 0,
                private_forwarded INTEGER NOT NULL DEFAULT 0,
                UNIQUE(order_id, otp),
                FOREIGN KEY(order_id) REFERENCES local_orders(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_local_otp_history_order
            ON local_otp_history(order_id, id DESC);
            """
        )

        # Migrate existing installations safely. These fields ensure that a
        # user refund can be tied to an explicit provider-side cancellation.
        local_columns = {
            row["name"]
            for row in DB.conn.execute("PRAGMA table_info(local_orders)")
        }
        local_migrations = (
            (
                "provider_order_id_verified",
                "ALTER TABLE local_orders "
                "ADD COLUMN provider_order_id_verified INTEGER NOT NULL DEFAULT 0",
            ),
            (
                "provider_cancelled",
                "ALTER TABLE local_orders "
                "ADD COLUMN provider_cancelled INTEGER NOT NULL DEFAULT 0",
            ),
            (
                "provider_cancelled_at",
                "ALTER TABLE local_orders "
                "ADD COLUMN provider_cancelled_at INTEGER",
            ),
            (
                "provider_cancel_response",
                "ALTER TABLE local_orders "
                "ADD COLUMN provider_cancel_response TEXT",
            ),
        )
        for column_name, ddl in local_migrations:
            if column_name not in local_columns:
                DB.conn.execute(ddl)

        DB.conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES('local_price_bdt',?)",
            (DEFAULT_LOCAL_PRICE_BDT,),
        )
        DB.conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES('local_session_minutes',?)",
            (str(DEFAULT_LOCAL_SESSION_MINUTES),),
        )
        DB.conn.commit()


ensure_local_usa_schema()


def local_walk_values(value: Any, prefix: str = "") -> Iterable[tuple[str, Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            path = f"{prefix}.{key}" if prefix else str(key)
            yield path.casefold(), child
            yield from local_walk_values(child, path)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            path = f"{prefix}[{index}]"
            yield path.casefold(), child
            yield from local_walk_values(child, path)


def local_first_value(data: Any, wanted: set[str]) -> Any:
    if data is None:
        return None
    wanted = {item.casefold() for item in wanted}
    for path, value in local_walk_values(data):
        leaf = re.split(r"[.\[]", path)[-1].rstrip("]")
        if leaf in wanted and value not in (None, "", [], {}):
            return value
    return None


def local_first_value_priority(data: Any, keys: Iterable[str]) -> Any:
    """Find values in priority order instead of accepting an unrelated generic id."""
    for key in keys:
        value = local_first_value(data, {key})
        if value not in (None, "", [], {}):
            return value
    return None


def local_clean_order_id(value: Any) -> str:
    candidate = str(value or "").strip()
    if not candidate:
        return ""
    if len(candidate) > 160:
        return ""
    if not re.fullmatch(r"[A-Za-z0-9._:-]+", candidate):
        return ""
    return candidate


def local_extract_provider_order_id(
    data: Any,
    raw: str,
    phone: str = "",
    *,
    allow_generic_id: bool = True,
) -> tuple[str, bool]:
    """
    Return (order_id, verified).

    A verified ID must be explicitly labelled as an order/activation/request ID
    in JSON or plain text. A phone-number fallback is never considered verified.
    """
    explicit_keys = (
        "order_id",
        "orderid",
        "activation_id",
        "activationid",
        "request_id",
        "requestid",
        "transaction_id",
        "transactionid",
    )
    explicit = local_first_value_priority(data, explicit_keys)
    candidate = local_clean_order_id(explicit)
    if candidate:
        return candidate, True

    labelled_patterns = (
        r"(?:order[_\s-]*id|orderid)\s*[:=]\s*['\"]?([A-Za-z0-9._:-]+)",
        r"(?:activation[_\s-]*id|activationid)\s*[:=]\s*['\"]?([A-Za-z0-9._:-]+)",
        r"(?:request[_\s-]*id|requestid)\s*[:=]\s*['\"]?([A-Za-z0-9._:-]+)",
        r"(?:transaction[_\s-]*id|transactionid)\s*[:=]\s*['\"]?([A-Za-z0-9._:-]+)",
    )
    for pattern in labelled_patterns:
        match = re.search(pattern, raw or "", re.IGNORECASE)
        if match:
            candidate = local_clean_order_id(match.group(1))
            if candidate:
                return candidate, True

    # Common activation API plain-text format: ACCESS_NUMBER:<order_id>:<phone>
    access = re.search(
        r"(?:ACCESS_NUMBER|SUCCESS)\s*[:|]\s*([^:|\s]+)\s*[:|]\s*(\+?\d{10,15})",
        raw or "",
        re.IGNORECASE,
    )
    if access:
        candidate = local_clean_order_id(access.group(1))
        if candidate and local_normalize_phone(candidate) != local_normalize_phone(phone):
            return candidate, True

    if allow_generic_id:
        # Generic JSON "id" is accepted only as a last resort and only when it
        # is not the purchased phone number.
        generic = local_clean_order_id(local_first_value(data, {"id"}))
        if generic and local_normalize_phone(generic) != local_normalize_phone(phone):
            return generic, True

    return "", False


def local_status_text(data: Any) -> str:
    values: list[str] = []
    for key in ("status", "success", "message", "msg", "error", "detail"):
        value = local_first_value(data, {key})
        if value is not None:
            values.append(str(value))
    return " ".join(values).casefold()


def local_response_failed(raw: str, data: Any) -> bool:
    combined = f"{raw} {local_status_text(data)}".casefold()
    negative = (
        "error",
        "failed",
        "failure",
        "invalid",
        "insufficient",
        "not available",
        "no number",
        "out of stock",
        "bad key",
        "unauthorized",
        "forbidden",
    )
    if any(word in combined for word in negative):
        return True
    success = local_first_value(data, {"success"})
    return success is False or str(success).casefold() in {"false", "0", "no"}


def local_normalize_phone(value: Any) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    if len(digits) == 10:
        return "1" + digits
    if 10 <= len(digits) <= 15:
        return digits
    return ""


def local_parse_purchase(result: LocalApiResult) -> LocalPurchaseResult:
    if not result.ok:
        return LocalPurchaseResult(False, message=result.error, raw=result.raw)
    if local_response_failed(result.raw, result.data):
        message = local_first_value(
            result.data,
            {"message", "msg", "error", "detail", "description"},
        )
        return LocalPurchaseResult(
            False,
            message=str(message or result.raw or "Purchase rejected."),
            raw=result.raw,
        )

    phone_value = local_first_value_priority(
        result.data,
        ("phone", "number", "phone_number", "phonenumber", "mobile"),
    )
    price_value = local_first_value_priority(
        result.data,
        ("price", "cost", "amount", "charge", "charged"),
    )

    phone = local_normalize_phone(phone_value)
    if not phone:
        match = re.search(r"(?<!\d)(?:\+?1)?([2-9]\d{9})(?!\d)", result.raw)
        if match:
            phone = local_normalize_phone(match.group(0))

    if not phone:
        return LocalPurchaseResult(
            False,
            message="Provider response did not contain a phone number.",
            raw=result.raw,
        )

    provider_order_id, verified = local_extract_provider_order_id(
        result.data,
        result.raw,
        phone,
    )

    # Preserve a completed provider purchase even when the provider omitted an
    # order ID. It is marked unverified and cannot release an automatic refund.
    if not provider_order_id:
        provider_order_id = f"PHONE:{phone}"

    return LocalPurchaseResult(
        True,
        provider_order_id=provider_order_id,
        provider_order_id_verified=verified,
        phone=phone,
        provider_price=str(price_value) if price_value is not None else None,
        message="Purchased.",
        raw=result.raw,
    )


LOCAL_OTP_FIELD_NAMES = {
    "otp",
    "code",
    "sms_code",
    "smscode",
    "verification_code",
    "verificationcode",
    "pin",
}


def local_extract_otp_candidates(data: Any, raw: str) -> list[str]:
    """Collect every plausible OTP while preserving the provider response order."""
    candidates: list[str] = []
    known: set[str] = set()
    token_pattern = re.compile(
        r"\b(?=[A-Za-z0-9-]{3,12}\b)(?=[A-Za-z0-9-]*\d)[A-Za-z0-9-]+\b"
    )

    def add_value(value: Any) -> None:
        if value is None or isinstance(value, (dict, list, tuple, set)):
            return
        for match in token_pattern.finditer(str(value).strip()):
            candidate = match.group(0).strip()
            key = candidate.casefold()
            if key not in known:
                known.add(key)
                candidates.append(candidate)

    if data is not None:
        for path, value in local_walk_values(data):
            path_parts = {
                part.casefold()
                for part in re.split(r"[.\[\]]+", path)
                if part and not part.isdigit()
            }
            if path_parts.intersection(LOCAL_OTP_FIELD_NAMES):
                add_value(value)

    # Plain-text APIs often return STATUS_OK:123456, OTP=123456, or
    # "verification code is 123456". Collect all labelled codes so a response
    # containing OTP history can still yield the newest unseen code.
    labelled_pattern = re.compile(
        r"(?:STATUS_OK|OTP|SMS[_\s-]*CODE|VERIFICATION[_\s-]*CODE|CODE)"
        r"\s*(?::|=|\||\bis\b)\s*([A-Za-z0-9-]{3,12})",
        re.IGNORECASE,
    )
    for match in labelled_pattern.finditer(raw or ""):
        add_value(match.group(1))

    return candidates


def local_parse_otp(
    result: LocalApiResult,
    seen_otps: Optional[Iterable[str]] = None,
) -> LocalOtpResult:
    if not result.ok:
        return LocalOtpResult("error", message=result.error, raw=result.raw)

    combined = f"{result.raw} {local_status_text(result.data)}".casefold()
    if any(word in combined for word in ("cancelled", "canceled", "expired")):
        return LocalOtpResult(
            "cancelled",
            message="Provider reports the number is inactive.",
            raw=result.raw,
        )
    if local_response_failed(result.raw, result.data):
        message = local_first_value(
            result.data,
            {"message", "msg", "error", "detail", "description"},
        )
        return LocalOtpResult(
            "error",
            message=str(message or result.raw or "OTP check failed."),
            raw=result.raw,
        )

    already_seen = {
        str(value).strip().casefold()
        for value in (seen_otps or ())
        if str(value).strip()
    }
    candidates = local_extract_otp_candidates(result.data, result.raw)

    # Providers may return the complete SMS history. Walk backwards so an
    # appended/latest code is preferred, but accept any unseen code.
    for candidate in reversed(candidates):
        if candidate.casefold() not in already_seen:
            return LocalOtpResult(
                "received",
                otp=candidate,
                message="New OTP.",
                raw=result.raw,
            )

    if candidates:
        return LocalOtpResult(
            "waiting",
            message="No new OTP.",
            raw=result.raw,
        )

    waiting = (
        "waiting",
        "pending",
        "no otp",
        "not received",
        "status_wait",
        "processing",
        "empty",
    )
    if (
        any(word in combined for word in waiting)
        or result.raw.casefold() in {"", "0", "null", "[]"}
    ):
        return LocalOtpResult(
            "waiting",
            message="No new OTP.",
            raw=result.raw,
        )

    return LocalOtpResult(
        "waiting",
        message="No new OTP.",
        raw=result.raw,
    )


def local_parse_cancel(result: LocalApiResult) -> tuple[bool, str]:
    """
    Accept cancellation only when the provider explicitly confirms it.

    Loose substring matching is intentionally avoided because text such as
    "not cancelled" or "not refunded" must never release a user refund.
    """
    if not result.ok:
        return False, result.error

    raw = (result.raw or "").strip()
    raw_folded = re.sub(r"\s+", " ", raw).casefold()
    status_value = local_first_value_priority(
        result.data,
        ("status", "state", "result"),
    )
    message_value = local_first_value_priority(
        result.data,
        ("message", "msg", "detail", "description", "error"),
    )
    success_value = local_first_value(result.data, {"success"})

    combined = " ".join(
        value
        for value in (
            raw_folded,
            str(status_value or "").casefold(),
            str(message_value or "").casefold(),
        )
        if value
    )

    negative_phrases = (
        "not cancelled",
        "not canceled",
        "not refunded",
        "cannot cancel",
        "can't cancel",
        "unable to cancel",
        "cancel failed",
        "cancellation failed",
        "no active order",
        "order not found",
        "not found",
        "too early",
        "not allowed",
        "not refundable",
        "already completed",
        "already finished",
        "already expired",
        "unsupported",
        "invalid",
        "denied",
        "failure",
        "failed",
        "error",
    )
    if any(phrase in combined for phrase in negative_phrases):
        return False, str(message_value or raw or "Cancellation rejected.")

    if success_value is True or str(success_value).strip().casefold() in {
        "true",
        "1",
        "yes",
    }:
        return True, str(message_value or "Provider confirmed cancellation.")

    accepted_statuses = {
        "success",
        "successful",
        "ok",
        "cancelled",
        "canceled",
        "refunded",
        "cancelled_refunded",
        "canceled_refunded",
        "order_cancelled",
        "order_canceled",
        "access_cancel",
    }
    normalized_status = re.sub(
        r"[\s-]+",
        "_",
        str(status_value or "").strip().casefold(),
    )
    if normalized_status in accepted_statuses:
        return True, str(message_value or "Provider confirmed cancellation.")

    positive_message_patterns = (
        r"\b(?:order\s+)?(?:cancelled|canceled)\s+(?:and\s+)?refunded\b",
        r"\brefund(?:ed)?\s+(?:and\s+)?(?:cancelled|canceled)\b",
        r"\bsuccessfully\s+(?:cancelled|canceled)\b",
        r"\b(?:cancelled|canceled)\s+successfully\b",
    )
    if any(re.search(pattern, combined) for pattern in positive_message_patterns):
        return True, str(message_value or "Provider confirmed cancellation.")

    exact_token = re.sub(r"[\s-]+", "_", raw_folded)
    if exact_token in accepted_statuses:
        return True, "Provider confirmed cancellation."

    return False, str(
        message_value
        or raw
        or "Provider did not explicitly confirm cancellation/refund."
    )


def local_parse_balance(result: LocalApiResult) -> tuple[bool, str]:
    if not result.ok:
        return False, result.error
    if local_response_failed(result.raw, result.data):
        return False, result.raw or "Balance request failed."
    value = local_first_value(
        result.data,
        {"balance", "amount", "credit", "wallet", "funds"},
    )
    if value is not None:
        return True, str(value)
    match = re.search(
        r"(?:balance|credit|funds?)\s*[:=]\s*([$৳]?\s*-?\d+(?:\.\d+)?)",
        result.raw,
        re.IGNORECASE,
    )
    if match:
        return True, match.group(1).strip()
    match = re.fullmatch(r"\s*[$৳]?\s*(-?\d+(?:\.\d+)?)\s*", result.raw)
    if match:
        return True, match.group(1)
    return False, "Balance format was not recognized."


def local_price_units() -> int:
    try:
        units = bdt_to_units(
            DB.get_setting("local_price_bdt", DEFAULT_LOCAL_PRICE_BDT)
        )
        return max(1, units)
    except Exception:
        return max(1, bdt_to_units(DEFAULT_LOCAL_PRICE_BDT))


def local_session_minutes() -> int:
    try:
        minutes = int(
            DB.get_setting(
                "local_session_minutes",
                str(DEFAULT_LOCAL_SESSION_MINUTES),
            )
        )
    except Exception:
        minutes = DEFAULT_LOCAL_SESSION_MINUTES
    return max(10, min(minutes, 7 * 24 * 60))


def local_session_seconds() -> int:
    return local_session_minutes() * 60


def local_format_hm(seconds: int) -> str:
    total_minutes = max(0, math.ceil(int(seconds) / 60))
    hours, minutes = divmod(total_minutes, 60)
    if hours:
        return f"{hours}h {minutes:02d}m"
    return f"{minutes}m"


def local_format_hms(seconds: int) -> str:
    seconds = max(0, int(seconds))
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours}h {minutes:02d}m {secs:02d}s"
    return f"{minutes:02d}m {secs:02d}s"


def local_expire_due_orders() -> None:
    """Expire only sessions that already received an OTP.

    Before the first OTP, expires_at is the 60-minute no-OTP refund deadline.
    Those orders must be cancelled at the provider and refunded before being
    removed from Active Numbers, so the async shared poller handles them.
    """
    ts = now_ts()
    DB.execute(
        """
        UPDATE local_orders
        SET status='expired', updated_at=?
        WHERE status='active'
          AND has_otp=1
          AND expires_at<=?
        """,
        (ts, ts),
    )


def local_get_order(order_id: int) -> Optional[sqlite3.Row]:
    local_expire_due_orders()
    return DB.fetchone("SELECT * FROM local_orders WHERE id=?", (order_id,))


def local_get_user_order(
    order_id: int, user_id: int
) -> Optional[sqlite3.Row]:
    local_expire_due_orders()
    return DB.fetchone(
        "SELECT * FROM local_orders WHERE id=? AND user_id=?",
        (order_id, user_id),
    )


def local_active_orders(user_id: int) -> list[sqlite3.Row]:
    local_expire_due_orders()
    return DB.fetchall(
        """
        SELECT * FROM local_orders
        WHERE user_id=? AND status='active' AND expires_at>?
        ORDER BY expires_at ASC,id DESC
        """,
        (user_id, now_ts()),
    )


def local_all_active_orders() -> list[sqlite3.Row]:
    local_expire_due_orders()
    return DB.fetchall(
        """
        SELECT * FROM local_orders
        WHERE status='active' AND expires_at>?
        ORDER BY id ASC
        """,
        (now_ts(),),
    )


def local_save_purchase(
    user_id: int,
    purchase: LocalPurchaseResult,
    charged_units: int,
) -> tuple[Optional[int], str]:
    ts = now_ts()
    session_seconds = local_session_seconds()
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            user = DB.conn.execute(
                "SELECT balance_units,is_blocked FROM users WHERE user_id=?",
                (user_id,),
            ).fetchone()
            if not user or int(user["is_blocked"] or 0):
                DB.conn.rollback()
                return None, "Account unavailable."
            if int(user["balance_units"] or 0) < charged_units:
                DB.conn.rollback()
                return None, "Insufficient balance."

            DB.conn.execute(
                """
                UPDATE users
                SET balance_units=balance_units-?,updated_at=?
                WHERE user_id=?
                """,
                (charged_units, ts, user_id),
            )
            balance_after = int(user["balance_units"] or 0) - charged_units
            cursor = DB.conn.execute(
                """
                INSERT INTO local_orders(
                    user_id,provider_order_id,provider_order_id_verified,
                    phone,provider_price,charged_units,balance_after_units,
                    status,latest_otp,has_otp,created_at,cancel_after,expires_at,
                    updated_at,refunded,raw_buy_response
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    user_id,
                    purchase.provider_order_id,
                    1 if purchase.provider_order_id_verified else 0,
                    purchase.phone,
                    purchase.provider_price,
                    charged_units,
                    balance_after,
                    "active",
                    None,
                    0,
                    ts,
                    ts + LOCAL_CANCEL_LOCK_SECONDS,
                    ts + LOCAL_CANCEL_WINDOW_SECONDS,
                    ts,
                    0,
                    purchase.raw[:10000],
                ),
            )
            order_id = int(cursor.lastrowid)
            DB.conn.execute(
                """
                INSERT INTO transactions(
                    user_id,amount_units,kind,reference,note,created_at
                ) VALUES(?,?,?,?,?,?)
                """,
                (
                    user_id,
                    -charged_units,
                    "local_purchase",
                    str(order_id),
                    "USA local number purchase",
                    ts,
                ),
            )
            DB.conn.commit()
            return order_id, ""
        except Exception:
            DB.conn.rollback()
            raise


def local_resolve_cancel_order_id(order: sqlite3.Row) -> tuple[str, bool]:
    """
    Resolve the exact provider order ID for cancel_order.

    Older versions sometimes stored the phone number as provider_order_id.
    This repairs those rows from raw_buy_response when possible.
    """
    stored = local_clean_order_id(order["provider_order_id"])
    phone = local_normalize_phone(order["phone"])
    verified = bool(int(order["provider_order_id_verified"] or 0))

    if verified and stored and not stored.startswith("PHONE:"):
        return stored, True

    raw = str(order["raw_buy_response"] or "")
    data: Any = None
    try:
        data = json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        data = None

    recovered, recovered_verified = local_extract_provider_order_id(
        data,
        raw,
        phone,
        allow_generic_id=True,
    )
    if recovered and recovered_verified:
        DB.execute(
            """
            UPDATE local_orders
            SET provider_order_id=?,provider_order_id_verified=1,updated_at=?
            WHERE id=?
            """,
            (recovered, now_ts(), order["id"]),
        )
        return recovered, True

    return "", False


async def local_notify_cancel_problem(
    application: Application,
    order: sqlite3.Row,
    reason: str,
    raw: str = "",
) -> None:
    safe_reason = html.escape(short_text(reason, 1200))
    safe_raw = html.escape(short_text(raw, 1600))
    text_value = (
        "<b>⚠️ Local provider cancellation failed</b>\n"
        f"Local order: <code>{order['id']}</code>\n"
        f"Provider ID: <code>{html.escape(str(order['provider_order_id']))}</code>\n"
        f"Phone: <code>+{html.escape(str(order['phone']))}</code>\n"
        f"Reason: {safe_reason}"
    )
    if safe_raw:
        text_value += f"\nResponse: <code>{safe_raw}</code>"

    for admin_id in ADMIN_IDS:
        try:
            await application.bot.send_message(
                chat_id=admin_id,
                text=text_value,
                parse_mode=ParseMode.HTML,
            )
        except TelegramError as exc:
            logger.warning(
                "Could not notify admin %s about local cancellation: %s",
                admin_id,
                exc,
            )


def local_refund_order(
    order_id: int,
    api_status: str,
    *,
    provider_confirmed: bool = False,
) -> bool:
    if not provider_confirmed:
        logger.error(
            "Blocked local refund for order %s without provider confirmation",
            order_id,
        )
        return False

    ts = now_ts()
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            order = DB.conn.execute(
                "SELECT * FROM local_orders WHERE id=?",
                (order_id,),
            ).fetchone()
            if (
                not order
                or order["status"] != "active"
                or int(order["refunded"] or 0)
                or int(order["has_otp"] or 0)
            ):
                DB.conn.rollback()
                return False

            DB.conn.execute(
                """
                UPDATE users
                SET balance_units=balance_units+?,updated_at=?
                WHERE user_id=?
                """,
                (order["charged_units"], ts, order["user_id"]),
            )
            DB.conn.execute(
                """
                UPDATE local_orders
                SET status='cancelled',
                    refunded=1,
                    provider_cancelled=1,
                    provider_cancelled_at=?,
                    provider_cancel_response=?,
                    raw_last_response=?,
                    updated_at=?
                WHERE id=?
                """,
                (
                    ts,
                    api_status[:10000],
                    api_status[:10000],
                    ts,
                    order_id,
                ),
            )
            DB.conn.execute(
                """
                INSERT INTO transactions(
                    user_id,amount_units,kind,reference,note,created_at
                ) VALUES(?,?,?,?,?,?)
                """,
                (
                    order["user_id"],
                    order["charged_units"],
                    "local_cancel_refund",
                    str(order_id),
                    "USA local cancellation refund",
                    ts,
                ),
            )
            DB.conn.commit()
            return True
        except Exception:
            DB.conn.rollback()
            raise


def local_record_otp(
    order_id: int,
    otp: str,
    raw_response: str,
) -> Optional[int]:
    ts = now_ts()
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            cursor = DB.conn.execute(
                """
                INSERT OR IGNORE INTO local_otp_history(
                    order_id,otp,received_at,raw_response
                ) VALUES(?,?,?,?)
                """,
                (order_id, otp, ts, raw_response[:10000]),
            )
            is_new = cursor.rowcount == 1
            DB.conn.execute(
                """
                UPDATE local_orders
                SET latest_otp=?,has_otp=1,raw_last_response=?,updated_at=?
                WHERE id=? AND status='active'
                """,
                (otp, raw_response[:10000], ts, order_id),
            )
            history_id: Optional[int] = None
            if is_new:
                row = DB.conn.execute(
                    """
                    SELECT id FROM local_otp_history
                    WHERE order_id=? AND otp=?
                    """,
                    (order_id, otp),
                ).fetchone()
                history_id = int(row["id"]) if row else None
            DB.conn.commit()
            return history_id
        except Exception:
            DB.conn.rollback()
            raise


def local_seen_otps(order_id: int) -> set[str]:
    """Return OTPs already saved for this local-number session."""
    rows = DB.fetchall(
        "SELECT otp FROM local_otp_history WHERE order_id=?",
        (order_id,),
    )
    return {
        str(row["otp"]).strip()
        for row in rows
        if str(row["otp"] or "").strip()
    }


def local_set_status(
    order_id: int,
    status: str,
    raw_response: str = "",
) -> None:
    DB.execute(
        """
        UPDATE local_orders
        SET status=?,raw_last_response=?,updated_at=?
        WHERE id=?
        """,
        (status, raw_response[:10000], now_ts(), order_id),
    )


def local_cancel_view_task(
    store: dict[tuple[int, int], asyncio.Task[Any]],
    key: tuple[int, int],
) -> None:
    task = store.pop(key, None)
    if task and not task.done():
        task.cancel()


def local_unregister_view(
    order_id: int,
    key: tuple[int, int],
) -> None:
    views = LOCAL_ORDER_VIEWS.get(order_id)
    if not views:
        return
    views.pop(key, None)
    if not views:
        LOCAL_ORDER_VIEWS.pop(order_id, None)


async def local_edit_message(
    application: Application,
    chat_id: int,
    message_id: int,
    text: str,
    markup: InlineKeyboardMarkup,
) -> bool:
    try:
        with suppress(TelegramError):
            await application.bot.delete_message(chat_id=chat_id, message_id=message_id)
        await application.bot.send_message(
            chat_id=chat_id, text=text, parse_mode=ParseMode.HTML, reply_markup=markup
        )
        return True
    except (Forbidden, TelegramError) as exc:
        logger.debug("Local message replace failed %s/%s: %s", chat_id, message_id, exc)
        return False


def local_active_payload(
    user_id: int,
    page: int,
) -> tuple[str, InlineKeyboardMarkup, int]:
    """
    Active-number status colors:
      danger  = first 10 minutes; cancellation locked
      primary = cancellation/refund currently available
      success = OTP received, or the 60-minute refund window has ended

    Status text is also included so the list remains understandable on older
    Telegram clients or older python-telegram-bot versions without button styles.
    """
    orders = local_active_orders(user_id)
    total = len(orders)
    max_page = max(
        0,
        math.ceil(total / LOCAL_ACTIVE_PAGE_SIZE) - 1,
    )
    page = max(0, min(page, max_page))
    start = page * LOCAL_ACTIVE_PAGE_SIZE
    page_orders = orders[start : start + LOCAL_ACTIVE_PAGE_SIZE]

    current = now_ts()
    buttons: list[InlineKeyboardButton] = []

    for order in page_orders:
        session_remaining = max(0, int(order["expires_at"]) - current)
        lock_remaining = max(0, int(order["cancel_after"]) - current)
        refund_window_end = int(order["created_at"]) + LOCAL_CANCEL_WINDOW_SECONDS
        refund_remaining = max(0, refund_window_end - current)
        has_otp = bool(int(order["has_otp"] or 0))

        if has_otp:
            # OTP received at any point: green and cancellation stays disabled.
            style = "success"
            label = (
                f"✅ +{order['phone']} "
                f"({local_format_hm(session_remaining)})"
            )
        elif refund_remaining <= 0:
            # The shared poller is finalising provider cancellation/refund.
            style = "danger"
            label = f"⏳ +{order['phone']} (refunding)"
        elif lock_remaining > 0:
            # First 10 minutes: red/danger and show unlock countdown.
            style = "danger"
            label = (
                f"🔒 +{order['phone']} "
                f"({local_format_hm(lock_remaining)})"
            )
        else:
            # Minute 10 through minute 60: blue/primary and show how much
            # cancellation/refund time remains.
            style = "primary"
            label = (
                f"↩️ +{order['phone']} "
                f"({local_format_hm(refund_remaining)})"
            )

        buttons.append(
            InlineKeyboardButton(
                short_text(label, 48),
                callback_data=f"local:detail:{order['id']}:{page}",
                style=style,
            )
        )

    rows = [
        buttons[index : index + 2]
        for index in range(0, len(buttons), 2)
    ]

    text_value = (
        f"<b>📱 Active: {total}</b>\n"
        f"🔒 Lock  •  ↩️ Refund  •  ✅ Done\n"
        f"Page {page + 1}/{max_page + 1}"
        if total
        else "<b>📱 Active: 0</b>"
    )

    nav: list[InlineKeyboardButton] = []
    if page > 0:
        nav.append(
            InlineKeyboardButton(
                "⬅️ Prev",
                callback_data=f"local:page:{page - 1}",
                style="secondary",
            )
        )

    nav.append(
        InlineKeyboardButton(
            f"🔄 {page + 1}/{max_page + 1}",
            callback_data=f"local:page:{page}",
            style="secondary",
        )
    )

    if page < max_page:
        nav.append(
            InlineKeyboardButton(
                "Next ➡️",
                callback_data=f"local:page:{page + 1}",
                style="secondary",
            )
        )

    rows.append(nav)
    return text_value, InlineKeyboardMarkup(rows), page


async def local_detail_loop(
    application: Application,
    order_id: int,
    chat_id: int,
    message_id: int,
    page: int,
) -> None:
    key = (chat_id, message_id)
    try:
        while True:
            order = local_get_order(order_id)
            if not order:
                return
            ok = await local_edit_message(
                application,
                chat_id,
                message_id,
                local_detail_text(order),
                local_detail_markup(order, page),
            )
            if not ok or order["status"] != "active":
                return
            await asyncio.sleep(LOCAL_DETAIL_REFRESH_SECONDS)
    except asyncio.CancelledError:
        raise
    finally:
        LOCAL_DETAIL_TASKS.pop(key, None)
        local_unregister_view(order_id, key)


async def local_order_poll_loop(
    application: Application,
    order_id: int,
) -> None:
    try:
        while True:
            order = local_get_order(order_id)
            if not order or order["status"] != "active":
                return
            if int(order["expires_at"]) <= now_ts():
                local_set_status(order_id, "expired")
                await local_refresh_all_views(application, order_id)
                return
            if LOCAL_API:
                result = await LOCAL_API.get_local_otp(
                    str(order["phone"]),
                    seen_otps=local_seen_otps(order_id),
                )
                await local_apply_otp(application, order_id, result)
            await asyncio.sleep(LOCAL_OTP_POLL_SECONDS)
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.exception("Local OTP poll crashed for order %s", order_id)
    finally:
        LOCAL_ORDER_POLL_TASKS.pop(order_id, None)




# =============================================================================
# VSHOP — VPN / PROXY SELLING SYSTEM (full spec implementation)
# =============================================================================

# (VSHOP_MIN_QTY / VSHOP_MAX_QTY are defined in the config section above)

# ── helpers ──────────────────────────────────────────────────────────────────

def _vshop_notice_group() -> Optional[int]:
    return _setting_group_id("vshop_notice_group_id")


async def _vshop_notify(bot: Any, text: str) -> Optional[Any]:
    """Send a notification and return the Telegram message for later cleanup."""
    gid = _vshop_notice_group()
    if not gid:
        return None
    try:
        return await bot.send_message(chat_id=gid, text=text, parse_mode=ParseMode.HTML)
    except TelegramError:
        return None


def _vshop_notice_text(
    *,
    order_id: int,
    user_id: int,
    username: str,
    category: str,
    service: str,
    quantity: int,
    total: int,
    status: str,
    status_icon: str,
    extra_lines: Optional[list[str]] = None,
    balance_units: Optional[int] = None,
) -> str:
    """Build the compact notice-group format used in the reference screenshot."""
    detail_lines = [
        f"🆔 Order: <code>#{order_id}</code>",
        f"👤 User: {html.escape(username)}",
        f"🆔 Id : (<code>{user_id}</code>)",
        f"📂 Category: {html.escape(category)}",
        f"📦 Service: {html.escape(service)}",
        f"🔢 Quantity: {quantity}",
        f"💰 Total: {money(total)}",
    ]
    if balance_units is not None:
        detail_lines.append(f"💳 Current Balance: {money(balance_units)}")
    if extra_lines:
        detail_lines.extend(extra_lines)
    lines = [
        f"📋 <b>VPN/Proxy Order — {html.escape(status)}</b>",
        "━━━━━━━━━━━━━━━━",
        "<blockquote>" + "\n".join(detail_lines) + "</blockquote>",
        "━━━━━━━━━━━━━━━━",
        f"<blockquote>📌 Status: {status_icon} <b>{html.escape(status)}</b></blockquote>",
    ]
    return "\n".join(lines)


async def _vshop_delete_notice(bot: Any, order: Any) -> None:
    """Remove the old pending notice after an order is completed or cancelled."""
    chat_id = order["notice_chat_id"] if "notice_chat_id" in order.keys() else None
    message_id = order["notice_message_id"] if "notice_message_id" in order.keys() else None
    if not chat_id or not message_id:
        return
    with suppress(TelegramError):
        await bot.delete_message(chat_id=int(chat_id), message_id=int(message_id))
    DB.execute(
        "UPDATE vshop_orders SET notice_chat_id=NULL,notice_message_id=NULL WHERE id=?",
        (order["id"],),
    )


def _vshop_save_notice_reference(order_id: int, message: Any) -> None:
    """Persist the notice message ID so it can be deleted on completion."""
    if not message or not getattr(message, "chat_id", None) or not getattr(message, "message_id", None):
        return
    DB.execute(
        "UPDATE vshop_orders SET notice_chat_id=?,notice_message_id=? WHERE id=?",
        (int(message.chat_id), int(message.message_id), order_id),
    )


def _vshop_order_id_label(order: Any) -> str:
    """Compact label for an order inline button."""
    cat = DB.fetchone("SELECT name FROM vshop_categories WHERE id=?", (order["category_id"],))
    svc = DB.fetchone("SELECT name FROM vshop_services WHERE id=?", (order["service_id"],))
    cat_name = str(cat["name"]) if cat else "?"
    svc_name = str(svc["name"]) if svc else "?"
    pending = int(order["pending_quantity"] or 0)
    return f"📂 {cat_name} / {svc_name} / {pending} pcs"


def _deliver_inline_text(
    order_id: int,
    svc_name: str,
    qty: int,
    total: int,
    new_balance: int,
    duration_days: int,
    stock: list,
    kind: str,
    processed: str,
) -> str:
    lines = [
        "✅ <b>Purchase Successful!</b>",
        "━━━━━━━━━━━━━━━━━━",
        f"📦 Package : {html.escape(svc_name)}",
        f"🧾 Order : #{order_id}",
        f"🕒 Processed : {processed}",
        f"🔢 Quantity : {qty}",
        f"💎 Total : {money(total)} | Balance : {money(new_balance)}",
    ]
    if duration_days and duration_days > 0:
        lines.append(f"⏳ Validity : {duration_days} Days")
    lines.append("━━━━━━━━━━━━━━━━━━")
    if kind == "proxy":
        item = stock[0]
        raw = str(item["proxy_line"])
        parts = raw.split(":", 3)
        lines.append("🔐 <b>Proxy Details</b>")
        if len(parts) == 4:
            lines += [
                f"• Server : <code>{html.escape(parts[0])}</code>",
                f"• Port : <code>{html.escape(parts[1])}</code>",
                f"• User : <code>{html.escape(parts[2])}</code>",
                f"• Pass : <code>{html.escape(parts[3])}</code>",
            ]
        lines += ["📋 One-line", f"<code>{html.escape(raw)}</code>", "⚠️ <b>Save this info!</b>"]
    else:
        lines.append("🔐 <b>VPN Account Details</b>")
        for item in stock:
            lines += [
                f"📧 Email : <code>{html.escape(str(item['email'] or ''))}</code>",
                f"🔒 Password : <code>{html.escape(str(item['password'] or ''))}</code>",
                "",
            ]
        lines.append("⚠️ <b>Save this info!</b>")
    return "\n".join(lines)


# ── menus / markups ───────────────────────────────────────────────────────────

def vshop_admin_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("➕ Add Inline"), KeyboardButton("💰 Set Category")],
            [KeyboardButton("🛠 Manage Inline"), KeyboardButton("📜 VPN Order History")],
            [KeyboardButton("🔔 VPN Notice Group"), KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


def vshop_user_categories_markup() -> InlineKeyboardMarkup:
    rows = []
    cats = DB.fetchall("SELECT * FROM vshop_categories WHERE enabled=1 ORDER BY id")
    for cat in cats:
        rows.append([InlineKeyboardButton(str(cat["name"]), callback_data=f"vshop:cat:{cat['id']}")])
    if not rows:
        rows = [[InlineKeyboardButton("No services available", callback_data="vshop:noop")]]
    return InlineKeyboardMarkup(rows)


def vshop_admin_categories_markup() -> InlineKeyboardMarkup:
    """Manage Inline category list. Category buttons open service management directly."""
    enabled = DB.get_setting("vshop_manage_inline_enabled", "1") == "1"
    rows: list[list[InlineKeyboardButton]] = [[
        InlineKeyboardButton(
            "On" if enabled else "Off",
            callback_data="vshop:manage_toggle",
            style="primary" if enabled else "danger",
        )
    ]]
    buttons = []
    cats = DB.fetchall("SELECT * FROM vshop_categories ORDER BY id")
    for cat in cats:
        buttons.append(InlineKeyboardButton(str(cat["name"]), callback_data=f"vshop:acat:{cat['id']}"))
    for i in range(0, len(buttons), 2):
        rows.append(buttons[i:i+2])
    if not cats:
        rows.append([InlineKeyboardButton("No categories", callback_data="vshop:noop")])
    rows.append([InlineKeyboardButton("⬅️ Back", callback_data="vshop:adminmenu")])
    return InlineKeyboardMarkup(rows)


def vshop_set_category_markup() -> InlineKeyboardMarkup:
    """Set Category: category selection only."""
    cats = DB.fetchall("SELECT * FROM vshop_categories WHERE enabled=1 ORDER BY id")
    rows = []
    for cat in cats:
        rows.append([InlineKeyboardButton(str(cat["name"]), callback_data=f"vshop:setcat:{cat['id']}")])
    if not rows:
        rows = [[InlineKeyboardButton("No categories", callback_data="vshop:noop")]]
    rows.append([InlineKeyboardButton("⬅️ Back", callback_data="vshop:adminmenu")])
    return InlineKeyboardMarkup(rows)


def vshop_set_service_markup(category_id: int) -> InlineKeyboardMarkup:
    """Set Category service screen: add, select and cleanup services."""
    services = DB.fetchall(
        "SELECT * FROM vshop_services WHERE category_id=? ORDER BY id",
        (category_id,),
    )
    rows = [[InlineKeyboardButton("➕ add ser", callback_data=f"vshop:addsvc:{category_id}")]]
    for svc in services:
        rows.append([InlineKeyboardButton(str(svc["name"]), callback_data=f"vshop:setsvc:{svc['id']}")])
    if not services:
        rows = [[InlineKeyboardButton("No services", callback_data="vshop:noop")]]
        rows.insert(0, [InlineKeyboardButton("➕ add ser", callback_data=f"vshop:addsvc:{category_id}")])
    rows.extend(
        [
            [InlineKeyboardButton("🗑 remove all", callback_data=f"vshop:removeall:{category_id}")],
            [InlineKeyboardButton("🗑 rem inline", callback_data=f"vshop:reminline:{category_id}")],
        ]
    )
    rows.append([InlineKeyboardButton("⬅️ Back", callback_data="vshop:setcats")])
    return InlineKeyboardMarkup(rows)


def vshop_manage_service_markup(category_id: int) -> InlineKeyboardMarkup:
    """Manage Inline service list: service buttons + Back only."""
    services = DB.fetchall(
        "SELECT * FROM vshop_services WHERE category_id=? ORDER BY id",
        (category_id,),
    )
    rows = []
    for svc in services:
        status = "🟢" if int(svc["enabled"] or 0) else "🔴"
        rows.append([InlineKeyboardButton(f"{status} {svc['name']}", callback_data=f"vshop:asvc:{svc['id']}")])
    if not rows:
        rows = [[InlineKeyboardButton("No services", callback_data="vshop:noop")]]
    rows.append([InlineKeyboardButton("⬅️ Back", callback_data=f"vshop:managecat:{category_id}")])
    return InlineKeyboardMarkup(rows)


def vshop_service_markup(
    category_id: int, admin: bool = False, back_data: str | None = None
) -> InlineKeyboardMarkup:
    services = DB.fetchall(
        "SELECT * FROM vshop_services WHERE category_id=? ORDER BY id",
        (category_id,),
    )
    rows = []
    for svc in services:
        stock = int(DB.fetchone("SELECT COUNT(*) AS n FROM vshop_inventory WHERE service_id=? AND sold=0", (svc["id"],))["n"])
        automation = int(svc["automation"] or 0)
        if admin:
            status = "🟢" if int(svc["enabled"] or 0) else "🔴"
            auto_label = "🤖" if automation else "👤"
            price_label = f"৳{units_to_decimal(int(svc['price_units'])):.2f}"
            rows.append([InlineKeyboardButton(
                f"{status} {auto_label} {svc['name']} • {price_label} • Stock {stock}",
                callback_data=f"vshop:asvc:{svc['id']}",
            )])
        elif int(svc["enabled"] or 0):
            # Service ON/OFF controls visibility only. Even with zero stock, users may order
            # and the order will be stored as pending for admin fulfillment.
            rows.append([InlineKeyboardButton(
                f"{svc['name']} • ৳{units_to_decimal(int(svc['price_units'])):.2f}",
                callback_data=f"vshop:svc:{svc['id']}",
            )])
    if admin:
        rows.insert(0, [InlineKeyboardButton("➕ Add Service", callback_data=f"vshop:addsvc:{category_id}")])
    if not rows:
        rows = [[InlineKeyboardButton("No services available", callback_data="vshop:noop")]]
    if admin:
        rows.append([
            InlineKeyboardButton(
                "⬅️ Back",
                callback_data=back_data or "vshop:setcats",
            )
        ])
    else:
        rows.append([InlineKeyboardButton("⬅️ Back", callback_data="vshop:backcats")])
    return InlineKeyboardMarkup(rows)


def vshop_parse_inventory(path: str, ext: str) -> list[tuple[str, str, str]]:
    items: list[tuple[str, str, str]] = []
    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) < 2:
                    continue
                a, b = row[0].strip(), row[1].strip()
                if not a or not b:
                    continue
                if a.lower() in {"username", "email", "user", "email/username"} and b.lower() in {"password", "pass"}:
                    continue
                items.append(("account", a, b))
    else:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            for line in f:
                line=line.strip()
                if line and line.count(":") >= 3:
                    items.append(("proxy", "", line))
    return items


def vshop_duration_label(service: Any) -> str:
    """Use the admin's original date label, with legacy DB fallback."""
    label = str(service["duration_label"] or "").strip()
    if label:
        return label
    days = int(service["duration_days"] or 0)
    return f"{days} days" if days else ""


def vshop_quantity_limits(service: Any) -> tuple[int, int]:
    """Read a service's configured quantity bounds, with legacy-safe defaults."""
    try:
        minimum = max(1, int(service["min_quantity"] or VSHOP_MIN_QTY))
        maximum = max(minimum, int(service["max_quantity"] or VSHOP_MAX_QTY))
        return minimum, maximum
    except (KeyError, IndexError, TypeError, ValueError):
        return VSHOP_MIN_QTY, VSHOP_MAX_QTY


async def show_vshop_user(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_terms(update):
        return
    if DB.get_setting("vshop_manage_inline_enabled", "1") != "1":
        await update.effective_message.reply_text("🛡️ VPN / Proxy service is currently Off.", reply_markup=main_menu(update.effective_user.id))
        return
    await update.effective_message.reply_text(
        "🛡️ <b>VPN / PROXIES</b>\n\nChoose a category:",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_user_categories_markup(),
    )


async def vshop_admin_settings(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    context.user_data["admin_mode"] = True
    await update.effective_message.reply_text("🛡️ <b>SET 𝚅𝙿𝙽/𝙿𝚁𝙾𝚇𝙸𝙴𝚂</b>", parse_mode=ParseMode.HTML, reply_markup=vshop_admin_menu())


async def vshop_add_category_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update): return ConversationHandler.END
    begin_workflow(context, admin_mode=True)
    await update.effective_message.reply_text("Send the category name (example: vpn or Proxy).", reply_markup=workflow_keyboard(update.effective_user.id, admin=True))
    return VSHOP_ADD_CATEGORY


async def vshop_add_category_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name=(update.effective_message.text or "").strip()
    if not name: return VSHOP_ADD_CATEGORY
    try:
        DB.execute("INSERT INTO vshop_categories(name,created_at) VALUES(?,?)", (name, now_ts()))
    except sqlite3.IntegrityError:
        await update.effective_message.reply_text("❌ This category already exists.", reply_markup=vshop_admin_menu())
        return ConversationHandler.END
    await update.effective_message.reply_text(f"✅ Category <b>{html.escape(name)}</b> added.", parse_mode=ParseMode.HTML, reply_markup=vshop_admin_menu())
    return ConversationHandler.END


async def vshop_set_category(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Set Category: category -> service -> price + stock upload."""
    if not await require_admin(update):
        return
    await update.effective_message.reply_text(
        "💰 <b>Set Category</b>\n\nSelect a category:",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_set_category_markup(),
    )


async def vshop_manage_inline(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Manage Inline: ON/OFF, rename and remove categories."""
    if not await require_admin(update):
        return
    enabled = DB.get_setting("vshop_manage_inline_enabled", "1") == "1"
    await update.effective_message.reply_text(
        f"🛠 <b>Manage Inline</b>\nMode: <b>{'On' if enabled else 'Off'}</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_admin_categories_markup(),
    )


async def vshop_set_category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        cid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid category", show_alert=True)
        return
    cat = DB.fetchone("SELECT * FROM vshop_categories WHERE id=?", (cid,))
    if not cat:
        await q.answer("Category not found", show_alert=True)
        return
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"<b>{html.escape(str(cat['name']))}</b>\n\nChoose service:",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_set_service_markup(cid),
    )


async def vshop_set_service_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Set Category service screen for price, date, visibility and stock."""
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        sid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid service", show_alert=True)
        return
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Service not found", show_alert=True)
        return
    stock = int(DB.fetchone(
        "SELECT COUNT(*) AS n FROM vshop_inventory WHERE service_id=? AND sold=0",
        (sid,),
    )["n"])
    duration_days = int(svc["duration_days"] or 0)
    duration_label = str(svc["duration_label"] or "").strip()
    validity = duration_label or (f"{duration_days} days" if duration_days else "Not set")
    enabled = int(svc["enabled"] or 0)
    min_qty, max_qty = vshop_quantity_limits(svc)
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"<b>{html.escape(str(svc['name']))}</b>\n"
        f"date: <b>{html.escape(validity)}</b>\n"
        f"price: <b>{money(int(svc['price_units']))}</b>\n"
        f"quantity: <b>{min_qty}–{max_qty}</b>\n"
        f"Status: <b>{'ON' if enabled else 'OFF'}</b>\n"
        f"Stock: <b>{stock}</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("📅 date", callback_data=f"vshop:setdur:{sid}"),
             InlineKeyboardButton("💰 price", callback_data=f"vshop:setprice:{sid}")],
            [InlineKeyboardButton("⬅️ Back", callback_data=f"vshop:setcatback:{svc['category_id']}")],
        ]),
    )


async def vshop_set_category_back(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        cid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid category", show_alert=True)
        return
    cat = DB.fetchone("SELECT name FROM vshop_categories WHERE id=?", (cid,))
    if not cat:
        await q.answer("Category not found", show_alert=True)
        return
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"<b>{html.escape(str(cat['name']))}</b>\n\nChoose service:",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_set_service_markup(cid),
    )


async def vshop_add_service_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update): return ConversationHandler.END
    try: category_id=int(update.callback_query.data.split(":")[-1])
    except Exception: return ConversationHandler.END
    context.user_data["vshop_category_id"] = category_id
    await update.callback_query.answer()
    await update.callback_query.message.reply_text("Send the service name (example: Nord vpn [7days]).", reply_markup=workflow_keyboard(update.effective_user.id, admin=True))
    return VSHOP_ADD_SERVICE


async def vshop_add_service_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name=(update.effective_message.text or "").strip()
    cid=int(context.user_data.get("vshop_category_id",0))
    if not name or not cid: return VSHOP_ADD_SERVICE
    DB.execute("INSERT OR IGNORE INTO vshop_services(category_id,name,created_at) VALUES(?,?,?)", (cid,name,now_ts()))
    await update.effective_message.reply_text(
        f"✅ Service <b>{html.escape(name)}</b> added.\n\nNow select it to set the price and upload stock.",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_set_service_markup(cid),
    )
    return ConversationHandler.END


async def vshop_service_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    q = update.callback_query
    sid = int(q.data.split(":")[-1])
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Service not found", show_alert=True)
        return
    stock = int(DB.fetchone(
        "SELECT COUNT(*) AS n FROM vshop_inventory WHERE service_id=? AND sold=0", (sid,)
    )["n"])
    enabled = int(svc["enabled"] or 0)
    automation = int(svc["automation"] or 0)
    min_qty, max_qty = vshop_quantity_limits(svc)
    validity = vshop_duration_label(svc) or "Not set"
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"<b>{html.escape(str(svc['name']))}</b>\n"
        f"Mode: <b>{'On' if enabled else 'Off'}</b>\n"
        f"Automation: <b>{'On' if automation else 'Off'}</b>\n"
        f"Price: <b>{money(int(svc['price_units']))}</b>\n"
        f"Quantity limit: <b>{min_qty}–{max_qty}</b>\n"
        f"Stock: <b>{stock}</b>\n"
        f"Validity: <b>{html.escape(validity)}</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("On" if enabled else "Off", callback_data=f"vshop:stoggle:{sid}", style="primary" if enabled else "danger"),
             InlineKeyboardButton("On" if automation else "Off", callback_data=f"vshop:sauto:{sid}", style="primary" if automation else "secondary")],
            [InlineKeyboardButton("📤 Upload Stock", callback_data=f"vshop:upload:{sid}"),
             InlineKeyboardButton("🧹 Clear Stock", callback_data=f"vshop:clearstock:{sid}", style="danger")],
            [InlineKeyboardButton("💵 Set Price", callback_data=f"vshop:setprice:{sid}"),
             InlineKeyboardButton("⏳ Set Duration", callback_data=f"vshop:setdur:{sid}")],
            [InlineKeyboardButton("🔢 Qty Limit", callback_data=f"vshop:setqty:{sid}")],
            [InlineKeyboardButton("⬅️ Back", callback_data=f"vshop:admincats")],
        ]),
    )


async def vshop_set_quantity_limits_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    query = update.callback_query
    try:
        sid = int(query.data.rsplit(":", 1)[1])
    except (AttributeError, ValueError):
        await query.answer("Invalid service", show_alert=True)
        return ConversationHandler.END
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await query.answer("Service not found", show_alert=True)
        return ConversationHandler.END
    minimum, maximum = vshop_quantity_limits(svc)
    context.user_data["vshop_service_id"] = sid
    await query.answer()
    await query.message.reply_text(
        f"Current quantity limit for <b>{html.escape(str(svc['name']))}</b>: "
        f"<b>{minimum}–{maximum}</b>\n\n"
        "Send minimum and maximum separated by a space, for example: <code>3 100</code>.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return VSHOP_SET_QTY


async def vshop_set_quantity_limits_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    raw = (update.effective_message.text or "").strip().replace("–", "-").replace("—", "-")
    parts = re.split(r"[\s,;:-]+", raw)
    try:
        if len(parts) != 2:
            raise ValueError
        minimum, maximum = map(int, parts)
        if minimum < 1 or maximum < minimum or maximum > 1_000_000:
            raise ValueError
    except (TypeError, ValueError):
        await update.effective_message.reply_text(
            "❌ Enter two whole numbers with minimum ≥ 1, maximum ≥ minimum, "
            "and maximum ≤ 1,000,000. Example: <code>3 100</code>.",
            parse_mode=ParseMode.HTML,
        )
        return VSHOP_SET_QTY

    sid = int(context.user_data.get("vshop_service_id", 0))
    svc = DB.fetchone("SELECT name FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await update.effective_message.reply_text(
            "❌ Service not found.", reply_markup=vshop_admin_menu()
        )
        return ConversationHandler.END
    DB.execute(
        "UPDATE vshop_services SET min_quantity=?,max_quantity=? WHERE id=?",
        (minimum, maximum, sid),
    )
    await update.effective_message.reply_text(
        f"✅ Quantity limit for <b>{html.escape(str(svc['name']))}</b> set to "
        f"<b>{minimum}–{maximum}</b>.",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_admin_menu(),
    )
    return ConversationHandler.END


async def vshop_set_price_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update): return ConversationHandler.END
    q = update.callback_query
    sid = int(q.data.split(":")[-1])
    context.user_data["vshop_service_id"] = sid
    await q.answer()
    await q.message.reply_text(
        "💰 Send the new service price in BDT (example: 25).",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return VSHOP_SET_PRICE


async def vshop_set_price_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        price = bdt_to_units((update.effective_message.text or "").strip())
    except Exception:
        await update.effective_message.reply_text("❌ Enter a valid price.")
        return VSHOP_SET_PRICE
    if price <= 0:
        await update.effective_message.reply_text("❌ Price must be greater than 0.")
        return VSHOP_SET_PRICE

    sid = int(context.user_data.get("vshop_service_id", 0))
    svc = DB.fetchone("SELECT id, name FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await update.effective_message.reply_text("❌ Service not found.", reply_markup=vshop_admin_menu())
        return ConversationHandler.END
    DB.execute("UPDATE vshop_services SET price_units=?, enabled=1 WHERE id=?", (price, sid))
    await update.effective_message.reply_text(
        f"✅ Price updated for <b>{html.escape(str(svc['name']))}</b>.\n💰 Price: <b>{money(price)}</b>",
        parse_mode=ParseMode.HTML, reply_markup=vshop_admin_menu()
    )
    return ConversationHandler.END


async def vshop_upload_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update): return ConversationHandler.END
    q = update.callback_query
    sid = int(q.data.split(":")[-1])
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Service not found", show_alert=True)
        return ConversationHandler.END
    context.user_data["vshop_service_id"] = sid
    await q.answer()
    await q.message.reply_text(
        "📤 <b>Upload Stock</b>\n\n"
        "• <b>VPN / Account:</b> CSV — Column A = Email/Username, Column B = Password.\n"
        "• <b>Proxy:</b> TXT — one <code>server:port:username:password</code> per line.\n\n"
        "ℹ️ New stock will be <b>added</b> to existing unsold stock.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return VSHOP_UPLOAD_FILE


async def vshop_upload_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    doc=update.effective_message.document
    if not doc: await update.effective_message.reply_text("❌ Upload a .csv or .txt file."); return VSHOP_UPLOAD_FILE
    name=(doc.file_name or "").lower(); ext=".csv" if name.endswith(".csv") else ".txt" if name.endswith(".txt") else ""
    if not ext: await update.effective_message.reply_text("❌ Only CSV or TXT files are supported."); return VSHOP_UPLOAD_FILE
    sid=int(context.user_data.get("vshop_service_id",0))
    svc=DB.fetchone("SELECT * FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await update.effective_message.reply_text("❌ Service not found.", reply_markup=vshop_admin_menu())
        return ConversationHandler.END
    price=int(svc["price_units"] or 0)
    fd,path=tempfile.mkstemp(prefix="vshop_",suffix=ext); os.close(fd)
    try:
        tf=await doc.get_file(); await tf.download_to_drive(custom_path=path)
        items=vshop_parse_inventory(path,ext)
        if not items: raise ValueError("No valid stock rows found.")
        kind=items[0][0]
        with DB.lock:
            DB.conn.execute("BEGIN IMMEDIATE")
            existing=DB.conn.execute("SELECT DISTINCT kind FROM vshop_inventory WHERE service_id=? AND sold=0", (sid,)).fetchall()
            existing_kinds={str(r["kind"]) for r in existing}
            incoming_kind=str(kind)
            if existing_kinds and existing_kinds != {incoming_kind}:
                raise ValueError("This service already contains a different stock type. Use a separate service for VPN/Proxy stock.")
            DB.conn.execute("UPDATE vshop_services SET enabled=1 WHERE id=?",(sid,))
            for k,a,b in items:
                DB.conn.execute("INSERT INTO vshop_inventory(service_id,kind,email,password,proxy_line,created_at) VALUES(?,?,?,?,?,?)",(sid,k,a if k=="account" else None,b if k=="account" else None,b if k=="proxy" else None,now_ts()))
            DB.conn.commit()
        await update.effective_message.reply_text(f"✅ Stock uploaded.\nPrice: <b>{money(price)}</b>\nStock: <b>{len(items)}</b>\nType: <b>{'Proxy' if kind=='proxy' else 'Account'}</b>",parse_mode=ParseMode.HTML,reply_markup=vshop_admin_menu())
    except Exception as exc:
        with suppress(Exception): DB.conn.rollback()
        await update.effective_message.reply_text(f"❌ Upload failed: {short_text(str(exc),500)}",reply_markup=vshop_admin_menu())
    finally:
        with suppress(FileNotFoundError): os.remove(path)
    return ConversationHandler.END


async def vshop_category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q=update.callback_query; cid=int(q.data.split(":")[-1]); cat=DB.fetchone("SELECT * FROM vshop_categories WHERE id=?",(cid,))
    if not cat: await q.answer("Category not found",show_alert=True); return
    await q.answer(); await replace_callback_message(q, context.bot, f"<b>{html.escape(str(cat['name']))}</b>\nChoose service:",parse_mode=ParseMode.HTML,reply_markup=vshop_service_markup(cid))


async def vshop_remove_all_services(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Ask before deleting every service and stock row in a category."""
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        cid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid category", show_alert=True)
        return
    cat = DB.fetchone("SELECT name FROM vshop_categories WHERE id=?", (cid,))
    if not cat:
        await q.answer("Category not found", show_alert=True)
        return
    count = int(DB.fetchone(
        "SELECT COUNT(*) AS n FROM vshop_services WHERE category_id=?", (cid,)
    )["n"])
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"⚠️ <b>remove all</b>\n\n"
        f"This will permanently remove <b>{count}</b> service(s), their stock, "
        "and related VPN order records from this category.\n\nContinue?",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Yes, remove all", callback_data=f"vshop:confirmremoveall:{cid}")],
            [InlineKeyboardButton("❌ Cancel", callback_data=f"vshop:setcat:{cid}")],
        ]),
    )


async def vshop_remove_all_services_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    q = update.callback_query
    cid = int(q.data.split(":")[-1])
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            DB.conn.execute(
                """
                DELETE FROM vshop_order_items
                WHERE order_id IN (
                    SELECT id FROM vshop_orders
                    WHERE service_id IN (
                        SELECT id FROM vshop_services WHERE category_id=?
                    )
                )
                """,
                (cid,),
            )
            DB.conn.execute(
                "DELETE FROM vshop_orders WHERE service_id IN "
                "(SELECT id FROM vshop_services WHERE category_id=?)",
                (cid,),
            )
            DB.conn.execute(
                "DELETE FROM vshop_inventory WHERE service_id IN "
                "(SELECT id FROM vshop_services WHERE category_id=?)",
                (cid,),
            )
            DB.conn.execute("DELETE FROM vshop_services WHERE category_id=?", (cid,))
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            raise
    await q.answer("All services removed", show_alert=True)
    await replace_callback_message(q, context.bot, 
        "✅ All services removed.\n\nYou can add a new service below.",
        reply_markup=vshop_set_service_markup(cid),
    )


async def vshop_remove_category_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Ask before removing the whole inline category."""
    if not await require_admin(update):
        return
    q = update.callback_query
    cid = int(q.data.split(":")[-1])
    cat = DB.fetchone("SELECT name FROM vshop_categories WHERE id=?", (cid,))
    if not cat:
        await q.answer("Category not found", show_alert=True)
        return
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"⚠️ <b>rem inline</b>\n\n"
        f"This will permanently remove <b>{html.escape(str(cat['name']))}</b>, "
        "all its services, stock, and related VPN order records.\n\nContinue?",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Yes, rem inline", callback_data=f"vshop:confirmreminline:{cid}")],
            [InlineKeyboardButton("❌ Cancel", callback_data=f"vshop:setcat:{cid}")],
        ]),
    )


async def vshop_remove_category_confirmed(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    # Reuse the existing atomic category-delete implementation.
    await vshop_category_delete(update, context)


async def vshop_admin_category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Manage Services for a category."""
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        cid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid category", show_alert=True)
        return
    cat = DB.fetchone("SELECT * FROM vshop_categories WHERE id=?", (cid,))
    if not cat:
        await q.answer("Category not found", show_alert=True)
        return
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"<b>{html.escape(str(cat['name']))}</b>\n\nChoose service:",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_manage_service_markup(cid),
    )


async def vshop_category_manage(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update): return
    q=update.callback_query; cid=int(q.data.split(":")[-1]); cat=DB.fetchone("SELECT * FROM vshop_categories WHERE id=?",(cid,))
    if not cat: await q.answer("Category not found",show_alert=True); return
    status="🟢 ON" if int(cat["enabled"] or 0) else "🔴 OFF"
    await q.answer(); await replace_callback_message(q, context.bot, f"<b>{html.escape(str(cat['name']))}</b>\nStatus: <b>{status}</b>",parse_mode=ParseMode.HTML,reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🟢 ON / 🔴 OFF",callback_data=f"vshop:ctoggle:{cid}")],[InlineKeyboardButton("✏️ Rename",callback_data=f"vshop:crename:{cid}")],[InlineKeyboardButton("🗑 Remove Category",callback_data=f"vshop:cdelete:{cid}")],[InlineKeyboardButton("📦 Manage Services",callback_data=f"vshop:acat:{cid}")],[InlineKeyboardButton("⬅️ Back",callback_data="vshop:admincats")]]))

async def vshop_manage_inline_toggle(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    q = update.callback_query
    current = DB.get_setting("vshop_manage_inline_enabled", "1") == "1"
    new_val = "0" if current else "1"
    DB.set_setting("vshop_manage_inline_enabled", new_val)
    await q.answer("On" if new_val == "1" else "Off")
    await replace_callback_message(q, context.bot, 
        f"🛠 <b>Manage Inline</b>\nMode: <b>{'On' if new_val == '1' else 'Off'}</b>",
        parse_mode=ParseMode.HTML, reply_markup=vshop_admin_categories_markup()
    )


async def vshop_clear_stock(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        sid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid service", show_alert=True); return
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Service not found", show_alert=True); return
    DB.execute("DELETE FROM vshop_inventory WHERE service_id=? AND sold=0", (sid,))
    await q.answer("Stock cleared")
    await vshop_service_admin(update, context)


async def vshop_category_toggle(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update): return
    q=update.callback_query; cid=int(q.data.split(":")[-1]); cat=DB.fetchone("SELECT enabled FROM vshop_categories WHERE id=?",(cid,))
    if not cat: await q.answer("Category not found",show_alert=True); return
    enabled=0 if int(cat["enabled"] or 0) else 1; DB.execute("UPDATE vshop_categories SET enabled=? WHERE id=?",(enabled,cid)); await q.answer("Category ON" if enabled else "Category OFF",show_alert=True); await vshop_category_manage(update,context)

async def vshop_category_rename_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update): return ConversationHandler.END
    q=update.callback_query; context.user_data["vshop_category_id"]=int(q.data.split(":")[-1]); await q.answer(); await q.message.reply_text("✏️ Send the new category name.",reply_markup=workflow_keyboard(update.effective_user.id,admin=True)); return VSHOP_RENAME_CATEGORY

async def vshop_category_rename_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name=(update.effective_message.text or "").strip(); cid=int(context.user_data.get("vshop_category_id",0))
    if not name or not cid: return VSHOP_RENAME_CATEGORY
    try: DB.execute("UPDATE vshop_categories SET name=? WHERE id=?",(name,cid))
    except sqlite3.IntegrityError:
        await update.effective_message.reply_text("❌ This category name already exists."); return VSHOP_RENAME_CATEGORY
    await update.effective_message.reply_text("✅ Category renamed.",reply_markup=vshop_admin_menu()); return ConversationHandler.END

async def vshop_category_delete(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update): return
    q=update.callback_query; cid=int(q.data.split(":")[-1]); cat=DB.fetchone("SELECT name FROM vshop_categories WHERE id=?",(cid,))
    if not cat: await q.answer("Category not found",show_alert=True); return
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            # Remove dependent order rows first.  vshop_orders.service_id and
            # vshop_order_items.inventory_id are intentionally kept for
            # history, so SQLite otherwise rejects the service/category
            # delete when the service has ever been purchased.
            DB.conn.execute(
                """
                DELETE FROM vshop_order_items
                WHERE order_id IN (
                    SELECT id FROM vshop_orders
                    WHERE service_id IN (
                        SELECT id FROM vshop_services WHERE category_id=?
                    )
                )
                """,
                (cid,),
            )
            DB.conn.execute(
                """
                DELETE FROM vshop_orders
                WHERE service_id IN (
                    SELECT id FROM vshop_services WHERE category_id=?
                )
                """,
                (cid,),
            )
            DB.conn.execute(
                """
                DELETE FROM vshop_inventory
                WHERE service_id IN (
                    SELECT id FROM vshop_services WHERE category_id=?
                )
                """,
                (cid,),
            )
            DB.conn.execute("DELETE FROM vshop_services WHERE category_id=?", (cid,))
            DB.conn.execute("DELETE FROM vshop_categories WHERE id=?",(cid,)); DB.conn.commit()
        except Exception:
            DB.conn.rollback(); raise
    await q.answer("Category removed",show_alert=True); await replace_callback_message(q, context.bot, "✅ Category removed.",reply_markup=vshop_admin_categories_markup())

async def vshop_service_toggle(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update): return
    q=update.callback_query; sid=int(q.data.split(":")[-1]); svc=DB.fetchone("SELECT enabled FROM vshop_services WHERE id=?",(sid,))
    if not svc: await q.answer("Service not found",show_alert=True); return
    enabled=0 if int(svc["enabled"] or 0) else 1
    DB.execute("UPDATE vshop_services SET enabled=? WHERE id=?",(enabled,sid))
    await q.answer("Service ON" if enabled else "Service OFF",show_alert=True)
    if str(q.data).startswith("vshop:settoggle:"):
        await vshop_set_service_admin(update, context)
    else:
        await vshop_service_admin(update,context)

async def vshop_service_rename_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update): return ConversationHandler.END
    q=update.callback_query; context.user_data["vshop_rename_service_id"]=int(q.data.split(":")[-1]); await q.answer(); await q.message.reply_text("✏️ Send the new service name.",reply_markup=workflow_keyboard(update.effective_user.id,admin=True)); return VSHOP_RENAME_SERVICE

async def vshop_service_rename_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name=(update.effective_message.text or "").strip(); sid=int(context.user_data.get("vshop_rename_service_id",0))
    svc=DB.fetchone("SELECT category_id FROM vshop_services WHERE id=?",(sid,))
    if not name or not svc: return VSHOP_RENAME_SERVICE
    try: DB.execute("UPDATE vshop_services SET name=? WHERE id=?",(name,sid))
    except sqlite3.IntegrityError:
        await update.effective_message.reply_text("❌ This service name already exists in this category."); return VSHOP_RENAME_SERVICE
    await update.effective_message.reply_text("✅ Service renamed.",reply_markup=vshop_admin_menu()); return ConversationHandler.END

async def vshop_remove_service(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Permanently remove one service and all of its unsold inventory.

    This callback is registered before the main ConversationHandler so an active
    admin workflow cannot swallow the Remove Service click.  The deletion is
    performed atomically and the service list is refreshed immediately.
    """
    if not await require_admin(update):
        return
    q = update.callback_query
    if not q:
        return
    try:
        sid = int(q.data.rsplit(":", 1)[1])
    except (ValueError, IndexError, AttributeError):
        await q.answer("Invalid service.", show_alert=True)
        return

    svc = DB.fetchone("SELECT id, category_id, name FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Service already removed.", show_alert=True)
        return

    category_id = int(svc["category_id"])
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            # Clear rows that reference sold inventory/service before deleting
            # the service.  Without this, a previously purchased service
            # remains protected by SQLite foreign keys and the button looks
            # like it did nothing.
            DB.conn.execute(
                """
                DELETE FROM vshop_order_items
                WHERE order_id IN (
                    SELECT id FROM vshop_orders WHERE service_id=?
                )
                """,
                (sid,),
            )
            DB.conn.execute("DELETE FROM vshop_orders WHERE service_id=?", (sid,))
            DB.conn.execute("DELETE FROM vshop_inventory WHERE service_id=?", (sid,))
            DB.conn.execute("DELETE FROM vshop_services WHERE id=?", (sid,))
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            raise

    await q.answer("🗑 Service removed", show_alert=True)
    is_set_category = str(q.data).startswith("vshop:setremove:")
    await replace_callback_message(q, context.bot, 
        "🛠 <b>Manage Services</b>\n\nService removed successfully.",
        parse_mode=ParseMode.HTML,
        reply_markup=(
            vshop_set_service_markup(category_id)
            if is_set_category
            else vshop_service_markup(
                category_id,
                admin=True,
                back_data=f"vshop:managecat:{category_id}",
            )
        ),
    )


async def vshop_show_service(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Show service details with quantity controls bounded by service settings."""
    q = update.callback_query
    sid = int(q.data.split(":")[-1])
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=? AND enabled=1", (sid,))
    if not svc:
        await q.answer("Service unavailable", show_alert=True)
        return
    price = int(svc["price_units"])
    validity = vshop_duration_label(svc)
    automation = int(svc["automation"] or 0)
    # Service ON/OFF is visibility only. Zero stock is allowed and becomes pending.
    min_qty, _ = vshop_quantity_limits(svc)
    qty_default = min_qty
    total = price * qty_default
    dur_line = f"\n⏳ Validity: <b>{html.escape(validity)}</b>" if validity else ""
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"📦 <b>Service:</b> {html.escape(str(svc['name']))}\n"
        f"💵 <b>Price:</b> {money(price)} / pcs\n"
        f"🔢 <b>Quantity:</b> {qty_default}\n"
        f"💰 <b>Total:</b> {money(total)}"
        f"{dur_line}",
        parse_mode=ParseMode.HTML,
        reply_markup=_vshop_qty_markup(sid, qty_default),
    )


def _vshop_qty_markup(sid: int, qty: int) -> InlineKeyboardMarkup:
    svc = DB.fetchone("SELECT min_quantity,max_quantity FROM vshop_services WHERE id=?", (sid,))
    min_qty, max_qty = vshop_quantity_limits(svc) if svc else (VSHOP_MIN_QTY, VSHOP_MAX_QTY)
    qty = max(min_qty, min(max_qty, qty))
    minus_style = "secondary" if qty <= min_qty else "danger"
    plus_style  = "secondary" if qty >= max_qty else "success"
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("➖", callback_data=f"vshop:qty:{sid}:{max(min_qty, qty-1)}", style=minus_style),
            InlineKeyboardButton(f"  {qty}  ", callback_data="vshop:noop"),
            InlineKeyboardButton("➕", callback_data=f"vshop:qty:{sid}:{min(max_qty, qty+1)}", style=plus_style),
        ],
        [InlineKeyboardButton("✏️ Custom Quantity", callback_data=f"vshop:enter:{sid}")],
        [
            InlineKeyboardButton("✅ Confirm", callback_data=f"vshop:confirm:{sid}:{qty}", style="primary"),
            InlineKeyboardButton("❌ Cancel", callback_data="vshop:cancel", style="danger"),
        ],
        [InlineKeyboardButton("⬅️ Back", callback_data=f"vshop:catback:{sid}")],
    ])


async def vshop_enter_quantity(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """User tapped Custom Quantity."""
    q = update.callback_query
    sid = int(q.data.split(":")[-1])
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=? AND enabled=1", (sid,))
    if not svc:
        await q.answer("Service unavailable", show_alert=True)
        return ConversationHandler.END
    min_qty, max_qty = vshop_quantity_limits(svc)
    context.user_data["vshop_service_id"] = sid
    await q.answer()
    await q.message.reply_text(
        f"📦 Send the quantity you want (min {min_qty}, max {max_qty}):",
        reply_markup=workflow_keyboard(update.effective_user.id),
    )
    return VSHOP_QUANTITY


async def vshop_quantity_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        qty = int((update.effective_message.text or "").strip())
    except Exception:
        qty = 0
    sid = int(context.user_data.get("vshop_service_id", 0))
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=? AND enabled=1", (sid,))
    if not svc:
        await update.effective_message.reply_text("❌ Service unavailable.")
        return ConversationHandler.END
    min_qty, max_qty = vshop_quantity_limits(svc)
    if qty < min_qty or qty > max_qty:
        await update.effective_message.reply_text(
            f"❌ Quantity must be between {min_qty} and {max_qty}."
        )
        return VSHOP_QUANTITY
    await vshop_send_confirmation(update.effective_message, context, sid, qty)
    return ConversationHandler.END


async def vshop_qty_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles ➖ / ➕ quantity buttons by editing the existing service panel."""
    q = update.callback_query
    _, _, sid_raw, qty_raw = q.data.split(":")
    sid, qty = int(sid_raw), int(qty_raw)
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=? AND enabled=1", (sid,))
    if not svc:
        await q.answer("Service unavailable", show_alert=True)
        return
    min_qty, max_qty = vshop_quantity_limits(svc)
    qty = max(min_qty, min(max_qty, qty))
    price = int(svc["price_units"])
    validity = vshop_duration_label(svc)
    dur_line = f"\n⏳ Validity: <b>{html.escape(validity)}</b>" if validity else ""
    await q.answer()
    try:
        await q.message.edit_text(
            f"📦 <b>Service:</b> {html.escape(str(svc['name']))}\n"
            f"💵 <b>Price:</b> {money(price)} / pcs\n"
            f"🔢 <b>Quantity:</b> {qty}\n"
            f"💰 <b>Total:</b> {money(price * qty)}"
            f"{dur_line}",
            parse_mode=ParseMode.HTML,
            reply_markup=_vshop_qty_markup(sid, qty),
        )
    except TelegramError:
        pass


async def vshop_send_confirmation(message: Any, context: ContextTypes.DEFAULT_TYPE, sid: int, qty: int) -> None:
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=? AND enabled=1", (sid,))
    if not svc:
        await message.reply_text("❌ Service unavailable.")
        return
    min_qty, max_qty = vshop_quantity_limits(svc)
    if qty < min_qty or qty > max_qty:
        await message.reply_text(f"❌ Quantity must be between {min_qty} and {max_qty}.")
        return
    total = int(svc["price_units"]) * qty
    user = DB.fetchone("SELECT balance_units FROM users WHERE user_id=?", (message.chat_id,))
    bal = int(user["balance_units"] or 0) if user else 0
    validity = vshop_duration_label(svc)
    dur_line = f"\n⏳ Validity: <b>{html.escape(validity)}</b>" if validity else ""
    context.user_data["vshop_confirm"] = {"sid": sid, "qty": qty}
    await message.reply_text(
        f"📦 <b>Service:</b> {html.escape(str(svc['name']))}\n"
        f"💵 <b>Price:</b> {money(int(svc['price_units']))} / pcs\n"
        f"🔢 <b>Quantity:</b> {qty}\n"
        f"💰 <b>Total:</b> {money(total)}"
        f"{dur_line}\n"
        f"💳 <b>Your Balance:</b> {money(bal)}\n\n"
        f"{'⚠️ <b>Insufficient balance!</b>' if bal < total else 'Confirm purchase?'}",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [
                InlineKeyboardButton("✅ Confirm", callback_data=f"vshop:confirm:{sid}:{qty}", style="primary"),
                InlineKeyboardButton("❌ Cancel", callback_data="vshop:cancel", style="danger"),
            ],
        ]),
    )


async def vshop_confirm_purchase(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Full purchase flow: balance deduct, auto/manual/partial delivery, notice group."""
    q = update.callback_query
    user_id = q.from_user.id
    user_obj = q.from_user
    try:
        _, _, sid_raw, qty_raw = q.data.split(":")
        sid, qty = int(sid_raw), int(qty_raw)
    except Exception:
        await q.answer("Invalid order", show_alert=True)
        return

    # ── Atomic: balance check + deduct + create order ──────────────────────
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            svc = DB.conn.execute(
                "SELECT * FROM vshop_services WHERE id=? AND enabled=1", (sid,)
            ).fetchone()
            if not svc:
                raise ValueError("Service unavailable.")

            min_qty, max_qty = vshop_quantity_limits(svc)
            if qty < min_qty or qty > max_qty:
                raise ValueError(f"Quantity must be between {min_qty} and {max_qty}.")

            total = int(svc["price_units"]) * qty
            user = DB.conn.execute(
                "SELECT balance_units FROM users WHERE user_id=?", (user_id,)
            ).fetchone()
            if not user or int(user["balance_units"] or 0) < total:
                raise ValueError(
                    f"❌ Insufficient balance.\nNeed: {money(total)}\n"
                    f"Balance: {money(int(user['balance_units'] or 0) if user else 0)}"
                )

            automation = int(svc["automation"] or 0)
            ts = now_ts()
            new_balance = int(user["balance_units"]) - total
            DB.conn.execute(
                "UPDATE users SET balance_units=balance_units-?,updated_at=? WHERE user_id=?",
                (total, ts, user_id),
            )
            cat = DB.conn.execute(
                "SELECT id FROM vshop_categories WHERE id=?", (svc["category_id"],)
            ).fetchone()
            category_id = int(cat["id"]) if cat else None

            # How many stock items do we have right now?
            if automation:
                avail_stock = DB.conn.execute(
                    "SELECT * FROM vshop_inventory WHERE service_id=? AND sold=0 ORDER BY id LIMIT ?",
                    (sid, qty),
                ).fetchall()
            else:
                avail_stock = []

            auto_count = len(avail_stock)
            pending_qty = qty - auto_count

            cur = DB.conn.execute(
                """INSERT INTO vshop_orders(
                    user_id, service_id, category_id, quantity,
                    auto_delivered, pending_quantity, total_units, status, created_at
                ) VALUES(?,?,?,?,?,?,?,?,?)""",
                (
                    user_id, sid, category_id, qty,
                    auto_count, pending_qty, total,
                    "success" if pending_qty == 0 else ("partial" if auto_count > 0 else "pending"),
                    ts,
                ),
            )
            oid = cur.lastrowid

            for item in avail_stock:
                updated = DB.conn.execute(
                    "UPDATE vshop_inventory SET sold=1,sold_to=?,sold_order_id=?,sold_at=? WHERE id=? AND sold=0",
                    (user_id, oid, ts, item["id"]),
                ).rowcount
                if updated != 1:
                    raise ValueError("Stock sold simultaneously. Please try again.")
                DB.conn.execute(
                    "INSERT INTO vshop_order_items(order_id,inventory_id) VALUES(?,?)",
                    (oid, item["id"]),
                )

            DB.conn.execute(
                "INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at) VALUES(?,?,?,?,?,?)",
                (user_id, -total, "vshop_purchase", str(oid), f"VPN/Proxy: {svc['name']}", ts),
            )
            DB.conn.commit()
        except Exception as exc:
            DB.conn.rollback()
            msg = str(exc)
            await q.answer("Order rejected.", show_alert=False)
            if "Insufficient balance" in msg:
                await replace_callback_message(
                    q, context.bot,
                    "আপনার পর্যাপ্ত ব্যালেন্স নেই। অর্ডার করতে আগে ব্যালেন্স যোগ করুন।",
                )
            else:
                await replace_callback_message(q, context.bot, f"❌ {html.escape(msg[:500])}")
            return

    # ── Variables for messaging ────────────────────────────────────────────
    service_name = str(svc["name"])
    duration = int(svc["duration_days"] or 0)
    processed = datetime.fromtimestamp(ts).strftime("%Y-%m-%d %I:%M %p")
    uname_display = f"@{user_obj.username}" if user_obj.username else str(user_id)
    cat_name_row = DB.fetchone("SELECT name FROM vshop_categories WHERE id=?", (category_id,))
    cat_name = str(cat_name_row["name"]) if cat_name_row else "?"

    await q.answer("Purchase successful!")
    with suppress(TelegramError):
        await q.message.delete()

    # ── Case 1: Full auto delivery ─────────────────────────────────────────
    if auto_count == qty:
        stock = avail_stock
        kind = str(stock[0]["kind"]).strip().lower()
        await _vshop_deliver_stock(
            context.bot, user_id, oid, service_name, qty, total, new_balance,
            duration, stock, kind, processed, auto_count, 0,
        )
        await _vshop_notify(
            context.bot,
            _vshop_notice_text(
                order_id=oid,
                user_id=user_id,
                username=uname_display,
                category=cat_name,
                service=service_name,
                quantity=qty,
                total=total,
                status="SUCCESS (Auto Provide)",
                status_icon="✅",
                balance_units=new_balance,
            ),
        )
        return

    # ── Case 2: Partial auto + pending ────────────────────────────────────
    if auto_count > 0:
        stock = avail_stock
        kind = str(stock[0]["kind"]).strip().lower()
        await _vshop_deliver_stock(
            context.bot, user_id, oid, service_name, auto_count, total, new_balance,
            duration, stock, kind, processed, auto_count, pending_qty,
        )
        await q.message.reply_text(
            f"⏳ <b>Remaining {pending_qty} pcs Pending.</b>\n"
            "অবশিষ্ট পরে Provide করা হবে।",
            parse_mode=ParseMode.HTML,
        )
        notice = await _vshop_notify(
            context.bot,
            _vshop_notice_text(
                order_id=oid,
                user_id=user_id,
                username=uname_display,
                category=cat_name,
                service=service_name,
                quantity=qty,
                total=total,
                status="PARTIAL / PENDING",
                status_icon="⏳",
                extra_lines=[
                    f"🤖 Auto Provided: {auto_count}",
                    f"⏳ Pending: {pending_qty}",
                ],
                balance_units=new_balance,
            ),
        )
        _vshop_save_notice_reference(oid, notice)
        return

    # ── Case 3: Full pending (automation OFF or stock empty) ───────────────
    await q.message.reply_text(
        f"✅ <b>Order Confirmed!</b>\n"
        f"🆔 Order ID: <code>#{oid}</code>\n"
        f"📦 Service: {html.escape(service_name)}\n"
        f"🔢 Quantity: {qty}\n"
        f"💰 Total: {money(total)}\n"
        f"⏳ <b>Status: PENDING</b>\n\n"
        "আপনার Order Pending আছে। Admin শীঘ্রই Provide করবে।",
        parse_mode=ParseMode.HTML,
    )
    notice = await _vshop_notify(
        context.bot,
        _vshop_notice_text(
            order_id=oid,
            user_id=user_id,
            username=uname_display,
            category=cat_name,
            service=service_name,
            quantity=qty,
            total=total,
            status="PENDING",
            status_icon="⏳",
            balance_units=new_balance,
        ),
    )
    _vshop_save_notice_reference(oid, notice)


async def _vshop_deliver_stock(
    bot: Any,
    user_id: int,
    oid: int,
    service_name: str,
    qty: int,
    total: int,
    new_balance: int,
    duration: int,
    stock: list,
    kind: str,
    processed: str,
    auto_delivered: int,
    pending_qty: int,
) -> None:
    """Deliver stock items to user (file or inline)."""
    prefix = ""
    if pending_qty > 0:
        prefix = f"✅ <b>{auto_delivered} pcs Automatically Provided.</b>\n\n"

    if kind == "proxy" and qty >= 2:
        proxy_lines = [str(item["proxy_line"]).strip() for item in stock if str(item["proxy_line"] or "").strip()]
        fd, out_path = tempfile.mkstemp(prefix="proxy_order_", suffix=".txt")
        os.close(fd)
        try:
            with open(out_path, "w", encoding="utf-8", newline="\n") as f:
                f.write("\n".join(proxy_lines) + "\n")
            dur_line = f"⏳ Validity: {duration} Days\n" if duration else ""
            await bot.send_message(
                chat_id=user_id,
                text=(
                    prefix +
                    "✅ <b>Purchase Successful!</b>\n━━━━━━━━━━━━━━━━━━\n"
                    f"📦 Package : {html.escape(service_name)}\n"
                    f"🧾 Order : #{oid}\n"
                    f"🔢 Quantity : {qty}\n"
                    f"💎 Total : {money(total)} | Balance : {money(new_balance)}\n"
                    f"{dur_line}━━━━━━━━━━━━━━━━━━\n"
                    "📄 <b>Your proxy list is attached as a TXT file.</b>\n\n"
                    "⚠️ <b>Save this info!</b>"
                ),
                parse_mode=ParseMode.HTML,
            )
            with open(out_path, "rb") as f:
                await bot.send_document(
                    chat_id=user_id,
                    document=f,
                    filename=f"{service_name}_proxies.txt",
                    caption="📄 Proxy List",
                )
        finally:
            with suppress(FileNotFoundError):
                os.remove(out_path)
        return

    if kind == "account" and qty >= 4:
        fd, out_path = tempfile.mkstemp(prefix="vpn_order_", suffix=".csv")
        os.close(fd)
        try:
            with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Email/Username", "Password"])
                for item in stock:
                    writer.writerow([str(item["email"] or ""), str(item["password"] or "")])
            dur_line = f"⏳ Validity: {duration} Days\n" if duration else ""
            await bot.send_message(
                chat_id=user_id,
                text=(
                    prefix +
                    "✅ <b>Purchase Successful!</b>\n━━━━━━━━━━━━━━━━━━\n"
                    f"📦 Package : {html.escape(service_name)}\n"
                    f"🧾 Order : #{oid}\n"
                    f"🔢 Quantity : {qty}\n"
                    f"💎 Total : {money(total)} | Balance : {money(new_balance)}\n"
                    f"{dur_line}━━━━━━━━━━━━━━━━━━\n"
                    "📄 <b>Your VPN accounts are attached as a CSV file.</b>\n\n"
                    "⚠️ <b>Save this info!</b>"
                ),
                parse_mode=ParseMode.HTML,
            )
            with open(out_path, "rb") as f:
                await bot.send_document(
                    chat_id=user_id,
                    document=f,
                    filename=f"{service_name}_accounts.csv",
                    caption="📄 VPN Account List",
                )
        finally:
            with suppress(FileNotFoundError):
                os.remove(out_path)
        return

    # Inline for small quantities
    dur_line = f"\n⏳ Validity: <b>{duration} Days</b>" if duration else ""
    lines = [
        prefix + "✅ <b>Purchase Successful!</b>",
        "━━━━━━━━━━━━━━━━━━",
        f"🆔 Order ID: <code>#{oid}</code>",
        f"📦 Service: {html.escape(service_name)}",
        f"🔢 Quantity: {qty}",
        f"💎 Total: {money(total)} | Balance: {money(new_balance)}",
    ]
    if duration:
        lines.append(f"⏳ Validity: {duration} Days")
    lines.append("━━━━━━━━━━━━━━━━━━")
    if kind == "proxy":
        item = stock[0]
        raw = str(item["proxy_line"])
        parts = raw.split(":", 3)
        lines.append("🔐 <b>Proxy Details</b>")
        if len(parts) == 4:
            lines += [
                f"• Server: <code>{html.escape(parts[0])}</code>",
                f"• Port: <code>{html.escape(parts[1])}</code>",
                f"• User: <code>{html.escape(parts[2])}</code>",
                f"• Pass: <code>{html.escape(parts[3])}</code>",
            ]
        lines += ["📋 One-line:", f"<code>{html.escape(raw)}</code>"]
    else:
        lines.append("🔐 <b>VPN Account Details</b>")
        for item in stock:
            lines += [
                f"📧 Email: <code>{html.escape(str(item['email'] or ''))}</code>",
                f"🔒 Password: <code>{html.escape(str(item['password'] or ''))}</code>",
                "",
            ]
    lines.append("⚠️ <b>Save this info!</b>")
    await bot.send_message(
        chat_id=user_id,
        text="\n".join(lines),
        parse_mode=ParseMode.HTML,
    )


# ── Automation toggle ─────────────────────────────────────────────────────────

async def vshop_service_automation_toggle(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        sid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid", show_alert=True)
        return
    svc = DB.fetchone("SELECT automation FROM vshop_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Not found", show_alert=True)
        return
    new_val = 0 if int(svc["automation"] or 0) else 1
    DB.execute("UPDATE vshop_services SET automation=? WHERE id=?", (new_val, sid))
    await q.answer("On" if new_val else "Off")
    await vshop_service_admin(update, context)


# ── Duration (validity) ───────────────────────────────────────────────────────

async def vshop_set_duration_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    q = update.callback_query
    sid = int(q.data.split(":")[-1])
    context.user_data["vshop_service_id"] = sid
    await q.answer()
    svc = DB.fetchone(
        "SELECT name, duration_days, duration_label FROM vshop_services WHERE id=?",
        (sid,),
    )
    cur_dur = int(svc["duration_days"] or 0) if svc else 0
    cur_label = str(svc["duration_label"] or "").strip() if svc else ""
    await q.message.reply_text(
        f"⏳ Current date: <b>{html.escape(cur_label or (f'{cur_dur} days' if cur_dur else 'Not set'))}</b>\n\n"
        "Send the package date/validity, for example: 7d, 1m, or 10 days.\n"
        "Send 0 to remove it.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return VSHOP_SET_DURATION


def parse_vshop_duration(raw: str) -> tuple[int, str]:
    """Return a usable day count and the label users should see."""
    value = re.sub(r"\s+", " ", (raw or "").strip().lower())
    if value == "0":
        return 0, ""
    match = re.fullmatch(r"(\d+)\s*(d|day|days|w|week|weeks|m|month|months)?", value)
    if not match:
        raise ValueError("Use 7d, 1m, 10 days, or 0.")
    amount = int(match.group(1))
    unit = match.group(2) or "days"
    if amount < 1:
        raise ValueError("Validity must be greater than 0, or 0 to remove it.")
    if unit in {"m", "month", "months"}:
        days = amount * 30
        label = f"{amount}m"
    elif unit in {"w", "week", "weeks"}:
        days = amount * 7
        label = f"{amount}w"
    elif unit == "d":
        days = amount
        label = f"{amount}d"
    else:
        days = amount
        label = f"{amount} days"
    return days, label


async def vshop_set_duration_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        days, label = parse_vshop_duration(update.effective_message.text or "")
    except Exception as exc:
        await update.effective_message.reply_text(f"❌ {short_text(str(exc), 160)}")
        return VSHOP_SET_DURATION
    sid = int(context.user_data.get("vshop_service_id", 0))
    DB.execute(
        "UPDATE vshop_services SET duration_days=?, duration_label=? WHERE id=?",
        (days, label, sid),
    )
    display_label = label or "Removed (no date shown)"
    await update.effective_message.reply_text(
        f"✅ date set to <b>{html.escape(display_label)}</b>.",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_admin_menu(),
    )
    return ConversationHandler.END


# ── Notice Group ──────────────────────────────────────────────────────────────

async def vshop_set_notice_group_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    begin_workflow(context, admin_mode=True)
    current = DB.get_setting("vshop_notice_group_id", "") or "Not set"
    await update.effective_message.reply_text(
        f"🔔 Current VPN Notice Group ID: <code>{html.escape(current)}</code>\n\n"
        "Send the new group Chat ID (e.g. -1001234567890). Send 0 to clear.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return VSHOP_SET_NOTICE_GROUP


async def vshop_set_notice_group_value(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw = (update.effective_message.text or "").strip()
    if not re.fullmatch(r"-?\d+", raw):
        await update.effective_message.reply_text("❌ Enter a valid Chat ID (e.g. -1001234567890) or 0 to clear.")
        return VSHOP_SET_NOTICE_GROUP
    val = raw if raw != "0" else ""
    DB.set_setting("vshop_notice_group_id", val)
    label = val if val else "Cleared"
    await update.effective_message.reply_text(
        f"✅ VPN Notice Group set to <code>{html.escape(label)}</code>.",
        parse_mode=ParseMode.HTML,
        reply_markup=vshop_admin_menu(),
    )
    return ConversationHandler.END


# ── Admin: Order History / Provide / Cancel+Refund ────────────────────────────

async def vshop_order_history_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    begin_workflow(context, admin_mode=True)
    await update.effective_message.reply_text(
        "📜 <b>VPN Order History</b>\n\nEnter user's Chat ID or @username to view their Pending orders:",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return VSHOP_ORDER_LOOKUP


async def vshop_order_history_lookup(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    row = user_profile_row_by_query(update.effective_message.text or "")
    if not row:
        await update.effective_message.reply_text("❌ User not found. Try again:")
        return VSHOP_ORDER_LOOKUP

    target_id = int(row["user_id"])
    orders = DB.fetchall(
        "SELECT * FROM vshop_orders WHERE user_id=? AND status IN ('pending','partial') ORDER BY id DESC LIMIT 20",
        (target_id,),
    )
    if not orders:
        await update.effective_message.reply_text(
            f"✅ No pending VPN orders for user <code>{target_id}</code>.",
            parse_mode=ParseMode.HTML,
            reply_markup=vshop_admin_menu(),
        )
        return ConversationHandler.END

    buttons = []
    for o in orders:
        label = _vshop_order_id_label(o)
        buttons.append([InlineKeyboardButton(label, callback_data=f"vshop:adminorder:{o['id']}")])

    await update.effective_message.reply_text(
        f"📜 Pending VPN orders for <code>{target_id}</code>:",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(buttons),
    )
    return ConversationHandler.END


async def vshop_admin_order_detail(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Admin clicks on a pending order to Provide or Cancel+Refund."""
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        oid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid order", show_alert=True)
        return
    order = DB.fetchone("SELECT * FROM vshop_orders WHERE id=?", (oid,))
    if not order:
        await q.answer("Order not found", show_alert=True)
        return
    svc = DB.fetchone("SELECT * FROM vshop_services WHERE id=?", (order["service_id"],))
    cat = DB.fetchone("SELECT name FROM vshop_categories WHERE id=?", (order["category_id"],))
    user_row = DB.fetchone("SELECT first_name,username FROM users WHERE user_id=?", (order["user_id"],))
    uname = f"@{user_row['username']}" if user_row and user_row["username"] else str(order["user_id"])
    status = str(order["status"] or "pending").upper()
    await q.answer()
    await replace_callback_message(q, context.bot, 
        f"📦 <b>VPN Order #{oid}</b>\n"
        f"📂 Category: {html.escape(str(cat['name'] if cat else '?'))}\n"
        f"📦 Service: {html.escape(str(svc['name'] if svc else '?'))}\n"
        f"👤 User: {html.escape(uname)} (<code>{order['user_id']}</code>)\n"
        f"🔢 Ordered: {order['quantity']}\n"
        f"🤖 Auto Delivered: {order['auto_delivered']}\n"
        f"⏳ Pending: {order['pending_quantity']}\n"
        f"💰 Total: {money(int(order['total_units'] or 0))}\n"
        f"📌 Status: {status}",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("📤 Provide", callback_data=f"vshop:provide:{oid}", style="primary")],
            [InlineKeyboardButton("❌ Cancel + Refund", callback_data=f"vshop:cancelorder:{oid}", style="danger")],
            [InlineKeyboardButton("⬅️ Back", callback_data="vshop:adminmenu")],
        ]),
    )


async def vshop_provide_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Admin taps Provide — ask for delivery content."""
    if not await require_admin(update):
        return ConversationHandler.END
    q = update.callback_query
    try:
        oid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid", show_alert=True)
        return ConversationHandler.END
    order = DB.fetchone("SELECT * FROM vshop_orders WHERE id=? AND status IN ('pending','partial')", (oid,))
    if not order:
        await q.answer("Order not found or already completed.", show_alert=True)
        return ConversationHandler.END
    context.user_data["vshop_provide_oid"] = oid
    await q.answer()
    pending = int(order["pending_quantity"] or 0)
    await q.message.reply_text(
        f"📤 <b>Provide for Order #{oid}</b>\n"
        f"⏳ Pending: <b>{pending} pcs</b>\n\n"
        "Send the Account/Details/Message you want to deliver to the user:",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return VSHOP_PROVIDE_CONTENT


async def vshop_provide_content(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Admin sent the delivery content — ask for confirmation."""
    content = (update.effective_message.text or "").strip()
    if not content:
        await update.effective_message.reply_text("❌ Send the content to deliver (accounts, messages, etc.)")
        return VSHOP_PROVIDE_CONTENT
    oid = int(context.user_data.get("vshop_provide_oid", 0))
    context.user_data["vshop_provide_content"] = content
    await update.effective_message.reply_text(
        f"⚠️ <b>Confirm Delivery?</b>\n\n"
        f"📤 Content preview:\n<code>{html.escape(content[:800])}</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [
                InlineKeyboardButton("✅ Confirm Provide", callback_data=f"vshop:doconfirmprovide:{oid}", style="primary"),
                InlineKeyboardButton("❌ Cancel", callback_data=f"vshop:provide:{oid}", style="danger"),
            ],
        ]),
    )
    return VSHOP_PROVIDE_CONFIRM


async def vshop_provide_confirm_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Admin confirmed delivery."""
    if not await require_admin(update):
        return ConversationHandler.END
    q = update.callback_query
    try:
        oid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid", show_alert=True)
        return ConversationHandler.END

    content = context.user_data.pop("vshop_provide_content", "")
    order = DB.fetchone("SELECT * FROM vshop_orders WHERE id=?", (oid,))
    if not order:
        await q.answer("Order not found", show_alert=True)
        return ConversationHandler.END

    ts = now_ts()
    DB.execute(
        "UPDATE vshop_orders SET status='success', pending_quantity=0, delivery_info=?, delivered_at=? WHERE id=?",
        (content[:4000], ts, oid),
    )
    user_id = int(order["user_id"])
    svc = DB.fetchone("SELECT name, duration_days FROM vshop_services WHERE id=?", (order["service_id"],))
    svc_name = str(svc["name"]) if svc else "?"
    duration = int(svc["duration_days"] or 0) if svc else 0
    dur_line = f"\n⏳ Validity: {duration} Days" if duration else ""
    cat = DB.fetchone("SELECT name FROM vshop_categories WHERE id=?", (order["category_id"],))
    cat_name = str(cat["name"]) if cat else "?"
    user_row = DB.fetchone(
        "SELECT username, first_name, balance_units FROM users WHERE user_id=?",
        (user_id,),
    )
    username = (
        f"@{user_row['username']}"
        if user_row and user_row["username"]
        else str(user_row["first_name"] or user_id) if user_row else str(user_id)
    )
    await _vshop_delete_notice(context.bot, order)
    await q.answer("✅ Delivered!")
    # Notify user
    with suppress(TelegramError):
        await context.bot.send_message(
            chat_id=user_id,
            text=(
                f"✅ <b>Order #{oid} Delivered!</b>\n"
                f"📦 Service: {html.escape(svc_name)}\n"
                f"🔢 Quantity: {int(order['pending_quantity'])} pcs{dur_line}\n\n"
                f"📋 <b>Details:</b>\n<code>{html.escape(content)}</code>\n\n"
                "⚠️ Save this info!"
            ),
            parse_mode=ParseMode.HTML,
        )
    # Notify notice group
    await _vshop_notify(
        context.bot,
        _vshop_notice_text(
            order_id=oid,
            user_id=user_id,
            username=username,
            category=cat_name,
            service=svc_name,
            quantity=int(order["pending_quantity"] or 0),
            total=int(order["total_units"] or 0),
            status="SUCCESS (Manual Provide)",
            status_icon="✅",
            extra_lines=["📤 Provided: Complete"],
            balance_units=int(user_row["balance_units"] or 0) if user_row else 0,
        ),
    )
    await replace_callback_message(q, context.bot, 
        f"✅ Order #{oid} delivered successfully.",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("⬅️ Back", callback_data="vshop:adminmenu"),
        ]]),
    )
    return ConversationHandler.END


async def vshop_cancel_order_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Admin cancels a pending order and refunds."""
    if not await require_admin(update):
        return
    q = update.callback_query
    try:
        oid = int(q.data.split(":")[-1])
    except Exception:
        await q.answer("Invalid", show_alert=True)
        return
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            order = DB.conn.execute(
                "SELECT * FROM vshop_orders WHERE id=? AND status IN ('pending','partial')", (oid,)
            ).fetchone()
            if not order:
                DB.conn.rollback()
                await q.answer("Order not found or not cancellable.", show_alert=True)
                return
            user_id = int(order["user_id"])
            total = int(order["total_units"] or 0)
            ts = now_ts()
            DB.conn.execute(
                "UPDATE vshop_orders SET status='cancelled', pending_quantity=0 WHERE id=?",
                (oid,),
            )
            DB.conn.execute(
                "UPDATE users SET balance_units=balance_units+?,updated_at=? WHERE user_id=?",
                (total, ts, user_id),
            )
            DB.conn.execute(
                "INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at) VALUES(?,?,?,?,?,?)",
                (user_id, total, "vshop_refund", str(oid), f"VPN order #{oid} cancelled+refunded", ts),
            )
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            await q.answer("Refund failed. Try again.", show_alert=True)
            return

    svc = DB.fetchone("SELECT name FROM vshop_services WHERE id=?", (order["service_id"],))
    svc_name = str(svc["name"]) if svc else "?"
    balance_row = DB.fetchone(
        "SELECT balance_units FROM users WHERE user_id=?",
        (user_id,),
    )
    balance_after_refund = (
        int(balance_row["balance_units"] or 0) if balance_row else 0
    )
    await _vshop_delete_notice(context.bot, order)
    # Notify user
    with suppress(TelegramError):
        await context.bot.send_message(
            chat_id=user_id,
            text=(
                f"↩️ <b>Order #{oid} Cancelled & Refunded</b>\n"
                f"📦 Service: {html.escape(svc_name)}\n"
                f"💰 Refund: <b>{money(total)}</b>\n\n"
                "পরিশোধিত টাকা আপনার Balance-এ ফেরত দেওয়া হয়েছে।"
            ),
            parse_mode=ParseMode.HTML,
        )
    await _vshop_notify(
        context.bot,
        f"↩️ <b>VPN Order #{oid} — CANCELLED + REFUNDED</b>\n"
        f"👤 User: <code>{user_id}</code>\n"
        f"📦 Service: {html.escape(svc_name)}\n"
        f"💰 Refund: {money(total)}\n"
        f"💳 Current Balance: {money(balance_after_refund)}\n"
        f"📌 Status: ❌ CANCELLED",
    )
    await q.answer("✅ Cancelled and refunded.")
    await replace_callback_message(q, context.bot, 
        f"✅ Order #{oid} cancelled and {money(total)} refunded to user.",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("⬅️ Back", callback_data="vshop:adminmenu"),
        ]]),
    )



# =============================================================================
# SMSBower MAIL — separate 25-minute email activation workflow
# =============================================================================
MAIL_API_BASE = "https://smsbower.page/api/mail"
MAIL_BUY_LOCKS: set[int] = set()
MAIL_ACTION_LOCKS: set[int] = set()
MAIL_LATE_WATCHES: set[tuple[int, str, str]] = set()


def ensure_mail_schema() -> None:
    DB.execute("""CREATE TABLE IF NOT EXISTS mail_services (
        id INTEGER PRIMARY KEY AUTOINCREMENT, service_code TEXT NOT NULL UNIQUE,
        service_name TEXT NOT NULL, profit_percent REAL NOT NULL DEFAULT 0,
        profit_usd TEXT NOT NULL DEFAULT '0',
        enabled INTEGER NOT NULL DEFAULT 1, created_at INTEGER NOT NULL)""")
    mail_columns = {row["name"] for row in DB.fetchall("PRAGMA table_info(mail_services)")}
    if "profit_usd" not in mail_columns:
        DB.execute("ALTER TABLE mail_services ADD COLUMN profit_usd TEXT NOT NULL DEFAULT '0'")
    DB.execute("""CREATE TABLE IF NOT EXISTS mail_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL,
        service_id INTEGER NOT NULL, service_code TEXT NOT NULL, service_name TEXT NOT NULL,
        domain TEXT NOT NULL, email TEXT NOT NULL DEFAULT '', activation_id TEXT NOT NULL UNIQUE,
        provider_cost_units INTEGER NOT NULL, charged_units INTEGER NOT NULL,
        status TEXT NOT NULL DEFAULT 'waiting', created_at INTEGER NOT NULL,
        expires_at INTEGER NOT NULL, updated_at INTEGER NOT NULL,
        refunded INTEGER NOT NULL DEFAULT 0, last_code TEXT,
        FOREIGN KEY(user_id) REFERENCES users(user_id))""")
    DB.execute("CREATE INDEX IF NOT EXISTS idx_mail_orders_user ON mail_orders(user_id,id DESC)")
    DB.execute("CREATE INDEX IF NOT EXISTS idx_mail_orders_status ON mail_orders(status,expires_at)")


ensure_mail_schema()


def mail_api_key() -> str:
    return _panel_api_key("smsbower")


async def mail_api(action: str, **params: Any) -> tuple[Any, str]:
    key = mail_api_key()
    if not key or key.upper() in {"XXX", "PASTE_API_KEY"}:
        raise RuntimeError("SMSBower API key is not configured by admin.")
    url = f"{MAIL_API_BASE}/{action}"
    query = {"api_key": key, **{k: v for k, v in params.items() if v is not None}}
    timeout = aiohttp.ClientTimeout(total=HTTP_TIMEOUT_SECONDS)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        async with session.get(url, params=query) as response:
            raw = (await response.text()).strip()
            if response.status >= 400:
                raise RuntimeError(f"SMSBower mail HTTP {response.status}: {short_text(raw,250)}")
    try:
        data = json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        data = raw
    if isinstance(data, dict) and (data.get("status") in (0, "0", "error", "failed")):
        raise RuntimeError(str(data.get("error") or data.get("message") or "Mail provider rejected request"))
    return data, raw


def _mail_data_list(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        for key in ("services", "data", "domains", "result"):
            value = data.get(key)
            if isinstance(value, list): return [x for x in value if isinstance(x, dict)]
            if isinstance(value, dict):
                # Service/domain maps returned by getPriceRests.
                return [{"domain": k, **(v if isinstance(v, dict) else {"price": v})} for k,v in value.items()]
    return []


def mail_menu_markup() -> InlineKeyboardMarkup:
    services = DB.fetchall("SELECT id,service_name FROM mail_services WHERE enabled=1 ORDER BY service_name COLLATE NOCASE")
    rows = [[InlineKeyboardButton("📧 " + short_text(r["service_name"], 32), callback_data=f"mail:svc:{r['id']}")] for r in services]
    rows += [[InlineKeyboardButton("📋 My Mail Orders", callback_data="mail:orders")]]
    return InlineKeyboardMarkup(rows)


async def show_mail_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_terms(update): return
    msg = update.effective_message
    if not msg: return
    if not mail_api_key():
        await msg.reply_text("📧 Mail service is temporarily unavailable. Please contact admin.", reply_markup=number_buy_menu()); return
    services = DB.fetchall("SELECT COUNT(*) AS n FROM mail_services WHERE enabled=1")
    if not services or not int(services[0]["n"]):
        await msg.reply_text("No Mail services are available yet.", reply_markup=number_buy_menu()); return
    await msg.reply_text("📧 <b>Mail Services</b>\nSelect a service:", parse_mode=ParseMode.HTML, reply_markup=mail_menu_markup())


async def mail_home_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q=update.callback_query
    await q.answer()
    await q.edit_message_text("📧 <b>Mail Services</b>",parse_mode=ParseMode.HTML,reply_markup=mail_menu_markup())


async def mail_service_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q=update.callback_query; await q.answer()
    try: sid=int(q.data.rsplit(":",1)[1])
    except Exception: return
    svc=DB.fetchone("SELECT * FROM mail_services WHERE id=? AND enabled=1",(sid,))
    if not svc: await q.answer("Service unavailable",show_alert=True); return
    try:
        data,raw=await mail_api("getPriceRests",service=svc["service_code"])
        domains=_mail_data_list(data)
        # Accept the documented nested Service -> Domain price/count shape.
        if isinstance(data,dict) and isinstance(data.get("data"),dict):
            root=data["data"].get(svc["service_code"],data["data"].get(svc["service_name"],{}))
            if isinstance(root,dict): domains=[{"domain":k,**(v if isinstance(v,dict) else {"price":v})} for k,v in root.items()]
        rows=[]
        for d in domains:
            domain=str(d.get("domain") or d.get("Domain") or d.get("name") or d.get("id") or "").strip()
            if not domain: continue
            cost=decimal_value(d.get("price",d.get("Price",d.get("cost",0))))
            count=d.get("count",d.get("Count",d.get("stock","?")))
            # SMSBower returns USD prices. Apply the configured percentage margin
            # in USD first, then convert the final amount using the admin USD→BDT rate.
            usd_to_bdt = exchange_rate()
            charged=bdt_to_units((cost+decimal_value(svc["profit_usd"]))*usd_to_bdt)
            rows.append([InlineKeyboardButton(f"🌐 {domain} • {money(charged)} • stock {count}",callback_data=f"mail:buy:{sid}:{domain[:30]}")])
        if not rows:
            await q.edit_message_text("No available domains for this service right now.\n\n"+html.escape(short_text(raw,500)),reply_markup=mail_menu_markup()); return
        rows.append([InlineKeyboardButton("⬅️ Mail Services",callback_data="mail:home")])
        await q.edit_message_text(f"📧 <b>{html.escape(svc['service_name'])}</b>\nChoose a domain:",parse_mode=ParseMode.HTML,reply_markup=InlineKeyboardMarkup(rows))
    except Exception as e:
        logger.exception("Mail price/domain lookup failed")
        await q.edit_message_text("Could not load mail domains: "+html.escape(short_text(str(e),350)),reply_markup=mail_menu_markup())


async def mail_watch_and_cancel_late_stock(application: Any, user_id: int, service_code: str, domain: str) -> None:
    """Watch temporarily for stock that appears after No mails yet, then cancel it without charging the user."""
    watch_key = (user_id, service_code, domain)
    if watch_key in MAIL_LATE_WATCHES:
        return
    MAIL_LATE_WATCHES.add(watch_key)
    try:
        # Check once per minute for up to 25 minutes. If the provider allocates
        # a late mailbox, immediately cancel its returned activation ID.
        for _ in range(25):
            await asyncio.sleep(60)
            try:
                data, raw = await mail_api("getActivation", service=service_code, domain=domain)
            except Exception as exc:
                if str(exc).strip().casefold() == "no mails yet":
                    continue
                logger.warning("Late mail stock watch error for %s/%s: %s", service_code, domain, exc)
                continue
            src = data if isinstance(data, dict) else {}
            activation_id = str(src.get("mailId") or src.get("mail_id") or src.get("id") or "").strip()
            if not activation_id:
                logger.warning("Late mail watch got response without activation ID: %s", short_text(raw, 250))
                continue
            cancelled = False
            for attempt in range(3):
                try:
                    cancel_data, cancel_raw = await mail_api("setStatus", id=activation_id, status=2)
                    cancel_text = cancel_raw.upper()
                    cancelled = (
                        "SUCCESS" in cancel_text or "CANCEL" in cancel_text or
                        (isinstance(cancel_data, dict) and str(cancel_data.get("status")).lower() in {"1", "success", "ok", "cancelled"})
                    )
                    if cancelled:
                        break
                except Exception as cancel_error:
                    logger.error("Could not cancel late mail activation %s (attempt %s): %s", activation_id, attempt + 1, cancel_error)
                if attempt < 2:
                    await asyncio.sleep(5)
            if cancelled:
                logger.info("Late mail activation %s auto-cancelled for user %s", activation_id, user_id)
                with suppress(Exception):
                    await application.bot.send_message(
                        chat_id=user_id,
                        text="ℹ️ A mailbox became available late after your earlier unavailable request. It was automatically cancelled; your bot balance was never charged."
                    )
            else:
                logger.critical("Late mail activation %s was detected but provider cancellation was not confirmed", activation_id)
                with suppress(Exception):
                    await application.bot.send_message(
                        chat_id=user_id,
                        text="⚠️ A mailbox appeared late, but the provider did not confirm its cancellation. Your bot balance was not charged. Please contact support."
                    )
            return
    finally:
        MAIL_LATE_WATCHES.discard(watch_key)


async def mail_buy_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q=update.callback_query; user=q.from_user
    try: _,_,sid_s,domain=q.data.split(":",3); sid=int(sid_s)
    except Exception: await q.answer("Invalid selection",show_alert=True); return
    if user.id in MAIL_BUY_LOCKS: await q.answer("Purchase already processing",show_alert=True); return
    svc=DB.fetchone("SELECT * FROM mail_services WHERE id=? AND enabled=1",(sid,))
    if not svc: await q.answer("Service unavailable",show_alert=True); return
    MAIL_BUY_LOCKS.add(user.id)
    await q.answer("Requesting mail activation…")
    try:
        # Recheck live price immediately before purchase.
        data,_=await mail_api("getPriceRests",service=svc["service_code"])
        domains=_mail_data_list(data)
        if isinstance(data,dict) and isinstance(data.get("data"),dict):
            root=data["data"].get(svc["service_code"],data["data"].get(svc["service_name"],{}))
            if isinstance(root,dict): domains=[{"domain":k,**(v if isinstance(v,dict) else {"price":v})} for k,v in root.items()]
        chosen=next((d for d in domains if str(d.get("domain") or d.get("Domain") or d.get("name") or d.get("id") or "")==domain),None)
        if not chosen: await q.message.reply_text("Domain is no longer available. Please select again."); return
        cost=decimal_value(chosen.get("price",chosen.get("Price",chosen.get("cost",0))))
        # Keep Mail order accounting in BDT units, matching the user's wallet.
        # Provider price and configured percentage margin are calculated in USD,
        # then the configured exchange rate converts the total to BDT.
        usd_to_bdt = exchange_rate()
        cost_units=bdt_to_units(cost*usd_to_bdt)
        charge=bdt_to_units((cost+decimal_value(svc["profit_usd"]))*usd_to_bdt)
        with DB.lock:
            row=DB.conn.execute("SELECT balance_units FROM users WHERE user_id=?",(user.id,)).fetchone()
        if not row or int(row["balance_units"] or 0)<charge:
            await q.message.reply_text(f"Insufficient balance. Required: {money(charge)}"); return
        # Give the provider at most 30 seconds to allocate a mailbox. Show
        # visible progress and refresh it every 10 seconds while retrying stock.
        purchase = None
        raw = ""
        last_stock_error = None
        processing_msg = await q.message.reply_text("⏳ Processing your mail request...")
        deadline = asyncio.get_running_loop().time() + 30
        next_progress_at = 10
        started_at = asyncio.get_running_loop().time()
        while asyncio.get_running_loop().time() < deadline:
            remaining = deadline - asyncio.get_running_loop().time()
            try:
                purchase, raw = await asyncio.wait_for(
                    mail_api("getActivation", service=svc["service_code"], domain=domain),
                    timeout=min(10, max(0.1, remaining)),
                )
                last_stock_error = None
                break
            except asyncio.TimeoutError:
                last_stock_error = RuntimeError("Mail provider request timed out")
            except Exception as activation_error:
                if str(activation_error).strip().casefold() != "no mails yet":
                    raise
                last_stock_error = activation_error

            now_loop = asyncio.get_running_loop().time()
            elapsed = now_loop - started_at
            # Avoid hammering the provider when it immediately answers "No mails yet".
            wait_for_progress = max(0, next_progress_at - elapsed)
            if now_loop < deadline and wait_for_progress > 0:
                await asyncio.sleep(min(wait_for_progress, deadline - now_loop))
            elapsed = asyncio.get_running_loop().time() - started_at
            if elapsed >= next_progress_at and elapsed < 30:
                dots = "." * ((int(elapsed // 10) % 3) + 1)
                with suppress(Exception):
                    await processing_msg.edit_text(f"⏳ Processing your mail request{dots}")
                next_progress_at += 10

        if last_stock_error is not None:
            logger.info("Mail stock did not appear within 30-second window: service=%s domain=%s", svc["service_code"], domain)
            context.application.create_task(
                mail_watch_and_cancel_late_stock(context.application, user.id, svc["service_code"], domain)
            )
            with suppress(Exception):
                await processing_msg.edit_text(
                    "📭 Mail not available right now. Your balance was not charged. "
                    "If the provider allocates this mailbox late, the bot will try to cancel it automatically."
                )
            return
        with suppress(Exception):
            await processing_msg.edit_text("✅ Mail received. Preparing your order...")
        src=purchase if isinstance(purchase,dict) else {}
        activation_id=str(src.get("mailId") or src.get("mail_id") or src.get("id") or "").strip()
        email_addr=str(src.get("mail") or src.get("email") or "").strip()
        if not activation_id or not email_addr:
            await q.message.reply_text("Provider did not return a valid mail activation. No balance was charged.\n"+html.escape(short_text(raw,350))); return
        now=now_ts(); expires=now+25*60
        try:
            with DB.lock:
                DB.conn.execute("BEGIN IMMEDIATE")
                cur=DB.conn.execute("UPDATE users SET balance_units=balance_units-?,updated_at=? WHERE user_id=? AND balance_units>=?",(charge,now,user.id,charge))
                if cur.rowcount!=1: raise RuntimeError("Balance changed; purchase stopped")
                DB.conn.execute("INSERT INTO mail_orders(user_id,service_id,service_code,service_name,domain,email,activation_id,provider_cost_units,charged_units,status,created_at,expires_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,'waiting',?,?,?)",(user.id,sid,svc["service_code"],svc["service_name"],domain,email_addr,activation_id,cost_units,charge,now,expires,now))
                oid=DB.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                DB.conn.execute("INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at) VALUES(?,?,?,?,?,?)",(user.id,-charge,"mail_purchase",str(oid),f"Mail order #{oid}",now))
                DB.conn.commit()
        except Exception:
            with DB.lock: DB.conn.rollback()
            # Provider activation exists but could not be safely charged/saved: request cancellation.
            with suppress(Exception): await mail_api("setStatus",id=activation_id,status=2)
            raise
        order=DB.fetchone("SELECT * FROM mail_orders WHERE id=?",(oid,))
        await q.message.reply_text(mail_order_text(order),parse_mode=ParseMode.HTML,reply_markup=mail_order_markup(order))
    except Exception as e:
        error_text = str(e).strip()
        if error_text.casefold() == "no mails yet":
            # SMSBower can return this when the selected domain has no
            # allocatable mailbox. This is an availability response, not a
            # bot crash; the wallet is debited only after activation succeeds.
            logger.warning(
                "Mail activation unavailable: provider has no mail for service=%s domain=%s",
                svc["service_code"],
                domain,
            )
            await q.message.reply_text(
                "📭 No mail is available for this domain right now. "
                "Your balance was not charged. Please choose another domain or try again later."
            )
        else:
            logger.exception("Mail purchase failed")
            await q.message.reply_text(
                "❌ Mail purchase failed. " + html.escape(short_text(error_text,350))
            )
    finally: MAIL_BUY_LOCKS.discard(user.id)


def mail_order_text(o: sqlite3.Row) -> str:
    remaining=max(0,int(o["expires_at"])-now_ts())
    return (f"📧 <b>Mail Order #{o['id']}</b>\n\n✉️ Email: <code>{html.escape(o['email'])}</code>\n"
            f"📋 Service: {html.escape(o['service_name'])}\n🌐 Domain: {html.escape(o['domain'])}\n"
            f"💳 Paid: <b>{money(int(o['charged_units']))}</b>\n⏳ Time left: <b>{remaining//60:02d}:{remaining%60:02d}</b>\n"
            f"📌 Status: <b>{html.escape(o['status'].upper())}</b>"+
            (f"\n\n🔐 Latest OTP: <code>{html.escape(o['last_code'])}</code>" if o["last_code"] else "\n\nWaiting for OTP…"))


def mail_order_markup(o: sqlite3.Row) -> InlineKeyboardMarkup:
    rows=[[copy_button("📋 Copy Email",o["email"])],
          [InlineKeyboardButton("🔄 Receive OTP",callback_data=f"mail:code:{o['id']}")],
          [InlineKeyboardButton("📨 Another Message",callback_data=f"mail:again:{o['id']}")]]
    if o["status"]=="waiting" and not int(o["refunded"]) and not o["last_code"]:
        rows.append([InlineKeyboardButton("↩️ Cancel & Refund",callback_data=f"mail:cancel:{o['id']}")])
    rows.append([InlineKeyboardButton("⬅️ My Mail Orders",callback_data="mail:orders")])
    return InlineKeyboardMarkup(rows)


async def mail_order_action(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q=update.callback_query; uid=q.from_user.id
    parts=(q.data or "").split(":"); action=parts[1]
    try: oid=int(parts[2])
    except Exception: await q.answer("Invalid order",show_alert=True); return
    o=DB.fetchone("SELECT * FROM mail_orders WHERE id=? AND user_id=?",(oid,uid))
    if not o: await q.answer("Order not found",show_alert=True); return
    if action=="copy":
        await q.answer("Email address is shown above; tap and copy it.",show_alert=True); return
    if uid in MAIL_ACTION_LOCKS: await q.answer("Please wait…",show_alert=True); return
    MAIL_ACTION_LOCKS.add(uid)
    try:
        if action in ("code","again"):
            if o["status"]!="waiting" or now_ts()>=int(o["expires_at"]):
                await q.answer("This mail activation has expired.",show_alert=True); return
            if action=="again":
                await mail_api("setStatus",id=o["activation_id"],status=5)
            data,raw=await mail_api("getCode",mailId=o["activation_id"])
            src=data if isinstance(data,dict) else {}
            code=str(src.get("code") or src.get("sms") or src.get("message") or "").strip()
            if not code and isinstance(data,str): code=data.strip()
            if code and code not in {"0","null","None"} and not (isinstance(src,dict) and src.get("status")==0):
                DB.execute("UPDATE mail_orders SET last_code=?,updated_at=? WHERE id=?",(code,now_ts(),oid))
                o=DB.fetchone("SELECT * FROM mail_orders WHERE id=?",(oid,))
                await q.answer("OTP received!")
                await q.message.reply_text("🔐 <b>New Mail OTP</b>\n<code>"+html.escape(code)+"</code>",parse_mode=ParseMode.HTML,reply_markup=InlineKeyboardMarkup([[copy_button("📋 Copy OTP",code)]]))
            else:
                await q.answer("No new OTP received yet. Try again later.",show_alert=True)
            await q.edit_message_text(mail_order_text(o),parse_mode=ParseMode.HTML,reply_markup=mail_order_markup(o)); return
        if action=="cancel":
            if o["status"]!="waiting" or int(o["refunded"]) or o["last_code"]:
                await q.answer("An OTP has already arrived. This order is no longer refundable.",show_alert=True); return
            if now_ts()>=int(o["expires_at"]):
                DB.execute("UPDATE mail_orders SET status='expired',updated_at=? WHERE id=?",(now_ts(),oid)); await q.answer("The 25-minute activation period has expired; refund is unavailable.",show_alert=True); return
            data,raw=await mail_api("setStatus",id=o["activation_id"],status=2)
            upper=raw.upper()
            ok=("SUCCESS" in upper or "CANCEL" in upper or (isinstance(data,dict) and str(data.get("status")).lower() in {"1","success","ok","cancelled"}))
            if not ok: await q.answer("Provider did not confirm cancellation. No refund was issued.",show_alert=True); return
            with DB.lock:
                DB.conn.execute("BEGIN IMMEDIATE")
                cur=DB.conn.execute("UPDATE mail_orders SET status='cancelled',refunded=1,updated_at=? WHERE id=? AND status='waiting' AND refunded=0",(now_ts(),oid))
                if cur.rowcount:
                    DB.conn.execute("UPDATE users SET balance_units=balance_units+?,updated_at=? WHERE user_id=?",(int(o["charged_units"]),now_ts(),uid))
                    DB.conn.execute("INSERT INTO transactions(user_id,amount_units,kind,reference,note,created_at) VALUES(?,?,?,?,?,?)",(uid,int(o["charged_units"]),"mail_refund",str(oid),f"Mail refund #{oid}",now_ts()))
                DB.conn.commit()
            await q.answer("Cancelled and refunded.",show_alert=True)
            await q.edit_message_text(f"↩️ Mail order #{oid} cancelled. {money(int(o['charged_units']))} refunded to your balance.")
    except Exception as e:
        error_text = str(e).strip()
        normalized = error_text.casefold()
        if "code has not been received yet" in normalized:
            logger.info("Mail OTP still pending for order %s", oid)
            with suppress(TelegramError):
                await q.answer("⌛ OTP has not arrived yet. Please try again shortly.", show_alert=True)
        elif action == "again" and "bad actual activation status" in normalized:
            logger.warning("Provider rejected another-message request for mail order %s: %s", oid, short_text(error_text, 180))
            with suppress(TelegramError):
                await q.answer("Provider cannot request another message in the current activation status. Please wait for the current OTP.", show_alert=True)
        else:
            logger.exception("Mail order action failed")
            with suppress(TelegramError):
                await q.answer("Mail provider error: "+short_text(error_text,150),show_alert=True)
    finally: MAIL_ACTION_LOCKS.discard(uid)


async def mail_orders_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Show only currently active mail activations, like Active Number List."""
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    now = now_ts()

    # Expired orders stay in the database for accounting/history, but are
    # removed from this active list.
    DB.execute(
        "UPDATE mail_orders SET status='expired', updated_at=? "
        "WHERE user_id=? AND status='waiting' AND expires_at<=? AND refunded=0",
        (now, uid, now),
    )
    rows = DB.fetchall(
        "SELECT * FROM mail_orders WHERE user_id=? AND status='waiting' "
        "AND refunded=0 AND expires_at>? ORDER BY created_at DESC LIMIT 50",
        (uid, now),
    )
    if not rows:
        await q.edit_message_text(
            "📧 <b>My Mail Service — Active Mails</b>\n\nএখন কোনো active mail নেই।",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Mail Services", callback_data="mail:home")]]),
        )
        return

    buttons = []
    for order in rows:
        remaining = max(0, int(order["expires_at"]) - now)
        label = f"{'🔐' if order['last_code'] else '⏳'} {order['email'] or order['domain']}"
        if order["last_code"]:
            label += f" • OTP {order['last_code']}"
        label += f" • {remaining // 60:02d}:{remaining % 60:02d}"
        buttons.append([InlineKeyboardButton(label[:60], callback_data=f"mail:view:{order['id']}")])
    buttons.append([InlineKeyboardButton("⬅️ Mail Services", callback_data="mail:home")])
    await q.edit_message_text(
        "📧 <b>My Mail Service — Active Mails</b>\n"
        "শুধু ২৫ মিনিটের মধ্যে থাকা active mail-গুলো এখানে দেখাবে।\n"
        "মেইল নির্বাচন করে OTP নিতে পারবেন।",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(buttons),
    )


async def mail_view_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    try:
        oid = int(q.data.rsplit(":", 1)[1])
    except Exception:
        await q.answer("Invalid mail activation.", show_alert=True)
        return
    o = DB.fetchone(
        "SELECT * FROM mail_orders WHERE id=? AND user_id=? AND status='waiting' "
        "AND refunded=0 AND expires_at>?",
        (oid, q.from_user.id, now_ts()),
    )
    if not o:
        await q.answer("এই mail activation-এর ২৫ মিনিটের সময় শেষ হয়েছে অথবা এটি আর active নেই।", show_alert=True)
        return
    await q.answer()
    await q.edit_message_text(mail_order_text(o), parse_mode=ParseMode.HTML, reply_markup=mail_order_markup(o))


async def mail_admin_add_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update): return
    context.user_data["mail_admin_step"]="code"
    await update.effective_message.reply_text("Mail Service form opened.", reply_markup=ReplyKeyboardRemove())
    await update.effective_message.reply_text(
        "Send SMSBower Mail Service code (e.g. service ID):",
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("✖️ Cancel", callback_data="mail:admin:cancel")]])
    )


async def mail_admin_cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q: await q.answer("Admin only", show_alert=True)
        return
    for key in ("mail_admin_step", "mail_admin_code", "mail_admin_name", "mail_admin_prices", "mail_admin_edit_id"):
        context.user_data.pop(key, None)
    await q.answer("Add service cancelled")
    await q.edit_message_text("✖️ Adding Mail Service cancelled. No changes were saved.")
    await q.message.reply_text("📧 Mail Service Management", reply_markup=admin_menu())


async def mail_admin_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # This handler is registered before unknown_text in the same PTB handler
    # group, so it must pass pending Panel Management API-key input through.
    # Otherwise every text message is consumed here and panel_api_key_input()
    # never gets a chance to validate/save the new key.
    if await panel_api_key_input(update, context):
        return
    step=context.user_data.get("mail_admin_step")
    if not step: return
    if not await require_admin(update): return
    value=(update.effective_message.text or "").strip()
    cancel_markup = InlineKeyboardMarkup([[InlineKeyboardButton("✖️ Cancel", callback_data="mail:admin:cancel")]])
    if step == "edit_margin":
        try:
            margin = decimal_value(value.replace("$", "").strip())
            if margin < 0 or margin > Decimal("1000"):
                raise ValueError("margin out of range")
        except Exception:
            await update.effective_message.reply_text("Invalid amount. Send a USD amount such as 0.01.")
            return
        sid = int(context.user_data.get("mail_admin_edit_id", 0))
        DB.execute("UPDATE mail_services SET profit_usd=? WHERE id=?", (str(margin), sid))
        context.user_data.pop("mail_admin_step", None)
        context.user_data.pop("mail_admin_edit_id", None)
        await update.effective_message.reply_text(
            f"✅ Profit margin updated to ${margin:.4f} USD.", reply_markup=admin_menu()
        )
        return
    if step=="code":
        if not value or len(value)>40:
            await update.effective_message.reply_text("Invalid service code. Try again.", reply_markup=cancel_markup); return
        context.user_data["mail_admin_code"]=value
        # Prices are domain-specific in SMSBower; show every live base price before asking for margin.
        try:
            price_data, _ = await mail_api("getPriceRests", service=value)
            price_rows = _mail_data_list(price_data)
            if isinstance(price_data, dict) and isinstance(price_data.get("data"), dict):
                root = price_data["data"].get(value, {})
                if isinstance(root, dict):
                    price_rows = [{"domain": k, **(v if isinstance(v, dict) else {"price": v})} for k, v in root.items()]
            lines=[]
            for item in price_rows:
                dom=str(item.get("domain") or item.get("Domain") or item.get("name") or item.get("id") or "").strip()
                price=item.get("price", item.get("Price", item.get("cost")))
                if dom and price is not None:
                    lines.append(f"• {html.escape(dom)}: ${decimal_value(price):.4f} USD")
            context.user_data["mail_admin_prices"] = lines
            price_text = "\n".join(lines) if lines else "No live domain price was returned for this service."
        except Exception as exc:
            logger.warning("Could not fetch default mail prices for service %s: %s", value, exc)
            context.user_data["mail_admin_prices"] = []
            price_text = "Could not load the live default price right now."
        context.user_data["mail_admin_step"]="name"
        await update.effective_message.reply_text(
            "📊 <b>Provider base prices (before profit margin)</b>\n" + price_text +
            "\n\nSend display name for this Mail service:",
            parse_mode=ParseMode.HTML, reply_markup=cancel_markup
        )
        return
    if step=="name":
        if not value:
            await update.effective_message.reply_text("Service name cannot be empty. Send a display name:", reply_markup=cancel_markup); return
        context.user_data["mail_admin_name"]=value[:80]
        context.user_data["mail_admin_step"]="profit"
        price_lines = context.user_data.get("mail_admin_prices") or []
        price_text = "\n".join(price_lines) if price_lines else "Live base price: unavailable; it varies by domain."
        await update.effective_message.reply_text(
            "📊 <b>Default/base price</b> (before margin):\n" + price_text +
            "\n\nNow send fixed profit margin in USD (e.g. 0.01 for one cent):",
            parse_mode=ParseMode.HTML, reply_markup=cancel_markup
        )
        return
    if step=="profit":
        try:
            margin = decimal_value(value.replace("$", "").strip())
            if margin < 0 or margin > Decimal("1000"):
                raise ValueError("margin out of range")
        except Exception:
            await update.effective_message.reply_text(
                "Enter a valid USD amount, e.g. 0.01 (one cent). Percentage is not used.",
                reply_markup=cancel_markup
            )
            return
        try:
            DB.execute(
                "INSERT INTO mail_services(service_code,service_name,profit_usd,enabled,created_at) VALUES(?,?,?,1,?)",
                (context.user_data["mail_admin_code"], context.user_data["mail_admin_name"], str(margin), now_ts())
            )
        except sqlite3.IntegrityError:
            await update.effective_message.reply_text("That service code already exists. Send another service code or cancel.", reply_markup=cancel_markup)
            context.user_data["mail_admin_step"] = "code"
            return
        for k in ("mail_admin_step","mail_admin_code","mail_admin_name","mail_admin_prices"):
            context.user_data.pop(k,None)
        await update.effective_message.reply_text(
            "✅ Mail service added. Domains and their base prices are fetched automatically from SMSBower.",
            reply_markup=admin_menu()
        )


def _mail_admin_list_markup() -> InlineKeyboardMarkup:
    rows = DB.fetchall("SELECT * FROM mail_services ORDER BY id DESC")
    buttons = [
        [InlineKeyboardButton(
            f"{'✅' if r['enabled'] else '⛔'} {short_text(r['service_name'], 28)} • +${decimal_value(r['profit_usd']):.4f}",
            callback_data=f"mail:admin:detail:{r['id']}"
        )]
        for r in rows
    ]
    buttons.append([InlineKeyboardButton("➕ Add Mail Service", callback_data="mail:admin:add")])
    return InlineKeyboardMarkup(buttons)


async def mail_admin_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    msg = update.effective_message
    if msg:
        await msg.reply_text(
            "📧 <b>Mail Service Management</b>\\nSelect a service to manage it:",
            parse_mode=ParseMode.HTML, reply_markup=_mail_admin_list_markup()
        )


async def mail_admin_open_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Open the mail-service admin list from the admin panel."""
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q:
            await q.answer("Admin only", show_alert=True)
        return
    await q.answer()
    await q.edit_message_text(
        "📧 <b>Mail Service Management</b>\nSelect a service to manage it:",
        parse_mode=ParseMode.HTML,
        reply_markup=_mail_admin_list_markup(),
    )


async def mail_admin_detail_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q: await q.answer("Admin only", show_alert=True)
        return
    await q.answer()
    sid = int(q.data.rsplit(":", 1)[1])
    svc = DB.fetchone("SELECT * FROM mail_services WHERE id=?", (sid,))
    if not svc:
        await q.edit_message_text("Service not found.", reply_markup=_mail_admin_list_markup())
        return
    await q.edit_message_text(
        f"📧 <b>{html.escape(svc['service_name'])}</b>\\n"
        f"Service code: <code>{html.escape(svc['service_code'])}</code>\\n"
        f"Status: {'✅ Enabled' if svc['enabled'] else '⛔ Disabled'}\\n"
        f"Fixed profit margin: <b>${decimal_value(svc['profit_usd']):.4f} USD</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔴 Turn OFF" if svc["enabled"] else "🟢 Turn ON", callback_data=f"mail:admin:toggle:{sid}")],
            [InlineKeyboardButton("✏️ Edit profit margin", callback_data=f"mail:admin:edit:{sid}")],
            [InlineKeyboardButton("🗑 Delete service", callback_data=f"mail:admin:delete:{sid}")],
            [InlineKeyboardButton("⬅️ Back to services", callback_data="mail:admin:list")]
        ])
    )


async def mail_admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q: await q.answer("Admin only", show_alert=True)
        return
    sid = int(q.data.rsplit(":", 1)[1])
    DB.execute("UPDATE mail_services SET enabled=CASE enabled WHEN 1 THEN 0 ELSE 1 END WHERE id=?", (sid,))
    await q.answer("Service status updated")
    svc = DB.fetchone("SELECT * FROM mail_services WHERE id=?", (sid,))
    if not svc:
        await q.edit_message_text("Service not found.", reply_markup=_mail_admin_list_markup())
        return
    await q.edit_message_text(
        f"📧 <b>{html.escape(svc['service_name'])}</b>\n"
        f"Service code: <code>{html.escape(svc['service_code'])}</code>\n"
        f"Status: {'✅ Enabled' if svc['enabled'] else '⛔ Disabled'}\n"
        f"Fixed profit margin: <b>${decimal_value(svc['profit_usd']):.4f} USD</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔴 Turn OFF" if svc["enabled"] else "🟢 Turn ON", callback_data=f"mail:admin:toggle:{sid}")],
            [InlineKeyboardButton("✏️ Edit profit margin", callback_data=f"mail:admin:edit:{sid}")],
            [InlineKeyboardButton("🗑 Delete service", callback_data=f"mail:admin:delete:{sid}")],
            [InlineKeyboardButton("⬅️ Back to services", callback_data="mail:admin:list")]
        ])
    )


async def mail_admin_edit_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q: await q.answer("Admin only", show_alert=True)
        return
    sid = int(q.data.rsplit(":", 1)[1])
    svc = DB.fetchone("SELECT id, service_name, profit_usd FROM mail_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Service not found", show_alert=True); return
    context.user_data["mail_admin_step"] = "edit_margin"
    context.user_data["mail_admin_edit_id"] = sid
    await q.answer()
    await q.message.reply_text(
        f"Current margin for {svc['service_name']}: ${decimal_value(svc['profit_usd']):.4f} USD\\n"
        "Send the new fixed margin in USD (example: 0.01)."
    )


async def mail_admin_delete_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q: await q.answer("Admin only", show_alert=True)
        return
    sid = int(q.data.rsplit(":", 1)[1])
    svc = DB.fetchone("SELECT service_name FROM mail_services WHERE id=?", (sid,))
    if not svc:
        await q.answer("Service not found", show_alert=True); return
    await q.answer()
    await q.edit_message_text(
        f"Delete <b>{html.escape(svc['service_name'])}</b>? This removes it from the Mail service list. Existing orders remain in history.",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Yes, delete", callback_data=f"mail:admin:deleteconfirm:{sid}")],
            [InlineKeyboardButton("Cancel", callback_data=f"mail:admin:detail:{sid}")]
        ])
    )


async def mail_admin_delete_confirm_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q: await q.answer("Admin only", show_alert=True)
        return
    sid = int(q.data.rsplit(":", 1)[1])
    DB.execute("DELETE FROM mail_services WHERE id=?", (sid,))
    await q.answer("Service deleted")
    await q.edit_message_text("🗑 Service deleted.", reply_markup=_mail_admin_list_markup())


async def mail_admin_back_list_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q: await q.answer("Admin only", show_alert=True)
        return
    await q.answer()
    await q.edit_message_text("📧 <b>Mail Service Management</b>\\nSelect a service:",
                              parse_mode=ParseMode.HTML, reply_markup=_mail_admin_list_markup())


def main_menu(user_id: Optional[int] = None) -> ReplyKeyboardMarkup:
    rows = [
        [
            KeyboardButton("🔢 Number Buy"),
            KeyboardButton("💰 My Balance", style="success"),
        ],
        [KeyboardButton("💳 Add Balance", style="danger"), KeyboardButton("📚 How to use")],
        [
            KeyboardButton("🆘 Support"),
            KeyboardButton("🛡️ 𝙱𝚄𝚈 𝚅𝙿𝙽/𝙿𝚁𝙾𝚇𝙸𝙴𝚂", style="danger"),
        ],
    ]
    if is_admin(user_id):
        rows.append([KeyboardButton("🛠 Admin Panel")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


def admin_referral_settings_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("💵 Referral Reward"), KeyboardButton("📏 Referral Deposit Limit")],
            [KeyboardButton("🔙 General Settings")],
        ],
        resize_keyboard=True,
    )


def admin_how_to_use_settings_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("➕ Set Button"), KeyboardButton("🎬 Set Video")],
            [KeyboardButton("🔙 General Settings")],
        ],
        resize_keyboard=True,
    )


async def show_admin_referral_settings(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    reward = bdt_to_units(DB.get_setting("referral_reward_bdt", "0"))
    threshold = bdt_to_units(DB.get_setting("referral_min_deposit_bdt", "0"))
    await update.effective_message.reply_text(
        "<b>Referral Settings</b>\n\n"
        f"Per referral reward: <b>{money(reward)}</b>\n"
        f"Minimum qualifying deposit: <b>{money(threshold)}</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(
                        "💵 Set Referral Reward",
                        callback_data="admin:referral:reward",
                    )
                ],
                [
                    InlineKeyboardButton(
                        "📏 Set Deposit Limit",
                        callback_data="admin:referral:threshold",
                    )
                ],
                [
                    InlineKeyboardButton(
                        "🔙 General Settings",
                        callback_data="admin:general",
                    )
                ],
            ]
        ),
    )


async def set_referral_reward_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    if update.callback_query:
        await update.callback_query.answer()
    current = bdt_to_units(DB.get_setting("referral_reward_bdt", "0"))
    await update.effective_message.reply_text(
        f"Current referral reward: <b>{money(current)}</b>\n"
        "Send the amount in BDT that the referrer should earn for each qualifying user.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_REFERRAL_REWARD


async def set_referral_reward_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    try:
        value = bdt_input_value(update.effective_message.text)
        if value < 0:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text(
            "Send a valid non-negative BDT amount."
        )
        return SET_REFERRAL_REWARD
    DB.set_setting("referral_reward_bdt", str(value))
    await update.effective_message.reply_text(
        f"✅ Referral reward saved: <b>{money(bdt_to_units(value))}</b>",
        parse_mode=ParseMode.HTML,
    )
    await show_admin_referral_settings(update, context)
    return ConversationHandler.END


async def set_referral_threshold_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    if update.callback_query:
        await update.callback_query.answer()
    current = bdt_to_units(DB.get_setting("referral_min_deposit_bdt", "0"))
    await update.effective_message.reply_text(
        f"Current qualifying deposit limit: <b>{money(current)}</b>\n"
        "Send the minimum deposit in BDT. A referred user must deposit at least this amount.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_REFERRAL_THRESHOLD


async def set_referral_threshold_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    try:
        value = bdt_input_value(update.effective_message.text)
        if value <= 0:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text(
            "Send a valid positive BDT amount, for example 30 or 40."
        )
        return SET_REFERRAL_THRESHOLD
    DB.set_setting("referral_min_deposit_bdt", str(value))
    await update.effective_message.reply_text(
        f"✅ Referral deposit limit saved: <b>{money(bdt_to_units(value))}</b>",
        parse_mode=ParseMode.HTML,
    )
    await show_admin_referral_settings(update, context)
    return ConversationHandler.END


async def show_admin_how_to_use_settings(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    rows = DB.fetchall(
        "SELECT id,name,video_file_id FROM how_to_use_buttons ORDER BY id ASC"
    )
    if rows:
        listing = "\n".join(
            f"• {html.escape(str(row['name']))} — "
            f"{'video set' if row['video_file_id'] else 'video not set'}"
            for row in rows
        )
    else:
        listing = "No buttons saved yet."
    action_rows = [
        [
            InlineKeyboardButton(
                "➕ Add Inline Button",
                callback_data="howadmin:add",
            ),
            InlineKeyboardButton(
                "🎬 Set Video",
                callback_data="howadmin:video_menu",
            ),
        ],
    ]
    action_rows.extend(
        [
            [
                InlineKeyboardButton(
                    f"🗑 Delete {short_text(str(row['name']), 30)}",
                    callback_data=f"howadmin:delete:{row['id']}",
                )
            ]
            for row in rows
        ]
    )
    action_rows.append(
        [InlineKeyboardButton("🔙 General Settings", callback_data="admin:general")]
    )
    await update.effective_message.reply_text(
        "<b>How to use settings</b>\n\n"
        f"{listing}\n\n"
        "The saved topics appear as inline buttons under the user's "
        "📚 How to use menu.",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup(action_rows),
    )


async def set_how_button_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    if update.callback_query:
        await update.callback_query.answer()
    await update.effective_message.reply_text(
        "What name should this How to use button have?\n"
        "Example: <code>How to buy a number</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_HOW_BUTTON_NAME


async def set_how_button_name(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    name = (update.effective_message.text or "").strip()
    if not name or len(name) > 64:
        await update.effective_message.reply_text(
            "Send a button name between 1 and 64 characters."
        )
        return SET_HOW_BUTTON_NAME
    cur = DB.execute(
        "INSERT INTO how_to_use_buttons(name,video_file_id,created_at) VALUES(?,?,?)",
        (name, None, now_ts()),
    )
    saved = DB.fetchone(
        "SELECT id FROM how_to_use_buttons WHERE id=?",
        (cur.lastrowid,),
    )
    if not saved:
        await update.effective_message.reply_text(
            "❌ The button could not be saved. Please try again."
        )
        return ConversationHandler.END
    await update.effective_message.reply_text(
        f"✅ Inline button <b>{html.escape(name)}</b> saved.\n"
        "Use Set Video to attach its video.",
        parse_mode=ParseMode.HTML,
    )
    await show_admin_how_to_use_settings(update, context)
    return ConversationHandler.END


async def set_how_video_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    if update.callback_query:
        await update.callback_query.answer()
    rows = DB.fetchall(
        "SELECT id,name,video_file_id FROM how_to_use_buttons ORDER BY id ASC"
    )
    if not rows:
        await update.effective_message.reply_text(
            "Add a How to use button first.",
        )
        await show_admin_how_to_use_settings(update, context)
        return
    buttons = [
        [
            InlineKeyboardButton(
                f"{row['name']} — {'replace video' if row['video_file_id'] else 'set video'}",
                callback_data=f"howadmin:video:{row['id']}",
            )
        ]
        for row in rows
    ]
    await update.effective_message.reply_text(
        "Choose the button whose video you want to set:",
        reply_markup=InlineKeyboardMarkup(buttons),
    )


async def admin_select_how_video(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    query = update.callback_query
    if not query or not await require_admin(update):
        return ConversationHandler.END
    await query.answer()
    try:
        button_id = int(str(query.data).rsplit(":", 1)[1])
    except (TypeError, ValueError):
        return ConversationHandler.END
    row = DB.fetchone(
        "SELECT name FROM how_to_use_buttons WHERE id=?", (button_id,)
    )
    if not row:
        await replace_callback_message(query, context.bot, "That button no longer exists.")
        return ConversationHandler.END
    context.user_data["how_video_button_id"] = button_id
    await replace_callback_message(query, context.bot, 
        f"Send the video for <b>{html.escape(str(row['name']))}</b> now.",
        parse_mode=ParseMode.HTML,
    )
    return SET_HOW_VIDEO_FILE


async def save_how_video(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    message = update.effective_message
    file_id = message.video.file_id if message and message.video else None
    if not file_id:
        await message.reply_text("Please send a video file, or use /cancel.")
        return SET_HOW_VIDEO_FILE
    button_id = int(context.user_data.pop("how_video_button_id", 0) or 0)
    cur = DB.execute(
        "UPDATE how_to_use_buttons SET video_file_id=? WHERE id=?",
        (file_id, button_id),
    )
    if cur.rowcount != 1:
        await message.reply_text("That button no longer exists.")
        return ConversationHandler.END
    await message.reply_text(
        "✅ Video saved.",
    )
    await show_admin_how_to_use_settings(update, context)
    return ConversationHandler.END


async def delete_how_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    await query.answer()
    try:
        button_id = int(str(query.data).rsplit(":", 1)[1])
    except (TypeError, ValueError):
        return
    DB.execute("DELETE FROM how_to_use_buttons WHERE id=?", (button_id,))
    await replace_callback_message(query, context.bot, "✅ Button and its saved video were deleted.")
    await show_admin_how_to_use_settings(update, context)


def number_buy_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("🛒 Buy Number"), KeyboardButton("🇺🇸 USA Local")],
            [KeyboardButton("📧 Mail", style="danger")],
            [KeyboardButton("🔙 Main Menu")],
        ],
        resize_keyboard=True,
    )


async def show_number_buy_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_terms(update):
        return
    await update.effective_message.reply_text(
        "🔢 <b>Number Buy</b>", parse_mode=ParseMode.HTML, reply_markup=number_buy_menu()
    )


def local_user_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("🛒 Buy Local"), KeyboardButton("📱 Active Numbers")],
            [KeyboardButton("🔙 Main Menu")],
        ],
        resize_keyboard=True,
    )


def admin_balance_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("💵 Provider Balance"), KeyboardButton("🇺🇸 Local Balance")],
            [KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


def admin_usa_settings_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("💵 Local Price"), KeyboardButton("⏱ Local Session")],
            [KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


def admin_deposit_settings_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("💵 Minimum Deposit")],
            [KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


async def show_admin_group_settings(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    await update.effective_message.reply_text(
        "Group Settings",
        reply_markup=admin_group_settings_menu(),
    )


async def show_admin_balance_menu(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    await update.effective_message.reply_text(
        "Balance",
        reply_markup=admin_balance_menu(),
    )


async def show_admin_lists_menu(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    await update.effective_message.reply_text(
        "Lists",
        reply_markup=admin_lists_menu(),
    )


async def show_admin_usa_settings(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    await update.effective_message.reply_text(
        "USA Settings",
        reply_markup=admin_usa_settings_menu(),
    )


async def show_admin_deposit_settings(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    await update.effective_message.reply_text(
        "Deposit Settings",
        reply_markup=admin_deposit_settings_menu(),
    )


async def show_admin_general_settings(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    if update.callback_query:
        await update.callback_query.answer()
    await update.effective_message.reply_text(
        "General Settings",
        reply_markup=admin_general_settings_menu(),
    )


async def show_admin_database_settings(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    await update.effective_message.reply_text(
        "Database Settings",
        reply_markup=admin_database_settings_menu(),
    )


async def show_local_menu(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    if not await require_terms(update):
        return
    await update.effective_message.reply_text(
        "🇺🇸 USA Local",
        reply_markup=local_user_menu(),
    )


async def local_buy_abort(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    if update.callback_query:
        await update.callback_query.answer("Cancelled.")
        await replace_callback_message(
            update.callback_query, context.bot, "Cancelled.", reply_markup=local_user_menu()
        )


async def local_active_numbers(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    if not await require_terms(update):
        return
    text, markup, page = local_active_payload(
        update.effective_user.id,
        0,
    )
    sent = await update.effective_message.reply_text(
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=markup,
    )
    local_start_active_task(
        context.application,
        update.effective_user.id,
        sent.chat_id,
        sent.message_id,
        page,
    )


async def local_active_page_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    await query.answer()
    try:
        page = int(query.data.rsplit(":", 1)[1])
    except Exception:
        page = 0
    key = (query.message.chat_id, query.message.message_id)
    local_cancel_view_task(LOCAL_DETAIL_TASKS, key)
    for order_id in list(LOCAL_ORDER_VIEWS):
        local_unregister_view(order_id, key)
    text, markup, page = local_active_payload(user.id, page)
    await replace_callback_message(query, context.bot, 
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=markup,
    )
    local_start_active_task(
        context.application,
        user.id,
        query.message.chat_id,
        query.message.message_id,
        page,
    )


async def local_detail_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    await query.answer()
    try:
        _, _, order_raw, page_raw = query.data.split(":", 3)
        order_id = int(order_raw)
        page = int(page_raw)
    except Exception:
        await query.answer("Invalid number.", show_alert=True)
        return

    order = local_get_user_order(order_id, user.id)
    if not order:
        await query.answer("Number not found.", show_alert=True)
        return

    await replace_callback_message(query, context.bot, 
        local_detail_text(order),
        parse_mode=ParseMode.HTML,
        reply_markup=local_detail_markup(order, page),
    )
    local_start_detail_task(
        context.application,
        order_id,
        query.message.chat_id,
        query.message.message_id,
        page,
    )
    if order["status"] == "active":
        ensure_local_poll_task(context.application, order_id)


async def admin_local_balance(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    if not await require_admin(update):
        return
    if not LOCAL_API:
        await update.effective_message.reply_text("Local API is not ready.")
        return
    waiting = await update.effective_message.reply_text("⏳ Checking…")
    ok, value = await LOCAL_API.get_balance()
    await waiting.edit_text(
        f"🇺🇸 Local balance: <b>{html.escape(value)}</b>"
        if ok
        else f"❌ {html.escape(value)}",
        parse_mode=ParseMode.HTML,
    )


async def set_local_price_start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    await update.effective_message.reply_text(
        f"Current local price: <b>{money(local_price_units())}</b>\n"
        "Enter new BDT price.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_LOCAL_PRICE


async def set_local_price_value(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    try:
        value = decimal_value(update.effective_message.text)
        if value <= 0:
            raise ValueError
        units = bdt_to_units(value)
    except Exception:
        await update.effective_message.reply_text(
            "Enter a valid positive BDT amount.",
            reply_markup=current_workflow_keyboard(update, context),
        )
        return SET_LOCAL_PRICE
    DB.set_setting("local_price_bdt", str(units_to_decimal(units)))
    await update.effective_message.reply_text(
        f"✅ Local price: {money(units)}",
        reply_markup=admin_usa_settings_menu(),
    )
    return ConversationHandler.END


def parse_local_session_input(raw: str) -> int:
    value = raw.strip().casefold().replace(" ", "")
    if value.endswith("h"):
        hours = decimal_value(value[:-1])
        minutes = int(
            (hours * Decimal(60)).quantize(
                Decimal("1"),
                rounding=ROUND_HALF_UP,
            )
        )
    elif value.endswith("m"):
        minutes = int(decimal_value(value[:-1]))
    else:
        minutes = int(decimal_value(value))
    if not 10 <= minutes <= 7 * 24 * 60:
        raise ValueError
    return minutes


async def set_local_session_start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    current = local_session_minutes()
    await update.effective_message.reply_text(
        f"Current session: <b>{local_format_hm(current * 60)}</b>\n"
        "Enter minutes or hours, for example: <code>720</code> or <code>12h</code>.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SET_LOCAL_SESSION


async def set_local_session_value(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    try:
        minutes = parse_local_session_input(
            update.effective_message.text or ""
        )
    except Exception:
        await update.effective_message.reply_text(
            "Use 10–10080 minutes, or values such as 12h.",
            reply_markup=current_workflow_keyboard(update, context),
        )
        return SET_LOCAL_SESSION
    DB.set_setting("local_session_minutes", str(minutes))
    await update.effective_message.reply_text(
        f"✅ Local session: {local_format_hm(minutes * 60)}",
        reply_markup=admin_usa_settings_menu(),
    )
    return ConversationHandler.END



# =============================================================================
# HIGH-LOAD RUNTIME OVERRIDES
# =============================================================================

_ORDER_PROCESS_LOCKS: dict[int, asyncio.Lock] = {}


def _order_process_lock(order_id: int) -> asyncio.Lock:
    lock = _ORDER_PROCESS_LOCKS.get(order_id)
    if lock is None:
        lock = asyncio.Lock()
        _ORDER_PROCESS_LOCKS[order_id] = lock
    return lock


async def _poll_one_order(
    application: Application,
    order: sqlite3.Row,
    semaphore: asyncio.Semaphore,
) -> None:
    async with semaphore:
        try:
            await process_single_order(application, order)
            refreshed = DB.fetchone(
                "SELECT * FROM orders WHERE id=?",
                (int(order["id"]),),
            )
            if refreshed:
                await maybe_refresh_countdown(application, refreshed)
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Order poll failed for order %s", order["id"])


_LOCAL_POLL_SUPERVISOR_KEY = 0
_LOCAL_ORDER_CHECK_LOCKS: dict[int, asyncio.Lock] = {}


async def _local_poll_one(
    application: Application,
    order: sqlite3.Row,
    semaphore: asyncio.Semaphore,
) -> None:
    order_id = int(order["id"])
    lock = _LOCAL_ORDER_CHECK_LOCKS.setdefault(order_id, asyncio.Lock())
    async with semaphore:
        async with lock:
            current = local_get_order(order_id)
            if not current or current["status"] != "active":
                _LOCAL_ORDER_CHECK_LOCKS.pop(order_id, None)
                return
            if int(current["expires_at"]) <= now_ts():
                local_set_status(order_id, "expired")
                await local_refresh_all_views(application, order_id)
                _LOCAL_ORDER_CHECK_LOCKS.pop(order_id, None)
                return
            if LOCAL_API:
                result = await LOCAL_API.get_local_otp(
                    str(current["phone"]),
                    seen_otps=local_seen_otps(order_id),
                )
                await local_apply_otp(application, order_id, result)


async def local_poll_supervisor(application: Application) -> None:
    """Use one bounded scheduler instead of one permanent task per local order."""
    semaphore = asyncio.Semaphore(LOCAL_PROVIDER_MAX_CONCURRENCY)
    try:
        while True:
            cycle_started = time.monotonic()
            orders = local_all_active_orders()
            if orders:
                await asyncio.gather(
                    *(
                        _local_poll_one(application, order, semaphore)
                        for order in orders
                    )
                )
            elapsed = time.monotonic() - cycle_started
            await asyncio.sleep(max(0.5, LOCAL_OTP_POLL_SECONDS - elapsed))
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.exception("Local OTP poll supervisor crashed")
    finally:
        LOCAL_ORDER_POLL_TASKS.pop(_LOCAL_POLL_SUPERVISOR_KEY, None)


DB_USER_IMPORT_FILE = 117
BROADCAST_CONTENT = 118
BROADCAST_CONFIRM = 119
ADMIN_USER_RESTRICT_DURATION = 120

BROADCAST_DELAY_SECONDS = max(0.06, float(os.getenv("BROADCAST_DELAY_SECONDS", "0.10")))


def ensure_latest_feature_schema() -> None:
    """Apply non-destructive migrations for warnings and timed restrictions."""
    with DB.lock:
        user_columns = {
            row["name"] for row in DB.conn.execute("PRAGMA table_info(users)")
        }
        migrations = (
            (
                "restricted_until",
                "ALTER TABLE users ADD COLUMN restricted_until INTEGER NOT NULL DEFAULT 0",
            ),
            (
                "warning_count",
                "ALTER TABLE users ADD COLUMN warning_count INTEGER NOT NULL DEFAULT 0",
            ),
            (
                "last_warning_at",
                "ALTER TABLE users ADD COLUMN last_warning_at INTEGER",
            ),
        )
        for column, ddl in migrations:
            if column not in user_columns:
                DB.conn.execute(ddl)
        DB.conn.commit()


ensure_latest_feature_schema()


def _remaining_text(seconds: int) -> str:
    seconds = max(0, int(seconds))
    days, seconds = divmod(seconds, 86400)
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    parts: list[str] = []
    if days:
        parts.append(f"{days}d")
    if hours:
        parts.append(f"{hours}h")
    if minutes:
        parts.append(f"{minutes}m")
    if seconds or not parts:
        parts.append(f"{seconds}s")
    return " ".join(parts[:3])


def user_restriction_until(user_id: int) -> int:
    row = DB.fetchone(
        "SELECT restricted_until FROM users WHERE user_id=?",
        (user_id,),
    )
    return int(row["restricted_until"] or 0) if row else 0


def purchase_restriction_message(user_id: int) -> str:
    restricted_until = user_restriction_until(user_id)
    remaining = restricted_until - now_ts()
    if remaining <= 0:
        if restricted_until:
            DB.execute(
                "UPDATE users SET restricted_until=0,updated_at=? WHERE user_id=?",
                (now_ts(), user_id),
            )
        return ""
    return (
        "⛔ <b>Number purchase temporarily restricted</b>\n\n"
        f"Remaining time: <b>{_remaining_text(remaining)}</b>\n"
        "Please contact support if you believe this is a mistake."
    )


# -----------------------------------------------------------------------------
# Short user order history (no XLSX/TXT file)
# -----------------------------------------------------------------------------


def _rotate_backup_files(prefix: str) -> None:
    os.makedirs(BACKUP_DIRECTORY, exist_ok=True)
    paths = sorted(
        (
            os.path.join(BACKUP_DIRECTORY, name)
            for name in os.listdir(BACKUP_DIRECTORY)
            if name.startswith(prefix) and not name.endswith(("-wal", "-shm"))
        ),
        key=lambda path: os.path.getmtime(path),
        reverse=True,
    )
    for old_path in paths[BACKUP_KEEP_COUNT:]:
        for candidate in (old_path, old_path + "-wal", old_path + "-shm"):
            with suppress(OSError):
                os.remove(candidate)


def _recovery_backup_score(summary: dict[str, int]) -> tuple[int, int]:
    return (
        int(summary.get("users", 0)),
        _database_activity_count(summary),
    )


def _update_recovery_anchor(snapshot_path: str, summary: dict[str, int]) -> None:
    """Keep one strong recovery point outside the rolling backup rotation."""
    anchor_path = os.path.join(BACKUP_DIRECTORY, DATABASE_RECOVERY_ANCHOR)
    candidates = _valid_database_backups()
    best_candidate = _select_recovery_backup(candidates)
    source_path = snapshot_path
    best_summary = summary
    if best_candidate:
        candidate_path, candidate_summary = best_candidate
        if _recovery_backup_score(candidate_summary) >= _recovery_backup_score(summary):
            source_path = candidate_path
            best_summary = candidate_summary

    if os.path.isfile(anchor_path):
        try:
            existing_summary = _database_file_summary(anchor_path, immutable=True)
            if _recovery_backup_score(best_summary) < _recovery_backup_score(existing_summary):
                return
        except Exception as exc:
            logger.warning("Replacing invalid recovery anchor: %s", exc)

    temp_path = anchor_path + ".tmp"
    for suffix in ("", "-wal", "-shm"):
        with suppress(FileNotFoundError):
            os.remove(temp_path + suffix)
    shutil.copy2(source_path, temp_path)
    _database_file_summary(temp_path, immutable=True)
    os.replace(temp_path, anchor_path)


def create_persistent_database_export(prefix: str = "database_export") -> str:
    os.makedirs(BACKUP_DIRECTORY, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(BACKUP_DIRECTORY, f"{prefix}_{timestamp}.db")
    database_snapshot_to(path)
    summary = database_validate_file(path)
    if prefix == "database_export":
        _update_recovery_anchor(path, summary)
    _rotate_backup_files(prefix + "_")
    return path


def create_users_uid_txt() -> str:
    os.makedirs(BACKUP_DIRECTORY, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(BACKUP_DIRECTORY, f"users_uid_list_{timestamp}.txt")
    rows = DB.fetchall(
        "SELECT user_id,username,first_name,created_at FROM users ORDER BY user_id"
    )
    with open(path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write("# Paid Number Bot user UID list\n")
        handle.write(f"# Exported: {datetime.now().isoformat(timespec='seconds')}\n")
        handle.write("# Format: user_id<TAB>username<TAB>first_name\n")
        for row in rows:
            username = str(row["username"] or "").replace("\t", " ").replace("\n", " ")
            first_name = str(row["first_name"] or "").replace("\t", " ").replace("\n", " ")
            handle.write(f"{int(row['user_id'])}\t{username}\t{first_name}\n")
    _rotate_backup_files("users_uid_list_")
    return path


async def admin_database_download(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    message = update.effective_message
    export_path = ""
    try:
        await message.reply_text("⏳ Creating database_export.db…")
        export_path = await run_file_job(create_persistent_database_export)
        with open(export_path, "rb") as handle:
            await context.bot.send_document(
                chat_id=message.chat_id,
                document=handle,
                filename="database_export.db",
                caption=(
                    "✅ Full database export\n"
                    "Includes users, balances, deposits, orders, settings and histories."
                ),
            )
        await message.reply_text(
            "✅ DB export completed.",
            reply_markup=admin_database_settings_menu(),
        )
    except RetryAfter as exc:
        logger.warning(
            "Telegram blocked DB download for %s seconds; export kept at %s",
            retry_after_seconds(exc),
            export_path,
        )
    except Exception as exc:
        logger.exception("Raw database export failed")
        with suppress(TelegramError):
            await message.reply_text(
                f"❌ DB export failed: {short_text(str(exc), 500)}",
                reply_markup=admin_database_settings_menu(),
            )


async def database_restore_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    await update.effective_message.reply_text(
        "Upload database_export.db (also accepts .sqlite or .sqlite3).\n"
        "The complete current database will be replaced after confirmation.",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return DB_RESTORE_FILE


def import_users_uid_txt(path: str) -> dict[str, int]:
    added = 0
    existing = 0
    invalid = 0
    seen: set[int] = set()
    ts = now_ts()
    with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
        lines = handle.readlines()

    parsed: list[tuple[int, str, str]] = []
    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        parts = re.split(r"[\t,|;]", line, maxsplit=2)
        match = re.search(r"-?\d{5,20}", parts[0])
        if not match:
            invalid += 1
            continue
        user_id = int(match.group(0))
        if user_id <= 0 or user_id in seen:
            invalid += 1
            continue
        seen.add(user_id)
        username = parts[1].strip().lstrip("@") if len(parts) > 1 else ""
        first_name = parts[2].strip() if len(parts) > 2 else ""
        parsed.append((user_id, username[:64], first_name[:128]))

    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            for user_id, username, first_name in parsed:
                old = DB.conn.execute(
                    "SELECT 1 FROM users WHERE user_id=?", (user_id,)
                ).fetchone()
                if old:
                    existing += 1
                    DB.conn.execute(
                        """
                        UPDATE users SET
                            username=CASE WHEN ?<>'' THEN ? ELSE username END,
                            first_name=CASE WHEN ?<>'' THEN ? ELSE first_name END,
                            updated_at=?
                        WHERE user_id=?
                        """,
                        (username, username, first_name, first_name, ts, user_id),
                    )
                else:
                    DB.conn.execute(
                        """
                        INSERT INTO users(
                            user_id,username,first_name,balance_units,accepted_terms,
                            is_blocked,created_at,updated_at,restricted_until,warning_count
                        ) VALUES(?,?,?,?,?,?,?,?,?,?)
                        """,
                        (user_id, username or None, first_name or None, 0, 0, 0, ts, ts, 0, 0),
                    )
                    added += 1
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            raise
    return {"added": added, "existing": existing, "invalid": invalid}


async def database_restore_confirm(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    """Restore a DB while keeping the safety copy on the VPS, not Telegram."""
    query = update.callback_query
    if not query or not await require_admin(update):
        return ConversationHandler.END
    await query.answer("Restoring…")

    candidate = context.user_data.pop("db_restore_candidate", None)
    filename = context.user_data.pop("db_restore_filename", "uploaded database")
    if not candidate or not os.path.exists(candidate):
        await replace_callback_message(query, context.bot, "❌ Restore file expired. Upload it again.")
        return ConversationHandler.END

    safety_db = ""
    stopped = False
    try:
        await replace_callback_message(query, context.bot, "⏳ Creating VPS safety backup…")
        safety_db = await run_file_job(
            create_persistent_database_export,
            "pre_restore",
        )
        await replace_callback_message(query, context.bot, "⏳ Restoring database…")
        await database_stop_background_tasks()
        stopped = True
        try:
            await run_file_job(database_reopen_from, candidate)
            ensure_latest_feature_schema()
        except Exception:
            logger.exception("Restore failed; rolling back")
            await run_file_job(database_reopen_from, safety_db)
            ensure_latest_feature_schema()
            raise
        finally:
            database_restart_background_tasks(context.application)
            stopped = False

        summary = await run_file_job(database_validate_file, database_absolute_path())
        await replace_callback_message(query, context.bot, 
            "✅ Database restored.\n"
            f"Source: {html.escape(filename)}\n"
            f"Users: {summary['users']}\n"
            f"Balance: {money(summary['balance_units'])}\n"
            f"VPS safety copy: <code>{html.escape(safety_db)}</code>",
            parse_mode=ParseMode.HTML,
        )
        await query.message.reply_text(
            "Database Settings",
            reply_markup=admin_database_settings_menu(),
        )
    except Exception as exc:
        logger.exception("Database restore failed")
        if stopped:
            with suppress(Exception):
                database_restart_background_tasks(context.application)
        await replace_callback_message(query, context.bot, 
            "❌ Restore failed. Existing data was kept or rolled back.\n"
            f"{short_text(str(exc), 700)}"
        )
    finally:
        with suppress(FileNotFoundError):
            os.remove(candidate)
    return ConversationHandler.END


# -----------------------------------------------------------------------------
# Broadcast with preview + confirmation and flood-safe pacing
# -----------------------------------------------------------------------------


async def broadcast_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    await update.effective_message.reply_text(
        "Send the text, photo, video, document or message you want to broadcast.",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return BROADCAST_CONTENT


async def broadcast_receive_content(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    message = update.effective_message
    context.user_data["broadcast_source_chat_id"] = message.chat_id
    context.user_data["broadcast_source_message_id"] = message.message_id

    await message.reply_text("<b>Broadcast preview</b>", parse_mode=ParseMode.HTML)
    with suppress(TelegramError):
        await context.bot.copy_message(
            chat_id=message.chat_id,
            from_chat_id=message.chat_id,
            message_id=message.message_id,
        )
    await message.reply_text(
        "Send this message to all users?",
        reply_markup=InlineKeyboardMarkup(
            [[
                InlineKeyboardButton("✅ Confirm Broadcast", callback_data="broadcast:confirm"),
                InlineKeyboardButton("❌ Cancel", callback_data="broadcast:cancel"),
            ]]
        ),
    )
    return BROADCAST_CONFIRM


async def broadcast_cancel_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    query = update.callback_query
    if query:
        await query.answer("Broadcast cancelled.")
        await replace_callback_message(query, context.bot, "Broadcast cancelled.")
    context.user_data.pop("broadcast_source_chat_id", None)
    context.user_data.pop("broadcast_source_message_id", None)
    await update.effective_message.reply_text(
        "Admin Panel",
        reply_markup=admin_menu(),
    )
    return ConversationHandler.END


async def _copy_broadcast_with_retry(
    bot: Any,
    user_id: int,
    source_chat_id: int,
    source_message_id: int,
) -> bool:
    try:
        await bot.copy_message(
            chat_id=user_id,
            from_chat_id=source_chat_id,
            message_id=source_message_id,
        )
        return True
    except RetryAfter as exc:
        wait = min(retry_after_seconds(exc) + 1, 120)
        logger.warning("Broadcast rate limit; sleeping %s seconds", wait)
        await asyncio.sleep(wait)
        try:
            await bot.copy_message(
                chat_id=user_id,
                from_chat_id=source_chat_id,
                message_id=source_message_id,
            )
            return True
        except TelegramError:
            return False
    except (Forbidden, BadRequest, TelegramError):
        return False


def user_profile_row_by_query(value: str) -> Optional[sqlite3.Row]:
    raw = value.strip()
    if re.fullmatch(r"\d+", raw):
        return user_profile_row(int(raw))
    username = raw.lstrip("@").strip()
    if not username:
        return None
    row = DB.fetchone(
        "SELECT user_id FROM users WHERE lower(username)=lower(?) LIMIT 1",
        (username,),
    )
    return user_profile_row(int(row["user_id"])) if row else None


def render_user_profile(row: sqlite3.Row) -> str:
    username = f"@{row['username']}" if row["username"] else "Not set"
    restricted_until = int(row["restricted_until"] or 0) if "restricted_until" in row.keys() else 0
    remaining = max(0, restricted_until - now_ts())
    restriction = f"Yes — {_remaining_text(remaining)} left" if remaining else "No"
    warning_count = int(row["warning_count"] or 0) if "warning_count" in row.keys() else 0
    return (
        "<b>User Profile</b>\n\n"
        f"👤 Name: <b>{html.escape(row['first_name'] or 'Unknown')}</b>\n"
        f"🔗 Username: <code>{html.escape(username)}</code>\n"
        f"🆔 Chat ID: <code>{row['user_id']}</code>\n"
        f"💰 Balance: <b>{money(int(row['balance_units'] or 0))}</b>\n"
        f"💳 Approved deposits: <b>{money(int(row['total_deposit_units'] or 0))}</b>\n"
        f"📦 Orders: <b>{int(row['total_orders'] or 0)}</b>\n"
        f"✅ Successful: <b>{int(row['successful_orders'] or 0)}</b>\n"
        f"📩 Today's OTP received: <b>{int(row['today_otp_count'] or 0)}</b>\n"
        f"💸 Today's spent: <b>{money(int(row['today_spent_units'] or 0))}</b>\n"
        f"⚠️ Warnings: <b>{warning_count}</b>\n"
        f"⏳ Restricted: <b>{html.escape(restriction)}</b>\n"
        f"🚫 Blocked: <b>{'Yes' if row['is_blocked'] else 'No'}</b>"
    )


def user_profile_keyboard(user_id: int) -> InlineKeyboardMarkup:
    row = DB.fetchone("SELECT restricted_until FROM users WHERE user_id=?", (user_id,))
    restricted = bool(row and int(row["restricted_until"] or 0) > now_ts())
    rows = [
        [
            InlineKeyboardButton("➕ Add Balance", callback_data=f"userbal:add:{user_id}"),
            InlineKeyboardButton("➖ Subtract", callback_data=f"userbal:sub:{user_id}"),
        ],
        [
            InlineKeyboardButton("⚠️ Warning", callback_data=f"userwarn:{user_id}"),
            InlineKeyboardButton("⏳ Restrict", callback_data=f"userrestrict:{user_id}"),
        ],
    ]
    if restricted:
        rows.append([
            InlineKeyboardButton("✅ Remove Restriction", callback_data=f"userunrestrict:{user_id}")
        ])
    rows.append([
        InlineKeyboardButton(
            "📜 My History",
            callback_data=f"umhistory:{user_id}",
        )
    ])
    rows.append([InlineKeyboardButton("🔎 Find Another User", callback_data="user:lookup")])
    return InlineKeyboardMarkup(rows)


async def admin_user_lookup_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    if update.callback_query:
        await update.callback_query.answer()
        message = update.callback_query.message
    else:
        message = update.effective_message
    await message.reply_text(
        "Enter a Telegram Chat ID or username (for example @username).",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return ADMIN_USER_LOOKUP_ID


async def admin_user_lookup_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    row = user_profile_row_by_query(update.effective_message.text or "")
    if not row:
        await update.effective_message.reply_text(
            "User not found. Enter a valid Chat ID or exact username."
        )
        return ADMIN_USER_LOOKUP_ID
    await update.effective_message.reply_text(
        render_user_profile(row),
        parse_mode=ParseMode.HTML,
        reply_markup=user_profile_keyboard(int(row["user_id"])),
    )
    await update.effective_message.reply_text("Admin menu", reply_markup=admin_menu())
    return ConversationHandler.END


async def admin_user_warning_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    try:
        user_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        return
    row = user_profile_row(user_id)
    if not row:
        await query.answer("User not found.", show_alert=True)
        return
    DB.execute(
        """
        UPDATE users SET warning_count=warning_count+1,last_warning_at=?,updated_at=?
        WHERE user_id=?
        """,
        (now_ts(), now_ts(), user_id),
    )
    warning_text = (
        "⚠️ <b>Account Warning</b>\n\n"
        "You are cancelling/refunding an unusually high number of phone numbers. "
        "Please request a number only when you intend to use it. Continued misuse "
        "may result in a temporary purchase restriction."
    )
    delivered = True
    try:
        await context.bot.send_message(user_id, warning_text, parse_mode=ParseMode.HTML)
    except TelegramError:
        delivered = False
    await query.answer("Warning sent." if delivered else "Warning saved; delivery failed.", show_alert=True)
    refreshed = user_profile_row(user_id)
    if refreshed:
        await replace_callback_message(query, context.bot, 
            render_user_profile(refreshed),
            parse_mode=ParseMode.HTML,
            reply_markup=user_profile_keyboard(user_id),
        )


async def admin_user_restrict_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    query = update.callback_query
    if not query or not await require_admin(update):
        return ConversationHandler.END
    try:
        user_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        return ConversationHandler.END
    if not user_profile_row(user_id):
        await query.answer("User not found.", show_alert=True)
        return ConversationHandler.END
    await query.answer()
    context.user_data["restrict_user_id"] = user_id
    await query.message.reply_text(
        "How long should purchases be restricted?\n"
        "Examples: <code>30</code> (minutes), <code>2h</code>, <code>1d</code>.",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return ADMIN_USER_RESTRICT_DURATION


def parse_restriction_duration(raw: str) -> int:
    value = raw.strip().lower().replace(" ", "")
    match = re.fullmatch(r"(\d+)(m|min|minute|minutes|h|hr|hour|hours|d|day|days)?", value)
    if not match:
        raise ValueError("Invalid duration")
    amount = int(match.group(1))
    unit = match.group(2) or "m"
    if amount <= 0:
        raise ValueError("Invalid duration")
    multiplier = 60
    if unit in {"h", "hr", "hour", "hours"}:
        multiplier = 3600
    elif unit in {"d", "day", "days"}:
        multiplier = 86400
    seconds = amount * multiplier
    return min(seconds, 365 * 86400)


async def admin_user_restrict_duration(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not is_admin(update.effective_user.id if update.effective_user else None):
        return ConversationHandler.END
    try:
        seconds = parse_restriction_duration(update.effective_message.text or "")
    except ValueError:
        await update.effective_message.reply_text("Enter 30, 2h or 1d.")
        return ADMIN_USER_RESTRICT_DURATION
    user_id = int(context.user_data.pop("restrict_user_id", 0) or 0)
    if not user_id:
        await update.effective_message.reply_text("Restriction session expired.")
        return ConversationHandler.END
    until = now_ts() + seconds
    DB.execute(
        "UPDATE users SET restricted_until=?,updated_at=? WHERE user_id=?",
        (until, now_ts(), user_id),
    )
    with suppress(TelegramError):
        await context.bot.send_message(
            user_id,
            "⛔ <b>Purchase Restriction Applied</b>\n\n"
            f"You cannot buy a new number for <b>{_remaining_text(seconds)}</b>.\n"
            "This restriction does not remove your existing balance or order history.",
            parse_mode=ParseMode.HTML,
        )
    row = user_profile_row(user_id)
    await update.effective_message.reply_text(
        f"✅ User restricted for {_remaining_text(seconds)}.",
        reply_markup=admin_menu(),
    )
    if row:
        await update.effective_message.reply_text(
            render_user_profile(row),
            parse_mode=ParseMode.HTML,
            reply_markup=user_profile_keyboard(user_id),
        )
    return ConversationHandler.END


async def admin_user_unrestrict_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    if not query or not await require_admin(update):
        return
    try:
        user_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        return
    DB.execute(
        "UPDATE users SET restricted_until=0,updated_at=? WHERE user_id=?",
        (now_ts(), user_id),
    )
    with suppress(TelegramError):
        await context.bot.send_message(
            user_id,
            "✅ Your number-purchase restriction has been removed.",
        )
    await query.answer("Restriction removed.", show_alert=True)
    row = user_profile_row(user_id)
    if row:
        await replace_callback_message(query, context.bot, 
            render_user_profile(row),
            parse_mode=ParseMode.HTML,
            reply_markup=user_profile_keyboard(user_id),
        )


# -----------------------------------------------------------------------------
# Standard-number quick buy / refund-and-new
# -----------------------------------------------------------------------------


async def _purchase_catalog_provider(
    user_id: int,
    item_id: int,
    provider_id: str,
) -> tuple[Optional[sqlite3.Row], str]:
    restriction = purchase_restriction_message(user_id)
    if restriction:
        return None, re.sub(r"<[^>]+>", "", restriction)
    if not API:
        return None, "Provider API is not ready."
    item = get_catalog_item(item_id)
    if not item or not item["enabled"]:
        return None, "This item is no longer available."
    try:
        offers = await API.get_provider_offers(
            item["country_id"], item["service_code"], catalog_extra(item)
        )
        selected = next(value for value in offers if value.provider_id == provider_id)
    except StopIteration:
        return None, "The same provider is not available now."
    except Exception as exc:
        return None, f"Live price error: {short_text(str(exc), 350)}"
    if selected.available <= 0:
        return None, "The same provider is out of stock."

    charged_units = usd_units_to_bdt_units(
        selected.provider_cost_units + item_markup_units(item)
    )
    reference = f"catalog:{item_id}:provider:{provider_id}:{now_ts()}"
    if not DB.hold_balance(user_id, charged_units, reference):
        balance = DB.fetchone("SELECT balance_units FROM users WHERE user_id=?", (user_id,))
        current = int(balance["balance_units"] or 0) if balance else 0
        return None, f"Insufficient balance. Need {money(charged_units)}, balance {money(current)}."
    try:
        purchased = await API.get_number(
            item["country_id"],
            item["service_code"],
            selected.provider_cost_units + 10,
            catalog_extra(item),
            provider_id=provider_id,
        )
    except Exception as exc:
        DB.credit_balance(
            user_id,
            charged_units,
            "purchase_release",
            reference,
            "Provider purchase failed; hold released",
        )
        return None, f"Purchase failed; balance restored. {short_text(str(exc), 350)}"

    balance_row = DB.fetchone("SELECT balance_units FROM users WHERE user_id=?", (user_id,))
    balance_after = int(balance_row["balance_units"] or 0) if balance_row else 0
    provider_cost_units = purchased.provider_cost_units or selected.provider_cost_units
    # Standard Buy Number is fixed at 20 minutes when no OTP has arrived.
    expires_in = 20 * 60
    ts = now_ts()
    try:
        cursor = DB.execute(
            """
            INSERT INTO orders(
                user_id,catalog_id,activation_id,phone,provider_id,country_name,
                service_name,provider_cost_units,charged_units,balance_after_units,
                status,created_at,expires_at,updated_at,last_api_status,message_id
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,NULL)
            """,
            (
                user_id,item_id,purchased.activation_id,purchased.phone,provider_id,
                item["country_name"],item["service_name"],provider_cost_units,
                charged_units,balance_after,"waiting",ts,ts+expires_in,ts,"PURCHASED",
            ),
        )
        order_id = int(cursor.lastrowid)
    except Exception as exc:
        cancelled = False
        with suppress(Exception):
            data, text = await API.set_status(purchased.activation_id, 8, {})
            cancelled = API.cancellation_succeeded(data, text)
        if cancelled:
            DB.credit_balance(
                user_id,charged_units,"purchase_release",reference,
                "Order save failed; provider cancellation succeeded",
            )
        logger.exception("Could not save provider order")
        return None, f"Could not save purchased number: {short_text(str(exc), 300)}"
    return DB.fetchone("SELECT * FROM orders WHERE id=?", (order_id,)), ""


async def _send_new_order_message(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    order: sqlite3.Row,
) -> Any:
    sent = await context.bot.send_message(
        chat_id=chat_id,
        text=render_order(order),
        parse_mode=ParseMode.HTML,
        reply_markup=order_keyboard(order),
    )
    DB.execute("UPDATE orders SET message_id=? WHERE id=?", (sent.message_id, order["id"]))
    # Start automatic OTP polling immediately for the number just purchased.
    # Without this selection, the shared poller may ignore the new order until
    # the user opens it from Active Numbers or presses another action.
    _select_order_for_auto_otp(int(order["user_id"]), int(order["id"]))
    COUNTDOWN_MINUTE_CACHE[int(order["id"])] = max(
        0, (int(order["expires_at"]) - now_ts()) // 60
    )
    return sent


async def buy_same_provider_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        return
    source = DB.fetchone("SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id))
    if not source or source["status"] != "otp_received":
        await query.answer("This order is not ready for quick buy.", show_alert=True)
        return
    lock = PURCHASE_LOCKS.setdefault(user.id, asyncio.Lock())
    if lock.locked():
        await query.answer("Another purchase is processing.", show_alert=True)
        return
    await query.answer("Buying from same provider…")
    async with lock:
        order, error = await _purchase_catalog_provider(
            user.id, int(source["catalog_id"]), str(source["provider_id"])
        )
        if not order:
            await query.message.reply_text(f"❌ {html.escape(error)}", parse_mode=ParseMode.HTML)
            return
        with suppress(TelegramError):
            await query.message.delete()
        await _send_new_order_message(context, query.message.chat_id, order)


def _local_sms_body(raw_response: str, otp: str) -> str:
    raw = str(raw_response or "").strip()
    data = safe_json_loads(raw, None)
    if data is not None:
        value = local_first_value_priority(
            data,
            ("sms", "sms_text", "message_text", "body", "text", "message"),
        )
        if value is not None:
            text = str(value).strip()
            if text and text != otp:
                return short_text(text, 500)
    if raw and len(raw) <= 500 and raw.casefold() not in {otp.casefold(), "success", "ok"}:
        return raw
    return "New SMS received."


async def local_buy_prompt(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    if not await require_terms(update):
        return
    # Local USA is optional. If no API client/key is configured, show a clear
    # user-facing notice and never continue to balance checks or provider calls.
    if LOCAL_API is None:
        await update.effective_message.reply_text(
            "🇺🇸 <b>USA Local service is currently unavailable.</b>\n\n"
            "Admin has not added the Local USA API yet. Please try again later.",
            parse_mode=ParseMode.HTML,
            reply_markup=local_user_menu(),
        )
        return
    restriction = purchase_restriction_message(update.effective_user.id)
    if restriction:
        await update.effective_message.reply_text(restriction, parse_mode=ParseMode.HTML)
        return
    user = DB.fetchone("SELECT balance_units FROM users WHERE user_id=?", (update.effective_user.id,))
    price = local_price_units()
    balance = int(user["balance_units"] or 0) if user else 0
    if balance < price:
        await update.effective_message.reply_text(
            f"❌ Need {money(price)}\nBalance: {money(balance)}",
            reply_markup=local_user_menu(),
        )
        return
    await update.effective_message.reply_text(
        f"Buy USA Local?\n💳 {money(price)} • ⏱ {local_format_hm(local_session_seconds())}",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Buy", callback_data="local:buy"),
            InlineKeyboardButton("❌ Cancel", callback_data="local:abort"),
        ]]),
    )


def _panel_api_key(provider: str) -> str:
    """Read a provider API key from persistent runtime settings without exposing it."""
    setting_key = "panel_smsbower_api_key" if provider == "smsbower" else "panel_localusa_api_key"
    default = SMS_API_KEY if provider == "smsbower" else LOCAL_API_KEY
    return DB.get_setting(setting_key, default)


def panel_management_markup() -> InlineKeyboardMarkup:
    sms_key = _panel_api_key("smsbower")
    local_key = _panel_api_key("localusa")
    def status(value: str) -> str:
        return "✅ API set" if value and value.strip() and value.upper() != "XXX" and not value.startswith("PASTE_") else "⚠️ API not set"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(f"🌐 SMSBower — {status(sms_key)}", callback_data="panel:provider:smsbower")],
        [InlineKeyboardButton("➕ Add API", callback_data="panel:add:smsbower"), InlineKeyboardButton("♻️ Replace API", callback_data="panel:replace:smsbower"), InlineKeyboardButton("🗑 Remove API", callback_data="panel:remove:smsbower")],
        [InlineKeyboardButton(f"🇺🇸 Local USA — {status(local_key)}", callback_data="panel:provider:localusa")],
        [InlineKeyboardButton("➕ Add API", callback_data="panel:add:localusa"), InlineKeyboardButton("♻️ Replace API", callback_data="panel:replace:localusa"), InlineKeyboardButton("🗑 Remove API", callback_data="panel:remove:localusa")],
        [InlineKeyboardButton("🔄 Refresh", callback_data="panel:home"), InlineKeyboardButton("⬅️ Admin Menu", callback_data="panel:back")],
    ])


async def show_panel_management(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    context.user_data.pop("panel_api_input", None)
    message = update.effective_message
    await message.reply_text(
        "🧩 <b>Panel Management</b>\n\nManage API keys for SMSBower and Local USA here. Changes are saved in the bot database and applied without restarting the bot. API keys are never displayed in this menu.",
        parse_mode=ParseMode.HTML,
        reply_markup=panel_management_markup(),
    )


async def panel_management_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    global API, LOCAL_API, _PROVIDER_BALANCE_CACHE_VALUE, _PROVIDER_BALANCE_CACHE_AT
    q = update.callback_query
    if not q or not q.from_user or not is_admin(q.from_user.id):
        if q:
            await q.answer("Admin only", show_alert=True)
        return
    await q.answer()
    data = q.data or ""
    if data in ("panel:home", "panel:back"):
        context.user_data.pop("panel_api_input", None)
        if data == "panel:back":
            await q.edit_message_text("Admin panel", reply_markup=admin_menu_markup())
        else:
            await q.edit_message_text(
                "🧩 <b>Panel Management</b>\nManage SMSBower and Local USA API keys below.",
                parse_mode=ParseMode.HTML,
                reply_markup=panel_management_markup(),
            )
        return

    if data.startswith("panel:provider:"):
        provider = data.rsplit(":", 1)[-1]
        label = "SMSBower" if provider == "smsbower" else "Local USA"
        key = _panel_api_key(provider)
        configured = bool(key and key.strip() and key.upper() != "XXX" and not key.startswith("PASTE_"))
        status_text = "⚠️ API key not configured" if not configured else "✅ API key configured"
        balance_text = "Balance check unavailable."
        try:
            if configured and provider == "smsbower":
                candidate = SmsPowerAPI(SMS_API_URL, key)
                try:
                    balance = await candidate.get_balance()
                    balance_text = f"💰 Live balance: <b>${html.escape(str(balance))}</b>"
                finally:
                    await candidate.close()
            elif configured and provider == "localusa":
                candidate = LocalUsaAPI(LOCAL_API_URL, key)
                try:
                    ok, value = await candidate.get_balance()
                    balance_text = (
                        f"💰 Live balance: <b>{html.escape(value)}</b>"
                        if ok else f"❌ Balance check failed: {html.escape(short_text(value, 250))}"
                    )
                finally:
                    await candidate.close()
            else:
                balance_text = "Add an API key to check its balance."
        except Exception as exc:
            balance_text = f"❌ Balance check failed: <code>{html.escape(short_text(str(exc), 250))}</code>"
        await q.edit_message_text(
            f"<b>{label}</b>\n{status_text}\n{balance_text}\n\nAPI key: {'•••••••• (hidden)' if configured else 'not set'}",
            parse_mode=ParseMode.HTML,
            reply_markup=panel_management_markup(),
        )
        return

    parts = data.split(":")
    if len(parts) != 3 or parts[1] not in ("add", "replace", "remove") or parts[2] not in ("smsbower", "localusa"):
        return
    action, provider = parts[1], parts[2]
    label = "SMSBower" if provider == "smsbower" else "Local USA"
    setting_key = "panel_smsbower_api_key" if provider == "smsbower" else "panel_localusa_api_key"
    current = _panel_api_key(provider)
    configured = bool(current and current.strip() and current.upper() != "XXX" and not current.startswith("PASTE_"))
    if action == "remove":
        DB.set_setting(setting_key, "")
        if provider == "smsbower":
            old, API = API, None
            _PROVIDER_BALANCE_CACHE_VALUE = None
            _PROVIDER_BALANCE_CACHE_AT = 0.0
        else:
            old, LOCAL_API = LOCAL_API, None
        if old:
            with suppress(Exception):
                await old.close()
        await q.edit_message_text(
            f"🗑 <b>{label} API removed.</b>\nProvider disabled until a valid key is added.",
            parse_mode=ParseMode.HTML,
            reply_markup=panel_management_markup(),
        )
        return
    if action == "add" and configured:
        await q.edit_message_text(
            f"{label} API is already configured. Use ♻️ Replace API to change it.",
            reply_markup=panel_management_markup(),
        )
        return
    context.user_data["panel_api_input"] = {"provider": provider, "action": action}
    verb = "add" if action == "add" else "replace"
    await q.edit_message_text(
        f"Send the new <b>{label} API key</b> in your next message to {verb} it.\n\n"
        "The key will be tested first and saved only if the provider accepts it.\n"
        "Send /cancel to stop.",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Cancel", callback_data="panel:home")]]),
    )


async def panel_api_key_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """Validate a replacement provider key before committing it to runtime settings."""
    pending = context.user_data.get("panel_api_input")
    user = update.effective_user
    message = update.effective_message
    if not pending or not user or not is_admin(user.id) or not message or not message.text:
        return False
    raw = message.text.strip()
    if not raw or len(raw) > 500 or any(ch.isspace() for ch in raw):
        await message.reply_text("❌ Invalid API key. Send the key as one line, or /cancel to stop.")
        return True

    provider = pending["provider"]
    setting_key = "panel_smsbower_api_key" if provider == "smsbower" else "panel_localusa_api_key"
    label = "SMSBower" if provider == "smsbower" else "Local USA"
    candidate = None
    try:
        await message.reply_text(f"⏳ Testing {label} API and checking live balance...")
        if provider == "smsbower":
            candidate = SmsPowerAPI(SMS_API_URL, raw)
            balance = await candidate.get_balance()
            balance_text = f"💰 Provider balance: ${balance}"
        else:
            candidate = LocalUsaAPI(LOCAL_API_URL, raw)
            ok, value = await candidate.get_balance()
            if not ok:
                raise RuntimeError(value or "Provider balance check failed")
            balance_text = f"💰 Provider balance: {value}"

        # Only persist/swap after a successful live balance response. A bad key
        # therefore cannot overwrite the last working key or client.
        DB.set_setting(setting_key, raw)
        global API, LOCAL_API, _PROVIDER_BALANCE_CACHE_VALUE, _PROVIDER_BALANCE_CACHE_AT
        if provider == "smsbower":
            old, API = API, candidate
            _PROVIDER_BALANCE_CACHE_VALUE = balance
            _PROVIDER_BALANCE_CACHE_AT = time.monotonic()
        else:
            old, LOCAL_API = LOCAL_API, candidate
        candidate = None  # ownership transferred to the runtime global
        if old:
            with suppress(Exception):
                await old.close()
        context.user_data.pop("panel_api_input", None)
        await message.reply_text(
            f"✅ <b>{label} API updated successfully.</b>\n{balance_text}\n\n"
            "The new API is active immediately. No bot restart is required.",
            parse_mode=ParseMode.HTML,
            reply_markup=admin_menu(),
        )
        await message.reply_text("🧩 Panel Management", reply_markup=panel_management_markup())
        return True
    except Exception as exc:
        await message.reply_text(
            f"❌ <b>{label} API was NOT changed.</b>\n"
            f"The previous working key remains active.\n\n"
            f"Reason: <code>{html.escape(short_text(str(exc), 350))}</code>\n\n"
            "Check the API key and provider connection, then try again.",
            parse_mode=ParseMode.HTML,
            reply_markup=panel_management_markup(),
        )
        return True
    finally:
        if candidate is not None:
            with suppress(Exception):
                await candidate.close()



def admin_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("📊 Statistics"), KeyboardButton("🌍 Catalog")],
            [KeyboardButton("➕ Add Service"), KeyboardButton("👤 User Manager")],
            [KeyboardButton("📣 Broadcast"), KeyboardButton("👥 Group Settings")],
            [KeyboardButton("💰 Balance"), KeyboardButton("📋 Lists")],
            [KeyboardButton("🇺🇸 USA Settings"), KeyboardButton("💳 Deposit Settings")],
            [KeyboardButton("🛡️ SET 𝚅𝙿𝙽/𝙿𝚁𝙾𝚇𝙸𝙴𝚂"), KeyboardButton("⚙️ General Settings")],
            [KeyboardButton("📧 Mail Services"), KeyboardButton("🧩 Panel Management")],
            [KeyboardButton("⭐ Premium Emoji")],
            [KeyboardButton("🗄 Database Settings"), KeyboardButton("🔙 Main Menu")],
        ],
        resize_keyboard=True,
    )


ADMIN_ONE_SHOT_ACTIONS = {
    "🛠 Admin Panel": admin_panel,
    "🔙 Admin Panel": admin_panel,
    "🌍 Catalog": admin_catalog_reply,

    # Compact submenu navigation.
    "👥 Group Settings": show_admin_group_settings,
    "💰 Balance": show_admin_balance_menu,
    "📋 Lists": show_admin_lists_menu,
    "🇺🇸 USA Settings": show_admin_usa_settings,
    "💳 Deposit Settings": show_admin_deposit_settings,
    "⚙️ General Settings": show_admin_general_settings,
    "🗄 Database Settings": show_admin_database_settings,
    "🧩 Panel Management": show_panel_management,
    "⭐ Premium Emoji": premium_emoji_menu_reply,
    "📧 Mail Services": mail_admin_list,
    "🛡️ SET 𝚅𝙿𝙽/𝙿𝚁𝙾𝚇𝙸𝙴𝚂": vshop_admin_settings,
    "Add Inline": vshop_add_category_start,
    "Set Category": vshop_set_category,
    "➕ Add Inline": vshop_add_category_start,
    "💰 Set Category": vshop_set_category,
    "🛠 Manage Inline": vshop_set_category,
    "🔗 Referral Settings": show_admin_referral_settings,
    "📚 How to use Settings": show_admin_how_to_use_settings,
    "🔙 General Settings": show_admin_general_settings,
    "💵 Referral Reward": set_referral_reward_start,
    "📏 Referral Deposit Limit": set_referral_threshold_start,
    "➕ Set Button": set_how_button_start,
    "🎬 Set Video": set_how_video_start,

    # One-tap actions inside submenus.
    "🇺🇸 Local Balance": admin_local_balance,
    "⬇️ Download DB": admin_database_download,
}

USER_ONE_SHOT_ACTIONS = {
    "🔢 Number Buy": show_number_buy_menu,
    "🛒 Buy Number": show_buy_menu,
    "🇺🇸 USA Local": show_local_menu,
    "📧 Mail": show_mail_menu,
    "🛒 Buy Local": local_buy_prompt,
    "📱 Active Numbers": local_active_numbers,
    "💰 My Balance": show_balance,
    "💳 Add Balance": deposit_start,
    "🆘 Support": show_support,
    "📚 How to use": show_how_to_use,
    "🛡️ 𝙱𝚄𝚈 𝚅𝙿𝙽/𝙿𝚁𝙾𝚇𝙸𝙴𝚂": show_vshop_user,
}

ONE_SHOT_MENU_PATTERN = r"^(?:" + "|".join(
    re.escape(label)
    for label in [
        *ADMIN_ONE_SHOT_ACTIONS.keys(),
        *USER_ONE_SHOT_ACTIONS.keys(),
        "🔙 Main Menu",
    ]
) + r")$"


async def workflow_start_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=False)
    await start(update, context)
    return ConversationHandler.END


async def workflow_admin_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    await admin_panel(update, context)
    return ConversationHandler.END


async def workflow_id_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    clear_transient_workflow(context)
    await show_id(update, context)
    return ConversationHandler.END


# =============================================================================
# LOW-PRESSURE / ANTI-SPAM OVERRIDES
# =============================================================================

# New admin workflow states.
SPAM_INTERVAL_VALUE = 121
SPAM_LIMIT_VALUE = 122
SPAM_BLOCK_VALUE = 123

# Runtime-only click tracking. Persistent policy values are stored in settings.
# A burst is reset whenever the gap between two clicks is greater than or equal
# to the configured fast-click interval.
_SPAM_LAST_CLICK: dict[int, float] = {}
_SPAM_BURST_COUNTS: dict[int, int] = {}
_SPAM_BLOCKED_UNTIL: dict[int, float] = {}
_PROVIDER_BALANCE_CACHE_VALUE: Optional[Decimal] = None
_PROVIDER_BALANCE_CACHE_AT = 0.0
_PROVIDER_BALANCE_CACHE_SECONDS = 120.0


def ensure_low_pressure_schema() -> None:
    with DB.lock:
        for key, value in (
            ("spam_fast_interval_seconds", "5"),
            ("spam_click_limit", "5"),
            ("spam_cooldown_seconds", "5"),
        ):
            DB.conn.execute(
                "INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)",
                (key, value),
            )
        # Group 1 is intentionally disabled. Deposit requests still use their
        # separate deposit_group_id setting and OTP logs use Group 2 only.
        DB.conn.execute(
            "INSERT INTO settings(key,value) VALUES('otp_forward_group_1','') "
            "ON CONFLICT(key) DO UPDATE SET value=''"
        )
        DB.conn.commit()


ensure_low_pressure_schema()
TRANSIENT_WORKFLOW_KEYS.update({"spam_fast_interval_pending", "spam_click_limit_pending"})


def spam_fast_interval_seconds() -> int:
    try:
        return max(1, min(60, int(DB.get_setting("spam_fast_interval_seconds", "5"))))
    except Exception:
        return 5


def spam_click_limit() -> int:
    try:
        return max(0, min(100, int(DB.get_setting("spam_click_limit", "5"))))
    except Exception:
        return 5


def spam_block_seconds() -> int:
    try:
        return max(1, min(3600, int(DB.get_setting("spam_cooldown_seconds", "5"))))
    except Exception:
        return 5


def _clear_spam_user_state(user_id: int) -> None:
    _SPAM_LAST_CLICK.pop(user_id, None)
    _SPAM_BURST_COUNTS.pop(user_id, None)


def _spam_remaining(user_id: int) -> int:
    """Return remaining block seconds, or zero when this click is allowed.

    Normal clicks do not display any warning. Only clicks made faster than the
    configured interval build a burst. The click that reaches the configured
    burst limit is blocked and starts the short cooldown.
    """
    if is_admin(user_id):
        return 0

    limit = spam_click_limit()
    if limit <= 0:
        _clear_spam_user_state(user_id)
        _SPAM_BLOCKED_UNTIL.pop(user_id, None)
        return 0

    now = time.monotonic()
    blocked_until = _SPAM_BLOCKED_UNTIL.get(user_id, 0.0)
    if blocked_until > now:
        return max(1, math.ceil(blocked_until - now))
    if blocked_until:
        _SPAM_BLOCKED_UNTIL.pop(user_id, None)
        _clear_spam_user_state(user_id)

    interval = float(spam_fast_interval_seconds())
    last_click = _SPAM_LAST_CLICK.get(user_id)
    if last_click is not None and (now - last_click) < interval:
        burst_count = _SPAM_BURST_COUNTS.get(user_id, 1) + 1
    else:
        # A normal gap resets the burst. This click becomes click number one of
        # a possible new burst, but no message or delay is shown.
        burst_count = 1

    _SPAM_LAST_CLICK[user_id] = now
    _SPAM_BURST_COUNTS[user_id] = burst_count

    if burst_count >= limit:
        block_seconds = spam_block_seconds()
        _SPAM_BLOCKED_UNTIL[user_id] = now + block_seconds
        _clear_spam_user_state(user_id)
        return block_seconds
    return 0

def _spam_callback_protected(data: str) -> bool:
    """Protect normal inline callbacks, but keep recovery actions available.

    The speed limit applies to inline buttons, as requested.  The Reply
    Keyboard main menu is handled by ``workflow_one_shot_menu`` and is not
    passed through this guard, so it remains usable during the cooldown.
    Cancellation and terms acceptance stay available so a user can recover
    from an active flow.
    """
    return not data.startswith((
        "workflow:cancel",
        "order:cancel:",
        "local:cancel:",
        "terms:accept",
    ))


_OriginalCallbackQueryHandler = CallbackQueryHandler


def CallbackQueryHandler(callback: Any, *args: Any, **kwargs: Any) -> Any:
    """Wrap purchase callbacks with the configurable user click limiter."""
    async def guarded(update: Update, context: ContextTypes.DEFAULT_TYPE) -> Any:
        query = update.callback_query
        user = update.effective_user
        data = str(query.data or "") if query else ""
        if query and user and _spam_callback_protected(data):
            remaining = _spam_remaining(int(user.id))
            if remaining:
                with suppress(TelegramError):
                    await query.answer(
                        f"⏳ Please wait {remaining} seconds before taking another number.",
                        show_alert=True,
                    )
                return None
        result = callback(update, context)
        if inspect.isawaitable(result):
            return await result
        return result

    guarded.__name__ = getattr(callback, "__name__", "guarded_callback")
    return _OriginalCallbackQueryHandler(guarded, *args, **kwargs)


async def spam_control_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    await update.effective_message.reply_text(
        "<b>🛡 Spam Control</b>\n\n"
        f"দ্রুত-click ব্যবধান: <b>{spam_fast_interval_seconds()} সেকেন্ড</b>\n"
        f"দ্রুত click limit: <b>{spam_click_limit()} বার</b>\n"
        f"Block সময়: <b>{spam_block_seconds()} সেকেন্ড</b>\n\n"
        "কত সেকেন্ডের কম ব্যবধানে click করলে দ্রুত-click হিসেবে গণনা হবে?\n"
        "উদাহরণ: <code>5</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return SPAM_INTERVAL_VALUE


async def spam_control_interval_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    try:
        seconds = int((update.effective_message.text or "").strip())
        if seconds < 1 or seconds > 60:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text("১ থেকে ৬০ সেকেন্ডের মধ্যে দিন।")
        return SPAM_INTERVAL_VALUE

    context.user_data["spam_fast_interval_pending"] = seconds
    await update.effective_message.reply_text(
        "একটানা কতটি দ্রুত click-এর পর সাময়িক block হবে?\n"
        "উদাহরণ: <code>5</code>\n"
        "Spam Control বন্ধ রাখতে <code>0</code> দিন।",
        parse_mode=ParseMode.HTML,
    )
    return SPAM_LIMIT_VALUE


async def spam_control_limit_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    try:
        value = int((update.effective_message.text or "").strip())
        if value < 0 or value > 100:
            raise ValueError
        if value == 1:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text(
            "বন্ধ রাখতে ০ দিন, অথবা ২ থেকে ১০০-এর মধ্যে একটি সংখ্যা দিন।"
        )
        return SPAM_LIMIT_VALUE

    context.user_data["spam_click_limit_pending"] = value
    await update.effective_message.reply_text(
        "Limit হলে কত সেকেন্ড সব button বন্ধ থাকবে?\n"
        "উদাহরণ: <code>5</code>",
        parse_mode=ParseMode.HTML,
    )
    return SPAM_BLOCK_VALUE


async def spam_control_block_value(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    try:
        seconds = int((update.effective_message.text or "").strip())
        if seconds < 1 or seconds > 3600:
            raise ValueError
    except Exception:
        await update.effective_message.reply_text("১ থেকে ৩৬০০ সেকেন্ডের মধ্যে দিন।")
        return SPAM_BLOCK_VALUE

    interval = int(context.user_data.pop("spam_fast_interval_pending", 5))
    limit = int(context.user_data.pop("spam_click_limit_pending", 5))
    DB.set_setting("spam_fast_interval_seconds", str(interval))
    DB.set_setting("spam_click_limit", str(limit))
    DB.set_setting("spam_cooldown_seconds", str(seconds))
    _SPAM_LAST_CLICK.clear()
    _SPAM_BURST_COUNTS.clear()
    _SPAM_BLOCKED_UNTIL.clear()

    status = "বন্ধ" if limit == 0 else "চালু"
    await update.effective_message.reply_text(
        f"✅ Spam Control {status}।\n"
        f"ব্যবধান: {interval} সেকেন্ড | Limit: {limit} বার | Block: {seconds} সেকেন্ড",
        reply_markup=admin_menu(),
    )
    return ConversationHandler.END


# -------------------------- compact admin menus ------------------------------

def admin_group_settings_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("📥 Deposit Group ID"), KeyboardButton("🔐 OTP Group 2")],
            [KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


def admin_lists_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("📋 Payment List"), KeyboardButton("🧾 Order Logs")],
            [KeyboardButton("👥 User List"), KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


def admin_general_settings_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("🆘 Support ID/Link"), KeyboardButton("🛡 Spam Control")],
            [KeyboardButton("💱 Exchange Rate"), KeyboardButton("⚙️ Profit Markup")],
            [KeyboardButton("🔗 Referral Settings"), KeyboardButton("📚 How to use Settings")],
            [KeyboardButton("📧 Mail Services")],
            [KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


# ------------------------------- TXT reports ---------------------------------

def _new_txt_path(prefix: str) -> str:
    fd, path = tempfile.mkstemp(
        prefix=f"{prefix}_", suffix=".txt", dir=tempfile.gettempdir()
    )
    os.close(fd)
    return path


def create_user_list_txt() -> str:
    rows = DB.fetchall(
        "SELECT user_id,username,first_name,balance_units,created_at,is_blocked "
        "FROM users ORDER BY user_id"
    )
    path = _new_txt_path("user_list")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("USER LIST\n" + "=" * 58 + "\n")
        for index, row in enumerate(rows, 1):
            username = f"@{row['username']}" if row["username"] else "-"
            name = str(row["first_name"] or "Unknown").replace("\n", " ")
            handle.write(
                f"{index}. {name} | {username} | ID {row['user_id']} | "
                f"Balance {money(int(row['balance_units'] or 0))} | "
                f"Blocked {'Yes' if row['is_blocked'] else 'No'}\n"
            )
    return path


async def _send_txt_file(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    generator: Any,
    filename: str,
    caption: str,
) -> None:
    if not await require_admin(update):
        return
    path = ""
    try:
        path = await run_file_job(generator)
        with open(path, "rb") as handle:
            await context.bot.send_document(
                chat_id=update.effective_chat.id,
                document=handle,
                filename=filename,
                caption=caption,
                read_timeout=60,
                write_timeout=BOT_MEDIA_WRITE_TIMEOUT_SECONDS,
            )
    except RetryAfter as exc:
        logger.warning("TXT export blocked by Telegram; retry after %s seconds", retry_after_seconds(exc))
    except TelegramError as exc:
        logger.warning("Could not send %s: %s", filename, exc)
    except Exception:
        logger.exception("Could not build/send %s", filename)
    finally:
        if path:
            with suppress(OSError):
                os.remove(path)


async def admin_payment_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _send_txt_file(
        update, context, create_payment_list_txt,
        "payment_list.txt", "Approved amount by user"
    )


async def admin_order_logs(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _send_txt_file(
        update, context, create_order_logs_txt,
        "order_logs.txt", "Complete order logs"
    )


async def admin_user_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _send_txt_file(
        update, context, create_user_list_txt,
        "user_list.txt", "Registered users"
    )


# -------------------------- low-edit provider balance ------------------------

async def admin_provider_balance(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    global _PROVIDER_BALANCE_CACHE_VALUE, _PROVIDER_BALANCE_CACHE_AT
    if not await require_admin(update):
        return
    try:
        if API is None:
            raise RuntimeError("SMSBower API is not configured. Add it from Admin → Panel Management.")
        now_mono = time.monotonic()
        if (
            _PROVIDER_BALANCE_CACHE_VALUE is None
            or now_mono - _PROVIDER_BALANCE_CACHE_AT >= _PROVIDER_BALANCE_CACHE_SECONDS
        ):
            _PROVIDER_BALANCE_CACHE_VALUE = await API.get_balance()
            _PROVIDER_BALANCE_CACHE_AT = now_mono
        await update.effective_message.reply_text(
            f"<b>Balance: ${_PROVIDER_BALANCE_CACHE_VALUE}</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=admin_balance_menu(),
        )
    except RetryAfter as exc:
        logger.warning("Provider balance reply blocked; retry after %s seconds", retry_after_seconds(exc))
    except Exception as exc:
        with suppress(TelegramError):
            await update.effective_message.reply_text(
                f"Balance error: <code>{html.escape(short_text(str(exc), 400))}</code>",
                parse_mode=ParseMode.HTML,
                reply_markup=admin_balance_menu(),
            )


# ---------------------- standard-number refresh behavior ---------------------

async def maybe_refresh_countdown(application: Application, order: sqlite3.Row) -> None:
    """Countdown messages are manual-only; background polling never edits timers."""
    if order["status"] != "waiting":
        COUNTDOWN_MINUTE_CACHE.pop(int(order["id"]), None)


# -------------------------- Group 2 forwarding only --------------------------

async def local_apply_otp(
    application: Application,
    order_id: int,
    result: LocalOtpResult,
) -> tuple[bool, str]:
    order = local_get_order(order_id)
    if not order or order["status"] != "active":
        return False, "Session ended."
    if result.state == "received" and result.otp:
        had_otp = bool(int(order["has_otp"] or 0))
        history_id = local_record_otp(order_id, result.otp, result.raw)
        if history_id and not had_otp:
            # The 60-minute pre-OTP window ends here. From the first OTP onward,
            # keep the number for the admin-configured Local Session duration.
            ts = now_ts()
            DB.execute(
                """
                UPDATE local_orders
                SET has_otp=1,
                    expires_at=?,
                    updated_at=?
                WHERE id=? AND status='active'
                """,
                (ts + local_session_seconds(), ts, order_id),
            )
        refreshed = local_get_order(order_id)
        if history_id and refreshed:
            await local_update_user_otp_message(
                application, refreshed, result.otp, result.raw
            )
            await local_forward_otp_groups(
                application, order_id, history_id, result.otp
            )
            return True, f"New OTP: {result.otp}"
        return False, "No newer OTP."
    if result.state == "cancelled":
        local_set_status(order_id, "cancelled", result.raw)
        return False, "Provider reports inactive."
    if result.state == "error":
        DB.execute(
            "UPDATE local_orders SET raw_last_response=?,updated_at=? WHERE id=?",
            (result.raw[:10000], now_ts(), order_id),
        )
        return False, result.message or "OTP check failed."
    return False, result.message or "No new OTP."


async def local_forward_otp_groups(
    application: Application,
    order_id: int,
    history_id: int,
    otp: str,
) -> None:
    """Forward local OTP only to Group 2, including username when available."""
    order = DB.fetchone("SELECT * FROM local_orders WHERE id=?", (order_id,))
    history = DB.fetchone("SELECT * FROM local_otp_history WHERE id=?", (history_id,))
    group_id = _setting_group_id("otp_admin_group_2")
    if not order or not history or not group_id or int(history["private_forwarded"] or 0):
        return
    user = DB.fetchone("SELECT * FROM users WHERE user_id=?", (order["user_id"],))
    username = f"@{user['username']}" if user and user["username"] else "Not set"
    name = str(user["first_name"] or "Unknown") if user else "Unknown"
    balance = int(user["balance_units"] or 0) if user else 0
    text_value = (
        "🔐 <b>USA LOCAL OTP</b>\n"
        f"👤 {html.escape(name)} • <code>{html.escape(username)}</code>\n"
        f"🆔 <code>{order['user_id']}</code> • 💰 <b>{money(balance)}</b>\n"
        f"📦 Local #{order_id}\n"
        f"☎️ <code>+{html.escape(str(order['phone']))}</code>\n"
        f"?? <code>{html.escape(str(otp))}</code> • 💳 <b>{money(int(order['charged_units'] or 0))}</b>"
    )
    try:
        await application.bot.send_message(
            chat_id=group_id,
            text=text_value,
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([[copy_button(f"🔐 {otp}", otp)]]),
        )
        DB.execute(
            "UPDATE local_otp_history SET private_forwarded=1 WHERE id=?",
            (history_id,),
        )
    except TelegramError as exc:
        logger.warning("Local OTP Group 2 forward failed: %s", exc)


# --------------------------- reduced-edit broadcast --------------------------

async def _run_broadcast_background(
    bot,
    chat_id: int,
    status_message_id: int,
    source_chat_id: int,
    source_message_id: int,
) -> None:
    """Run the actual broadcast loop in the background so the bot stays responsive."""
    users = DB.fetchall("SELECT user_id FROM users ORDER BY user_id")
    sent = failed = 0
    for row in users:
        ok = await _copy_broadcast_with_retry(
            bot, int(row["user_id"]), source_chat_id, source_message_id
        )
        sent += int(ok)
        failed += int(not ok)
        await asyncio.sleep(max(BROADCAST_DELAY_SECONDS, 0.15))
    with suppress(TelegramError):
        await bot.delete_message(chat_id=chat_id, message_id=status_message_id)
        await bot.send_message(
            chat_id=chat_id,
            text=f"✅ Broadcast শেষ।\nSent: {sent}\nFailed: {failed}",
        )
    with suppress(TelegramError):
        await bot.send_message(
            chat_id=chat_id,
            text="Admin Panel",
            reply_markup=admin_menu(),
        )


async def broadcast_confirm_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    query = update.callback_query
    if not query or not await require_admin(update):
        return ConversationHandler.END
    source_chat_id = int(context.user_data.pop("broadcast_source_chat_id", 0) or 0)
    source_message_id = int(context.user_data.pop("broadcast_source_message_id", 0) or 0)
    if not source_chat_id or not source_message_id:
        await query.answer("Broadcast session expired.", show_alert=True)
        return ConversationHandler.END
    await query.answer("Broadcast শুরু হয়েছে!")
    chat_id = query.message.chat_id
    status_message_id = query.message.message_id
    try:
        status_msg = await replace_callback_message(query, context.bot, 
            "⏳ Broadcast চলছে… (background এ। বট স্বাভাবিক কাজ করবে।)"
        )
        status_message_id = status_msg.message_id
    except TelegramError:
        pass

    # Launch broadcast in background — bot stays fully responsive
    asyncio.create_task(
        _run_broadcast_background(
            context.bot,
            chat_id,
            status_message_id,
            source_chat_id,
            source_message_id,
        )
    )
    return ConversationHandler.END


# ---------------------- lifecycle without local auto-poll --------------------

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    error = context.error
    if isinstance(error, RetryAfter):
        logger.warning(
            "Telegram flood control active; retry after %s seconds",
            retry_after_seconds(error),
        )
        return
    if isinstance(error, NetworkError):
        # Telegram occasionally returns HTTP 502/Bad Gateway or has a short
        # network interruption. PTB's polling loop retries polling requests;
        # log this as a warning rather than a full traceback for every blip.
        logger.warning("Temporary Telegram network/API error; the bot will continue: %s", error)
        return
    logger.error(
        "Unhandled exception while processing update",
        exc_info=(type(error), error, error.__traceback__) if error else None,
    )
    # Do not send another Telegram message from the global error handler. When
    # Telegram itself is rate-limiting or unavailable, that can create more errors.



# =============================================================================
# MANUAL-ONLY, API-AUTHORITATIVE RUNTIME OVERRIDES
# =============================================================================
# These final definitions intentionally replace older implementations above.
# Provider status is checked only after a user presses a button. Telegram order
# and timer messages are sent as fresh messages instead of being edited by
# background jobs. This avoids long-running flood-control/429 pressure.


def _manual_minutes_left(expires_at: Any) -> int:
    remaining = max(0, int(expires_at or 0) - now_ts())
    if remaining <= 0:
        return 0
    return max(1, int(math.ceil(remaining / 60)))


async def _send_standard_snapshot(
    bot: Any,
    chat_id: int,
    order: sqlite3.Row,
    prefix: str = "",
) -> Any:
    text = render_order(order)
    if prefix:
        text = f"{prefix}\n\n{text}"
    return await bot.send_message(
        chat_id=chat_id,
        text=text,
        parse_mode=ParseMode.HTML,
        reply_markup=order_keyboard(order),
    )


async def refresh_order_message(application: Application, order_id: int) -> None:
    """Compatibility helper: send a fresh snapshot; never edit an old message."""
    order = DB.fetchone("SELECT * FROM orders WHERE id=?", (order_id,))
    if not order:
        return
    try:
        sent = await _send_standard_snapshot(
            application.bot, int(order["user_id"]), order
        )
        DB.execute(
            "UPDATE orders SET message_id=? WHERE id=?",
            (sent.message_id, order_id),
        )
    except (Forbidden, TelegramError) as exc:
        logger.debug("Could not send order snapshot %s: %s", order_id, exc)


async def _manual_standard_status_check(
    application: Application,
    order: sqlite3.Row,
) -> tuple[str, str, sqlite3.Row]:
    """Make exactly one provider status request for a manual Check SMS click."""
    if not API:
        return "error", "Provider API is not ready.", order
    item = get_catalog_item(order["catalog_id"]) if order["catalog_id"] else None
    extra = catalog_extra(item) if item else {}
    result = await API.get_status(str(order["activation_id"]), extra)
    ts = now_ts()

    if result.state == "code" and result.code:
        DB.execute(
            """
            UPDATE orders
            SET status='otp_received',otp_code=?,last_api_status=?,updated_at=?,ever_received_otp=1
            WHERE id=? AND status='waiting'
            """,
            (str(result.code), result.raw_text, ts, order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return "code", str(result.code), refreshed or order

    if result.state == "expired":
        DB.execute(
            """
            UPDATE orders SET status='expired',last_api_status=?,updated_at=?
            WHERE id=? AND status='waiting'
            """,
            (result.raw_text, ts, order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return "expired", "Provider reports that the number expired.", refreshed or order

    if result.state == "cancelled":
        refunded = DB.refund_order(int(order["id"]), result.raw_text)
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        message = (
            "Provider reports cancellation; balance refunded."
            if refunded
            else "Provider reports cancellation; local order state was already changed."
        )
        return "cancelled", message, refreshed or order

    DB.execute(
        "UPDATE orders SET last_api_status=?,updated_at=? WHERE id=?",
        (short_text(result.raw_text, 500), ts, order["id"]),
    )
    refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
    if result.state == "waiting":
        return "waiting", "No OTP yet. Press Check SMS again later.", refreshed or order
    return "error", (
        result.raw_text or "Provider did not return a usable SMS status."
    ), refreshed or order


async def process_single_order(application: Application, order: sqlite3.Row) -> None:
    """Compatibility entry point; status checks happen only when explicitly called."""
    if order["status"] != "waiting":
        return
    state, value, refreshed = await _manual_standard_status_check(application, order)
    if state == "code":
        await send_otp_once(application, int(refreshed["id"]), value)


def standard_cancel_without_refund(order_id: int, api_status: str) -> bool:
    """Mark a 20-minute no-OTP standard order cancelled without refunding it."""
    ts = now_ts()
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            order = DB.conn.execute(
                "SELECT * FROM orders WHERE id=?",
                (order_id,),
            ).fetchone()
            if not order or order["status"] != "waiting" or order["refunded"]:
                DB.conn.rollback()
                return False
            DB.conn.execute(
                """
                UPDATE orders
                SET status='cancelled', refunded=0,
                    last_api_status=?, updated_at=?
                WHERE id=? AND status='waiting'
                """,
                (api_status[:10000], ts, order_id),
            )
            DB.conn.commit()
            return True
        except Exception:
            DB.conn.rollback()
            raise


async def _api_expire_standard_order_without_refund(
    application: Application,
    order: sqlite3.Row,
) -> tuple[str, str, sqlite3.Row]:
    """At 20 minutes with no OTP: cancel at provider and keep the charge."""
    if not API:
        return "error", "Provider API is not ready.", order

    item = get_catalog_item(order["catalog_id"]) if order["catalog_id"] else None
    extra = catalog_extra(item) if item else {}
    current = await API.get_status(str(order["activation_id"]), extra)

    if current.state == "code" and current.code:
        DB.execute(
            """
            UPDATE orders
            SET status='otp_received', otp_code=?, last_api_status=?, updated_at=?, ever_received_otp=1
            WHERE id=? AND status='waiting'
            """,
            (str(current.code), current.raw_text, now_ts(), order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        await send_otp_once(application, int(order["id"]), str(current.code))
        return "otp", "OTP arrived before the 20-minute cutoff.", refreshed or order

    if current.state == "expired":
        DB.execute(
            """
            UPDATE orders
            SET status='expired', last_api_status=?, updated_at=?
            WHERE id=? AND status='waiting'
            """,
            (current.raw_text, now_ts(), order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return "expired", "Provider reports expiry.", refreshed or order

    if current.state == "cancelled":
        standard_cancel_without_refund(int(order["id"]), current.raw_text)
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return "cancelled", "Provider reports cancellation; charge kept.", refreshed or order

    if current.state != "waiting":
        DB.execute(
            "UPDATE orders SET last_api_status=?,updated_at=? WHERE id=?",
            (short_text(current.raw_text, 500), now_ts(), order["id"]),
        )
        return "error", "Provider did not confirm a cancellable waiting order.", order

    data, raw = await API.set_status(str(order["activation_id"]), 8, extra)
    if not API.cancellation_succeeded(data, raw):
        DB.execute(
            "UPDATE orders SET last_api_status=?,updated_at=? WHERE id=?",
            (short_text(raw, 500), now_ts(), order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return "error", "Provider did not accept cancellation.", refreshed or order

    marked = standard_cancel_without_refund(int(order["id"]), raw)
    refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
    if not marked:
        return "error", "Provider cancelled, but the local order state changed.", refreshed or order
    return "cancelled", "Provider cancelled at the 20-minute cutoff; charge kept.", refreshed or order


async def _api_cancel_standard_order(
    application: Application,
    order: sqlite3.Row,
) -> tuple[str, str, sqlite3.Row]:
    """Cancel/refund only after an explicit provider-confirmed response."""
    if not API:
        return "error", "Provider API is not ready.", order
    item = get_catalog_item(order["catalog_id"]) if order["catalog_id"] else None
    extra = catalog_extra(item) if item else {}

    current = await API.get_status(str(order["activation_id"]), extra)
    if current.state == "code" and current.code:
        DB.execute(
            """
            UPDATE orders SET status='otp_received',otp_code=?,last_api_status=?,updated_at=?,ever_received_otp=1
            WHERE id=? AND status='waiting'
            """,
            (str(current.code), current.raw_text, now_ts(), order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        await send_otp_once(application, int(order["id"]), str(current.code))
        return "otp", "OTP already arrived. Cancellation was not attempted.", refreshed or order

    if current.state == "expired":
        DB.execute(
            """
            UPDATE orders SET status='expired',last_api_status=?,updated_at=?
            WHERE id=? AND status='waiting'
            """,
            (current.raw_text, now_ts(), order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return "expired", "Provider reports expiry. No refund was issued.", refreshed or order

    if current.state == "cancelled":
        refunded = DB.refund_order(int(order["id"]), current.raw_text)
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return (
            "cancelled",
            "Provider reports cancellation; balance refunded."
            if refunded
            else "Provider reports cancellation; order state had already changed.",
            refreshed or order,
        )

    if current.state != "waiting":
        DB.execute(
            "UPDATE orders SET last_api_status=?,updated_at=? WHERE id=?",
            (short_text(current.raw_text, 500), now_ts(), order["id"]),
        )
        return (
            "error",
            "Provider did not confirm that the number is cancellable. No refund was issued.",
            order,
        )

    data, text = await API.set_status(str(order["activation_id"]), 8, extra)
    if not API.cancellation_succeeded(data, text):
        DB.execute(
            "UPDATE orders SET last_api_status=?,updated_at=? WHERE id=?",
            (short_text(text, 500), now_ts(), order["id"]),
        )
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
        return (
            "error",
            "Provider did not accept cancellation. No refund was issued.",
            refreshed or order,
        )

    refunded = DB.refund_order(int(order["id"]), text)
    refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order["id"],))
    if not refunded:
        return "error", "Provider cancelled, but the local order state changed before refund.", refreshed or order
    return "cancelled", "Provider cancelled successfully. Balance refunded.", refreshed or order


async def cancel_order(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        await query.answer("Invalid order.", show_alert=True)
        return
    order = DB.fetchone(
        "SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id)
    )
    if not order:
        await query.answer("Order not found.", show_alert=True)
        return
    if order["status"] != "waiting":
        await query.answer("This order is no longer cancellable.")
        await _send_standard_snapshot(context.bot, query.message.chat_id, order)
        return

    await query.answer("Checking provider cancellation…")
    lock = _order_process_lock(order_id)
    async with lock:
        order = DB.fetchone(
            "SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id)
        )
        if not order or order["status"] != "waiting":
            return
        try:
            state, message, refreshed = await _api_cancel_standard_order(
                context.application, order
            )
        except Exception as exc:
            await query.message.reply_text(
                f"❌ Cancellation API failed. No refund was issued.\n"
                f"<code>{html.escape(short_text(str(exc), 500))}</code>",
                parse_mode=ParseMode.HTML,
            )
            return

    icon = "✅" if state == "cancelled" else ("⌛" if state == "expired" else "❌")
    await _send_standard_snapshot(
        context.bot,
        query.message.chat_id,
        refreshed,
        prefix=f"{icon} {message}",
    )


async def confirm_purchase(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Buy only after the provider returns a real activation ID and phone number."""
    query = update.callback_query
    user = update.effective_user
    if not query or not user or not await require_terms(update):
        return
    try:
        _, _, item_raw, provider_id = query.data.split(":", 3)
        item_id = int(item_raw)
    except Exception:
        await query.answer("Invalid selection.", show_alert=True)
        return
    restriction = purchase_restriction_message(user.id)
    if restriction:
        await query.answer("Purchase restricted.", show_alert=True)
        await query.message.reply_text(restriction, parse_mode=ParseMode.HTML)
        return
    lock = PURCHASE_LOCKS.setdefault(user.id, asyncio.Lock())
    if lock.locked():
        await query.answer("Another purchase is processing.", show_alert=True)
        return
    await query.answer("Buying from provider…")
    async with lock:
        order, error = await _purchase_catalog_provider(user.id, item_id, provider_id)
        if not order:
            text = (
                "আপনার পর্যাপ্ত ব্যালেন্স নেই। অর্ডার করতে আগে ব্যালেন্স যোগ করুন।"
                if "Insufficient balance" in error
                else "❌ Provider did not supply a number. Nothing was saved.\n" + html.escape(error)
            )
            stock_changed = any(
                phrase in error.casefold()
                for phrase in ("out of stock", "no stock", "no numbers", "not available")
            )
            reply_markup = None
            if stock_changed:
                reply_markup = InlineKeyboardMarkup(
                    [
                        [InlineKeyboardButton(
                            "🔄 Refresh prices",
                            callback_data=f"buy:service:{item_id}",
                        )],
                        [InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")],
                    ]
                )
            await query.edit_message_text(
                text,
                parse_mode=ParseMode.HTML,
                reply_markup=reply_markup,
            )
            return
        if not str(order["activation_id"] or "").strip() or not str(order["phone"] or "").strip():
            await query.message.reply_text(
                "❌ Provider response had no valid number. Nothing was shown in the panel."
            )
            return
        with suppress(TelegramError):
            await query.message.delete()
        await _send_new_order_message(context, query.message.chat_id, order)


async def refund_and_new_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        await query.answer("Invalid order.", show_alert=True)
        return
    lock = PURCHASE_LOCKS.setdefault(user.id, asyncio.Lock())
    if lock.locked():
        await query.answer("Another purchase is processing.", show_alert=True)
        return
    await query.answer("Checking provider and replacing…")
    async with lock:
        order = DB.fetchone(
            "SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id)
        )
        if (
            not order
            or order["status"] != "waiting"
            or int(order["ever_received_otp"] or 0)
        ):
            await query.message.reply_text("এই নম্বরে OTP একবার এসেছে, তাই Refund করা যাবে না। New নাম্বার নিন।")
            return
        try:
            state, message, refreshed = await _api_cancel_standard_order(
                context.application, order
            )
        except Exception as exc:
            await query.message.reply_text(
                f"❌ Provider cancellation failed. No refund/new purchase was made.\n"
                f"<code>{html.escape(short_text(str(exc), 500))}</code>",
                parse_mode=ParseMode.HTML,
            )
            return
        if state != "cancelled":
            await _send_standard_snapshot(
                context.bot,
                query.message.chat_id,
                refreshed,
                prefix=f"❌ {message}",
            )
            return
        if not order["catalog_id"] or not order["provider_id"]:
            await query.message.reply_text(
                "↩️ Old number refunded, but the original provider selection is unavailable."
            )
            return
        new_order, error = await _purchase_catalog_provider(
            user.id, int(order["catalog_id"]), str(order["provider_id"])
        )
        if not new_order:
            await query.message.reply_text(
                "↩️ Old number refunded.\n"
                "❌ Provider did not supply a replacement number.\n"
                f"{html.escape(error)}",
                parse_mode=ParseMode.HTML,
            )
            return
        await _send_new_order_message(context, query.message.chat_id, new_order)


# ------------------------ USA local manual-only flow -------------------------


async def local_active_loop(
    application: Application,
    user_id: int,
    chat_id: int,
    message_id: int,
    page: int,
) -> None:
    """Disabled: active-number panels never auto-edit in the background."""
    return None


def local_start_active_task(
    application: Application,
    user_id: int,
    chat_id: int,
    message_id: int,
    page: int,
) -> None:
    return None


def local_start_detail_task(
    application: Application,
    order_id: int,
    chat_id: int,
    message_id: int,
    page: int,
) -> None:
    return None


async def local_refresh_all_views(application: Application, order_id: int) -> None:
    """Disabled: no automatic edits of local order/list/detail messages."""
    return None


async def local_cancel_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        _, _, order_raw, page_raw = query.data.split(":", 3)
        order_id, page = int(order_raw), int(page_raw)
    except Exception:
        await query.answer("Invalid order.", show_alert=True)
        return
    order = local_get_user_order(order_id, user.id)
    if not order:
        await query.answer("Order not found.", show_alert=True)
        return
    if order["status"] != "active":
        await query.answer("Session ended.")
        await query.message.reply_text(
            local_detail_text(order),
            parse_mode=ParseMode.HTML,
            reply_markup=local_detail_markup(order, page),
        )
        return
    if int(order["has_otp"] or 0):
        await query.answer("OTP received. Cancel is unavailable.", show_alert=True)
        return

    current = now_ts()
    if current < int(order["cancel_after"]):
        await query.answer(
            f"Cancel is locked for about {_manual_minutes_left(order['cancel_after'])} minute(s).",
            show_alert=True,
        )
        return
    cancel_window_end = int(order["created_at"]) + LOCAL_CANCEL_WINDOW_SECONDS
    if current >= cancel_window_end:
        await query.answer("Cancel/refund window expired.", show_alert=True)
        return
    if not LOCAL_API:
        await query.answer("Local API is not ready.", show_alert=True)
        return

    await query.answer("Checking provider cancellation…")
    lock = _LOCAL_ORDER_CHECK_LOCKS.setdefault(order_id, asyncio.Lock())
    async with lock:
        try:
            latest = await LOCAL_API.get_local_otp(
                str(order["phone"]), seen_otps=local_seen_otps(order_id)
            )
            is_new, _ = await local_apply_otp(context.application, order_id, latest)
            order = local_get_order(order_id)
            if is_new or not order or int(order["has_otp"] or 0):
                await query.message.reply_text(
                    "❌ OTP arrived, so cancellation and refund were blocked."
                )
                return
            success, message, raw = await LOCAL_API.cancel_local(str(order["phone"]))
        except Exception as exc:
            await query.message.reply_text(
                f"❌ Provider cancellation API failed. No refund was issued.\n"
                f"<code>{html.escape(short_text(str(exc), 500))}</code>",
                parse_mode=ParseMode.HTML,
            )
            return

        if not success:
            DB.execute(
                """
                UPDATE local_orders
                SET provider_cancel_response=?,raw_last_response=?,updated_at=?
                WHERE id=?
                """,
                (raw[:10000], raw[:10000], now_ts(), order_id),
            )
            await query.message.reply_text(
                "❌ Provider did not confirm cancellation. No refund was issued.\n"
                f"{html.escape(short_text(message or raw, 500))}",
                parse_mode=ParseMode.HTML,
            )
            return

        refunded = local_refund_order(
            order_id, raw or message, provider_confirmed=True
        )
        refreshed = local_get_order(order_id) or order
        prefix = (
            "✅ Provider cancelled successfully. Balance refunded."
            if refunded
            else "⚠️ Provider cancelled, but local refund needs admin review."
        )
        await query.message.reply_text(
            f"{prefix}\n\n{local_detail_text(refreshed)}",
            parse_mode=ParseMode.HTML,
            reply_markup=local_detail_markup(refreshed, page),
        )


async def local_buy_confirm(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    restriction = purchase_restriction_message(user.id)
    if restriction:
        await query.answer("Purchase restricted.", show_alert=True)
        await query.message.reply_text(restriction, parse_mode=ParseMode.HTML)
        return
    if user.id in LOCAL_BUY_LOCKS:
        await query.answer("Processing…", show_alert=True)
        return
    if LOCAL_API is None:
        await query.answer("USA Local service is unavailable: API not added.", show_alert=True)
        with suppress(TelegramError):
            await query.message.reply_text(
                "🇺🇸 <b>USA Local service is currently unavailable.</b>\n\n"
                "Admin has not added the Local USA API yet. Please try again later.",
                parse_mode=ParseMode.HTML,
                reply_markup=local_user_menu(),
            )
        return
    price = local_price_units()
    balance_row = DB.fetchone(
        "SELECT balance_units FROM users WHERE user_id=?", (user.id,)
    )
    if not balance_row or int(balance_row["balance_units"] or 0) < price:
        await query.answer("Order rejected.")
        await replace_callback_message(
            query, context.bot,
            "আপনার পর্যাপ্ত ব্যালেন্স নেই। অর্ডার করতে আগে ব্যালেন্স যোগ করুন।",
            reply_markup=main_menu(user.id),
        )
        return

    await query.answer("Buying from provider…")
    LOCAL_BUY_LOCKS.add(user.id)
    try:
        purchase = await LOCAL_API.buy_local()
        if not purchase.ok or not str(purchase.phone or "").strip():
            await query.message.reply_text(
                "❌ Provider did not supply a number. Nothing was saved.\n"
                f"{html.escape(short_text(purchase.message or purchase.raw, 700))}",
                parse_mode=ParseMode.HTML,
            )
            return
        if purchase_restriction_message(user.id):
            success, message, raw = await LOCAL_API.cancel_local(purchase.phone)
            await query.message.reply_text(
                "⛔ Purchase restriction was applied after the API response.\n"
                + (
                    "Provider number was cancelled."
                    if success
                    else "Provider cancellation was not confirmed; admin review is required."
                )
            )
            return
        try:
            order_id, error = local_save_purchase(user.id, purchase, price)
        except Exception as exc:
            logger.exception("Could not save local purchase")
            order_id, error = None, str(exc)
        if not order_id:
            with suppress(Exception):
                await LOCAL_API.cancel_local(purchase.phone)
            await query.message.reply_text(
                f"❌ Purchased number could not be saved: {html.escape(error or 'Save failed.')}\n"
                "The bot did not create a panel for it.",
                parse_mode=ParseMode.HTML,
            )
            return
        order = local_get_order(order_id)
        if not order:
            await query.message.reply_text("❌ Number was saved but could not be loaded.")
            return
        await replace_callback_message(
            query, context.bot, local_detail_text(order),
            parse_mode=ParseMode.HTML, reply_markup=local_detail_markup(order, 0)
        )
    finally:
        LOCAL_BUY_LOCKS.discard(user.id)


# --------------------------- background lifecycle ----------------------------


async def post_shutdown(application: Application) -> None:
    global POLL_TASK, DAILY_REPORT_TASK, DAILY_BACKUP_TASK
    global PERSISTENT_DATABASE_BACKUP_TASK, API, LOCAL_API
    tasks = [
        POLL_TASK,
        DAILY_REPORT_TASK,
        DAILY_BACKUP_TASK,
        PERSISTENT_DATABASE_BACKUP_TASK,
        *list(_IPSHOP_PAYMENT_TASKS.values()),
        *list(LOCAL_ORDER_POLL_TASKS.values()),
        *list(LOCAL_DETAIL_TASKS.values()),
        *list(LOCAL_ACTIVE_TASKS.values()),
    ]
    for task in tasks:
        if task and not task.done():
            task.cancel()
    for task in tasks:
        if task:
            with suppress(asyncio.CancelledError):
                await task
    POLL_TASK = None
    DAILY_REPORT_TASK = None
    DAILY_BACKUP_TASK = None
    PERSISTENT_DATABASE_BACKUP_TASK = None
    _IPSHOP_PAYMENT_TASKS.clear()
    LOCAL_ORDER_POLL_TASKS.clear()
    LOCAL_DETAIL_TASKS.clear()
    LOCAL_ACTIVE_TASKS.clear()
    LOCAL_ORDER_VIEWS.clear()
    if API:
        await API.close()
        API = None
    if LOCAL_API:
        await LOCAL_API.close()
        LOCAL_API = None
    logger.info("Bot shut down")



# Register the new lightweight list action, then rebuild the one-shot pattern.
ADMIN_ONE_SHOT_ACTIONS["👥 User List"] = admin_user_list
ADMIN_ONE_SHOT_ACTIONS["📋 Payment List"] = admin_payment_list
ADMIN_ONE_SHOT_ACTIONS["🧾 Order Logs"] = admin_order_logs
ADMIN_ONE_SHOT_ACTIONS["💵 Provider Balance"] = admin_provider_balance
ONE_SHOT_MENU_PATTERN = r"^(?:" + "|".join(
    re.escape(label)
    for label in [
        *ADMIN_ONE_SHOT_ACTIONS.keys(),
        *USER_ONE_SHOT_ACTIONS.keys(),
        "🔙 Main Menu",
    ]
) + r")$"


async def workflow_one_shot_menu(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    """Execute a reply-keyboard action without purchase anti-spam.

    Reply-keyboard buttons are navigation and panel actions.  They are not
    purchase confirmations, so they must continue working during a purchase
    cooldown.
    """
    label = (update.effective_message.text or "").strip()

    clear_transient_workflow(context)
    if label == "🔙 Main Menu":
        await admin_back_main(update, context)
        return ConversationHandler.END
    if label in ADMIN_ONE_SHOT_ACTIONS:
        context.user_data["admin_mode"] = True
        result = ADMIN_ONE_SHOT_ACTIONS[label](update, context)
        if inspect.isawaitable(result):
            result = await result
        # Preserve the next ConversationHandler state for multi-step admin
        # actions such as Referral Reward, Set Button and Set Video.
        return result if isinstance(result, int) else ConversationHandler.END
    if label in USER_ONE_SHOT_ACTIONS:
        context.user_data.pop("admin_mode", None)
        result = USER_ONE_SHOT_ACTIONS[label](update, context)
        if inspect.isawaitable(result):
            result = await result
        return result if isinstance(result, int) else ConversationHandler.END
    return ConversationHandler.END


# =============================================================================
# 2026-07-29 — AUTO OTP / COMPACT PANELS / DAILY REPORTS / BALANCE BACKUP
# =============================================================================
# Final runtime overrides. One shared provider poller checks active orders every
# five seconds. It never edits Telegram messages and never sends "no OTP" spam.

# Automatic OTP polling for the single selected standard number per user.
# Use a 2-second interval as a speed/rate-limit compromise; never poll every second by default.
OTP_AUTO_POLL_SECONDS = max(1, int(os.getenv("OTP_AUTO_POLL_SECONDS", "1")))
_DHAKA_TIMEZONE_NAME = "Asia/Dhaka"

# Only the most recently opened standard active number per user is auto-polled.
# Other waiting numbers stay idle until the user opens/selects them; no manual check button is needed.
DB.execute("CREATE TABLE IF NOT EXISTS user_otp_poll_selection (user_id INTEGER PRIMARY KEY, order_id INTEGER NOT NULL, updated_at INTEGER NOT NULL)")

def _select_order_for_auto_otp(user_id: int, order_id: int) -> None:
    DB.execute(
        "INSERT INTO user_otp_poll_selection(user_id,order_id,updated_at) VALUES(?,?,?) "
        "ON CONFLICT(user_id) DO UPDATE SET order_id=excluded.order_id, updated_at=excluded.updated_at",
        (int(user_id), int(order_id), now_ts()),
    )


def _dhaka_day_bounds(ts: Optional[int] = None) -> tuple[int, int, str]:
    from datetime import timedelta
    from zoneinfo import ZoneInfo

    zone = ZoneInfo(_DHAKA_TIMEZONE_NAME)
    current = datetime.fromtimestamp(int(ts or now_ts()), zone)
    start = current.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=1)
    return int(start.timestamp()), int(end.timestamp()), start.date().isoformat()


def _seconds_left(value: Any) -> int:
    return max(0, int(value or 0) - now_ts())


def _dhaka_format_timestamp(value: Any) -> str:
    from zoneinfo import ZoneInfo

    try:
        return datetime.fromtimestamp(
            int(value), ZoneInfo(_DHAKA_TIMEZONE_NAME)
        ).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "-"


def ensure_daily_history_settings() -> None:
    with DB.lock:
        DB.conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES('archived_income_units','0')"
        )
        DB.conn.execute(
            "INSERT OR IGNORE INTO settings(key,value) VALUES('last_history_cleanup_day','')"
        )
        DB.conn.commit()


ensure_daily_history_settings()


def _income_units(start_ts: Optional[int] = None, end_ts: Optional[int] = None) -> int:
    clauses = []
    params: list[Any] = []
    if start_ts is not None:
        clauses.append("created_at>=?")
        params.append(int(start_ts))
    if end_ts is not None:
        clauses.append("created_at<?")
        params.append(int(end_ts))
    suffix = (" AND " + " AND ".join(clauses)) if clauses else ""

    global_row = DB.fetchone(
        "SELECT COALESCE(SUM(charged_units),0) AS total FROM orders "
        "WHERE refunded=0 AND status IN ('otp_received','expired')" + suffix,
        params,
    )
    local_row = DB.fetchone(
        "SELECT COALESCE(SUM(charged_units),0) AS total FROM local_orders "
        "WHERE refunded=0 AND status<>'cancelled'" + suffix,
        params,
    )
    return int(global_row["total"] or 0) + int(local_row["total"] or 0)


def _run_daily_cleanup_if_needed() -> None:
    """At Dhaka midnight archive income and remove old completed histories.

    Waiting/active numbers are retained. User accounts and current balances are
    never removed.
    """
    start_ts, _, day_key = _dhaka_day_bounds()
    with DB.lock:
        last_row = DB.conn.execute(
            "SELECT value FROM settings WHERE key='last_history_cleanup_day'"
        ).fetchone()
        if last_row and str(last_row[0] or "") == day_key:
            return

        # Preserve each user's latest product purchases before source order rows
        # are cleaned up below.
        sync_recent_purchase_history()

        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            global_old = DB.conn.execute(
                """
                SELECT COALESCE(SUM(charged_units),0)
                FROM orders
                WHERE created_at<? AND refunded=0
                  AND status IN ('otp_received','expired')
                """,
                (start_ts,),
            ).fetchone()[0]
            local_old = DB.conn.execute(
                """
                SELECT COALESCE(SUM(charged_units),0)
                FROM local_orders
                WHERE created_at<? AND refunded=0 AND status<>'cancelled'
                  AND status<>'active'
                """,
                (start_ts,),
            ).fetchone()[0]
            archived_row = DB.conn.execute(
                "SELECT value FROM settings WHERE key='archived_income_units'"
            ).fetchone()
            archived = int((archived_row[0] if archived_row else "0") or 0)
            archived += int(global_old or 0) + int(local_old or 0)
            DB.conn.execute(
                """
                INSERT INTO settings(key,value) VALUES('archived_income_units',?)
                ON CONFLICT(key) DO UPDATE SET value=excluded.value
                """,
                (str(archived),),
            )

            # Keep live financial/account state, remove only yesterday's finished logs.
            DB.conn.execute(
                "DELETE FROM orders WHERE created_at<? AND status<>'waiting'",
                (start_ts,),
            )
            DB.conn.execute(
                "DELETE FROM local_orders WHERE created_at<? AND status<>'active'",
                (start_ts,),
            )
            DB.conn.execute(
                """
                DELETE FROM deposits
                WHERE created_at<? AND status<>'pending'
                  AND NOT EXISTS (
                      SELECT 1 FROM referral_rewards
                      WHERE referral_rewards.deposit_id=deposits.id
                  )
                """,
                (start_ts,),
            )
            DB.conn.execute(
                "DELETE FROM transactions WHERE created_at<?",
                (start_ts,),
            )
            DB.conn.execute(
                """
                INSERT INTO settings(key,value) VALUES('last_history_cleanup_day',?)
                ON CONFLICT(key) DO UPDATE SET value=excluded.value
                """,
                (day_key,),
            )
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            raise


def _current_total_income_units() -> int:
    archived = int(DB.get_setting("archived_income_units", "0") or 0)
    return archived + _income_units()


async def admin_stats_reply(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_admin(update):
        return
    _run_daily_cleanup_if_needed()
    start_ts, end_ts, day_label = _dhaka_day_bounds()
    users = DB.fetchone("SELECT COUNT(*) AS n,COALESCE(SUM(balance_units),0) AS bal FROM users")
    today = DB.fetchone(
        """SELECT COUNT(*) AS bought,
                  SUM(CASE WHEN status='otp_received' THEN 1 ELSE 0 END) AS otp_count,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status IN ('otp_received','expired') THEN charged_units ELSE 0 END),0) AS revenue,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status IN ('otp_received','expired') THEN provider_cost_units ELSE 0 END),0) AS cost
           FROM orders WHERE created_at>=? AND created_at<?""", (start_ts, end_ts))
    local_today = DB.fetchone(
        """SELECT COUNT(*) AS bought,
                  COALESCE(SUM(CASE WHEN has_otp=1 THEN 1 ELSE 0 END),0) AS otp_count,
                  COALESCE(SUM(CASE WHEN refunded=0 AND status<>'cancelled' THEN charged_units ELSE 0 END),0) AS revenue
           FROM local_orders WHERE created_at>=? AND created_at<?""", (start_ts, end_ts))
    total = _current_total_income_units()
    today_revenue = int(today['revenue'] or 0) + int(local_today['revenue'] or 0)
    # Provider costs are stored in USD units; show and calculate them in BDT.
    cost = usd_units_to_bdt_units(int(today['cost'] or 0))
    await update.effective_message.reply_text(
        "<b>📊 Admin Statistics</b>\n\n"
        f"📅 <b>Today ({html.escape(day_label)})</b>\n"
        f"☎️ Numbers bought: <b>{int(today['bought'] or 0) + int(local_today['bought'] or 0)}</b>\n"
        f"📩 OTPs received: <b>{int(today['otp_count'] or 0) + int(local_today['otp_count'] or 0)}</b>\n"
        f"💵 User charges: <b>{money(today_revenue)}</b>\n"
        f"🏷️ Provider cost (standard orders): <b>{money(cost)}</b>\n"
        f"📈 Estimated profit (standard orders): <b>{money(int(today['revenue'] or 0) - cost)}</b>\n\n"
        f"👥 Total users: <b>{int(users['n'] or 0)}</b>\n"
        f"💰 Current user balances: <b>{money(int(users['bal'] or 0))}</b>\n"
        f"💳 All-time income: <b>{money(total)}</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=admin_menu(),
    )


async def show_orders(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await require_terms(update):
        return
    _run_daily_cleanup_if_needed()
    start_ts, end_ts, _ = _dhaka_day_bounds()
    user_id = int(update.effective_user.id)
    rows = DB.fetchall(
        """
        SELECT created_at,phone,charged_units,status,'global' AS source
        FROM orders WHERE user_id=? AND created_at>=? AND created_at<?
        UNION ALL
        SELECT created_at,phone,charged_units,status,'local' AS source
        FROM local_orders WHERE user_id=? AND created_at>=? AND created_at<?
        ORDER BY created_at DESC LIMIT 10
        """,
        (user_id, start_ts, end_ts, user_id, start_ts, end_ts),
    )
    if not rows:
        await update.effective_message.reply_text("No orders today.")
        return
    lines = ["<b>Today's Orders</b>"]
    for row in rows:
        phone = str(row["phone"] or "-")
        if row["source"] == "local" and phone and not phone.startswith("+"):
            phone = "+" + phone
        lines.append(
            f"\n☎️ <code>{html.escape(phone)}</code> • "
            f"<b>{money(int(row['charged_units'] or 0))}</b>\n"
            f"{html.escape(status_label(str(row['status'] or 'unknown')))}"
        )
    await update.effective_message.reply_text(
        "\n".join(lines), parse_mode=ParseMode.HTML
    )


# Country calling codes used only for the user's copy/display preference. The
# provider-facing phone number in the database is never changed.
_COUNTRY_DIAL_CODES = {
    "united states": "1", "usa": "1", "canada": "1", "united kingdom": "44", "uk": "44",
    "russia": "7", "kazakhstan": "7", "egypt": "20", "south africa": "27", "greece": "30",
    "netherlands": "31", "belgium": "32", "france": "33", "spain": "34", "hungary": "36",
    "italy": "39", "romania": "40", "switzerland": "41", "austria": "43", "guernsey": "44",
    "denmark": "45", "sweden": "46", "norway": "47", "poland": "48", "germany": "49",
    "peru": "51", "mexico": "52", "cuba": "53", "argentina": "54", "brazil": "55",
    "chile": "56", "colombia": "57", "venezuela": "58", "malaysia": "60", "australia": "61",
    "indonesia": "62", "philippines": "63", "new zealand": "64", "singapore": "65", "thailand": "66",
    "japan": "81", "south korea": "82", "vietnam": "84", "china": "86", "turkey": "90",
    "india": "91", "pakistan": "92", "afghanistan": "93", "sri lanka": "94", "myanmar": "95",
    "iran": "98", "south sudan": "211", "morocco": "212", "algeria": "213", "tunisia": "216",
    "libya": "218", "gambia": "220", "senegal": "221", "mauritania": "222", "mali": "223",
    "guinea": "224", "ivory coast": "225", "burkina faso": "226", "niger": "227", "togo": "228",
    "benin": "229", "mauritius": "230", "liberia": "231", "sierra leone": "232", "ghana": "233",
    "nigeria": "234", "chad": "235", "central african republic": "236", "cameroon": "237", "cape verde": "238",
    "sao tome and principe": "239", "equatorial guinea": "240", "gabon": "241", "republic of the congo": "242",
    "democratic republic of the congo": "243", "angola": "244", "guinea-bissau": "245", "seychelles": "248",
    "sudan": "249", "rwanda": "250", "ethiopia": "251", "somalia": "252", "djibouti": "253",
    "kenya": "254", "tanzania": "255", "uganda": "256", "burundi": "257", "mozambique": "258",
    "zambia": "260", "madagascar": "261", "reunion": "262", "zimbabwe": "263", "namibia": "264",
    "malawi": "265", "lesotho": "266", "botswana": "267", "eswatini": "268", "comoros": "269",
    "eritrea": "291", "aruba": "297", "faroe islands": "298", "greenland": "299", "gibraltar": "350",
    "portugal": "351", "luxembourg": "352", "ireland": "353", "iceland": "354", "albania": "355",
    "malta": "356", "cyprus": "357", "finland": "358", "bulgaria": "359", "lithuania": "370",
    "latvia": "371", "estonia": "372", "moldova": "373", "armenia": "374", "belarus": "375",
    "andorra": "376", "monaco": "377", "san marino": "378", "ukraine": "380", "serbia": "381",
    "montenegro": "382", "croatia": "385", "slovenia": "386", "bosnia and herzegovina": "387",
    "north macedonia": "389", "czech republic": "420", "slovakia": "421", "liechtenstein": "423",
    "falkland islands": "500", "belize": "501", "guatemala": "502", "el salvador": "503", "honduras": "504",
    "nicaragua": "505", "costa rica": "506", "panama": "507", "haiti": "509", "bolivia": "591",
    "guyana": "592", "ecuador": "593", "paraguay": "595", "uruguay": "598", "curacao": "599",
    "taiwan": "886", "bangladesh": "880", "united arab emirates": "971", "israel": "972",
    "bahrain": "973", "qatar": "974", "bhutan": "975", "mongolia": "976", "nepal": "977",
    "tajikistan": "992", "turkmenistan": "993", "azerbaijan": "994", "georgia": "995", "kyrgyzstan": "996",
    "uzbekistan": "998", "saudi arabia": "966", "yemen": "967", "oman": "968", "palestine": "970",
    "jordan": "962", "lebanon": "961", "iraq": "964", "kuwait": "965", "syria": "963",
}

def _cc_removed(user_id: int) -> bool:
    row = DB.fetchone("SELECT value FROM settings WHERE key=?", (f"cc_removed:{int(user_id)}",))
    return bool(row and str(row["value"]).strip() == "1")

def _set_cc_removed(user_id: int, removed: bool) -> None:
    DB.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (f"cc_removed:{int(user_id)}", "1" if removed else "0"))

def _display_phone(order: sqlite3.Row) -> str:
    phone = re.sub(r"\s+", "", str(order["phone"] or ""))
    if not _cc_removed(int(order["user_id"])):
        return phone
    country = re.sub(r"[^a-z0-9 ]", "", str(order["country_name"] or "").lower()).strip()
    dial = _COUNTRY_DIAL_CODES.get(country)
    if not dial:
        # Country name may contain a service/state suffix (e.g. USA State Virtual).
        for name, code in sorted(_COUNTRY_DIAL_CODES.items(), key=lambda x: len(x[0]), reverse=True):
            if country == name or country.startswith(name + " "):
                dial = code
                break
    if not dial:
        return phone
    digits = re.sub(r"\D", "", phone)
    if phone.startswith("+") and digits.startswith(dial) and len(digits) > len(dial) + 4:
        return digits[len(dial):]
    if not phone.startswith("+") and digits.startswith(dial) and len(digits) > len(dial) + 7:
        return digits[len(dial):]
    return phone

def _cc_toggle_button(order: sqlite3.Row) -> InlineKeyboardButton:
    removed = _cc_removed(int(order["user_id"]))
    # Button text describes the next action; Telegram button style provides the color.
    if removed:
        label = "CC Add"
        style = "success"
    else:
        label = "CC Remove"
        style = "danger"
    return InlineKeyboardButton(
        label,
        callback_data=f"order:cc:{order['id']}",
        style=style,
    )


def render_order(order: sqlite3.Row) -> str:
    user = DB.fetchone(
        "SELECT balance_units FROM users WHERE user_id=?", (order["user_id"],)
    )
    balance = int(user["balance_units"] or 0) if user else int(
        order["balance_after_units"] or 0
    )
    status = str(order["status"] or "unknown")
    lines = [
        f"💳 <b>{money(int(order['charged_units'] or 0))}</b>",
        f"💰 <b>{money(balance)}</b>",
    ]
    if status == "waiting":
        lines.append(f"⏳ <b>{_seconds_left(order['expires_at'])}s</b>")
    elif status == "otp_received":
        lines.append("✅ OTP received")
    elif status == "cancelled":
        lines.append("↩️ Cancelled + refunded")
    elif status == "expired":
        lines.append("⌛ Expired")
    else:
        lines.append(f"📌 {html.escape(status_label(status))}")
    return "\n".join(lines)


def order_keyboard(order: sqlite3.Row) -> Optional[InlineKeyboardMarkup]:
    rows: list[list[InlineKeyboardButton]] = []
    if order["status"] == "waiting":
        shown_phone = _display_phone(order)
        rows.append([
            copy_button(f"📋 {shown_phone}", shown_phone),
        ])
        if int(order["cc_controls_enabled"] or 0):
            rows.append([_cc_toggle_button(order)])
        rows.append([
            InlineKeyboardButton("⏱ Time", callback_data=f"order:time:{order['id']}"),
            InlineKeyboardButton("❌ Cancel", callback_data=f"order:cancel:{order['id']}"),
        ])
        if not int(order["ever_received_otp"] or 0):
            rows.append([InlineKeyboardButton("♻️ Refund & New", callback_data=f"order:refund_new:{order['id']}", style="secondary")])
    elif order["status"] == "otp_received":
        shown_phone = _display_phone(order)
        rows.append([
            copy_button(f"📋 {shown_phone}", shown_phone),
            copy_button(f"🔐 {order['otp_code'] or 'OTP'}", str(order["otp_code"] or "")),
        ])
        actions = []
        if order["catalog_id"] and order["provider_id"]:
            actions.append(InlineKeyboardButton("➕ New নাম্বার", callback_data=f"order:new:{order['id']}", style="secondary"))
            if now_ts() < int(order["expires_at"] or 0):
                actions.append(InlineKeyboardButton("📩 Another Message", callback_data=f"order:another:{order['id']}"))
        if actions:
            rows.append(actions)
    return InlineKeyboardMarkup(rows) if rows else None


async def order_time_refresh_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        await query.answer("Invalid order.", show_alert=True)
        return
    order = DB.fetchone(
        "SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id)
    )
    if not order:
        await query.answer("Order not found.", show_alert=True)
        return
    await query.answer(
        f"Time left: {_seconds_left(order['expires_at'])} seconds",
        show_alert=True,
    )


async def send_otp_once(
    application: Application,
    order_id: int,
    code: str,
    *,
    retry_only_missing: bool = False,
) -> None:
    """Deliver an OTP, optionally retrying only destinations that previously failed."""
    order = DB.fetchone("SELECT * FROM orders WHERE id=?", (order_id,))
    if not order:
        return
    # A Telegram chat-not-found/blocked response is permanent for this order.
    # Persist the state so the background poller cannot retry on every cycle.
    if int(order["otp_delivery_blocked"] or 0):
        logger.debug("OTP delivery skipped for order %s: destination previously marked unavailable", order_id)
        return
    # Keep the same number card and edit it in place when OTP arrives.
    rows = [[
        copy_button(f"📋 {_display_phone(order)}", _display_phone(order)),
        copy_button(f"🔐 {code}", str(code)),
    ]]
    actions = []
    if order["catalog_id"] and order["provider_id"]:
        actions.append(InlineKeyboardButton("➕ New নাম্বার", callback_data=f"order:new:{order_id}", style="secondary"))
        if now_ts() < int(order["expires_at"] or 0):
            actions.append(InlineKeyboardButton("📩 Another Message", callback_data=f"order:another:{order_id}"))
    if actions:
        rows.append(actions)
    user_otp_delivered = bool(retry_only_missing and order["otp_message_id"])
    sent_message_id = int(order["otp_message_id"] or 0) if user_otp_delivered else 0
    target_message_id = order["message_id"] or order["otp_message_id"]
    otp_text = (
        "✅ <b>OTP received</b>\n"
        f"☎️ <code>{html.escape(_display_phone(order))}</code>\n"
        f"🔐 <code>{html.escape(str(code))}</code>"
    )

    # First try to update the existing number card. If Telegram rejects the
    # edit (deleted/too old/wrong message), remove the stale card if possible
    # and send a fresh OTP card so the user never gets stuck with only the
    # old copy-number button.
    edit_failed = False
    if not user_otp_delivered and target_message_id:
        try:
            await application.bot.edit_message_text(
                chat_id=order["user_id"],
                message_id=int(target_message_id),
                text=otp_text,
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup(rows),
            )
            sent_message_id = int(target_message_id)
            user_otp_delivered = True
        except TelegramError as exc:
            error_text = str(exc).casefold()
            if any(marker in error_text for marker in (
                "chat not found", "bot was blocked by the user",
                "user is deactivated", "peer_id_invalid",
            )):
                DB.execute("UPDATE orders SET otp_delivery_blocked=1 WHERE id=?", (order_id,))
                logger.warning(
                    "OTP delivery disabled for order %s: Telegram destination %s is unavailable (%s)",
                    order_id, order["user_id"], short_text(str(exc), 180),
                )
                return
            edit_failed = True
            logger.warning(
                "Could not edit OTP card for order %s; will replace it: %s",
                order_id, exc,
            )
            try:
                await application.bot.delete_message(
                    chat_id=order["user_id"], message_id=int(target_message_id)
                )
            except TelegramError as delete_exc:
                logger.debug(
                    "Could not delete stale OTP card for order %s: %s",
                    order_id, delete_exc,
                )

    # Fresh send (also used when there was no previous card).
    if not user_otp_delivered:
        for attempt in range(1, 4):
            try:
                if attempt > 1:
                    await asyncio.sleep(1)
                sent = await application.bot.send_message(
                    chat_id=order["user_id"],
                    text=otp_text,
                    parse_mode=ParseMode.HTML,
                    reply_markup=InlineKeyboardMarkup(rows),
                )
                sent_message_id = sent.message_id
                user_otp_delivered = True
                break
            except TelegramError as exc:
                error_text = str(exc).casefold()
                # A missing/blocked Telegram chat is permanent for this delivery.
                # Retrying the same chat three times only creates noisy logs.
                if any(marker in error_text for marker in (
                    "chat not found", "bot was blocked by the user",
                    "user is deactivated", "peer_id_invalid",
                )):
                    DB.execute("UPDATE orders SET otp_delivery_blocked=1 WHERE id=?", (order_id,))
                    logger.warning(
                        "OTP delivery disabled for order %s: Telegram chat %s is unavailable (%s). "
                        "Future background retries are disabled for this order.",
                        order_id, order["user_id"], short_text(str(exc), 180),
                    )
                    break
                logger.warning(
                    "OTP user delivery attempt %s/3 failed for order %s: %s",
                    attempt, order_id, exc,
                )

    if user_otp_delivered:
        DB.execute(
            "UPDATE orders SET message_id=?, otp_message_id=? WHERE id=?",
            (sent_message_id, sent_message_id, order_id),
        )
    else:
        logger.error(
            "OTP %s was not delivered to user %s after retries",
            order_id, order["user_id"],
        )

    # Forward every newly received OTP, including codes requested through
    # "Another Message". The per-order flags now identify whether this is a
    # follow-up code; they no longer suppress later OTP forwarding.
    order = DB.fetchone("SELECT * FROM orders WHERE id=?", (order_id,))
    if not order:
        return

    public_group = _setting_group_id("otp_forward_group_1")
    public_was_forwarded = bool(int(order["otp_public_forwarded"] or 0))
    if public_group and not (retry_only_missing and public_was_forwarded):
        public_title = "🔁 <b>OTP ANOTHER</b>" if public_was_forwarded else "✅ <b>OTP RECEIVED</b>"
        try:
            await application.bot.send_message(
                chat_id=public_group,
                text=(
                    f"{public_title}\n"
                    f"☎️ <code>{html.escape(str(order['phone']))}</code>\n"
                    f"🔐 <code>{html.escape(str(code))}</code>"
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([[copy_button(f"🔐 {code}", str(code))]]),
            )
            DB.execute("UPDATE orders SET otp_public_forwarded=1 WHERE id=?", (order_id,))
        except TelegramError as exc:
            logger.warning("Could not forward OTP order %s to group 1: %s", order_id, exc)

    private_group = _setting_group_id("otp_admin_group_2")
    private_was_forwarded = bool(int(order["otp_private_forwarded"] or 0))
    if private_group and not (retry_only_missing and private_was_forwarded):
        user = DB.fetchone("SELECT * FROM users WHERE user_id=?", (order["user_id"],))
        username = f"@{user['username']}" if user and user["username"] else "-"
        balance = int(user["balance_units"] or 0) if user else 0
        private_title = "🔁 <b>OTP ANOTHER</b>" if private_was_forwarded else "🔐 <b>OTP</b>"
        try:
            await application.bot.send_message(
                chat_id=private_group,
                text=(
                    f"{private_title}\n"
                    f"👤 <code>{order['user_id']}</code> • {html.escape(username)}\n"
                    f"☎️ <code>{html.escape(str(order['phone']))}</code>\n"
                    f"🔐 <code>{html.escape(str(code))}</code>\n"
                    f"💰 <b>{money(balance)}</b>"
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([[copy_button(f"🔐 {code}", str(code))]]),
            )
            DB.execute("UPDATE orders SET otp_private_forwarded=1 WHERE id=?", (order_id,))
        except TelegramError as exc:
            logger.warning("Could not forward OTP order %s to group 2: %s", order_id, exc)



async def check_order(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Manual check returns only a popup; it never creates a status message."""
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        await query.answer("Invalid order.", show_alert=True)
        return
    order = DB.fetchone(
        "SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id)
    )
    if not order:
        await query.answer("Order not found.", show_alert=True)
        return
    if order["status"] == "otp_received" and order["otp_code"]:
        await query.answer(f"OTP: {order['otp_code']}", show_alert=True)
        return
    if order["status"] != "waiting":
        await query.answer(status_label(str(order["status"])), show_alert=True)
        return

    lock = _order_process_lock(order_id)
    async with lock:
        order = DB.fetchone(
            "SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id)
        )
        if not order:
            await query.answer("Order not found.", show_alert=True)
            return
        try:
            state, value, refreshed = await _manual_standard_status_check(
                context.application, order
            )
        except Exception as exc:
            await query.answer(
                "Provider error: " + short_text(str(exc), 150), show_alert=True
            )
            return

    if state == "code":
        await send_otp_once(context.application, order_id, value)
        await query.answer(f"OTP: {value}", show_alert=True)
    elif state == "waiting":
        await query.answer("No OTP received.", show_alert=True)
    elif state == "expired":
        await query.answer("Number expired.", show_alert=True)
    elif state == "cancelled":
        await query.answer("Cancelled and refunded.", show_alert=True)
    else:
        await query.answer(short_text(value, 180), show_alert=True)


async def _auto_poll_standard_order(
    application: Application, order_id: int
) -> None:
    lock = _order_process_lock(order_id)
    if lock.locked():
        return
    async with lock:
        order = DB.fetchone(
            "SELECT * FROM orders WHERE id=? AND status='waiting'", (order_id,)
        )
        if not order:
            return
        try:
            # Fixed 20-minute cutoff for standard Buy Number.
            if now_ts() >= int(order["expires_at"]):
                await _api_expire_standard_order_without_refund(application, order)
                # Deliberately no Telegram message for the automatic 20-minute
                # cancellation. The order simply disappears from Active Orders.
                return

            state, value, refreshed = await _manual_standard_status_check(
                application, order
            )
            if state == "code":
                await send_otp_once(application, int(refreshed["id"]), value)
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Auto standard OTP/expiry check failed for order %s", order_id)



async def active_number_list_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    await query.answer()
    rows = DB.fetchall(
        "SELECT * FROM orders WHERE user_id=? AND status IN ('waiting','otp_received') AND expires_at>? ORDER BY created_at DESC",
        (user.id, now_ts()),
    )
    if not rows:
        await query.message.reply_text("📋 Active Number List\n\nএখন কোনো active number নেই।")
        return
    buttons = []
    for order in rows[:50]:
        label = f"{'✅' if order['otp_code'] else '⏳'} {_display_phone(order)}"
        if order['otp_code']:
            label += f" • OTP {order['otp_code']}"
        label += f" • {_seconds_left(order['expires_at'])}s"
        buttons.append([InlineKeyboardButton(label[:60], callback_data=f"order:active:view:{order['id']}")])
    buttons.append([InlineKeyboardButton("⬅️ Services", callback_data="buy:menu")])
    await query.message.reply_text("📋 <b>Active Number List</b>\n২০ মিনিটের মধ্যে থাকা number-গুলো:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(buttons))


async def active_number_view_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        await query.answer("Invalid number.", show_alert=True)
        return
    order = DB.fetchone("SELECT * FROM orders WHERE id=? AND user_id=? AND status IN ('waiting','otp_received') AND expires_at>?", (order_id, user.id, now_ts()))
    if not order:
        await query.answer("এই number-এর ২০ মিনিটের সময় শেষ হয়েছে।", show_alert=True)
        return
    _select_order_for_auto_otp(int(user.id), int(order_id))
    await query.answer("এই নম্বরের OTP auto-check চালু হয়েছে। অন্য নম্বর auto-check হবে না।")
    text = f"☎️ <b>{html.escape(_display_phone(order))}</b>\n⏳ সময় বাকি: <b>{_seconds_left(order['expires_at'])}s</b>"
    if order["otp_code"]:
        text += f"\n🔐 OTP: <code>{html.escape(str(order['otp_code']))}</code>"
    await query.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=order_keyboard(order))


async def cc_toggle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        await query.answer("Invalid number.", show_alert=True)
        return
    order = DB.fetchone("SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id))
    if not order or order["status"] != "waiting" or not int(order["cc_controls_enabled"] or 0):
        await query.answer("এই নম্বরে CC সেটিং পরিবর্তন করা যাবে না।", show_alert=True)
        return
    removed = not _cc_removed(user.id)
    _set_cc_removed(user.id, removed)
    refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order_id,))
    try:
        await query.message.edit_reply_markup(reply_markup=order_keyboard(refreshed))
    except TelegramError as exc:
        logger.debug("Could not refresh CC button for order %s: %s", order_id, exc)
    await query.answer("Country code removed" if removed else "Country code added", show_alert=False)


async def another_message_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Request another SMS for an active Buy Number order.

    There is intentionally NO bot-side OTP/Another Message count limit. The
    provider decides whether another SMS can be issued; the order's existing
    20-minute expiry remains the only time limit for standard Buy Number.
    """
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        order_id = int(query.data.rsplit(":", 1)[1])
    except Exception:
        await query.answer("Invalid number.", show_alert=True)
        return
    lock = _order_process_lock(order_id)
    async with lock:
        order = DB.fetchone("SELECT * FROM orders WHERE id=? AND user_id=?", (order_id, user.id))
        if not order or order["status"] != "otp_received" or now_ts() >= int(order["expires_at"] or 0):
            await query.answer("এই number-এর ২০ মিনিটের সময় শেষ, অথবা number active নেই।", show_alert=True)
            return
        if not API:
            await query.answer("Provider API is not ready.", show_alert=True)
            return
        # The user explicitly requested another SMS for this number, so make it
        # their sole automatically-polled standard order until another is opened.
        _select_order_for_auto_otp(int(user.id), int(order_id))
        item = get_catalog_item(order["catalog_id"]) if order["catalog_id"] else None
        extra = catalog_extra(item) if item else {}
        try:
            data, raw = await API.set_status(str(order["activation_id"]), 3, extra)
        except Exception as exc:
            await query.answer("API error: " + short_text(str(exc), 120), show_alert=True)
            return
        upper = str(raw or "").upper()
        success = any(token in upper for token in ("ACCESS_RETRY_GET", "STATUS_WAIT_RETRY", "ACCESS_READY", "SUCCESS"))
        if not success and isinstance(data, dict):
            success = str(data.get("status", "")).lower() in {"success", "ok", "ready", "retry"}
        if not success:
            await query.answer("Provider did not accept another SMS: " + short_text(raw, 100), show_alert=True)
            return
        # Reuse the existing OTP/number card: edit it in place, never delete it.
        DB.execute("UPDATE orders SET status='waiting', otp_code=NULL, cc_controls_enabled=0, last_api_status=?, updated_at=? WHERE id=?", (short_text(raw, 500), now_ts(), order_id))
        refreshed = DB.fetchone("SELECT * FROM orders WHERE id=?", (order_id,))
        try:
            active_text = (
                f"☎️ <b>{html.escape(_display_phone(refreshed))}</b>\n"
                f"⏳ সময় বাকি: <b>{_seconds_left(refreshed['expires_at'])}s</b>"
            )
            await query.message.edit_text(
                active_text,
                parse_mode=ParseMode.HTML,
                reply_markup=order_keyboard(refreshed),
            )
            if order["otp_message_id"]:
                DB.execute("UPDATE orders SET message_id=?, otp_message_id=NULL WHERE id=?", (query.message.message_id, order_id))
        except TelegramError as exc:
            logger.warning("Could not edit Another Message card for order %s: %s", order_id, exc)
        await query.answer("📩 Another SMS requested. এই মেসেজটি আপডেট করা হয়েছে। নতুন code-এর জন্য অপেক্ষা করুন।")


def local_detail_text(order: sqlite3.Row) -> str:
    user = DB.fetchone(
        "SELECT balance_units FROM users WHERE user_id=?", (order["user_id"],)
    )
    balance = int(user["balance_units"] or 0) if user else int(
        order["balance_after_units"] or 0
    )
    session = _seconds_left(order["expires_at"])
    current = now_ts()
    cancel_end = int(order["created_at"]) + LOCAL_CANCEL_WINDOW_SECONDS
    if int(order["has_otp"] or 0):
        cancel_text = "OTP received"
    elif current < int(order["cancel_after"]):
        cancel_text = f"Cancel locked • {_seconds_left(order['cancel_after'])}"
    elif current < cancel_end:
        cancel_text = f"Cancel available • {_seconds_left(cancel_end)}"
    else:
        cancel_text = "Auto refund processing"
    return (
        "<b>🇺🇸 USA Local</b>\n"
        f"💳 <b>{money(int(order['charged_units'] or 0))}</b>\n"
        f"💰 <b>{money(balance)}</b>\n"
        f"⏳ <b>{session}s</b>\n"
        f"↩️ {html.escape(cancel_text)}"
    )


def local_detail_markup(
    order: sqlite3.Row, page: int = 0
) -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = [
        [copy_button(f"📋 +{order['phone']}", f"+{order['phone']}")]
    ]
    if order["status"] == "active":
        rows.append([
            InlineKeyboardButton(
                "📩 Check Message", callback_data=f"local:refresh:{order['id']}:{page}"
            ),
            InlineKeyboardButton(
                "⏱ Time", callback_data=f"local:time:{order['id']}:{page}"
            ),
        ])
        if not int(order["has_otp"] or 0):
            rows.append([
                InlineKeyboardButton(
                    "❌ Cancel & Refund",
                    callback_data=f"local:cancel:{order['id']}:{page}",
                )
            ])
    return InlineKeyboardMarkup(rows)


async def local_time_refresh_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        _, _, order_raw, _page_raw = query.data.split(":", 3)
        order_id = int(order_raw)
    except Exception:
        await query.answer("Invalid number.", show_alert=True)
        return
    order = local_get_user_order(order_id, user.id)
    if not order:
        await query.answer("Number not found.", show_alert=True)
        return
    current = now_ts()
    cancel_end = int(order["created_at"]) + LOCAL_CANCEL_WINDOW_SECONDS
    if int(order["has_otp"] or 0):
        cancel_text = "OTP received; cancel blocked"
    elif current < int(order["cancel_after"]):
        cancel_text = f"cancel in {_seconds_left(order['cancel_after'])}s"
    elif current < cancel_end:
        cancel_text = f"cancel available for {cancel_end-current}s"
    else:
        cancel_text = "cancel window expired"
    await query.answer(
        f"Session: {_seconds_left(order['expires_at'])}s • {cancel_text}",
        show_alert=True,
    )


async def local_update_user_otp_message(
    application: Application,
    order: sqlite3.Row,
    otp: str,
    raw_response: str = "",
) -> None:
    rows = [
        [copy_button(f"📋 +{order['phone']}", f"+{order['phone']}")],
        [copy_button(f"🔐 {otp}", str(otp))],
        [InlineKeyboardButton("➕ New Number", callback_data="local:buy", style="secondary")],
    ]
    try:
        sent = await application.bot.send_message(
            chat_id=order["user_id"],
            text="✅ <b>USA Local OTP received</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(rows),
        )
        DB.execute(
            "UPDATE local_orders SET otp_message_id=? WHERE id=?",
            (sent.message_id, order["id"]),
        )
    except TelegramError as exc:
        logger.warning("Could not send local OTP to user: %s", exc)


async def local_refresh_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Manual local check uses popup only; new OTP delivery remains one message."""
    query = update.callback_query
    user = update.effective_user
    if not query or not user:
        return
    try:
        _, _, order_raw, _page_raw = query.data.split(":", 3)
        order_id = int(order_raw)
    except Exception:
        await query.answer("Invalid number.", show_alert=True)
        return
    order = local_get_user_order(order_id, user.id)
    if not order:
        await query.answer("Number not found.", show_alert=True)
        return
    if int(order["has_otp"] or 0) and order["latest_otp"]:
        await query.answer(f"OTP: {order['latest_otp']}", show_alert=True)
        return
    if order["status"] != "active":
        await query.answer("Session ended.", show_alert=True)
        return
    if not LOCAL_API:
        await query.answer("Local API is not ready.", show_alert=True)
        return

    lock = _LOCAL_ORDER_CHECK_LOCKS.setdefault(order_id, asyncio.Lock())
    async with lock:
        try:
            result = await LOCAL_API.get_local_otp(
                str(order["phone"]), seen_otps=local_seen_otps(order_id)
            )
            is_new, message = await local_apply_otp(
                context.application, order_id, result
            )
        except Exception as exc:
            await query.answer(
                "Provider error: " + short_text(str(exc), 150), show_alert=True
            )
            return
    refreshed = local_get_order(order_id) or order
    if is_new and refreshed["latest_otp"]:
        await query.answer(f"OTP: {refreshed['latest_otp']}", show_alert=True)
    else:
        await query.answer("No OTP received.", show_alert=True)


async def _auto_expire_local_order(
    application: Application,
    order_id: int,
) -> None:
    """At 60 minutes without OTP, cancel provider-side, refund, and notify user."""
    lock = _LOCAL_ORDER_CHECK_LOCKS.setdefault(order_id, asyncio.Lock())
    if lock.locked():
        return

    async with lock:
        order = DB.fetchone(
            "SELECT * FROM local_orders WHERE id=? AND status='active'",
            (order_id,),
        )
        if not order or int(order["has_otp"] or 0):
            return
        if now_ts() < int(order["expires_at"]) or not LOCAL_API:
            return

        # Provider cancellation failures are stored with updated_at. The shared
        # poller runs every few seconds, so back off before retrying a failed
        # cancellation instead of hammering the provider and flooding the log.
        last_cancel_response = str(order["provider_cancel_response"] or "").strip()
        last_update = int(order["updated_at"] or 0)
        if last_cancel_response and now_ts() - last_update < 900:
            return

        try:
            # Final OTP check prevents a race exactly at the 60-minute boundary.
            latest = await LOCAL_API.get_local_otp(
                str(order["phone"]),
                seen_otps=local_seen_otps(order_id),
            )
            is_new, _ = await local_apply_otp(application, order_id, latest)
            refreshed = local_get_order(order_id)
            if is_new or not refreshed or int(refreshed["has_otp"] or 0):
                return

            success, message, raw = await LOCAL_API.cancel_local(
                str(refreshed["phone"])
            )
            if not success:
                # Log the failure once per retry window; the next attempt is
                # delayed by the 15-minute backoff above.
                logger.warning(
                    "Local auto-expiry cancellation failed for order %s: %s",
                    order_id, short_text(message or raw, 500),
                )
                DB.execute(
                    """
                    UPDATE local_orders
                    SET provider_cancel_response=?, raw_last_response=?, updated_at=?
                    WHERE id=? AND status='active'
                    """,
                    (raw[:10000], raw[:10000], now_ts(), order_id),
                )
                return

            refunded = local_refund_order(
                order_id, raw or message, provider_confirmed=True
            )
            if not refunded:
                return

            final_order = DB.fetchone(
                "SELECT * FROM local_orders WHERE id=?", (order_id,)
            ) or refreshed

            try:
                await application.bot.send_message(
                    chat_id=final_order["user_id"],
                    text=(
                        "✅ <b>USA Local Number Refunded</b>\n\n"
                        f"📱 <code>+{html.escape(str(final_order['phone']))}</code>\n"
                        f"↩️ <b>ফেরত: {money(int(final_order['charged_units'] or 0))}</b>\n\n"
                        "প্রথম ৬০ মিনিটের মধ্যে কোনো OTP না আসায় "
                        "নাম্বারটি automatically refund করা হয়েছে।"
                    ),
                    parse_mode=ParseMode.HTML,
                )
            except TelegramError as exc:
                logger.warning(
                    "Could not notify user about local auto-refund %s: %s",
                    order_id, exc,
                )
        except Exception as exc:
            logger.exception("Local auto-expiry failed for %s: %s", order_id, exc)


async def _auto_poll_local_order(application: Application, order_id: int) -> None:
    lock = _LOCAL_ORDER_CHECK_LOCKS.setdefault(order_id, asyncio.Lock())
    if lock.locked():
        return
    async with lock:
        order = local_get_order(order_id)
        if not order or order["status"] != "active":
            return
        try:
            result = await LOCAL_API.get_local_otp(
                str(order["phone"]), seen_otps=local_seen_otps(order_id)
            )
            await local_apply_otp(application, order_id, result)
        except Exception as exc:
            logger.debug("Auto local OTP check failed for %s: %s", order_id, exc)


def ensure_local_poll_task(application: Application, order_id: int) -> None:
    """The shared five-second poller owns all local OTP checks."""
    return None


async def _retry_pending_otp_delivery(application: Application, order_id: int) -> None:
    """Retry Telegram delivery for OTPs whose user/group delivery previously failed."""
    order = DB.fetchone(
        "SELECT * FROM orders WHERE id=? AND status='otp_received' AND otp_code IS NOT NULL",
        (order_id,),
    )
    if not order:
        return
    user_pending = not bool(order["otp_message_id"])
    public_group = _setting_group_id("otp_forward_group_1")
    private_group = _setting_group_id("otp_admin_group_2")
    public_pending = bool(public_group) and not bool(int(order["otp_public_forwarded"] or 0))
    private_pending = bool(private_group) and not bool(int(order["otp_private_forwarded"] or 0))
    if not (user_pending or public_pending or private_pending):
        return
    try:
        await send_otp_once(
            application, int(order_id), str(order["otp_code"]), retry_only_missing=True
        )
    except asyncio.CancelledError:
        raise
    except Exception:
        logger.exception("Pending OTP delivery retry failed for order %s", order_id)


async def polling_loop(application: Application) -> None:
    logger.info("Shared OTP auto-poller started: every %ss", OTP_AUTO_POLL_SECONDS)
    while True:
        try:
            _run_daily_cleanup_if_needed()
            now_value = now_ts()
            # Expiry processing still covers every waiting order, but provider OTP
            # status requests are limited to each user's selected active number.
            standard_due = DB.fetchall(
                "SELECT id FROM orders WHERE status='waiting' AND expires_at<=? ORDER BY id LIMIT ?",
                (now_value, SMS_POLL_BATCH_SIZE),
            )
            standard = DB.fetchall(
                "SELECT o.id FROM orders o JOIN user_otp_poll_selection s "
                "ON s.user_id=o.user_id AND s.order_id=o.id "
                "WHERE o.status='waiting' AND o.expires_at>? ORDER BY o.id LIMIT ?",
                (now_value, SMS_POLL_BATCH_SIZE),
            )
            pending_otp_delivery = DB.fetchall(
                "SELECT id FROM orders WHERE status='otp_received' AND otp_code IS NOT NULL "
                "ORDER BY updated_at DESC LIMIT ?",
                (SMS_POLL_BATCH_SIZE,),
            )
            local_expire_due_orders()
            local_due = DB.fetchall(
                "SELECT id FROM local_orders "
                "WHERE status='active' AND has_otp=0 AND expires_at<=? "
                "ORDER BY id LIMIT ?",
                (now_ts(), SMS_POLL_BATCH_SIZE),
            )
            local = DB.fetchall(
                "SELECT id FROM local_orders "
                "WHERE status='active' AND expires_at>? ORDER BY id LIMIT ?",
                (now_ts(), SMS_POLL_BATCH_SIZE),
            )
            # Only contact SMS providers when there is an order to monitor.
            # With no purchased/waiting or active local numbers, the loop stays
            # idle and does not make any OTP-status API requests.
            has_standard_orders = bool(standard)
            has_pending_otp_delivery = bool(pending_otp_delivery)
            has_standard_expiry_work = bool(standard_due)
            has_local_expiry_work = bool(local_due) and bool(LOCAL_API)
            has_local_active_orders = bool(local) and bool(LOCAL_API)
            if not (has_standard_orders or has_pending_otp_delivery or has_standard_expiry_work or has_local_expiry_work or has_local_active_orders):
                await asyncio.sleep(OTP_AUTO_POLL_SECONDS)
                continue

            jobs = [
                _auto_poll_standard_order(application, int(row["id"]))
                for row in standard
            ]
            jobs.extend(
                _retry_pending_otp_delivery(application, int(row["id"]))
                for row in pending_otp_delivery
            )
            jobs.extend(
                _api_expire_standard_order_without_refund(application, DB.fetchone("SELECT * FROM orders WHERE id=?", (int(row["id"]),)))
                for row in standard_due
            )
            if LOCAL_API:
                jobs.extend(
                    _auto_expire_local_order(application, int(row["id"]))
                    for row in local_due
                )
                jobs.extend(
                    _auto_poll_local_order(application, int(row["id"]))
                    for row in local
                )
            if jobs:
                await asyncio.gather(*jobs, return_exceptions=True)
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Shared OTP auto-poller cycle failed")
        await asyncio.sleep(OTP_AUTO_POLL_SECONDS)


def _next_dhaka_midnight(now: Optional[datetime] = None) -> datetime:
    from datetime import timedelta
    from zoneinfo import ZoneInfo

    zone = ZoneInfo("Asia/Dhaka")
    current = datetime.now(zone) if now is None else now.astimezone(zone)
    return current.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)


async def daily_database_backup_loop(application: Application) -> None:
    """Send one complete ZIP database backup to configured admins at Dhaka midnight."""
    from zoneinfo import ZoneInfo

    zone = ZoneInfo("Asia/Dhaka")
    while True:
        now = datetime.now(zone)
        next_midnight = _next_dhaka_midnight(now)
        await asyncio.sleep(max(0.1, (next_midnight - now).total_seconds()))

        if not ADMIN_IDS:
            logger.error("Daily database backup skipped: ADMIN_IDS is empty.")
            continue

        backup_date = datetime.now(zone).date()
        backup_path = ""
        try:
            # Keep SQLite snapshot/ZIP work off the Telegram update loop.
            backup_path = await run_file_job(create_database_backup_zip)
            filename = f"paid_number_bot_backup_{backup_date.isoformat()}.zip"
            caption = (
                "🗄️ Daily full database backup\n"
                f"Database date: {backup_date.strftime('%d %B %Y')} (Bangladesh time)"
            )
            sent = 0
            for admin_id in sorted(ADMIN_IDS):
                try:
                    with open(backup_path, "rb") as backup_file:
                        await application.bot.send_document(
                            chat_id=admin_id,
                            document=backup_file,
                            filename=filename,
                            caption=caption,
                        )
                    sent += 1
                except Exception:
                    logger.exception(
                        "Could not send daily database backup to admin %s",
                        admin_id,
                    )
            logger.info(
                "Daily database backup for %s sent to %s/%s admin(s)",
                backup_date.isoformat(),
                sent,
                len(ADMIN_IDS),
            )
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Daily database backup failed for %s", backup_date)
        finally:
            if backup_path:
                with suppress(FileNotFoundError):
                    os.remove(backup_path)


async def persistent_database_backup_loop() -> None:
    """Keep regular, validated SQLite snapshots on the server."""
    while True:
        try:
            backup_path = await run_file_job(create_persistent_database_export)
            summary = await run_file_job(database_validate_file, backup_path)
            logger.info(
                "Persistent database snapshot saved: users=%s, activity=%s, file=%s",
                summary["users"],
                _database_activity_count(summary),
                os.path.basename(backup_path),
            )
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Persistent database snapshot failed")
        await asyncio.sleep(DATABASE_AUTO_BACKUP_INTERVAL_SECONDS)


def start_daily_database_backup_task(application: Application) -> None:
    global DAILY_BACKUP_TASK
    if DAILY_BACKUP_TASK and not DAILY_BACKUP_TASK.done():
        return
    DAILY_BACKUP_TASK = asyncio.create_task(
        daily_database_backup_loop(application),
        name="daily-dhaka-database-backup",
    )


def start_persistent_database_backup_task() -> None:
    global PERSISTENT_DATABASE_BACKUP_TASK
    if (
        PERSISTENT_DATABASE_BACKUP_TASK
        and not PERSISTENT_DATABASE_BACKUP_TASK.done()
    ):
        return
    PERSISTENT_DATABASE_BACKUP_TASK = asyncio.create_task(
        persistent_database_backup_loop(),
        name="persistent-database-backup",
    )


def database_restart_background_tasks(application: Application) -> None:
    global POLL_TASK, DAILY_REPORT_TASK
    if POLL_TASK and not POLL_TASK.done():
        POLL_TASK.cancel()
    POLL_TASK = asyncio.create_task(
        polling_loop(application), name="shared-otp-auto-poller"
    )
    start_daily_database_backup_task(application)
    start_persistent_database_backup_task()
    DAILY_REPORT_TASK = None
    LOCAL_ORDER_POLL_TASKS.clear()
    LOCAL_DETAIL_TASKS.clear()
    LOCAL_ACTIVE_TASKS.clear()
    LOCAL_ORDER_VIEWS.clear()


async def post_init(application: Application) -> None:
    global API, LOCAL_API, POLL_TASK, DAILY_REPORT_TASK
    sms_runtime_key = DB.get_setting("panel_smsbower_api_key", SMS_API_KEY)
    local_runtime_key = DB.get_setting("panel_localusa_api_key", LOCAL_API_KEY)
    API = SmsPowerAPI(SMS_API_URL, sms_runtime_key)
    local_runtime_key = (local_runtime_key or "").strip()
    if local_runtime_key and local_runtime_key.upper() != "XXX" and not local_runtime_key.startswith("PASTE_"):
        LOCAL_API = LocalUsaAPI(LOCAL_API_URL, local_runtime_key)
        logger.info("Local USA API configured and enabled.")
    else:
        LOCAL_API = None
        logger.info("Local USA API key not configured; Local USA provider disabled without errors.")
    bot_info = await application.bot.get_me()
    if bot_info.username:
        DB.set_setting("bot_username", bot_info.username)
    await application.bot.set_my_commands([
        BotCommand("start", "Open the bot"),
        BotCommand("id", "Show your Telegram ID"),
        BotCommand("cancel", "Cancel current input"),
    ])
    if DATABASE_STARTUP_NOTICE:
        logger.warning("%s", DATABASE_STARTUP_NOTICE)
        for admin_id in sorted(ADMIN_IDS):
            try:
                await application.bot.send_message(
                    chat_id=admin_id,
                    text=DATABASE_STARTUP_NOTICE,
                )
            except Exception:
                logger.exception(
                    "Could not send database recovery notice to admin %s",
                    admin_id,
                )
    _run_daily_cleanup_if_needed()
    for invoice in DB.fetchall(
        "SELECT invoice_id,user_id,chat_id,amount_units FROM payment_invoices WHERE status='awaiting_payment'"
    ):
        invoice_id = str(invoice["invoice_id"])
        existing_task = _IPSHOP_PAYMENT_TASKS.get(invoice_id)
        if existing_task and not existing_task.done():
            continue
        _IPSHOP_PAYMENT_TASKS[invoice_id] = asyncio.create_task(
            monitor_ipshop_payment(
                application,
                invoice_id=invoice_id,
                user_id=int(invoice["user_id"]),
                chat_id=int(invoice["chat_id"]),
                amount_units=int(invoice["amount_units"]),
            ),
            name=f"ipshop-payment-{invoice_id}",
        )
    POLL_TASK = asyncio.create_task(
        polling_loop(application), name="shared-otp-auto-poller"
    )
    start_daily_database_backup_task(application)
    start_persistent_database_backup_task()
    DAILY_REPORT_TASK = None
    LOCAL_ORDER_POLL_TASKS.clear()
    LOCAL_DETAIL_TASKS.clear()
    LOCAL_ACTIVE_TASKS.clear()
    LOCAL_ORDER_VIEWS.clear()
    logger.info(
        "Bot initialized: %s-second auto OTP checks, popup manual checks, no panel edits",
        OTP_AUTO_POLL_SECONDS,
    )


# ------------------------ daily TXT reports ---------------------------------


def create_payment_list_txt() -> str:
    _run_daily_cleanup_if_needed()
    start_ts, end_ts, _ = _dhaka_day_bounds()
    rows = DB.fetchall(
        """
        SELECT u.user_id,u.username,u.first_name,
               COUNT(d.id) AS approved_count,
               COALESCE(SUM(d.amount_units),0) AS approved_units
        FROM deposits d JOIN users u ON u.user_id=d.user_id
        WHERE d.status='approved' AND d.created_at>=? AND d.created_at<?
        GROUP BY u.user_id,u.username,u.first_name
        ORDER BY approved_units DESC,u.user_id ASC
        """,
        (start_ts, end_ts),
    )
    path = _new_txt_path("payment_list_today")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("TODAY APPROVED PAYMENTS — ASIA/DHAKA\n" + "=" * 48 + "\n")
        if not rows:
            handle.write("No approved deposits today.\n")
        for index, row in enumerate(rows, 1):
            username = f"@{row['username']}" if row["username"] else "-"
            handle.write(
                f"{index}. ID {row['user_id']} | {username} | "
                f"{money(int(row['approved_units'] or 0))} | "
                f"Requests {int(row['approved_count'] or 0)}\n"
            )
    return path


def create_order_logs_txt() -> str:
    _run_daily_cleanup_if_needed()
    start_ts, end_ts, _ = _dhaka_day_bounds()
    global_rows = DB.fetchall(
        """
        SELECT o.id,o.created_at,o.user_id,u.username,o.phone,o.charged_units,o.status
        FROM orders o LEFT JOIN users u ON u.user_id=o.user_id
        WHERE o.created_at>=? AND o.created_at<? ORDER BY o.id DESC
        """,
        (start_ts, end_ts),
    )
    local_rows = DB.fetchall(
        """
        SELECT lo.id,lo.created_at,lo.user_id,u.username,lo.phone,
               lo.charged_units,lo.status
        FROM local_orders lo LEFT JOIN users u ON u.user_id=lo.user_id
        WHERE lo.created_at>=? AND lo.created_at<? ORDER BY lo.id DESC
        """,
        (start_ts, end_ts),
    )
    path = _new_txt_path("order_logs_today")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("TODAY ORDER LOGS — ASIA/DHAKA\n" + "=" * 64 + "\n")
        for row in global_rows:
            username = f"@{row['username']}" if row["username"] else "-"
            handle.write(
                f"#{row['id']} | {_dhaka_format_timestamp(row['created_at'])} | "
                f"{row['user_id']} {username} | {row['phone']} | "
                f"{money(int(row['charged_units'] or 0))} | {row['status']}\n"
            )
        for row in local_rows:
            username = f"@{row['username']}" if row["username"] else "-"
            phone = str(row["phone"] or "-")
            if not phone.startswith("+"):
                phone = "+" + phone
            handle.write(
                f"L#{row['id']} | {_dhaka_format_timestamp(row['created_at'])} | "
                f"{row['user_id']} {username} | {phone} | "
                f"{money(int(row['charged_units'] or 0))} | {row['status']}\n"
            )
        if not global_rows and not local_rows:
            handle.write("No orders today.\n")
    return path


# ----------------------- balance TXT backup/restore -------------------------


def admin_database_settings_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("⬇️ Download DB"), KeyboardButton("⬆️ Upload DB")],
            [KeyboardButton("⬇️ Download Balance"), KeyboardButton("⬆️ Upload Balance")],
            [KeyboardButton("📄 Download Error Log")],
            [KeyboardButton("🔙 Admin Panel")],
        ],
        resize_keyboard=True,
    )


def create_balance_backup_txt() -> str:
    # Read all balances from one consistent point while writes are paused by DB.lock.
    with DB.lock:
        rows = DB.conn.execute(
            "SELECT user_id,username,first_name,balance_units FROM users ORDER BY user_id"
        ).fetchall()
    path = _new_txt_path("balance_backup")
    total = 0
    with open(path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write("# BALANCE BACKUP\n")
        handle.write("# Format: user_id<TAB>balance_bdt<TAB>username<TAB>first_name\n")
        handle.write("# Upload sets the exact current balance; it does not add twice.\n")
        for row in rows:
            units = int(row["balance_units"] or 0)
            total += units
            amount = f"{units_to_decimal(units):.6f}".rstrip("0").rstrip(".") or "0"
            username = str(row["username"] or "").replace("\t", " ").replace("\n", " ")
            first_name = str(row["first_name"] or "").replace("\t", " ").replace("\n", " ")
            handle.write(f"{int(row['user_id'])}\t{amount}\t{username}\t{first_name}\n")
        handle.write(f"# Users: {len(rows)}\n")
        handle.write(f"# Total balance: {money(total)}\n")
    return path


async def admin_error_log_download(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Send the current and rotated error logs to an authorized admin."""
    if not await require_admin(update):
        return
    message = update.effective_message
    log_path = _ERROR_LOG_PATH
    try:
        if not os.path.isfile(log_path):
            await message.reply_text(
                "ℹ️ No error log has been created yet. The bot will create error.log automatically when an error is logged.",
                reply_markup=admin_database_settings_menu(),
            )
            return
        fd, bundle_path = tempfile.mkstemp(prefix="streamnova_error_logs_", suffix=".zip")
        os.close(fd)
        try:
            with zipfile.ZipFile(bundle_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                directory = os.path.dirname(log_path) or "."
                for name in sorted(os.listdir(directory)):
                    if name == "error.log" or re.fullmatch(r"error\.log\.\d+", name):
                        candidate = os.path.join(directory, name)
                        if os.path.isfile(candidate):
                            archive.write(candidate, arcname=name)
            with open(bundle_path, "rb") as handle:
                await context.bot.send_document(
                    chat_id=message.chat_id,
                    document=handle,
                    filename="streamnova_error_logs.zip",
                    caption="Bot error.log and available rotated error logs. Keep this file private; logs may contain operational details.",
                )
        finally:
            with suppress(FileNotFoundError):
                os.remove(bundle_path)
    except Exception as exc:
        logger.exception("Error log download failed")
        await message.reply_text(
            f"❌ Error log download failed: {short_text(str(exc), 500)}",
            reply_markup=admin_database_settings_menu(),
        )


async def admin_users_txt_download(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if not await require_admin(update):
        return
    path = ""
    try:
        path = await run_file_job(create_balance_backup_txt)
        with open(path, "rb") as handle:
            await context.bot.send_document(
                chat_id=update.effective_chat.id,
                document=handle,
                filename="balance_backup.txt",
                caption="Current user balances. Uploading this file restores exact balances.",
            )
    except Exception as exc:
        logger.exception("Balance TXT export failed")
        await update.effective_message.reply_text(
            f"❌ Balance export failed: {short_text(str(exc), 500)}",
            reply_markup=admin_database_settings_menu(),
        )


async def users_txt_import_start(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    begin_workflow(context, admin_mode=True)
    if not await require_admin(update):
        return ConversationHandler.END
    await update.effective_message.reply_text(
        "Upload balance_backup.txt.\n"
        "Each balance will be restored as an exact value. Missing users will be created.",
        reply_markup=workflow_keyboard(update.effective_user.id, admin=True),
    )
    return DB_USER_IMPORT_FILE


def import_balance_backup_txt(path: str) -> dict[str, int]:
    parsed: dict[int, tuple[int, str, str]] = {}
    invalid = 0
    with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
        for raw in handle:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if len(parts) < 2:
                parts = re.split(r"[,|;]", line, maxsplit=3)
            try:
                user_id = int(parts[0].strip())
                amount_text = parts[1].strip().replace(",", "")
                amount_text = re.sub(r"[^0-9.+-]", "", amount_text)
                units = bdt_to_units(amount_text)
                if user_id <= 0 or units < 0:
                    raise ValueError
                username = parts[2].strip().lstrip("@")[:64] if len(parts) > 2 else ""
                first_name = parts[3].strip()[:128] if len(parts) > 3 else ""
                parsed[user_id] = (units, username, first_name)
            except Exception:
                invalid += 1

    created = updated = 0
    ts = now_ts()
    with DB.lock:
        DB.conn.execute("BEGIN IMMEDIATE")
        try:
            for user_id, (units, username, first_name) in parsed.items():
                old = DB.conn.execute(
                    "SELECT 1 FROM users WHERE user_id=?", (user_id,)
                ).fetchone()
                if old:
                    DB.conn.execute(
                        """
                        UPDATE users SET balance_units=?,
                            username=CASE WHEN ?<>'' THEN ? ELSE username END,
                            first_name=CASE WHEN ?<>'' THEN ? ELSE first_name END,
                            updated_at=? WHERE user_id=?
                        """,
                        (units, username, username, first_name, first_name, ts, user_id),
                    )
                    updated += 1
                else:
                    DB.conn.execute(
                        """
                        INSERT INTO users(
                            user_id,username,first_name,balance_units,accepted_terms,
                            is_blocked,created_at,updated_at,restricted_until,warning_count
                        ) VALUES(?,?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            user_id, username or None, first_name or None, units,
                            0, 0, ts, ts, 0, 0,
                        ),
                    )
                    created += 1
                DB.conn.execute(
                    """
                    INSERT INTO transactions(
                        user_id,amount_units,kind,reference,note,created_at
                    ) VALUES(?,?,?,?,?,?)
                    """,
                    (user_id, 0, "balance_restore", "balance_backup", "Exact balance restored", ts),
                )
            DB.conn.commit()
        except Exception:
            DB.conn.rollback()
            raise
    return {"created": created, "updated": updated, "invalid": invalid, "rows": len(parsed)}


async def users_txt_import_receive(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    if not await require_admin(update):
        return ConversationHandler.END
    document = update.effective_message.document
    if not document:
        await update.effective_message.reply_text("Send balance_backup.txt.")
        return DB_USER_IMPORT_FILE
    if document.file_size and document.file_size > 10 * 1024 * 1024:
        await update.effective_message.reply_text("❌ TXT file is too large.")
        return DB_USER_IMPORT_FILE
    if not (document.file_name or "").lower().endswith(".txt"):
        await update.effective_message.reply_text("❌ Send a .txt file.")
        return DB_USER_IMPORT_FILE

    fd, path = tempfile.mkstemp(prefix="balance_import_", suffix=".txt")
    os.close(fd)
    try:
        telegram_file = await document.get_file()
        await telegram_file.download_to_drive(custom_path=path)
        # Always preserve the live balances before applying an uploaded file.
        os.makedirs(BACKUP_DIRECTORY, exist_ok=True)
        safety_copy = await run_file_job(create_balance_backup_txt)
        safety_name = f"balance_pre_upload_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        safety_path = os.path.join(BACKUP_DIRECTORY, safety_name)
        await run_file_job(shutil.copy2, safety_copy, safety_path)
        with suppress(FileNotFoundError):
            os.remove(safety_copy)
        result = await run_file_job(import_balance_backup_txt, path)
        await update.effective_message.reply_text(
            "✅ Balance restore completed.\n"
            f"Rows: <b>{result['rows']}</b>\n"
            f"Updated: <b>{result['updated']}</b>\n"
            f"Created: <b>{result['created']}</b>\n"
            f"Invalid: <b>{result['invalid']}</b>\n\n"
            f"🛟 Previous balances backed up: <code>{html.escape(safety_name)}</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=admin_database_settings_menu(),
        )
    except Exception as exc:
        logger.exception("Balance TXT import failed")
        await update.effective_message.reply_text(
            f"❌ Balance import failed: {short_text(str(exc), 600)}"
        )
        return DB_USER_IMPORT_FILE
    finally:
        with suppress(FileNotFoundError):
            os.remove(path)
    return ConversationHandler.END


# Refresh action maps after redefining handlers.
ADMIN_ONE_SHOT_ACTIONS["📊 Statistics"] = admin_stats_reply
ADMIN_ONE_SHOT_ACTIONS["🧾 Order Logs"] = admin_order_logs
ADMIN_ONE_SHOT_ACTIONS["📋 Payment List"] = admin_payment_list
ADMIN_ONE_SHOT_ACTIONS["⬇️ Download DB"] = admin_database_download
ADMIN_ONE_SHOT_ACTIONS.pop("⬇️ Download User TXT", None)
ADMIN_ONE_SHOT_ACTIONS["⬇️ Download Balance"] = admin_users_txt_download
ADMIN_ONE_SHOT_ACTIONS["📄 Download Error Log"] = admin_error_log_download
ONE_SHOT_MENU_PATTERN = r"^(?:" + "|".join(
    re.escape(label)
    for label in [
        *ADMIN_ONE_SHOT_ACTIONS.keys(),
        *USER_ONE_SHOT_ACTIONS.keys(),
        "🔙 Main Menu",
    ]
) + r")$"

def build_application() -> Application:
    # AIORateLimiter requires PTB's optional ``rate-limiter`` dependency.
    # Build it defensively so the bot can still start when ``aiolimiter`` is
    # missing or an older/newer PTB version exposes a different signature.
    limiter = None
    limiter_kwargs = {
        "overall_max_rate": TELEGRAM_OVERALL_RATE,
        "overall_time_period": 1,
        "group_max_rate": TELEGRAM_GROUP_RATE,
        "group_time_period": 60,
        "max_retries": 0,
    }
    try:
        supported = inspect.signature(AIORateLimiter).parameters
        limiter = AIORateLimiter(
            **{key: value for key, value in limiter_kwargs.items() if key in supported}
        )
    except (ImportError, RuntimeError, TypeError) as exc:
        logger.warning(
            "Telegram rate limiter disabled: %s. "
            "Install it with: python -m pip install -U "
            "\"python-telegram-bot[rate-limiter]>=22.8,<23\"",
            exc,
        )

    builder = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .post_shutdown(post_shutdown)
    )
    if limiter is not None:
        builder = builder.rate_limiter(limiter)

    app = (
        builder
        .concurrent_updates(BOT_MAX_CONCURRENT_UPDATES)
        .connection_pool_size(BOT_CONNECTION_POOL_SIZE)
        .pool_timeout(BOT_POOL_TIMEOUT_SECONDS)
        .connect_timeout(15.0)
        .read_timeout(30.0)
        .write_timeout(30.0)
        .media_write_timeout(BOT_MEDIA_WRITE_TIMEOUT_SECONDS)
        .build()
    )

    # A single ConversationHandler owns every multi-step form. With
    # allow_reentry=True keeps command/callback recovery safe. While a form is
    # active, the reply keyboard intentionally shows only the Cancel button.
    workflow_conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex(r"^❌ Cancel$"), cancel_conversation),
            CallbackQueryHandler(cancel_conversation, pattern=r"^workflow:cancel$"),
            CommandHandler("cancel", cancel_conversation),
            CommandHandler("start", workflow_start_command),
            CommandHandler("admin", workflow_admin_command),
            CommandHandler("id", workflow_id_command),
            CallbackQueryHandler(add_catalog_start, pattern=r"^admin:addcatalog$"),
            CallbackQueryHandler(
                country_search_start,
                pattern=r"^buy:countrysearch:\d+:\d+$",
            ),

            # Stateful user/admin workflows. Put these before the generic menu
            # entry point so their returned state becomes the new active state.
            MessageHandler(filters.Regex(r"^💳 Add Balance$"), deposit_start),
            MessageHandler(filters.Regex(r"^📣 Broadcast$"), broadcast_start),
            MessageHandler(
                filters.Regex(r"^(➕ Add Service|➕ Add Country/Service)$"),
                add_catalog_start_reply,
            ),
            MessageHandler(filters.Regex(r"^⚙️ Profit Markup$"), set_markup_start_reply),
            MessageHandler(filters.Regex(r"^🛡 Spam Control$"), spam_control_start),
            MessageHandler(filters.Regex(r"^💱 Exchange Rate$"), set_exchange_rate_start),
            MessageHandler(
                filters.Regex(r"^📥 Deposit Group ID$"), set_deposit_group_start
            ),
            MessageHandler(
                filters.Regex(r"^🆘 Support ID/Link$"), set_support_link_start
            ),
            MessageHandler(
                filters.Regex(r"^💵 Referral Reward$"), set_referral_reward_start
            ),
            MessageHandler(
                filters.Regex(r"^📏 Referral Deposit Limit$"),
                set_referral_threshold_start,
            ),
            MessageHandler(
                filters.Regex(r"^➕ Set Button$"), set_how_button_start
            ),
            MessageHandler(
                filters.Regex(r"^🎬 Set Video$"), set_how_video_start
            ),
            MessageHandler(
                filters.Regex(r"^💵 Minimum Deposit$"), set_minimum_deposit_start
            ),
            MessageHandler(
                filters.Regex(r"^💵 Local Price$"), set_local_price_start
            ),
            MessageHandler(
                filters.Regex(r"^⏱ Local Session$"), set_local_session_start
            ),
            MessageHandler(filters.Regex(r"^🛡️ SET 𝚅𝙿𝙽/𝙿𝚁𝙾𝚇𝙸𝙴𝚂$"), vshop_admin_settings),
            MessageHandler(filters.Regex(r"^(?:Add Inline|➕ Add Inline)$"), vshop_add_category_start),
            MessageHandler(filters.Regex(r"^(?:Set Category|💰 Set Category)$"), vshop_set_category),
            MessageHandler(filters.Regex(r"^🛠 Manage Inline$"), vshop_manage_inline),
            MessageHandler(filters.Regex(r"^📜 VPN Order History$"), vshop_order_history_start),
            MessageHandler(filters.Regex(r"^🔔 VPN Notice Group$"), vshop_set_notice_group_start),
            CallbackQueryHandler(vshop_provide_start, pattern=r"^vshop:provide:\d+$"),
            CallbackQueryHandler(vshop_provide_confirm_callback, pattern=r"^vshop:doconfirmprovide:\d+$"),
            CallbackQueryHandler(vshop_set_duration_start, pattern=r"^vshop:setdur:\d+$"),
            MessageHandler(
                filters.Regex(r"^⬆️ Upload DB$"), database_restore_start
            ),
            MessageHandler(
                filters.Regex(r"^⬆️ Upload Balance$"), users_txt_import_start
            ),
            MessageHandler(filters.Regex(r"^👤 User Manager$"), admin_user_lookup_start),
            MessageHandler(
                filters.Regex(r"^🔐 OTP Group 2$"),
                set_otp_group_start,
            ),
            # Inline admin controls for referral and How to use settings.
            # These start the same ConversationHandler states as the legacy
            # reply-keyboard labels, without relying on a generic menu route.
            CallbackQueryHandler(
                set_referral_reward_start,
                pattern=r"^admin:referral:reward$",
            ),
            CallbackQueryHandler(
                set_referral_threshold_start,
                pattern=r"^admin:referral:threshold$",
            ),
            CallbackQueryHandler(
                set_how_button_start,
                pattern=r"^howadmin:add$",
            ),
            CallbackQueryHandler(
                set_how_video_start,
                pattern=r"^howadmin:video_menu$",
            ),
            CallbackQueryHandler(
                show_admin_general_settings,
                pattern=r"^admin:general$",
            ),
            CallbackQueryHandler(admin_user_lookup_start, pattern=r"^user:lookup$"),
            CallbackQueryHandler(
                admin_user_balance_action_start, pattern=r"^userbal:(add|sub):\d+$"
            ),
            CallbackQueryHandler(
                admin_user_restrict_start, pattern=r"^userrestrict:\d+$"
            ),
            CallbackQueryHandler(
                admin_select_how_video, pattern=r"^howadmin:video:\d+$"
            ),
            CallbackQueryHandler(vshop_enter_quantity, pattern=r"^vshop:enter:\d+$"),
            CallbackQueryHandler(vshop_set_quantity_limits_start, pattern=r"^vshop:setqty:\d+$"),
            CallbackQueryHandler(vshop_set_price_start, pattern=r"^vshop:setprice:\d+$"),
            CallbackQueryHandler(vshop_upload_start, pattern=r"^vshop:upload:\d+$"),
            CallbackQueryHandler(vshop_add_service_start, pattern=r"^vshop:addsvc:\d+$"),

            # One-step buttons also interrupt any old form, but finish
            # immediately after executing their action.
            MessageHandler(filters.Regex(ONE_SHOT_MENU_PATTERN), workflow_one_shot_menu),
        ],
        states={
            DEP_AMOUNT_NEW: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, deposit_amount)
            ],
            CAT_COUNTRY_ID: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_catalog_country_id)
            ],
            CAT_COUNTRY_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_catalog_country_name)
            ],
            CAT_SERVICE_CODE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_catalog_service_code)
            ],
            CAT_SERVICE_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_catalog_service_name)
            ],
            CAT_MARKUP: [
                CallbackQueryHandler(add_catalog_markup_default, pattern=r"^cat:markup:default$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_catalog_markup),
            ],
            BUY_COUNTRY_SEARCH: [
                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    country_search_receive,
                )
            ],
            SET_MARKUP: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_markup_value)
            ],
            SET_EXCHANGE_RATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_exchange_rate_value)
            ],
            SET_DEPOSIT_GROUP: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_deposit_group_value)
            ],
            SET_SUPPORT_LINK: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_support_link_value)
            ],
            SET_REFERRAL_REWARD: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_referral_reward_value)
            ],
            SET_REFERRAL_THRESHOLD: [
                MessageHandler(
                    filters.TEXT & ~filters.COMMAND, set_referral_threshold_value
                )
            ],
            SET_HOW_BUTTON_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_how_button_name)
            ],
            SET_HOW_VIDEO_FILE: [
                MessageHandler(filters.VIDEO, save_how_video)
            ],
            SET_MINIMUM_DEPOSIT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_minimum_deposit_value)
            ],
            ADMIN_USER_LOOKUP_ID: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_user_lookup_value)
            ],
            ADMIN_USER_BALANCE_AMOUNT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_user_balance_amount)
            ],
            SET_OTP_FORWARD_GROUP: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_otp_group_value)
            ],
            SET_LOCAL_PRICE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_local_price_value)
            ],
            SET_LOCAL_SESSION: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, set_local_session_value)
            ],
            SPAM_INTERVAL_VALUE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, spam_control_interval_value)
            ],
            SPAM_LIMIT_VALUE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, spam_control_limit_value)
            ],
            SPAM_BLOCK_VALUE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, spam_control_block_value)
            ],
            VSHOP_ADD_CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_add_category_value)],
            VSHOP_ADD_SERVICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_add_service_value)],
            VSHOP_SET_PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_set_price_value)],
            VSHOP_SET_QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_set_quantity_limits_value)],
            VSHOP_UPLOAD_FILE: [MessageHandler(filters.Document.ALL, vshop_upload_file)],
            VSHOP_QUANTITY: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_quantity_value)],
            VSHOP_RENAME_CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_category_rename_value)],
            VSHOP_RENAME_SERVICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_service_rename_value)],
            VSHOP_SET_DURATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_set_duration_value)],
            VSHOP_SET_NOTICE_GROUP: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_set_notice_group_value)],
            VSHOP_ORDER_LOOKUP: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_order_history_lookup)],
            VSHOP_PROVIDE_CONTENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, vshop_provide_content)],
            VSHOP_PROVIDE_CONFIRM: [
                CallbackQueryHandler(vshop_provide_confirm_callback, pattern=r"^vshop:doconfirmprovide:\d+$"),
                CallbackQueryHandler(vshop_provide_start, pattern=r"^vshop:provide:\d+$"),
            ],
            DB_RESTORE_FILE: [
                MessageHandler(filters.Document.ALL, database_restore_receive),
                MessageHandler(
                    filters.ALL & ~filters.Document.ALL & ~filters.COMMAND,
                    database_restore_receive,
                ),
            ],
            DB_RESTORE_CONFIRM: [
                CallbackQueryHandler(
                    database_restore_confirm,
                    pattern=r"^dbrestore:confirm$",
                ),
                CallbackQueryHandler(
                    database_restore_cancel,
                    pattern=r"^dbrestore:cancel$",
                ),
            ],
            DB_USER_IMPORT_FILE: [
                MessageHandler(filters.Document.ALL, users_txt_import_receive),
                MessageHandler(
                    filters.ALL & ~filters.Document.ALL & ~filters.COMMAND,
                    users_txt_import_receive,
                ),
            ],
            BROADCAST_CONTENT: [
                MessageHandler(filters.ALL & ~filters.COMMAND & ~filters.Regex(r"^❌ Cancel$"), broadcast_receive_content)
            ],
            BROADCAST_CONFIRM: [
                CallbackQueryHandler(broadcast_confirm_callback, pattern=r"^broadcast:confirm$"),
                CallbackQueryHandler(broadcast_cancel_callback, pattern=r"^broadcast:cancel$"),
            ],
            ADMIN_USER_RESTRICT_DURATION: [
                MessageHandler(filters.TEXT & ~filters.COMMAND & ~filters.Regex(r"^❌ Cancel$"), admin_user_restrict_duration)
            ],
        },
        fallbacks=[
            MessageHandler(filters.Regex(r"^❌ Cancel$"), cancel_conversation),
            CallbackQueryHandler(cancel_conversation, pattern=r"^workflow:cancel$"),
            CommandHandler("cancel", cancel_conversation),
        ],
        allow_reentry=True,
    )

    # VShop destructive actions MUST be registered before the main
    # ConversationHandler.  If an admin is inside any workflow state, the
    # ConversationHandler can otherwise consume the callback first and the
    # Remove buttons appear completely unresponsive.
    app.add_handler(CallbackQueryHandler(vshop_remove_service, pattern=r"^vshop:remove:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_remove_service, pattern=r"^vshop:setremove:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_category_delete, pattern=r"^vshop:cdelete:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_remove_all_services, pattern=r"^vshop:removeall:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_remove_all_services_confirm, pattern=r"^vshop:confirmremoveall:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_remove_category_confirm, pattern=r"^vshop:reminline:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_remove_category_confirmed, pattern=r"^vshop:confirmreminline:\d+$"))

    # Panel Management controls and entry button are registered before workflows
    # so the menu remains accessible even if an admin was previously inside a form.
    app.add_handler(CallbackQueryHandler(panel_management_callback, pattern=r"^panel:(?:home|back|provider:(?:smsbower|localusa)|(?:add|replace|remove):(?:smsbower|localusa))$"))
    app.add_handler(MessageHandler(filters.Regex(r"^🧩 Panel Management$"), show_panel_management))

    app.add_handler(workflow_conv)

    # Commands that are not workflow/navigation commands.
    app.add_handler(CommandHandler("addbalance", add_balance_command))
    app.add_handler(CommandHandler("broadcast", broadcast_command))

    # Regular inline actions.
    app.add_handler(CallbackQueryHandler(accept_terms, pattern=r"^terms:accept$"))
    app.add_handler(CallbackQueryHandler(show_referral, pattern=r"^referral:show$"))
    app.add_handler(
        CallbackQueryHandler(show_my_purchase_history, pattern=r"^purchasehistory:mine$")
    )
    app.add_handler(
        CallbackQueryHandler(
            show_my_purchase_history_back, pattern=r"^purchasehistory:back$"
        )
    )
    app.add_handler(
        CallbackQueryHandler(admin_user_purchase_history, pattern=r"^umhistory:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(admin_user_profile_back, pattern=r"^umprofile:\d+$")
    )
    app.add_handler(CallbackQueryHandler(show_how_to_use_video, pattern=r"^howuse:view:\d+$"))
    app.add_handler(
        CallbackQueryHandler(delete_how_button, pattern=r"^howadmin:delete:\d+$")
    )
    app.add_handler(CallbackQueryHandler(mail_service_callback, pattern=r"^mail:svc:\d+$"))
    app.add_handler(CallbackQueryHandler(mail_buy_callback, pattern=r"^mail:buy:\d+:.+$"))
    app.add_handler(CallbackQueryHandler(mail_orders_callback, pattern=r"^mail:orders$"))
    app.add_handler(CallbackQueryHandler(mail_view_callback, pattern=r"^mail:view:\d+$"))
    app.add_handler(CallbackQueryHandler(mail_order_action, pattern=r"^mail:(?:copy|code|again|cancel):\d+$"))
    app.add_handler(CallbackQueryHandler(mail_home_callback, pattern=r"^mail:home$"))
    app.add_handler(CallbackQueryHandler(mail_admin_open_callback, pattern=r"^mail:admin:open$"))
    app.add_handler(CallbackQueryHandler(mail_admin_callback, pattern=r"^mail:admin:toggle:\d+$"))
    app.add_handler(CallbackQueryHandler(mail_admin_detail_callback, pattern=r"^mail:admin:detail:\d+$"))
    app.add_handler(CallbackQueryHandler(mail_admin_edit_start, pattern=r"^mail:admin:edit:\d+$"))
    app.add_handler(CallbackQueryHandler(mail_admin_delete_callback, pattern=r"^mail:admin:delete:\d+$"))
    app.add_handler(CallbackQueryHandler(mail_admin_delete_confirm_callback, pattern=r"^mail:admin:deleteconfirm:\d+$"))
    app.add_handler(CallbackQueryHandler(mail_admin_back_list_callback, pattern=r"^mail:admin:list$"))
    app.add_handler(CallbackQueryHandler(lambda u,c: mail_admin_add_start(u,c), pattern=r"^mail:admin:add$"))
    app.add_handler(CallbackQueryHandler(mail_admin_cancel_callback, pattern=r"^mail:admin:cancel$"))
    app.add_handler(MessageHandler(filters.Regex(r"^📧 Mail Services$"), mail_admin_list))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mail_admin_text), group=0)

    app.add_handler(CallbackQueryHandler(premium_emoji_callback, pattern=r"^premiumemoji:"), group=-1)
    app.add_handler(MessageHandler(filters.Document.ALL, premium_emoji_document_receive), group=-1)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, premium_emoji_single_receive), group=-1)

    app.add_handler(CallbackQueryHandler(show_buy_menu, pattern=r"^buy:menu$"))
    app.add_handler(
        CallbackQueryHandler(show_buy_menu, pattern=r"^buy:country:[^:]+$")
    )
    app.add_handler(
        CallbackQueryHandler(show_service_countries, pattern=r"^buy:svc:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(show_service_countries, pattern=r"^buy:svcpage:\d+:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(show_provider_prices, pattern=r"^buy:service:\d+(?::\d+)?$")
    )
    app.add_handler(
        CallbackQueryHandler(toggle_country_favorite, pattern=r"^buy:favorite:\d+$")
    )
    app.add_handler(CallbackQueryHandler(noop_callback, pattern=r"^noop$"))
    app.add_handler(
        CallbackQueryHandler(view_provider_offer, pattern=r"^buy:provider:\d+:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(confirm_purchase, pattern=r"^buy:confirm:\d+:\d+$")
    )
    app.add_handler(CallbackQueryHandler(active_number_list_callback, pattern=r"^order:active:list$"))
    app.add_handler(CallbackQueryHandler(active_number_view_callback, pattern=r"^order:active:view:\d+$"))
    app.add_handler(CallbackQueryHandler(cc_toggle_callback, pattern=r"^order:cc:\d+$"))
    app.add_handler(CallbackQueryHandler(another_message_callback, pattern=r"^order:another:\d+$"))
    app.add_handler(CallbackQueryHandler(check_order, pattern=r"^order:check:\d+$"))
    app.add_handler(CallbackQueryHandler(order_time_refresh_callback, pattern=r"^order:time:\d+$"))
    app.add_handler(CallbackQueryHandler(cancel_order, pattern=r"^order:cancel:\d+$"))
    app.add_handler(CallbackQueryHandler(buy_same_provider_callback, pattern=r"^order:new:\d+$"))
    app.add_handler(CallbackQueryHandler(refund_and_new_callback, pattern=r"^order:refund_new:\d+$"))
    app.add_handler(CallbackQueryHandler(admin_user_warning_callback, pattern=r"^userwarn:\d+$"))
    app.add_handler(CallbackQueryHandler(admin_user_unrestrict_callback, pattern=r"^userunrestrict:\d+$"))

    # USA local-number inline flow.
    app.add_handler(CallbackQueryHandler(local_buy_confirm, pattern=r"^local:buy$"))
    app.add_handler(CallbackQueryHandler(local_buy_abort, pattern=r"^local:abort$"))
    app.add_handler(
        CallbackQueryHandler(local_active_page_callback, pattern=r"^local:page:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(local_detail_callback, pattern=r"^local:detail:\d+:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(local_refresh_callback, pattern=r"^local:refresh:\d+:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(local_time_refresh_callback, pattern=r"^local:time:\d+:\d+$")
    )
    app.add_handler(
        CallbackQueryHandler(local_cancel_callback, pattern=r"^local:cancel:\d+:\d+$")
    )

    app.add_handler(CallbackQueryHandler(vshop_manage_inline_toggle, pattern=r"^vshop:manage_toggle$"))
    app.add_handler(CallbackQueryHandler(vshop_clear_stock, pattern=r"^vshop:clearstock:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_category_manage, pattern=r"^vshop:managecat:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_set_category_callback, pattern=r"^vshop:setcat:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_set_service_admin, pattern=r"^vshop:setsvc:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_set_category_back, pattern=r"^vshop:setcatback:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_service_automation_toggle, pattern=r"^vshop:sauto:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_admin_order_detail, pattern=r"^vshop:adminorder:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_cancel_order_callback, pattern=r"^vshop:cancelorder:\d+$"))
    async def _vshop_setcats_nav(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        q = update.callback_query
        await q.answer()
        await replace_callback_message(q, context.bot, "💰 <b>Set Category</b>\n\nSelect a category:", parse_mode=ParseMode.HTML, reply_markup=vshop_set_category_markup())

    async def _vshop_admincats_nav(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        q = update.callback_query
        await q.answer()
        await replace_callback_message(q, context.bot, "🛠 <b>Manage Inline</b>", parse_mode=ParseMode.HTML, reply_markup=vshop_admin_categories_markup())

    async def _vshop_adminmenu_nav(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        q = update.callback_query
        await q.answer()
        await replace_callback_message(q, context.bot, "🛡️ <b>SET 𝚅𝙿𝙽/𝙿𝚁𝙾𝚇𝙸𝙴𝚂</b>", parse_mode=ParseMode.HTML, reply_markup=vshop_admin_menu())

    app.add_handler(CallbackQueryHandler(_vshop_setcats_nav, pattern=r"^vshop:setcats$"))
    app.add_handler(CallbackQueryHandler(_vshop_admincats_nav, pattern=r"^vshop:admincats$"))
    app.add_handler(CallbackQueryHandler(_vshop_adminmenu_nav, pattern=r"^vshop:adminmenu$"))
    app.add_handler(CallbackQueryHandler(vshop_category_toggle, pattern=r"^vshop:ctoggle:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_category_rename_start, pattern=r"^vshop:crename:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_service_toggle, pattern=r"^vshop:stoggle:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_service_toggle, pattern=r"^vshop:settoggle:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_service_rename_start, pattern=r"^vshop:srename:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_category_callback, pattern=r"^vshop:cat:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_admin_category_callback, pattern=r"^vshop:acat:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_service_admin, pattern=r"^vshop:asvc:\d+$"))
    async def _vshop_category_back_from_service(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        q=update.callback_query
        try:
            sid=int(q.data.split(":")[-1])
            svc=DB.fetchone("SELECT category_id FROM vshop_services WHERE id=?", (sid,))
            if not svc:
                await q.answer("Category not found", show_alert=True); return
            await q.answer()
            await replace_callback_message(q, context.bot, "Choose service:", reply_markup=vshop_service_markup(int(svc["category_id"]), admin=True, back_data=f"vshop:managecat:{int(svc["category_id"])}"))
        except Exception:
            await q.answer("Unable to go back", show_alert=True)
    app.add_handler(CallbackQueryHandler(_vshop_category_back_from_service, pattern=r"^vshop:catback:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_show_service, pattern=r"^vshop:svc:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_qty_callback, pattern=r"^vshop:qty:\d+:\d+$"))
    app.add_handler(CallbackQueryHandler(vshop_confirm_purchase, pattern=r"^vshop:confirm:\d+:\d+$"))
    async def _vshop_back_categories(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        q=update.callback_query
        await q.answer()
        await replace_callback_message(q, context.bot, "Choose a category:", reply_markup=vshop_user_categories_markup())
    app.add_handler(CallbackQueryHandler(_vshop_back_categories, pattern=r"^vshop:backcats$"))
    app.add_handler(CallbackQueryHandler(lambda u,c: u.callback_query.answer(), pattern=r"^vshop:noop$"))
    app.add_handler(CallbackQueryHandler(lambda u,c: u.callback_query.answer("Cancelled"), pattern=r"^vshop:cancel$"))

    app.add_handler(
        CallbackQueryHandler(
            admin_catalog_action,
            pattern=(
                r"^admin:(catmenu|catservice|catserviceid|catcountries|catmargin|catpage|cattoggle|catdelete|"
                r"catservicedelete|catservicedeleteconfirm|catservicedeletecancel)(:|$)"
            ),
        )
    )

    

    # Outside an active form, /cancel should still return a clean menu.
    app.add_handler(CommandHandler("cancel", cancel_conversation))
    app.add_handler(MessageHandler(filters.Regex(r"^❌ Cancel$"), cancel_conversation))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, unknown_text))
    app.add_error_handler(error_handler)
    return app


# =============================================================================
# CREDENTIALS — environment values take priority; XXX marks file-based setup slots.
# =============================================================================

# Secrets are read from the hosting Environment panel first, then from .env.
# There are deliberately no real/default credentials in this source file.
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
SMS_API_KEY = os.getenv("SMS_API_KEY", "").strip()
LOCAL_API_KEY = os.getenv("LOCAL_API_KEY", "").strip()
IPSHOP_API_KEY = os.getenv("IPSHOP_API_KEY", "").strip()




def main() -> None:
    validate_config()
    application = build_application()
    logger.info("Starting bot...")
    application.run_polling(
        allowed_updates=Update.ALL_TYPES,
        drop_pending_updates=False,
    )


if __name__ == "__main__":
    main()