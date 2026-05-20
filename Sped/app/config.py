from __future__ import annotations

import re


COMPARE_KEY_PATTERN = re.compile(r"\b\d{44}\b")
COMPARE_NS_NFE = {"nfe": "http://www.portalfiscal.inf.br/nfe"}
COMPARE_NS_NFSE = {"nfse": "http://www.sped.fazenda.gov.br/nfse"}
COMPARE_MARK_UNCHECKED = "\u25cb"
COMPARE_MARK_CHECKED = "\u25cf"

MYSQL_DEFAULT_CONFIG = {
    "host": "127.0.0.1",
    "port": "3306",
    "user": "root",
    "password": "",
    "database": "sped_icms",
}
APP_DEFAULT_CONFIG = {
    "window_title": "Revisor de SPED - DZ Consultoria",
    "home_title": "Revisor de SPED - DZ Consultoria",
}
MYSQL_CONNECTION_TIMEOUT_SECONDS = 3



