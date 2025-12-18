#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlencode, urlsplit, urlunsplit, parse_qsl
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError


class IikoApiError(RuntimeError):
    pass


@dataclass
class IikoProduct:
    name: str
    price: str = ""
    weight: str = ""
    product_id: str = ""


def _safe_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _redact_url(url: str) -> str:
    """Убирает/маскирует чувствительные параметры (password) в URL для сообщений об ошибках."""
    try:
        parts = urlsplit(url)
        q = []
        for k, v in parse_qsl(parts.query, keep_blank_values=True):
            if k.lower() in ("password", "pwd", "pass"):
                q.append((k, "***"))
            else:
                q.append((k, v))
        new_query = urlencode(q)
        return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))
    except Exception:
        return url


def _http_get_json(url: str, timeout_sec: int = 20) -> Any:
    req = Request(url)
    req.add_header("Accept", "application/json")
    try:
        with urlopen(req, timeout=timeout_sec) as resp:
            raw = resp.read()
            text = raw.decode("utf-8", errors="replace")
    except HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        raise IikoApiError(f"HTTP {e.code} при запросе {_redact_url(url)}. {body[:200]}")
    except URLError as e:
        raise IikoApiError(f"Ошибка соединения с iiko: {e}")

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


def _http_get_text(url: str, timeout_sec: int = 20) -> str:
    req = Request(url)
    try:
        with urlopen(req, timeout=timeout_sec) as resp:
            raw = resp.read()
            return raw.decode("utf-8", errors="replace").strip()
    except HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        raise IikoApiError(f"HTTP {e.code} при запросе {_redact_url(url)}. {body[:200]}")
    except URLError as e:
        raise IikoApiError(f"Ошибка соединения с iiko: {e}")


def _http_post_text(url: str, timeout_sec: int = 20) -> str:
    # В urllib POST делается через передачу data (даже пустой)
    req = Request(url, data=b"")
    try:
        with urlopen(req, timeout=timeout_sec) as resp:
            raw = resp.read()
            return raw.decode("utf-8", errors="replace").strip()
    except HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        raise IikoApiError(f"HTTP {e.code} при запросе {_redact_url(url)}. {body[:200]}")
    except URLError as e:
        raise IikoApiError(f"Ошибка соединения с iiko: {e}")


class IikoRmsClient:
    """Мини-клиент для iikoOffice/iikoChain (iikoRMS) через RESTO API.

    Документация iikoRMS (resto API):
      POST /resto/api/auth?login=[login]&pass=[sha1(password)]

    base_url ожидается вида: https://<host>/resto
    """

    def __init__(self, base_url: str, login: str, pass_sha1: str):
        self.base_url = base_url.rstrip("/")
        self.login = login
        self.pass_sha1 = (pass_sha1 or "").strip().lower()

    def auth_key(self) -> str:
        """Получает auth key (строку-токен)."""
        if not self.login or not self.pass_sha1:
            raise IikoApiError("Не задан login или sha1-хэш пароля (pass).")

        url = f"{self.base_url}/api/auth?{urlencode({'login': self.login, 'pass': self.pass_sha1})}"

        try:
            key = _http_post_text(url)
        except IikoApiError as e:
            low = str(e).lower()
            if ("401" in low) or ("unauthorized" in low):
                raise IikoApiError(
                    "HTTP 401 Unauthorized. Либо sha1-хэш пароля неверный, либо у пользователя нет прав на REST API."
                )
            raise

        if key and ("error" not in key.lower()):
            return key

        raise IikoApiError(f"Не удалось получить ключ авторизации. Ответ: {key[:200]}")

    def _try_products_endpoints(self, key: str) -> Tuple[str, Any]:
        candidates = [
            f"{self.base_url}/api/v2/entities/products/list?{urlencode({'key': key})}",
            f"{self.base_url}/api/v2/entities/products/list?{urlencode({'key': key, 'includeDeleted': 'false'})}",
            f"{self.base_url}/api/v2/entities/nomenclature?{urlencode({'key': key})}",
            f"{self.base_url}/api/v2/entities/products?{urlencode({'key': key})}",
        ]
        last_err: Optional[str] = None
        for url in candidates:
            try:
                data = _http_get_json(url)
                return url, data
            except Exception as e:
                last_err = str(e)
                continue
        raise IikoApiError(f"Не удалось загрузить номенклатуру: {last_err}")

    def get_products(self) -> List[IikoProduct]:
        key = self.auth_key()
        url_used, data = self._try_products_endpoints(key)

        items: List[Dict[str, Any]] = []

        # API может вернуть:
        # - список dict
        # - dict с полем 'products'/'items'/'productCategories'
        if isinstance(data, list):
            items = [x for x in data if isinstance(x, dict)]
        elif isinstance(data, dict):
            for k in ("products", "items", "productItems"):
                v = data.get(k)
                if isinstance(v, list):
                    items = [x for x in v if isinstance(x, dict)]
                    break
            # иногда номенклатура лежит глубже
            if not items and "productCategories" in data and isinstance(data.get("productCategories"), list):
                # попробуем собрать продукты из категорий
                def walk(cat_list: list):
                    for cat in cat_list:
                        if not isinstance(cat, dict):
                            continue
                        prods = cat.get("products")
                        if isinstance(prods, list):
                            for p in prods:
                                if isinstance(p, dict):
                                    items.append(p)
                        children = cat.get("children")
                        if isinstance(children, list):
                            walk(children)
                walk(data.get("productCategories", []))
        else:
            raise IikoApiError(f"Непонятный ответ API от {url_used}: {str(data)[:200]}")

        out: List[IikoProduct] = []
        for it in items:
            name = _safe_str(it.get("name") or it.get("fullName"))
            if not name:
                continue

            # цена: у разных API по-разному
            price = ""
            for pk in ("price", "defaultPrice", "basePrice"):
                if it.get(pk) not in (None, ""):
                    price = _safe_str(it.get(pk))
                    break

            # часто цены лежат в sizePrices
            if not price:
                sp = it.get("sizePrices")
                if isinstance(sp, list) and sp:
                    # возьмём первую
                    if isinstance(sp[0], dict):
                        price = _safe_str(sp[0].get("price") or sp[0].get("value"))

            product_id = _safe_str(it.get("id") or it.get("productId"))
            out.append(IikoProduct(name=name, price=price, product_id=product_id))

        # уберём дубли по name
        seen = set()
        uniq: List[IikoProduct] = []
        for p in out:
            keyn = " ".join(p.name.lower().replace('ё', 'е').split())
            if keyn in seen:
                continue
            seen.add(keyn)
            uniq.append(p)

        uniq.sort(key=lambda x: " ".join(x.name.lower().replace('ё', 'е').split()))
        return uniq
