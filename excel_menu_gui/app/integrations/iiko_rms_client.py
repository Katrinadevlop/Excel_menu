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


def _extract_name_from_product_dict(it: Dict[str, Any]) -> str:
    """Пытается достать человекочитаемое название продукта из разных вариантов ответов API."""
    if not isinstance(it, dict):
        return ""

    for k in ("name", "fullName", "productName", "productFullName", "caption", "title"):
        v = it.get(k)
        if v not in (None, ""):
            s = _safe_str(v)
            if s:
                return s

    # иногда продукт лежит внутри обертки
    for k in ("product", "item", "data", "entity"):
        v = it.get(k)
        if isinstance(v, dict):
            s = _extract_name_from_product_dict(v)
            if s:
                return s

    return ""


def _extract_id_from_product_dict(it: Dict[str, Any]) -> str:
    """Пытается достать идентификатор продукта из разных вариантов ответов API."""
    if not isinstance(it, dict):
        return ""

    for k in ("id", "productId", "product_id", "productID", "itemId", "guid"):
        v = it.get(k)
        if v not in (None, ""):
            s = _safe_str(v)
            if s:
                return s

    # иногда продукт лежит внутри обертки
    for k in ("product", "item", "data", "entity"):
        v = it.get(k)
        if isinstance(v, dict):
            s = _extract_id_from_product_dict(v)
            if s:
                return s

    return ""


def _extract_price_value(raw: Any) -> str:
    """Нормализует цену к строке.

    В некоторых API цена приходит как:
    - число/строка
    - dict (например {currentPrice: 0.0, ...})
    - list/dict внутри sizePrices/prices
    """
    if raw is None:
        return ""

    if isinstance(raw, (int, float)):
        # 120.0 -> 120
        if isinstance(raw, float) and raw.is_integer():
            return str(int(raw))
        return str(raw)

    if isinstance(raw, str):
        return raw.strip()

    if isinstance(raw, dict):
        # частые варианты
        for k in (
            "currentPrice",
            "value",
            "price",
            "basePrice",
            "defaultPrice",
            "amount",
        ):
            if k in raw and raw.get(k) not in (None, ""):
                v = _extract_price_value(raw.get(k))
                if v:
                    return v

        # иногда цены лежат списком
        for k in ("prices", "sizePrices"):
            v = raw.get(k)
            if isinstance(v, list) and v:
                return _extract_price_value(v[0])

        return ""

    if isinstance(raw, list):
        if not raw:
            return ""
        return _extract_price_value(raw[0])

    return _safe_str(raw)


def _extract_price_from_product_dict(it: Dict[str, Any]) -> str:
    if not isinstance(it, dict):
        return ""

    # прямые ключи
    for pk in ("price", "defaultPrice", "basePrice"):
        if pk in it and it.get(pk) not in (None, ""):
            v = _extract_price_value(it.get(pk))
            if v:
                return v

    # sizePrices: берём первую цену
    sp = it.get("sizePrices")
    if isinstance(sp, list) and sp and isinstance(sp[0], dict):
        v = _extract_price_value(sp[0].get("price") or sp[0].get("value") or sp[0])
        if v:
            return v

    # иногда продукт внутри обертки
    for k in ("product", "item", "data", "entity"):
        v = it.get(k)
        if isinstance(v, dict):
            p = _extract_price_from_product_dict(v)
            if p:
                return p

    return ""


def _iter_dicts(obj: Any, *, max_depth: int = 6, _depth: int = 0):
    """Рекурсивно обходит структуру ответа и отдаёт все dict.

    Используется как fallback, когда API возвращает неожиданную форму.
    """
    if _depth > max_depth:
        return
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from _iter_dicts(v, max_depth=max_depth, _depth=_depth + 1)
    elif isinstance(obj, list):
        for v in obj:
            yield from _iter_dicts(v, max_depth=max_depth, _depth=_depth + 1)


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
        self._key_cache: Optional[str] = None

    def auth_key(self) -> str:
        """Получает auth key (строку-токен)."""
        if self._key_cache:
            return self._key_cache

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
            self._key_cache = key
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

    def open_product_from_stoplist(self, product_id: str) -> None:
        """Открыть блюдо = попытаться снять его со стоп-листа (сделать доступным).

        Точный эндпоинт зависит от версии iikoRMS, поэтому пробуем несколько вариантов.
        """
        product_id = _safe_str(product_id)
        if not product_id:
            raise IikoApiError("Не задан product_id для снятия со стоп-листа.")

        key = self.auth_key()

        candidates: List[Tuple[str, str]] = [
            # Часто встречается remove
            ("POST", f"{self.base_url}/api/stopList/remove?{urlencode({'key': key, 'productId': product_id})}"),
            ("GET", f"{self.base_url}/api/stopList/remove?{urlencode({'key': key, 'productId': product_id})}"),
            # Иногда снятие = установка остатка 0 через add
            ("POST", f"{self.base_url}/api/stopList/add?{urlencode({'key': key, 'productId': product_id, 'balance': '0'})}"),
            ("GET", f"{self.base_url}/api/stopList/add?{urlencode({'key': key, 'productId': product_id, 'balance': '0'})}"),
        ]

        last_err: Optional[str] = None
        for method, url in candidates:
            try:
                resp = _http_post_text(url) if method == "POST" else _http_get_text(url)
                low = (resp or "").lower()
                if (resp == "") or ("ok" in low) or ("true" in low) or ("success" in low):
                    return
                # Некоторые API возвращают JSON/строку без "error"
                if resp and ("error" not in low):
                    return
                last_err = resp
            except IikoApiError as e:
                last_err = str(e)
                continue

        raise IikoApiError(
            "Не удалось снять блюдо со стоп-листа через REST API. "
            f"Последняя ошибка: {last_err or ''}"
        )

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
            name = _extract_name_from_product_dict(it)
            if not name:
                continue

            price = _extract_price_from_product_dict(it)
            product_id = _extract_id_from_product_dict(it)
            out.append(IikoProduct(name=name, price=price, product_id=product_id))

        # Fallback: если почему-то вытащили подозрительно мало, пробуем рекурсивно найти продукты в ответе.
        # Это помогает, когда API меняет структуру и нужные поля лежат глубже.
        if len(out) <= 1 and isinstance(data, (dict, list)):
            extra: List[IikoProduct] = []
            for dct in _iter_dicts(data, max_depth=7):
                nm = _extract_name_from_product_dict(dct)
                pid = _extract_id_from_product_dict(dct)
                if not nm or not pid:
                    continue
                pr = _extract_price_from_product_dict(dct)
                extra.append(IikoProduct(name=nm, price=pr, product_id=pid))

            if extra:
                out.extend(extra)

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
