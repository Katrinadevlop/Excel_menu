#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import time
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlencode, urlsplit, urlunsplit, parse_qsl
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

from app.integrations.iiko_rms_client import IikoApiError, IikoProduct


@dataclass
class IikoOrganization:
    id: str
    name: str


def _safe_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _redact_url(url: str) -> str:
    try:
        parts = urlsplit(url)
        q = []
        for k, v in parse_qsl(parts.query, keep_blank_values=True):
            if k.lower() in ("user_secret", "access_token", "token"):
                q.append((k, "***"))
            else:
                q.append((k, v))
        new_query = urlencode(q)
        return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))
    except Exception:
        return url


def _dump_html_debug(kind: str, url: str, html_text: str) -> Optional[str]:
    """Пишет HTML-ответ в файл для отладки (на Рабочий стол), возвращает путь или None."""
    try:
        # чуть-чуть редактируем, чтобы не сохранить секреты если они случайно попали в HTML
        safe_text = (html_text or "")
        safe_text = safe_text.replace("access_token=", "access_token=***")
        safe_text = safe_text.replace("user_secret=", "user_secret=***")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"iiko_debug_{kind}_{ts}.html"
        out_path = Path.home() / "Desktop" / file_name
        out_path.write_text(
            f"<!-- URL: {_redact_url(url)} -->\n" + safe_text,
            encoding="utf-8",
            errors="replace",
        )
        return str(out_path)
    except Exception:
        return None


def _is_timeout_error(err: Exception) -> bool:
    s = str(err).lower()
    return ("timed out" in s) or ("timeout" in s) or ("handshake" in s)


def _http_get_raw(url: str, timeout_sec: int = 60, retries: int = 2, accept_json: bool = True) -> Tuple[int, str, str, str]:
    """HTTP GET, возвращает (status_code, content_type, content_encoding, text)."""
    req = Request(url)
    if accept_json:
        req.add_header("Accept", "application/json")
    req.add_header("User-Agent", "excel_menu_gui")

    last_exc: Optional[Exception] = None
    for attempt in range(retries + 1):
        try:
            with urlopen(req, timeout=timeout_sec) as resp:
                raw = resp.read()
                text = raw.decode("utf-8", errors="replace")
                status = getattr(resp, "status", None) or resp.getcode() or 0
                ct = _safe_str(resp.headers.get("Content-Type"))
                ce = _safe_str(resp.headers.get("Content-Encoding"))
            return int(status), ct, ce, text
        except HTTPError as e:
            try:
                body = e.read().decode("utf-8", errors="replace")
            except Exception:
                body = ""
            ct = _safe_str(getattr(e, "headers", {}).get("Content-Type") if getattr(e, "headers", None) else "")
            raise IikoApiError(f"HTTP {e.code} при запросе {_redact_url(url)}. CT={ct}. {body[:200]}")
        except URLError as e:
            last_exc = e
            if (attempt < retries) and _is_timeout_error(e):
                time.sleep(0.6 * (attempt + 1))
                continue
            raise IikoApiError(f"Ошибка соединения с iiko: {e}")

    raise IikoApiError(f"Ошибка соединения с iiko: {last_exc}")


def _try_parse_json_lenient(text: str) -> Any:
    """Пытается распарсить JSON, даже если сервер вернул его как текст или с мусором."""
    t = (text or "")
    if not t:
        raise json.JSONDecodeError("empty", "", 0)

    s = t.strip()
    if s.startswith("\ufeff"):
        s = s.lstrip("\ufeff").strip()

    # 1) обычный JSON
    try:
        val = json.loads(s)
    except json.JSONDecodeError:
        val = None

    # 2) JSON строкой "{...}" -> сначала строка, затем JSON внутри
    if isinstance(val, str):
        inner = val.strip()
        if inner.startswith("{") or inner.startswith("["):
            return json.loads(inner)
        return val

    if val is not None:
        return val

    # 3) вырезаем JSON-фрагмент (если есть префикс/суффикс)
    first_obj = s.find("{")
    first_arr = s.find("[")
    starts = [p for p in (first_obj, first_arr) if p != -1]
    if not starts:
        raise json.JSONDecodeError("no json start", s, 0)

    start = min(starts)
    end_obj = s.rfind("}")
    end_arr = s.rfind("]")
    end = max(end_obj, end_arr)
    if end == -1 or end <= start:
        raise json.JSONDecodeError("no json end", s, start)

    frag = s[start : end + 1]
    return json.loads(frag)


def _http_get_json(url: str, timeout_sec: int = 60, retries: int = 2) -> Any:
    status, ct, ce, text = _http_get_raw(url, timeout_sec=timeout_sec, retries=retries, accept_json=True)
    try:
        return _try_parse_json_lenient(text)
    except json.JSONDecodeError:
        return text


def _http_get_text(url: str, timeout_sec: int = 60, retries: int = 2) -> str:
    req = Request(url)
    req.add_header("User-Agent", "excel_menu_gui")
    last_exc: Optional[Exception] = None

    for attempt in range(retries + 1):
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
            last_exc = e
            if (attempt < retries) and _is_timeout_error(e):
                time.sleep(0.6 * (attempt + 1))
                continue
            raise IikoApiError(f"Ошибка соединения с iiko: {e}")

    raise IikoApiError(f"Ошибка соединения с iiko: {last_exc}")


class IikoCloudClient:
    """Клиент для iikoTransport/iikoCloud API.

    Поддерживает:
      - user_id/user_secret -> access_token
      - либо заранее полученный access_token (вручную)

    Базовые методы:
      - получение access_token
      - список организаций
      - загрузка номенклатуры

    api_url обычно: https://iiko.biz:9900
    """

    def __init__(
        self,
        api_url: str,
        user_id: str = "",
        user_secret: str = "",
        organization_id: str = "",
        access_token: str = "",
    ):
        self.api_url = (api_url or "").strip().rstrip("/")
        self.user_id = (user_id or "").strip()
        self.user_secret = (user_secret or "").strip()
        self.organization_id = (organization_id or "").strip()
        self._token_cache: Optional[str] = (access_token or "").strip().strip('"') or None

    def access_token(self) -> str:
        if self._token_cache:
            return self._token_cache

        if not self.api_url:
            raise IikoApiError("Не задан api_url для iiko.")

        if not self.user_id or not self.user_secret:
            raise IikoApiError(
                "Не задан user_id/user_secret. "
                "Либо введите их, либо используйте ручной ввод access_token в приложении."
            )

        # Для iikoTransport токен берётся на iiko.biz:9900.
        # Портал iikoweb.ru — это управление ключами, но сам access_token чаще всего выдаёт iiko.biz.
        url = f"{self.api_url}/api/0/auth/access_token?{urlencode({'user_id': self.user_id, 'user_secret': self.user_secret})}"

        try:
            token = _http_get_text(url)
        except IikoApiError as e:
            msg = str(e)
            # Частая ошибка: пытаются дергать access_token на *.iikoweb.ru
            if ("404" in msg) and ("iikoweb.ru" in self.api_url.lower()):
                raise IikoApiError(
                    "Этот API URL не поддерживает выдачу access_token. "
                    "Для Cloud/Transport API обычно нужен https://iiko.biz:9900"
                )
            raise

        token = token.strip().strip('"')
        if not token or "error" in token.lower():
            raise IikoApiError("Не удалось получить access_token: неверный user_id/user_secret или ключ не активен/не привязан к ресторану.")

        self._token_cache = token
        return token

    def organizations(self) -> List[IikoOrganization]:
        token = self.access_token()

        candidates = [
            f"{self.api_url}/api/0/organization/list?{urlencode({'access_token': token})}",
            f"{self.api_url}/api/0/organizations?{urlencode({'access_token': token})}",
        ]

        last_err: Optional[str] = None
        data: Any = None
        for url in candidates:
            try:
                data = _http_get_json(url)
                last_err = None
                break
            except Exception as e:
                last_err = str(e)
                continue

        if last_err:
            raise IikoApiError(f"Не удалось получить список организаций: {last_err}")

        out: List[IikoOrganization] = []
        if isinstance(data, list):
            for it in data:
                if not isinstance(it, dict):
                    continue
                oid = _safe_str(it.get("id") or it.get("organizationId") or it.get("guid"))
                name = _safe_str(it.get("name") or it.get("organizationName"))
                if oid:
                    out.append(IikoOrganization(id=oid, name=name or oid))
        elif isinstance(data, dict):
            items = data.get("organizations") or data.get("items")
            if isinstance(items, list):
                for it in items:
                    if not isinstance(it, dict):
                        continue
                    oid = _safe_str(it.get("id") or it.get("organizationId") or it.get("guid"))
                    name = _safe_str(it.get("name") or it.get("organizationName"))
                    if oid:
                        out.append(IikoOrganization(id=oid, name=name or oid))

        if not out:
            # часто приходит текст/HTML (например, страница ошибки) или неожиданный JSON
            if isinstance(data, str):
                snippet = data[:250]
            else:
                try:
                    snippet = json.dumps(data, ensure_ascii=False)[:250]
                except Exception:
                    snippet = str(data)[:250]
            raise IikoApiError(
                "Не удалось получить организации: список пустой или формат ответа не распознан. "
                f"Ответ: {snippet}"
            )

        return out

    def _try_nomenclature(self, organization_id: str) -> Any:
        token = self.access_token()
        org_id = _safe_str(organization_id)
        if not org_id:
            raise IikoApiError("Не задан organization_id для номенклатуры.")

        candidates = [
            f"{self.api_url}/api/0/nomenclature/{org_id}?{urlencode({'access_token': token})}",
            f"{self.api_url}/api/0/nomenclature?{urlencode({'access_token': token, 'organizationId': org_id})}",
        ]
        last_err: Optional[str] = None
        for url in candidates:
            try:
                status, ct, ce, text = _http_get_raw(url, accept_json=True)
                if not (text or "").strip():
                    raise IikoApiError(
                        "Номенклатура вернулась пустым ответом. "
                        f"HTTP {status}. CT={ct}. CE={ce}. URL={_redact_url(url)}"
                    )

                try:
                    return _try_parse_json_lenient(text)
                except json.JSONDecodeError:
                    # Частый случай: сервер возвращает HTML-страницу (например, страницу ошибки/заглушку)
                    raw_text = (text or "")
                    # Убираем BOM и ведущие пробелы/переводы строк, чтобы в сообщении не было пустого Ответ=''
                    stripped = raw_text.lstrip("\ufeff").lstrip()
                    snippet = repr(stripped[:600])
                    extra = f"LEN={len(raw_text)}"

                    if "text/html" in (ct or "").lower() or stripped.lower().startswith("<"):
                        dump_path = _dump_html_debug("nomenclature", url, raw_text)
                        hint = f" HTML сохранён в файл: {dump_path}" if dump_path else ""
                        raise IikoApiError(
                            "Номенклатура вернулась как HTML (не API-ответ). "
                            f"HTTP {status}. CT={ct}. CE={ce}. {extra}.{hint} Ответ={snippet}"
                        )
                    raise IikoApiError(
                        "Номенклатура вернулась НЕ в JSON (или формат с мусором не распознан). "
                        f"HTTP {status}. CT={ct}. CE={ce}. {extra}. Ответ={snippet}"
                    )

            except Exception as e:
                last_err = str(e)
                continue
        raise IikoApiError(f"Не удалось загрузить номенклатуру: {last_err}")

    def get_products(self) -> List[IikoProduct]:
        org_id = self.organization_id
        if not org_id:
            orgs = self.organizations()
            if len(orgs) == 1:
                org_id = orgs[0].id
                self.organization_id = org_id
            else:
                raise IikoApiError("Не выбрана организация (organization_id).")

        data = self._try_nomenclature(org_id)

        # Пытаемся собрать продукты из разных форматов ответа
        items: List[Dict[str, Any]] = []
        if isinstance(data, dict):
            for k in ("products", "items", "productItems"):
                v = data.get(k)
                if isinstance(v, list):
                    items = [x for x in v if isinstance(x, dict)]
                    break
            if not items and "productCategories" in data and isinstance(data.get("productCategories"), list):
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
        elif isinstance(data, list):
            items = [x for x in data if isinstance(x, dict)]

        # Если не смогли выделить список товаров — покажем понятную ошибку,
        # чтобы можно было подстроить парсер под ваш формат ответа.
        if not items:
            if isinstance(data, dict):
                keys = ", ".join(sorted([str(k) for k in data.keys()]))
                try:
                    snippet = json.dumps(data, ensure_ascii=False)[:400]
                except Exception:
                    snippet = str(data)[:400]
                raise IikoApiError(
                    "Номенклатура получена, но товары не найдены (не распознан формат ответа). "
                    f"Ключи ответа: {keys}. Ответ: {snippet}"
                )
            if isinstance(data, str):
                raw_text = data
                stripped = raw_text.lstrip("\ufeff").lstrip()
                snippet = repr(stripped[:600])
                raise IikoApiError(
                    "Номенклатура вернулась текстом (не JSON) и не удалось выделить JSON. "
                    f"LEN={len(raw_text)}. Ответ={snippet}"
                )

        out: List[IikoProduct] = []
        for it in items:
            name = _safe_str(it.get("name") or it.get("fullName"))
            if not name:
                continue

            product_id = _safe_str(it.get("id") or it.get("productId"))

            # цена: по возможности
            price = ""
            for pk in ("price", "defaultPrice", "basePrice"):
                if it.get(pk) not in (None, ""):
                    price = _safe_str(it.get(pk))
                    break
            if not price:
                sp = it.get("sizePrices")
                if isinstance(sp, list) and sp and isinstance(sp[0], dict):
                    price = _safe_str(sp[0].get("price") or sp[0].get("value"))

            out.append(IikoProduct(name=name, price=price, product_id=product_id))

        # uniq by name
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
