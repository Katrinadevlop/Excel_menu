#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import time
from dataclasses import dataclass
from typing import Any, Dict, List, Optional
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from app.integrations.iiko_rms_client import IikoApiError, IikoProduct


@dataclass
class IikoOrganization:
    id: str
    name: str


def _safe_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _is_timeout_error(err: Exception) -> bool:
    s = str(err).lower()
    return ("timed out" in s) or ("timeout" in s) or ("handshake" in s)


def _http_request_json(
    method: str,
    url: str,
    body: Optional[dict] = None,
    headers: Optional[dict] = None,
    timeout_sec: int = 60,
    retries: int = 2,
) -> Any:
    data = None
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")

    req = Request(url, method=method.upper(), data=data)
    req.add_header("User-Agent", "excel_menu_gui")
    req.add_header("Accept", "application/json")
    if body is not None:
        req.add_header("Content-Type", "application/json")
    if headers:
        for k, v in headers.items():
            if v is None:
                continue
            req.add_header(str(k), str(v))

    last_exc: Optional[Exception] = None
    for attempt in range(retries + 1):
        try:
            with urlopen(req, timeout=timeout_sec) as resp:
                raw = resp.read()
                text = raw.decode("utf-8", errors="replace")
            try:
                return json.loads(text)
            except json.JSONDecodeError:
                return text
        except HTTPError as e:
            try:
                body_text = e.read().decode("utf-8", errors="replace")
            except Exception:
                body_text = ""
            ct = _safe_str(getattr(e, "headers", {}).get("Content-Type") if getattr(e, "headers", None) else "")
            raise IikoApiError(f"HTTP {e.code} при запросе {url}. CT={ct}. {body_text[:300]}")
        except URLError as e:
            last_exc = e
            if (attempt < retries) and _is_timeout_error(e):
                time.sleep(0.6 * (attempt + 1))
                continue
            raise IikoApiError(f"Ошибка соединения с iiko: {e}")

    raise IikoApiError(f"Ошибка соединения с iiko: {last_exc}")


class IikoCloudV1Client:
    """iikoCloud API v1 (api-ru.iiko.services) клиент.

    Поток:
      - POST /api/1/access_token  {apiLogin: ...} -> {token: ...}
      - GET /api/1/organizations  (Authorization: Bearer <token>)
      - POST /api/1/nomenclature  {organizationId: ...}

    api_url по умолчанию: https://api-ru.iiko.services
    """

    def __init__(
        self,
        api_url: str,
        api_login: str,
        organization_id: str = "",
        access_token: str = "",
    ):
        self.api_url = (api_url or "").strip().rstrip("/")
        self.api_login = (api_login or "").strip()
        self.organization_id = (organization_id or "").strip()
        self._token_cache: Optional[str] = (access_token or "").strip() or None

    def access_token(self) -> str:
        if self._token_cache:
            return self._token_cache

        if not self.api_url:
            raise IikoApiError("Не задан api_url для iikoCloud.")
        if not self.api_login:
            raise IikoApiError("Не задан apiLogin для iikoCloud.")

        url = f"{self.api_url}/api/1/access_token"
        data = _http_request_json("POST", url, body={"apiLogin": self.api_login})

        token = ""
        if isinstance(data, dict):
            token = _safe_str(data.get("token") or data.get("access_token") or data.get("accessToken"))
        elif isinstance(data, str):
            token = data.strip().strip('"')

        if not token:
            raise IikoApiError(f"Не удалось получить токен iikoCloud. Ответ: {str(data)[:300]}")

        self._token_cache = token
        return token

    def _auth_headers(self) -> Dict[str, str]:
        token = self.access_token()
        return {"Authorization": f"Bearer {token}"}

    def organizations(self) -> List[IikoOrganization]:
        if not self.api_url:
            raise IikoApiError("Не задан api_url.")

        url = f"{self.api_url}/api/1/organizations"

        # Встречается как GET; на всякий случай поддержим POST без тела тоже.
        last_err: Optional[str] = None
        data: Any = None
        for method, body in (("GET", None), ("POST", {})):
            try:
                data = _http_request_json(method, url, body=body, headers=self._auth_headers())
                last_err = None
                break
            except Exception as e:
                last_err = str(e)
                continue

        if last_err is not None:
            raise IikoApiError(f"Не удалось получить организации iikoCloud: {last_err}")

        out: List[IikoOrganization] = []
        if isinstance(data, list):
            for it in data:
                if not isinstance(it, dict):
                    continue
                oid = _safe_str(it.get("id") or it.get("organizationId"))
                name = _safe_str(it.get("name"))
                if oid:
                    out.append(IikoOrganization(id=oid, name=name or oid))
        elif isinstance(data, dict):
            items = data.get("organizations") or data.get("items")
            if isinstance(items, list):
                for it in items:
                    if not isinstance(it, dict):
                        continue
                    oid = _safe_str(it.get("id") or it.get("organizationId"))
                    name = _safe_str(it.get("name"))
                    if oid:
                        out.append(IikoOrganization(id=oid, name=name or oid))

        if not out:
            raise IikoApiError(f"Организации не найдены. Ответ: {str(data)[:300]}")

        return out

    def nomenclature(self, organization_id: str) -> Any:
        org_id = _safe_str(organization_id)
        if not org_id:
            raise IikoApiError("Не задан organizationId.")

        url = f"{self.api_url}/api/1/nomenclature"

        # Чаще всего это POST.
        candidates = [
            ("POST", {"organizationId": org_id}),
            ("POST", {"organizationId": org_id, "startRevision": 0}),
            ("GET", None),
        ]

        last_err: Optional[str] = None
        for method, body in candidates:
            try:
                if method == "GET":
                    # некоторые реализации допускают query param
                    url2 = f"{url}?{urlencode({'organizationId': org_id})}"
                    return _http_request_json("GET", url2, headers=self._auth_headers())
                return _http_request_json(method, url, body=body, headers=self._auth_headers())
            except Exception as e:
                last_err = str(e)
                continue

        raise IikoApiError(f"Не удалось загрузить номенклатуру iikoCloud: {last_err}")

    def get_products(self) -> List[IikoProduct]:
        org_id = self.organization_id
        if not org_id:
            orgs = self.organizations()
            if len(orgs) == 1:
                org_id = orgs[0].id
                self.organization_id = org_id
            else:
                raise IikoApiError("Не выбрана организация (organizationId).")

        data = self.nomenclature(org_id)

        items: List[Dict[str, Any]] = []
        if isinstance(data, dict):
            # стандартные варианты
            for k in ("products", "items", "productItems"):
                v = data.get(k)
                if isinstance(v, list):
                    items = [x for x in v if isinstance(x, dict)]
                    break

            # иногда лежит по категориям
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

        if not items:
            raise IikoApiError(f"Номенклатура загружена, но товары не найдены. Ответ: {str(data)[:400]}")

        out: List[IikoProduct] = []
        for it in items:
            name = _safe_str(it.get("name") or it.get("fullName"))
            if not name:
                continue

            product_id = _safe_str(it.get("id") or it.get("productId"))

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
            keyn = " ".join(p.name.lower().replace("ё", "е").split())
            if keyn in seen:
                continue
            seen.add(keyn)
            uniq.append(p)

        uniq.sort(key=lambda x: " ".join(x.name.lower().replace("ё", "е").split()))
        return uniq
