"""
JiraClient — reusable HTTP client for the Zephyr Scale (ATM) REST API.
Encapsulates authentication, SSL context, and all HTTP verbs.
"""

import json
import ssl
import urllib.request
import urllib.parse
import urllib.error

from main.config import AppConfig


class JiraClient:

    def __init__(self):
        cfg            = AppConfig()
        self._pat      = cfg.jira_token
        self._atm_base = cfg.base_url + "/rest/atm/1.0"
        self._ssl_ctx  = ssl.create_default_context()

    # — Private helpers ————————————————————————————————————————————————

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._pat}",
            "Content-Type" : "application/json",
            "Accept"       : "application/json",
        }

    def _url(self, path: str, params: dict = None) -> str:
        url = self._atm_base + path
        if params:
            url += "?" + urllib.parse.urlencode(params)
        return url

    # — Public HTTP methods ————————————————————————————————————————————

    def get(self, path: str, params: dict = None) -> tuple:
        """GET request. Returns (status_code, parsed_json)."""
        req = urllib.request.Request(
            self._url(path, params), headers=self._headers(), method="GET"
        )
        try:
            with urllib.request.urlopen(req, context=self._ssl_ctx) as resp:
                return resp.status, json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode()[:300]
            except Exception:
                pass
            return e.code, {"error": body}
        except urllib.error.URLError as e:
            return 0, {"error": f"Network error: {e.reason}"}

    def put(self, path: str, payload: dict) -> tuple:
        """PUT request. Returns (status_code, parsed_json_or_dict)."""
        data = json.dumps(payload).encode("utf-8")
        req  = urllib.request.Request(
            self._url(path), data=data, headers=self._headers(), method="PUT"
        )
        try:
            with urllib.request.urlopen(req, context=self._ssl_ctx) as resp:
                body = resp.read().decode()
                try:
                    return resp.status, json.loads(body)
                except Exception:
                    return resp.status, {"body": body}
        except urllib.error.HTTPError as e:
            body = e.read().decode()[:300]
            try:
                return e.code, json.loads(body)
            except Exception:
                return e.code, {"error": body}
        except urllib.error.URLError as e:
            return 0, {"error": f"Network error: {e.reason}"}

    def post_list(self, path: str, payload: list) -> tuple:
        """POST request with a JSON array body. Returns (status_code, parsed_json)."""
        data = json.dumps(payload).encode("utf-8")
        req  = urllib.request.Request(
            self._url(path), data=data, headers=self._headers(), method="POST"
        )
        try:
            with urllib.request.urlopen(req, context=self._ssl_ctx) as resp:
                return resp.status, json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            return e.code, {"error": e.read().decode()[:300]}
        except urllib.error.URLError as e:
            return 0, {"error": f"Network error: {e.reason}"}

    def post(self, path: str, payload: dict) -> tuple:
        """POST request. Returns (status_code, parsed_json)."""
        data = json.dumps(payload).encode("utf-8")
        req  = urllib.request.Request(
            self._url(path), data=data, headers=self._headers(), method="POST"
        )
        try:
            with urllib.request.urlopen(req, context=self._ssl_ctx) as resp:
                return resp.status, json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            return e.code, {"error": e.read().decode()[:300]}
        except urllib.error.URLError as e:
            return 0, {"error": f"Network error: {e.reason}"}

    def delete(self, path: str) -> tuple:
        """DELETE request. Returns (status_code, response_body_str)."""
        req = urllib.request.Request(
            self._url(path), headers=self._headers(), method="DELETE"
        )
        try:
            with urllib.request.urlopen(req, context=self._ssl_ctx) as resp:
                return resp.status, resp.read().decode()
        except urllib.error.HTTPError as e:
            return e.code, e.read().decode()[:300]
        except urllib.error.URLError as e:
            return 0, f"Network error: {e.reason}"
