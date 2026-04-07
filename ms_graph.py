from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import msal


GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


@dataclass
class GraphSettings:
    client_id: str
    tenant_id: str = "organizations"
    client_secret: Optional[str] = None
    user_id: Optional[str] = None

    @classmethod
    def from_env(cls) -> "GraphSettings":
        client_id = os.getenv("GRAPH_CLIENT_ID")
        if not client_id:
            raise ValueError("Missing GRAPH_CLIENT_ID environment variable.")
        return cls(
            client_id=client_id,
            tenant_id=os.getenv("GRAPH_TENANT_ID", "organizations"),
            client_secret=os.getenv("GRAPH_CLIENT_SECRET"),
            user_id=os.getenv("GRAPH_USER_ID"),
        )


class GraphClient:
    def __init__(self, settings: GraphSettings):
        self.settings = settings
        self.authority = f"https://login.microsoftonline.com/{settings.tenant_id}"

    def _acquire_token(self) -> str:
        if self.settings.client_secret:
            app = msal.ConfidentialClientApplication(
                client_id=self.settings.client_id,
                client_credential=self.settings.client_secret,
                authority=self.authority,
            )
            token_response = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        else:
            app = msal.PublicClientApplication(
                client_id=self.settings.client_id,
                authority=self.authority,
            )
            flow = app.initiate_device_flow(scopes=["Files.ReadWrite.All", "User.Read"])
            if "user_code" not in flow:
                raise RuntimeError(f"Unable to create device flow: {json.dumps(flow, indent=2)}")
            print(flow["message"])
            token_response = app.acquire_token_by_device_flow(flow)

        access_token = token_response.get("access_token")
        if not access_token:
            error = token_response.get("error_description") or json.dumps(token_response)
            raise RuntimeError(f"Unable to acquire Microsoft Graph access token: {error}")
        return access_token

    def _drive_prefix(self) -> str:
        if self.settings.user_id:
            return f"/users/{self.settings.user_id}/drive"
        return "/me/drive"

    def _request(self, method: str, path_or_url: str, *, data: bytes | None = None, expected_json: bool = True):
        token = self._acquire_token()
        url = path_or_url if path_or_url.startswith("http") else f"{GRAPH_BASE_URL}{path_or_url}"
        headers = {"Authorization": f"Bearer {token}"}
        if data is not None:
            headers["Content-Type"] = "application/octet-stream"
        request = urllib.request.Request(url=url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(request) as response:
                payload = response.read()
                if not expected_json:
                    return payload
                if not payload:
                    return {}
                return json.loads(payload.decode("utf-8"))
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"Graph API request failed ({exc.code}) for {url}: {body}") from exc

    def download_onedrive_file(self, onedrive_path: str, destination: Path) -> Path:
        graph_path = urllib.parse.quote(onedrive_path.lstrip("/"))
        payload = self._request(
            "GET",
            f"{self._drive_prefix()}/root:/{graph_path}:/content",
            expected_json=False,
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(payload)
        return destination

    def upload_onedrive_file(self, source: Path, onedrive_path: str) -> dict:
        graph_path = urllib.parse.quote(onedrive_path.lstrip("/"))
        file_bytes = source.read_bytes()
        return self._request(
            "PUT",
            f"{self._drive_prefix()}/root:/{graph_path}:/content",
            data=file_bytes,
            expected_json=True,
        )
