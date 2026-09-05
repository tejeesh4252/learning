"""
config_loader.py

Loads org_config.yaml once and exposes it as a singleton so every module
(importers, validation engine, GUI) reads from the same source of truth.

Nothing in this codebase should hardcode a column name, account name, or
threshold — it should come through this loader.
"""

from __future__ import annotations
import os
import yaml
from typing import Any, Dict, List, Optional
from threading import Lock


class ConfigLoader:
    _instance: Optional["ConfigLoader"] = None
    _lock = Lock()

    def __new__(cls, config_path: Optional[str] = None):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._load(config_path)
            return cls._instance

    def _load(self, config_path: Optional[str]) -> None:
        if config_path is None:
            project_root = os.path.abspath(
                os.path.join(os.path.dirname(__file__), "..", "..")
            )
            config_path = os.path.join(project_root, "config", "org_config.yaml")

        if not os.path.exists(config_path):
            raise FileNotFoundError(
                f"Config file not found at {config_path}. "
                f"The application cannot run without org_config.yaml."
            )

        self.config_path = config_path
        with open(config_path, "r", encoding="utf-8") as f:
            self._data: Dict[str, Any] = yaml.safe_load(f)

        self.project_root = os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..", "..")
        )

    def reload(self) -> None:
        """Re-read the config file from disk (e.g. after user edits it in-app)."""
        self._load(self.config_path)

    # ------------------------------------------------------------------
    # Generic accessors
    # ------------------------------------------------------------------
    @property
    def raw(self) -> Dict[str, Any]:
        return self._data

    def get(self, *keys: str, default: Any = None) -> Any:
        node = self._data
        for key in keys:
            if not isinstance(node, dict) or key not in node:
                return default
            node = node[key]
        return node

    # ------------------------------------------------------------------
    # Domain-specific accessors
    # ------------------------------------------------------------------
    def get_file_types(self) -> Dict[str, Any]:
        return self._data.get("file_types", {})

    def get_file_type_def(self, file_type_key: str) -> Dict[str, Any]:
        return self.get_file_types().get(file_type_key, {})

    def get_field_aliases(self, file_type_key: str) -> Dict[str, List[str]]:
        return self.get_file_type_def(file_type_key).get("fields", {})

    def get_required_fields(self, file_type_key: str) -> List[str]:
        return self.get_file_type_def(file_type_key).get("required_fields", [])

    def get_filename_patterns(self, file_type_key: str) -> List[str]:
        return self.get_file_type_def(file_type_key).get("filename_patterns", [])

    def get_account_crosswalk(self) -> List[Dict[str, Any]]:
        return self._data.get("account_crosswalk", [])

    def get_matching_rules(self) -> List[Dict[str, Any]]:
        return self._data.get("matching_rules", [])

    def get_validation_settings(self) -> Dict[str, Any]:
        return self._data.get("validation", {})

    def get_duplicate_check_fields(self, file_type_key: str) -> List[str]:
        """Returns the correct duplicate-fingerprint fields for a given
        file type. Falls back to a generic list only if the config's
        duplicate_check_fields somehow isn't a per-type mapping."""
        raw = self._data.get("validation", {}).get("duplicate_check_fields", {})
        if isinstance(raw, dict):
            return raw.get(file_type_key, [])
        if isinstance(raw, list):
            return raw  # legacy flat-list format, used as-is
        return []

    def get_exception_categories(self) -> Dict[str, Any]:
        return self._data.get("exception_categories", {})

    def get_account_hierarchy(self) -> List[Dict[str, Any]]:
        return self._data.get("account_hierarchy", [])

    def get_payee_loan_mapping(self) -> List[Dict[str, Any]]:
        """Maps payee names (as they appear in bank narratives) to the
        Yardi investment/loan they relate to — used by the matching
        engine's Rule 6 (Payee-Mapped Group Total Match) for aggregation-
        account disbursements where individual lender-to-payee pairing
        isn't available in the source data."""
        return self._data.get("payee_loan_mapping", [])

    def get_team_directory(self) -> Dict[str, str]:
        """Returns {name: email} for everyone who can be assigned exceptions."""
        entries = self._data.get("team_directory", [])
        return {e.get("name"): e.get("email") for e in entries if e.get("name") and e.get("email")}

    def get_db_path(self) -> str:
        rel_path = self.get("database", "path", default="data/acr_database.db")
        return os.path.join(self.project_root, rel_path)

    def get_org_name(self) -> str:
        return self.get("organization", "name", default="Organization")

    def get_system_name(self) -> str:
        return self.get(
            "organization", "system_name", default="Cash Reconciliation System"
        )


def get_config(config_path: Optional[str] = None) -> ConfigLoader:
    """Convenience accessor used throughout the codebase."""
    return ConfigLoader(config_path)
