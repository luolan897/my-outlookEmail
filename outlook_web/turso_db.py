#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Turso-only database access for OutlookEmail.

This module is the single entry point for all database connections in the
application. It connects exclusively to a remote Turso database; there is no
local SQLite mode. If the Turso configuration is missing or invalid,
connecting fails fast with a clear error instead of silently falling back
to a local file.

Configuration (environment variables, never hardcoded):
    TURSO_DATABASE_URL: Turso database URL, e.g. libsql://<name>-<org>.turso.io
        Only libsql:// and https:// URLs are accepted. Local file URLs are
        rejected on purpose: this build does not support a local database.
    TURSO_AUTH_TOKEN: Turso auth token for the database above.

Why an adapter instead of using libsql directly:
    The application was written against Python's sqlite3 module. libsql's
    Python client is close to sqlite3 but differs in ways that would break
    existing code if used raw:
      * rows are plain tuples, while the app reads rows by column name
        (row['email']) everywhere;
      * libsql connections do not support the context manager protocol,
        while the app uses ``with db:`` blocks;
      * libsql raises ValueError for constraint/syntax errors, while the
        app catches sqlite3.IntegrityError / sqlite3.OperationalError;
      * libsql parameters must be tuples, while the app sometimes passes
        lists.
    The adapter below bridges exactly those gaps so the rest of the
    application keeps working unchanged. It deliberately covers only the
    sqlite3 surface the application actually uses (execute, executemany,
    executescript, commit, rollback, close, context manager, lastrowid,
    rowcount, description, fetchone/fetchall/fetchmany, PRAGMA passthrough).
"""

from __future__ import annotations

import os
import sqlite3
from urllib.parse import urlparse

import libsql_experimental as libsql

TURSO_URL_ENV = "TURSO_DATABASE_URL"
TURSO_TOKEN_ENV = "TURSO_AUTH_TOKEN"

# Only remote Turso schemes. Anything else (file:, http:, empty) is refused
# so a local SQLite database can never sneak back in as a runtime mode.
_ACCEPTED_SCHEMES = ("libsql", "https")


class TursoConfigurationError(RuntimeError):
    """Raised when the Turso configuration is missing or invalid."""


def turso_url_and_token() -> tuple[str, str]:
    """Return (url, token) from the environment, or raise a clear error."""
    url = (os.getenv(TURSO_URL_ENV) or "").strip()
    token = (os.getenv(TURSO_TOKEN_ENV) or "").strip()
    if not url:
        raise TursoConfigurationError(
            f"{TURSO_URL_ENV} is not set. Set it to your Turso database URL "
            "(e.g. libsql://<database>-<organization>.turso.io)."
        )
    scheme = urlparse(url).scheme.lower()
    if scheme not in _ACCEPTED_SCHEMES:
        raise TursoConfigurationError(
            f"{TURSO_URL_ENV} must be a remote Turso URL (libsql:// or https://); "
            f"got scheme {scheme!r}. Local database URLs are not supported."
        )
    if not token:
        raise TursoConfigurationError(
            f"{TURSO_TOKEN_ENV} is not set. Set it to your Turso auth token."
        )
    return url, token


def masked_turso_url() -> str:
    """Turso URL with any embedded credentials stripped, for safe logging."""
    try:
        url, _token = turso_url_and_token()
    except TursoConfigurationError:
        return "<not configured>"
    parsed = urlparse(url)
    host = parsed.hostname or ""
    return f"{parsed.scheme}://{host}"


def _translate_error(exc: Exception) -> Exception:
    """Map a libsql error onto the sqlite3 exception the app expects."""
    if isinstance(exc, sqlite3.Error):
        return exc
    message = str(exc)
    lowered = message.lower()
    if any(
        key in lowered
        for key in (
            "unique constraint failed",
            "foreign key constraint failed",
            "not null constraint failed",
            "check constraint failed",
            "primary key",
        )
    ):
        return sqlite3.IntegrityError(message)
    if any(
        key in lowered
        for key in (
            "no such table",
            "no such column",
            "no such function",
            "syntax error",
            "database is locked",
            "database table is locked",
            "readonly",
            "read-only",
            "disk i/o error",
        )
    ):
        return sqlite3.OperationalError(message)
    return sqlite3.DatabaseError(message)


def _normalize_params(params):
    """libsql requires tuple parameters; the app sometimes passes lists."""
    if params is None:
        return None
    if isinstance(params, list):
        return tuple(params)
    return params


def _normalize_many(seq_params):
    """Normalize an executemany parameter sequence (lists -> tuples)."""
    normalized = []
    for row in seq_params:
        if isinstance(row, list):
            normalized.append(tuple(row))
        else:
            normalized.append(row)
    return normalized


class Row(tuple):
    """A tuple row with sqlite3.Row-like column-name access.

    Supports row[0] (index), row['email'] (name), row.keys() and dict(row),
    matching how the application reads query results.
    """

    def __new__(cls, values, columns):
        obj = super().__new__(cls, tuple(values))
        obj._columns = tuple(columns)
        obj._index = {}
        for position, name in enumerate(obj._columns):
            if name is not None and name not in obj._index:
                obj._index[name] = position
        return obj

    def __getitem__(self, key):
        if isinstance(key, str):
            try:
                position = self._index[key]
            except KeyError:
                raise IndexError(f"No column '{key}'") from None
            return super().__getitem__(position)
        return super().__getitem__(key)

    def keys(self):
        return list(self._columns)


class Cursor:
    """sqlite3-compatible cursor wrapping a libsql cursor."""

    def __init__(self, raw_cursor):
        self._cursor = raw_cursor
        self.description = None

    def _wrap_one(self, row):
        if row is None:
            return None
        columns = [d[0] for d in (self.description or [])]
        return Row(row, columns)

    def _wrap_many(self, rows):
        # libsql returns None (not []) for statements with no result rows
        # (e.g. PRAGMA setters); sqlite3 returns [], so normalize here.
        if rows is None:
            return []
        columns = [d[0] for d in (self.description or [])]
        return [Row(row, columns) for row in rows]

    def _run(self, method, sql, params, many=False):
        try:
            if many:
                result = getattr(self._cursor, method)(sql, params)
            elif params is None:
                result = getattr(self._cursor, method)(sql)
            else:
                result = getattr(self._cursor, method)(sql, params)
        except Exception as exc:
            raise _translate_error(exc) from exc
        self.description = self._cursor.description
        return result

    def execute(self, sql, params=None):
        self._run("execute", sql, _normalize_params(params))
        return self

    def executemany(self, sql, seq_params):
        self._run("executemany", sql, _normalize_many(seq_params), many=True)
        return self

    def executescript(self, sql):
        try:
            self._cursor.executescript(sql)
        except Exception as exc:
            raise _translate_error(exc) from exc
        self.description = self._cursor.description
        return self

    def fetchone(self):
        try:
            row = self._cursor.fetchone()
        except Exception as exc:
            raise _translate_error(exc) from exc
        return self._wrap_one(row)

    def fetchall(self):
        try:
            rows = self._cursor.fetchall()
        except Exception as exc:
            raise _translate_error(exc) from exc
        return self._wrap_many(rows)

    def fetchmany(self, size=None):
        try:
            rows = (
                self._cursor.fetchmany(size)
                if size is not None
                else self._cursor.fetchmany()
            )
        except Exception as exc:
            raise _translate_error(exc) from exc
        return self._wrap_many(rows)

    @property
    def lastrowid(self):
        return self._cursor.lastrowid

    @property
    def rowcount(self):
        return self._cursor.rowcount

    def close(self):
        try:
            self._cursor.close()
        except Exception:
            pass


class Connection:
    """sqlite3-compatible connection wrapping a libsql connection.

    Rows are always Row objects (column-name access), so callers must not
    set row_factory; assigning it is accepted and ignored for compatibility
    with code that still does so.
    """

    def __init__(self, raw_connection):
        self._conn = raw_connection
        self.row_factory = None

    def cursor(self):
        return Cursor(self._conn.cursor())

    def execute(self, sql, params=None):
        cursor = self.cursor()
        cursor.execute(sql, params)
        return cursor

    def executemany(self, sql, seq_params):
        cursor = self.cursor()
        cursor.executemany(sql, seq_params)
        return cursor

    def executescript(self, sql):
        cursor = self.cursor()
        cursor.executescript(sql)
        return cursor

    def commit(self):
        # sqlite3 commit/rollback without an open transaction are no-ops.
        if not self._conn.in_transaction:
            return
        try:
            self._conn.commit()
        except Exception as exc:
            raise _translate_error(exc) from exc

    def rollback(self):
        if not self._conn.in_transaction:
            return
        try:
            self._conn.rollback()
        except Exception as exc:
            raise _translate_error(exc) from exc

    def close(self):
        try:
            self._conn.close()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        # Mirror sqlite3: commit on clean exit, rollback on exception.
        # The connection itself is NOT closed here, same as sqlite3.
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        return False

    @property
    def in_transaction(self):
        return self._conn.in_transaction


def connect_database(**_kwargs) -> Connection:
    """Connect to the Turso database.

    Turso-only: raises TursoConfigurationError if TURSO_DATABASE_URL or
    TURSO_AUTH_TOKEN is missing or the URL is not a remote Turso URL.
    Extra keyword arguments (e.g. timeout=) are accepted and ignored so
    existing call sites keep working.
    """
    url, token = turso_url_and_token()
    try:
        raw = libsql.connect(url, auth_token=token)
    except Exception as exc:
        raise _translate_error(exc) from exc
    return Connection(raw)
