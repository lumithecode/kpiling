from __future__ import annotations

import os
from datetime import datetime, timezone
from functools import lru_cache
from io import StringIO
from pathlib import Path
from typing import Any

import pandas as pd
from sqlalchemy import DateTime, String, Text, create_engine
from sqlalchemy.engine import make_url
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column
import streamlit as st
from streamlit.errors import StreamlitSecretNotFoundError


class Base(DeclarativeBase):
    pass


class WorkspaceRecord(Base):
    __tablename__ = "user_workspaces"

    user_id: Mapped[str] = mapped_column(String(64), primary_key=True)
    raw_data_json: Mapped[str] = mapped_column(Text, nullable=False)
    targets_json: Mapped[str] = mapped_column(Text, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)


def _normalized_database_url() -> str:
    database_url = os.getenv("DATABASE_URL", "").strip()
    if not database_url:
        try:
            database_url = str(st.secrets.get("DATABASE_URL", "")).strip()
        except StreamlitSecretNotFoundError:
            database_url = ""
    database_url = database_url or "sqlite:///./data/kpi_scorecard.db"
    if database_url.startswith("postgres://"):
        return "postgresql+psycopg://" + database_url.removeprefix("postgres://")
    if database_url.startswith("postgresql://"):
        return "postgresql+psycopg://" + database_url.removeprefix("postgresql://")
    return database_url


@lru_cache(maxsize=1)
def get_engine():
    """Create the database engine once per process."""
    database_url = _normalized_database_url()
    parsed_url = make_url(database_url)
    connect_args: dict[str, Any] = {}

    if parsed_url.get_backend_name() == "sqlite":
        connect_args["check_same_thread"] = False
        if parsed_url.database not in (None, ":memory:"):
            Path(parsed_url.database).parent.mkdir(parents=True, exist_ok=True)

    engine = create_engine(database_url, pool_pre_ping=True, connect_args=connect_args)
    Base.metadata.create_all(engine)
    return engine


def load_workspace(user_id: str) -> dict[str, Any] | None:
    """Load one user's persisted tables, or return None for a new workspace."""
    with Session(get_engine()) as session:
        record = session.get(WorkspaceRecord, user_id)
        if record is None:
            return None
        return {
            "raw_data": pd.read_json(StringIO(record.raw_data_json), orient="records"),
            "targets": pd.read_json(StringIO(record.targets_json), orient="records"),
            "updated_at": record.updated_at,
        }


def save_workspace(user_id: str, raw_data: pd.DataFrame, targets: pd.DataFrame) -> None:
    """Atomically save one user's weekly data and target definitions."""
    raw_data_json = raw_data.to_json(orient="records", date_format="iso", double_precision=15)
    targets_json = targets.to_json(orient="records", double_precision=15)
    updated_at = datetime.now(timezone.utc)

    with Session(get_engine()) as session, session.begin():
        record = session.get(WorkspaceRecord, user_id)
        if record is None:
            record = WorkspaceRecord(user_id=user_id, raw_data_json=raw_data_json, targets_json=targets_json, updated_at=updated_at)
            session.add(record)
        else:
            record.raw_data_json = raw_data_json
            record.targets_json = targets_json
            record.updated_at = updated_at
