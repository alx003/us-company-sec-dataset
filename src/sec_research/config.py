from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class PipelineConfig:
    root: Path
    inputs: dict[str, Any]
    outputs: dict[str, Any]
    cache: dict[str, Any]
    sec: dict[str, Any]
    run: dict[str, Any]

    def path(self, section: str, key: str) -> Path:
        raw = self.__dict__[section][key]
        path = Path(raw)
        return path if path.is_absolute() else self.root / path


def load_config(path: str | Path) -> PipelineConfig:
    config_path = Path(path).resolve()
    data = _read_simple_yaml(config_path)
    return PipelineConfig(
        root=config_path.parent.parent,
        inputs=data.get("inputs", {}),
        outputs=data.get("outputs", {}),
        cache=data.get("cache", {}),
        sec=data.get("sec", {}),
        run=data.get("run", {}),
    )


def _read_simple_yaml(path: Path) -> dict[str, Any]:
    """Read the small YAML subset used by config/pipeline.yml."""
    result: dict[str, Any] = {}
    current: dict[str, Any] | None = None
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.split("#", 1)[0].rstrip()
        if not line:
            continue
        if not line.startswith(" ") and line.endswith(":"):
            key = line[:-1].strip()
            result[key] = {}
            current = result[key]
            continue
        if current is None or ":" not in line:
            continue
        key, raw_value = line.strip().split(":", 1)
        current[key.strip()] = _parse_scalar(raw_value.strip())
    return result


def _parse_scalar(value: str) -> Any:
    if value in {"null", "None", ""}:
        return None
    if value in {"true", "True"}:
        return True
    if value in {"false", "False"}:
        return False
    if value.startswith("[") and value.endswith("]"):
        inner = value[1:-1].strip()
        if not inner:
            return []
        return [_parse_scalar(part.strip()) for part in inner.split(",")]
    if value.startswith('"') and value.endswith('"'):
        return value[1:-1]
    if value.startswith("'") and value.endswith("'"):
        return value[1:-1]
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        return value
