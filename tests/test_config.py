"""Tests for ConversionConfig and CLI argument parsing."""
import json
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_to_pdf import ConversionConfig, build_config_from_args, build_parser


# ---------------------------------------------------------------------------
# ConversionConfig
# ---------------------------------------------------------------------------


def test_default_values():
    cfg = ConversionConfig()
    assert cfg.root_dir == "."
    assert cfg.backend == "auto"
    assert cfg.bookmarks == "headings"
    assert cfg.workers == 1
    assert cfg.validate_pdf is True
    assert cfg.recursive is False
    assert cfg.overwrite is False
    assert cfg.pdfa is False
    assert cfg.log_file is None


def test_validate_invalid_backend():
    cfg = ConversionConfig(backend="invalid")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="Invalid backend"):
        cfg.validate()


def test_validate_invalid_bookmarks():
    cfg = ConversionConfig(bookmarks="invalid")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="Invalid bookmarks"):
        cfg.validate()


def test_validate_workers_zero():
    cfg = ConversionConfig(workers=0)
    with pytest.raises(ValueError, match="workers must be >= 1"):
        cfg.validate()


def test_validate_workers_negative():
    cfg = ConversionConfig(workers=-1)
    with pytest.raises(ValueError, match="workers must be >= 1"):
        cfg.validate()


def test_validate_empty_root_dir():
    cfg = ConversionConfig(root_dir="")
    with pytest.raises(ValueError, match="root_dir"):
        cfg.validate()


def test_validate_invalid_log_level():
    cfg = ConversionConfig(log_level="VERBOSE")
    with pytest.raises(ValueError, match="Invalid log_level"):
        cfg.validate()


def test_from_file_ignores_unknown_keys(tmp_path):
    config_data = {
        "root_dir": str(tmp_path),
        "recursive": True,
        "unknown_key": "should_be_ignored",
        "another_unknown": 42,
    }
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps(config_data))

    cfg = ConversionConfig.from_file(str(config_file))
    assert cfg.root_dir == str(tmp_path)
    assert cfg.recursive is True
    assert not hasattr(cfg, "unknown_key")


def test_from_file_raises_on_invalid_values(tmp_path):
    config_data = {"backend": "invalid_backend"}
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps(config_data))

    with pytest.raises(ValueError):
        ConversionConfig.from_file(str(config_file))


def test_round_trip_json(tmp_path):
    cfg = ConversionConfig(root_dir=str(tmp_path), recursive=True, workers=4)
    config_file = str(tmp_path / "config.json")
    cfg.to_file(config_file)
    cfg2 = ConversionConfig.from_file(config_file)
    assert cfg == cfg2


# ---------------------------------------------------------------------------
# CLI argument parsing
# ---------------------------------------------------------------------------


def _parse(args):
    parser = build_parser()
    return parser.parse_args(args)


def test_cli_defaults():
    args = _parse([])
    cfg = build_config_from_args(args)
    assert cfg.backend == "auto"
    assert cfg.recursive is False
    assert cfg.workers == 1


def test_cli_backend_flag():
    args = _parse(["--use", "libreoffice"])
    cfg = build_config_from_args(args)
    assert cfg.backend == "libreoffice"


def test_cli_recursive_flag():
    args = _parse(["--recursive"])
    cfg = build_config_from_args(args)
    assert cfg.recursive is True


def test_cli_overwrite_flag():
    args = _parse(["--overwrite"])
    cfg = build_config_from_args(args)
    assert cfg.overwrite is True


def test_cli_workers():
    args = _parse(["--workers", "8"])
    cfg = build_config_from_args(args)
    assert cfg.workers == 8


def test_cli_overrides_config_file(tmp_path):
    """CLI args must take precedence over values loaded from --config."""
    config_data = {"backend": "word", "workers": 8, "recursive": True}
    config_file = tmp_path / "config.json"
    config_file.write_text(json.dumps(config_data))

    args = _parse(["--config", str(config_file), "--use", "libreoffice"])
    cfg = build_config_from_args(args)
    assert cfg.backend == "libreoffice"   # CLI overrides config
    assert cfg.workers == 8               # config value preserved
    assert cfg.recursive is True          # config value preserved


def test_cli_invalid_log_level_rejected():
    with pytest.raises(SystemExit):
        _parse(["--log-level", "INVALID"])
