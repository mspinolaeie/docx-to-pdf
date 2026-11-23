"""Regression tests for strict worker-count configuration types."""

import json

import pytest

from docx_to_pdf import ConversionConfig


@pytest.mark.parametrize("workers", [True, False])
def test_boolean_worker_count_is_rejected(workers):
    with pytest.raises(ValueError, match="workers must be >= 1"):
        ConversionConfig(workers=workers).validate()


@pytest.mark.parametrize("workers", [1, 2, 8])
def test_positive_integer_worker_count_is_accepted(workers):
    ConversionConfig(workers=workers).validate()


def test_json_boolean_worker_count_is_rejected(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(json.dumps({"workers": True}), encoding="utf-8")
    with pytest.raises(ValueError, match="workers must be >= 1"):
        ConversionConfig.from_file(str(config_path))
