"""Kedro project settings."""
from pathlib import Path

# Use OmegaConfigLoader so we can interpolate ${runtime_params:entity} into paths.
CONF_SOURCE = "conf"

from kedro.config import OmegaConfigLoader  # noqa: E402

CONFIG_LOADER_CLASS = OmegaConfigLoader
CONFIG_LOADER_ARGS = {
    "base_env": "base",
    "default_run_env": "base",
}
