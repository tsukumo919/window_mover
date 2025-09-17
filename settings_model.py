import re
from pydantic import BaseModel, Field, field_validator, AliasChoices
from typing import List, Dict, Any, Literal, Union

# --- 基底モデル ---
class BaseSettingsModel(BaseModel):
    class Config:
        extra = 'forbid'  # 未知のフィールドを禁止

# --- Condition ---
class Condition(BaseSettingsModel):
    title: str | None = None
    process: str | None = None
    class_name: str | None = Field(None, alias='class')
    case_sensitive: bool = False

class ConditionGroup(BaseSettingsModel):
    logic: Literal['AND', 'OR'] = 'AND'
    conditions: List[Condition]

# --- Action ---
class ResizeTo(BaseSettingsModel):
    width: Union[str, int, None] = Field(None, validation_alias=AliasChoices('w', 'width'))
    height: Union[str, int, None] = Field(None, validation_alias=AliasChoices('h', 'height'))

class MoveTo(BaseSettingsModel):
    x: Union[str, int]
    y: Union[str, int]

class Offset(BaseSettingsModel):
    x: int = 0
    y: int = 0

# --- 型エイリアスを定義 ---
AnchorPoint = Literal[
    "TopLeft", "TopCenter", "TopRight",
    "MiddleLeft", "MiddleCenter", "MiddleRight",
    "BottomLeft", "BottomCenter", "BottomRight"
]

class Action(BaseSettingsModel):
    anchor: AnchorPoint = "TopLeft"
    move_to: Union[AnchorPoint, MoveTo, None] = None
    resize_to: ResizeTo | None = None
    offset: Offset | None = None
    target_monitor: int | None = None
    maximize: Literal["ON", "OFF"] = "OFF"
    minimize: Literal["ON", "OFF"] = "OFF"

    execution_delay: int | None = None
    target_workspace: int | None = None

# --- Rule ---
class Rule(BaseSettingsModel):
    name: str
    condition: Union[Condition, ConditionGroup]
    action: Action

# --- Ignore ---
class Ignore(BaseSettingsModel):
    name: str
    logic: Literal['AND', 'OR'] = 'OR'
    conditions: List[Condition]

# --- Global ---
class GlobalSettings(BaseSettingsModel):
    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = "INFO"
    apply_on_startup: bool = True
    apply_on_reload: bool = True
    apply_on_resume: bool = False
    recheck_on_title_change: bool = False
    cleanup_interval_seconds: int = 300
    monitor_offsets: Dict[str, Dict[str, int]] = Field(default_factory=dict)

    @field_validator('monitor_offsets')
    @classmethod
    def validate_monitor_offsets(cls, v: Dict[str, Dict[str, int]]) -> Dict[str, Dict[str, int]]:
        for key in v.keys():
            if key != 'default' and not re.match(r'^monitor_[1-9]\d*$', key):
                raise ValueError(f"Invalid key in monitor_offsets: '{key}'. Keys must be 'default' or 'monitor_N' where N is a number greater than 0.")
        return v

# --- Top Level ---
class SettingsModel(BaseSettingsModel):
    globals: GlobalSettings = Field(default_factory=GlobalSettings, alias='global')
    ignores: List[Ignore] = []
    rules: List[Rule] = []

    @field_validator('rules', 'ignores', mode='before')
    @classmethod
    def ensure_list(cls, v):
        return v if v is not None else []