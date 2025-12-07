from pydantic import BaseModel, JsonValue, TypeAdapter, Field
from enum import IntEnum
from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from types import UnionType
from fairypptx.core.types import COMObject

from pprint import pprint

from fairypptx.object_utils import setattr as f_setattr
from fairypptx.object_utils import getattr as f_getattr


def get_discriminator_mapping(klass: UnionType | Annotated, field_name: str):
    """Return {discriminator_value: BaseModel class} mapping."""
    adapter = TypeAdapter(klass)
    core_schema = cast(dict[str, Any], adapter.core_schema)

    def _to_choices(schema):
        if schema.get("type") in ("union", "tagged-union") and "choices" in schema:
            return schema["choices"]

        if schema.get("type") == "definitions":
            return _to_choices(schema["schema"])

        raise ValueError(f"Unsupported schema type: {schema.get('type')}")

    choices = _to_choices(core_schema)

    # Case 1: tagged-union → Mapping
    if isinstance(choices, Mapping):
        return {
            tag: sub_schema["cls"]
            for tag, sub_schema in choices.items()
            if "cls" in sub_schema
        }

    # Case 2: plain union → Sequence
    if isinstance(choices, Sequence):
        mapping = {}
        for elem in choices:
            target_cls = elem["cls"]

            fields = elem["schema"]["fields"]
            field_info = fields[field_name]
            assert field_info["type"] == "model-field"

            literal_schema = field_info["schema"]

            # unwrap default → literal
            if literal_schema["type"] == "default":
                literal_schema = literal_schema["schema"]

            if literal_schema["type"] == "literal":
                for item in literal_schema["expected"]:
                    mapping[item] = target_cls

        return mapping

    raise TypeError(f"Unsupported schema type: {type(choices)}")



class CrudeApiAccesssor:
    def __init__(self, props: Sequence[str]) -> None:
        self._props = props
    @property
    def props(self) -> Sequence[str]:
        return self._props

    def write(self, api: COMObject, data: Mapping[str, Any]) -> COMObject:
        for prop in self.props:
            f_setattr(api, prop, data[prop])
        return api
    
    def read(self, api: COMObject) -> Mapping[str, Any]:
        return {key: f_getattr(api, key) for key in self.props}

def crude_api_read(api: COMObject, props: Sequence[str]) -> Mapping[str, Any]:  
    return {key: f_getattr(api, key) for key in props}

def crude_api_write(api: COMObject, data:Mapping[str, Any]) -> COMObject:  
    for prop, value in data.items():
        f_setattr(api, prop, value)
    return api

def remove_invalidity(api:COMObject, data: Mapping[str, Any]) -> Mapping[str, Any]:
    """Remain only the valid keys for `api` object. 
    """
    remove_keys = set()
    for key, value in data.items():
        try:
            f_setattr(api, key, value)
        except ValueError:
            remove_keys.add(key)
    return {key: value for key, value in data.items() if key not in remove_keys}



