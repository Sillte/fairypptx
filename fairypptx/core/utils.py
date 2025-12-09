from pydantic import TypeAdapter


from types import UnionType
from typing import Annotated, Any, Mapping, Sequence, cast

from fairypptx.core.types import COMObject
from fairypptx.object_utils import getattr as f_getattr, setattr as f_setattr


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


def swap_props(api1: COMObject, api2: COMObject, attrs: Sequence[str]) -> None:
    ps1 = [f_getattr(api1, attr) for attr in attrs]
    ps2 = [f_getattr(api2, attr) for attr in attrs]
    for attr, p1, p2 in zip(attrs, ps1, ps2):
        f_setattr(api1, attr, p2)
        f_setattr(api2, attr, p1)
