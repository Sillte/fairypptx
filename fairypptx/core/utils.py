from pydantic import TypeAdapter


from types import UnionType
from typing import Annotated, Any, Mapping, Sequence, cast

from fairypptx.core.types import COMObject
from fairypptx.object_utils import getattr as f_getattr, setattr as f_setattr


def get_discriminator_mapping(klass: UnionType | Annotated[Any, Any], field_name: str) -> dict[Any, type]:
    """Acquire the mapping of the discriminator to the class.
    """
    adapter = TypeAdapter(klass)
    core_schema = cast(dict[str, Any], adapter.core_schema)
    
    definitions = core_schema.get("definitions", [])
    def_map = {d["ref"]: d for d in definitions if "ref" in d}

    def _resolve_schema(schema: dict[str, Any]) -> dict[str, Any]:
        """definition-ref を解決する。解決先が自分自身なら再帰を止める。"""
        if not isinstance(schema, dict):
            return schema

        # 1. 参照キーを取得
        ref = schema.get("schema_ref") or schema.get("ref")
        
        # 2. definition-ref 型、もしくは schema_ref を持つ場合
        if (schema.get("type") == "definition-ref" or "schema_ref" in schema) and ref in def_map:
            resolved = def_map[ref]
            # 解決先が今の schema と同一オブジェクトなら、これ以上掘ると無限ループになる
            if resolved is schema:
                return resolved
            # 解決先に対して再帰（ここで ref が解決された実体になる）
            return _resolve_schema(resolved)
        
        # 3. default ラップの解決
        if schema.get("type") == "default" and "schema" in schema:
            return _resolve_schema(schema["schema"])
            
        return schema

    def _extract_choices(schema: dict[str, Any]):
        schema = _resolve_schema(schema)
        # union系を掘る
        if schema.get("type") in ("union", "tagged-union") and "choices" in schema:
            return schema["choices"]
        if schema.get("type") == "definitions" and "schema" in schema:
            return _extract_choices(schema["schema"])
        raise ValueError(f"Union schema not found. Type: {schema.get('type')}")

    choices = _extract_choices(core_schema)
    mapping = {}

    # --- 解析フェーズ ---
    if isinstance(choices, Mapping):
        for tag, sub_schema in choices.items():
            resolved = _resolve_schema(sub_schema)
            if "cls" in resolved:
                mapping[tag] = resolved["cls"]
        return mapping

    if isinstance(choices, Sequence):
        for elem in choices:
            resolved_elem = _resolve_schema(elem)
            target_cls = resolved_elem.get("cls")
            if not target_cls:
                continue

            inner_schema = _resolve_schema(resolved_elem.get("schema", {}))
            
            # モデルのフィールド定義 (fields) を探す
            fields = inner_schema.get("fields", {})
            if field_name not in fields:
                # model-fields 型でラップされている場合があるため再度解決
                if inner_schema.get("type") == "model-fields":
                    fields = inner_schema.get("fields", {})
                else:
                    # それでもなければ、モデルの中身 (schema) をもう一段掘る
                    inner_content = _resolve_schema(inner_schema.get("schema", {}))
                    fields = inner_content.get("fields", {})

            if field_name in fields:
                field_info = _resolve_schema(fields[field_name])
                literal_schema = _resolve_schema(field_info.get("schema", {}))

                if literal_schema.get("type") == "literal":
                    for item in literal_schema["expected"]:
                        mapping[item] = target_cls

        return mapping
    raise TypeError(f"Unsupported choices type: {type(choices)}")



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


def crude_api_read(api: COMObject, props: Sequence[str]) -> dict[str, Any]:
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
