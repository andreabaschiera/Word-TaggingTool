import os
from enum import StrEnum
from typing import Literal
import polars as pl
from pydantic import BaseModel, Field
import json


class Datatype(StrEnum):
    ENUM = "enumeration"
    TEXT = "narrative"
    BOOL = "boolean"
    ENUMSET = "enumerationset"
    PERC = "percent"
    MON = "monetary"
    DATE = "date"
    STR = "string"
    UNDEF = "undefined"
    YEAR = "year"
    GHG = "ghgemissions"
    DEC = "decimal"
    ENER = "energy"
    MASS = "mass"
    VOL = "volume"
    INT = "integer"


class ValueLabel(BaseModel):
    value: str
    label: str


class Dimension(BaseModel):
    required: bool
    dimension_id: str = Field(pattern=r"^dim_\d+$")
    dimension_mode: Literal["fixed", "typed"]
    dimension_title: str
    dimension_members: list[ValueLabel] | None = None


class Cube(BaseModel):
    dimensions: list[Dimension]


class Enumeration(BaseModel):
    enum_id: str = Field(pattern=r"^enum_\d+$")
    enum_members: list[ValueLabel]


class Datapoint(BaseModel):
    reference: str = Field(max_length=15)
    dr: str
    datatype: Datatype
    name: str
    qname: str
    enum: Enumeration | None = None
    dimensions: Cube | None = None


source_wb = "src/python/taxonomy_excel.xlsx"
dps = pl.read_excel(source=source_wb, sheet_name="datapoints")
enums = pl.read_excel(source=source_wb, sheet_name="enumerations")
dimensions = pl.read_excel(source=source_wb, sheet_name="dimensions")
dim_members = pl.read_excel(source=source_wb, sheet_name="members")


def build_enum(enum_id: str | None) -> Enumeration | None:
    if enum_id is None:
        return None

    enum_rows = enums.filter(pl.col("enum_id") == enum_id)
    if enum_rows.is_empty():
        return None

    enum_members = []
    for row in enum_rows.iter_rows(named=True):
        enum_members.append(
            ValueLabel(
                value=row["enum_value"],
                label=row["enum_label"],
            )
        )

    return Enumeration(
        enum_id=enum_id,
        enum_members=enum_members,
    )


def build_single_dimension(dimension_id: str, req: bool) -> Dimension | None:
    if not dimension_id:
        return None

    dimension_rows = dimensions.filter(pl.col("dimension_id") == dimension_id)
    if dimension_rows.is_empty():
        return None

    dimension_row = dimension_rows.row(0, named=True)
    dimension_members = []
    dim_members_filtered = dim_members.filter(pl.col("dimension_id") == dimension_id)
    for row in dim_members_filtered.iter_rows(named=True):
        dimension_members.append(
            ValueLabel(
                value=row["dim_value"],
                label=row["dim_label"],
            )
        )

    return Dimension(
        required=req,
        dimension_id=dimension_row["dimension_id"],
        dimension_mode=dimension_row["dimension_type"],
        dimension_title=dimension_row["dimension_name"],
        dimension_members=dimension_members or None,
    )


def build_dimensions(required_ids: str | None, optional_ids: str | None) -> Cube | None:
    if required_ids is None and optional_ids is None:
        return None

    parsed_dimensions = []
    if required_ids:
        for dimension_id in (part.strip() for part in required_ids.split(",")):
            single_dim = build_single_dimension(dimension_id, True)
            if single_dim:
                parsed_dimensions.append(single_dim)
            else:
                continue

    if optional_ids:
        for dimension_id in (part.strip() for part in optional_ids.split(",")):
            single_dim = build_single_dimension(dimension_id, False)
            if single_dim:
                parsed_dimensions.append(single_dim)
            else:
                continue

    if not parsed_dimensions:
        return None

    return Cube(dimensions=parsed_dimensions)


def write_json_file(data, filename):
    if not isinstance(data, (dict, list)):
        raise TypeError("Data must be a dictionary or list to be serialized to JSON.")

    try:
        os.makedirs(os.path.dirname(filename), exist_ok=True)

        with open(filename, "w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False)

        print(f"✅ JSON data successfully written to '{filename}'")

    except OSError as e:
        print(f"❌ File error: {e}")
    except (TypeError, ValueError) as e:
        print(f"❌ JSON serialization error: {e}")


list_dps: list[Datapoint] = []
for row in dps.iter_rows(named=True):
    datapoint_kwargs = {
        "reference": row["reference"],
        "dr": row["DR"],
        "datatype": row["type"],
        "name": row["name"],
        "qname": row["qname"],
    }

    enum = build_enum(row["enumerations_id"])
    if enum is not None:
        datapoint_kwargs["enum"] = enum

    cube = build_dimensions(row["dimensions_required"], row["dimensions_optional"])
    if cube is not None:
        datapoint_kwargs["dimensions"] = cube

    list_dps.append(Datapoint(**datapoint_kwargs))


dps_dict_list = [dp.model_dump(exclude_none=True) for dp in list_dps]
write_json_file(dps_dict_list, "src/python/taxonomy.json")
