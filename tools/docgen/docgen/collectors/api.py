"""API 采集器：合并两个来源。

1. ``output/release/<tfm>/*.xml`` —— 编译器产出的 XML 文档注释（中文 summary/param/returns）
2. ``build/api-snapshot-baseline.json`` —— 公共 API 身份快照（成员计数 + 逐行成员签名）

产出一份「公共 API 清单 + 注释覆盖率」，用于自动生成 API 参考文档，
解决「手写文档跟不上代码、注释与文档不同步」的问题。
"""

from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Bing.Offices.Core.2.0.0.nupkg -> 包名 + 版本
_NUPKG_RE = re.compile(r"^(?P<name>.+?)\.(?P<version>\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.\-]+)?)\.nupkg$")

# 成员前缀 -> 中文种类名
_MEMBER_KINDS = {
    "T": "类型",
    "M": "方法",
    "P": "属性",
    "F": "字段",
    "E": "事件",
    "N": "命名空间",
}

_SNAPSHOT_KIND_CN = {
    "constructor": "构造函数",
    "method": "方法",
    "property": "属性",
    "field": "字段",
    "event": "事件",
    "type": "类型",
    "operator": "运算符",
    "nestedtype": "嵌套类型",
    "interface": "接口",
}
_MEMBER_KINDS.update(_SNAPSHOT_KIND_CN)


@dataclass
class ApiMemberDoc:
    """一个成员的文档注释。"""

    member_id: str
    kind: str = ""
    summary: str = ""
    returns: str = ""
    params: List[Tuple[str, str]] = field(default_factory=list)

    @property
    def kind_cn(self) -> str:
        return _MEMBER_KINDS.get(self.kind, self.kind or "未知")

    @property
    def name(self) -> str:
        """去掉 X:/M: 前缀的完整名。"""
        return self.member_id.split(":", 1)[-1] if ":" in self.member_id else self.member_id

    @property
    def namespace(self) -> str:
        parts = self.name.split(".")
        return ".".join(parts[:-1]) if len(parts) > 2 else ""

    @property
    def short_name(self) -> str:
        return self.name.rsplit(".", 1)[-1]


@dataclass
class AssemblyApi:
    """单个程序集的 API 概览。"""

    name: str
    tfm: str
    member_count: int = 0
    hash_short: str = ""
    xml_path: Optional[str] = None
    documented: List[ApiMemberDoc] = field(default_factory=list)

    @property
    def summary_count(self) -> int:
        return sum(1 for m in self.documented if m.summary)

    @property
    def coverage(self) -> float:
        """注释覆盖率（已写 summary 的类型数 / 类型总数）。"""
        types = [m for m in self.documented if m.kind == "T"]
        if not types:
            return 0.0
        return sum(1 for t in types if t.summary) / len(types) * 100.0


@dataclass
class ApiInventory:
    """全仓库 API 清单。"""

    schema: str = ""
    generator_version: str = ""
    baseline_commit: str = ""
    approved_by: str = ""
    approved_at: str = ""
    package_version: str = ""
    assemblies: List[AssemblyApi] = field(default_factory=list)
    snapshot_kinds: Dict[str, int] = field(default_factory=dict)

    @property
    def total_members(self) -> int:
        return sum(a.member_count for a in self.assemblies)

    @property
    def all_documented(self) -> List[ApiMemberDoc]:
        return [m for a in self.assemblies for m in a.documented]

    def public_types(self) -> List[ApiMemberDoc]:
        """公共类型清单（按命名空间排序）。"""
        types = [m for m in self.all_documented if m.kind == "T"]
        unique: Dict[str, ApiMemberDoc] = {}
        for item in types:
            unique.setdefault(item.name, item)
        return sorted(unique.values(), key=lambda m: m.name)

    def by_namespace(self) -> Dict[str, List[ApiMemberDoc]]:
        grouped: Dict[str, List[ApiMemberDoc]] = {}
        for item in self.public_types():
            grouped.setdefault(item.namespace or "(全局)", []).append(item)
        return dict(sorted(grouped.items()))


# --------------------------------------------------------------------------
def _parse_xml_doc(path: Path) -> List[ApiMemberDoc]:
    members: List[ApiMemberDoc] = []
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError:
        return members

    for member in root.iter("member"):
        member_id = member.get("name", "")
        if not member_id:
            continue
        kind = member_id.split(":", 1)[0] if ":" in member_id else ""

        def text(tag: str) -> str:
            node = member.find(tag)
            if node is None:
                return ""
            return " ".join("".join(node.itertext()).split())

        params = []
        for node in member.findall("param"):
            name = (node.get("name") or "").strip()
            value = " ".join("".join(node.itertext()).split())
            if name:
                params.append((name, value))

        members.append(
            ApiMemberDoc(
                member_id=member_id,
                kind=kind,
                summary=text("summary"),
                returns=text("returns"),
                params=params,
            )
        )
    return members


def _load_snapshot(path: Path) -> Tuple[dict, Dict[str, int]]:
    with path.open("r", encoding="utf-8-sig") as handle:
        data = json.load(handle)

    kinds: Dict[str, int] = {}
    for tfm_assemblies in data.get("assemblies", {}).values():
        for assembly in tfm_assemblies.values():
            for line in assembly.get("lines", []):
                kind = line.split("|", 1)[0].strip()
                if kind:
                    kinds[kind] = kinds.get(kind, 0) + 1
    return data, kinds


def collect_api(release_output_dir: Path, baseline_path: Path, preferred_tfm: str = "net8.0") -> ApiInventory:
    """采集 API 清单。

    以 ``preferred_tfm`` 下的 XML 注释为文档来源，以快照为成员计数来源。
    """
    inventory = ApiInventory()

    if baseline_path.exists():
        data, kinds = _load_snapshot(baseline_path)
        inventory.schema = data.get("schema", "")
        inventory.generator_version = data.get("generatorVersion", "")
        inventory.baseline_commit = (data.get("baselineCommit") or "")[:12]
        inventory.approved_by = data.get("approvedBy", "")
        inventory.approved_at = data.get("approvedAt", "")
        inventory.snapshot_kinds = kinds

        nupkg_files = data.get("candidateIdentity", {}).get("nupkgFiles", {})
        for file_name in nupkg_files:
            match = _NUPKG_RE.match(file_name)
            if match:
                inventory.package_version = match.group(2)
                break

        tfm_assemblies = data.get("assemblies", {}).get(preferred_tfm, {})
        xml_dir = release_output_dir / preferred_tfm

        for name, meta in sorted(tfm_assemblies.items()):
            xml_path = xml_dir / f"{name}.xml"
            assembly = AssemblyApi(
                name=name,
                tfm=preferred_tfm,
                member_count=int(meta.get("memberCount", 0)),
                hash_short=(meta.get("hash") or "")[:12],
            )
            if xml_path.exists():
                assembly.xml_path = str(xml_path)
                assembly.documented = _parse_xml_doc(xml_path)
            inventory.assemblies.append(assembly)

    if not inventory.assemblies:
        # 快照缺失时退化为纯 XML 注释扫描
        xml_dir = release_output_dir / preferred_tfm
        for xml_path in sorted(xml_dir.glob("*.xml")) if xml_dir.exists() else []:
            assembly = AssemblyApi(
                name=xml_path.stem, tfm=preferred_tfm, xml_path=str(xml_path)
            )
            assembly.documented = _parse_xml_doc(xml_path)
            assembly.member_count = len(assembly.documented)
            inventory.assemblies.append(assembly)

    return inventory
