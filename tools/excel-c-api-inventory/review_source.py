"""Reproduce the independent agent technical review of the pinned SDK source.

Does not import production inventory tooling or establish ABI/runtime behavior.
Source-first expectations were recorded before implementation inspection.
This reader accepts only the reviewed grammar, rejects forward references and
leading-zero decimal-looking literals, and is not a general C parser.
"""
import argparse
import collections
import hashlib
import json
from pathlib import Path
import re
import subprocess

ROOT = Path(__file__).resolve().parents[2]
REVISION = "6c8354d6536e18b58559f9cc3134f25a9e58d478"
SOURCE = "tools/Excel2013XLLSDK/INCLUDE/XLCALL.H"
BLOB = "4eb50e8fd3aea40d283181b7dcc74871ba88d854"
SHA256 = "d69f860ca5f7be9779f69423276808845ebb2f9648e38aadc99908277a710c74"


class ReviewError(ValueError):
    pass


def read_source():
    data = subprocess.check_output(["git", "show", f"{REVISION}:{SOURCE}"], cwd=ROOT)
    blob = hashlib.sha1(b"blob " + str(len(data)).encode() + b"\0" + data).hexdigest()
    if blob != BLOB or hashlib.sha256(data).hexdigest() != SHA256:
        raise ReviewError("reviewed source identity changed")
    return data.decode("latin1")


def integer(expression, values):
    """Recursive-descent integer/identifier/unary/OR grammar; no eval or AST."""
    tokens = re.findall(r"0[xX][0-9a-fA-F]+|[0-9]+|[A-Za-z_]\w*|[()+|~-]|\S", expression)
    cursor = 0

    def atom():
        nonlocal cursor
        if cursor == len(tokens):
            raise ReviewError("missing integer operand")
        token = tokens[cursor]
        cursor += 1
        if token in ("+", "-"):
            result = atom()
            return result if token == "+" else -result
        if token == "(":
            result = expression_or()
            if cursor == len(tokens) or tokens[cursor] != ")":
                raise ReviewError("unbalanced integer parentheses")
            cursor += 1
            return result
        if re.fullmatch(r"0[xX][0-9a-fA-F]+|[0-9]+", token):
            if len(token) > 1 and token.startswith("0") and not token.lower().startswith("0x"):
                raise ReviewError("unsupported leading-zero C integer literal")
            return int(token, 16 if token.lower().startswith("0x") else 10)
        if token in values:
            return values[token]
        raise ReviewError(f"unsupported or unresolved integer token {token!r}")

    def expression_or():
        nonlocal cursor
        result = atom()
        while cursor < len(tokens) and tokens[cursor] == "|":
            cursor += 1
            result |= atom()
        return result

    result = expression_or()
    if cursor != len(tokens):
        raise ReviewError("unsupported integer expression suffix")
    return result


def normalized(value):
    return re.sub(r"\s+", "", value)


def expectations(text):
    # Preserve positions while removing comments; directives remain inspectable.
    clean = re.sub(r"/\*.*?\*/|//[^\n]*", lambda m: re.sub(r"[^\n]", " ", m[0]), text, flags=re.S)
    original = text.splitlines()
    lines = clean.splitlines()
    records = []
    values = {}
    guards = []
    linkage = False
    index = 0
    family = None
    categories = collections.Counter()
    aggregate_fields = collections.Counter()
    # Categories are source-comment labels, not a second symbol registry.
    headings = {
        "XLOPER and XLOPER12 data types": ("type-flag", None),
        "Error codes": ("error-code", None),
        "Flow data types": ("flow-code", None),
        "Return codes": ("return-code", None),
        "XLL events": ("event-id", None),
        "Cluster connector entry point return codes": ("other-constant", "cluster-return-code"),
        "Function number bits": ("function-modifier", None),
        "Auxiliary function numbers": ("special-function-id", None),
        "edit modes": ("other-constant", "edit-mode"),
        "document(page) types": ("other-constant", "document-type"),
        "hit test codes": ("other-constant", "mouse-hit-test"),
        "User defined function": ("other-constant", "user-defined-function-id"),
        "Excel function numbers": ("worksheet-function-id", None),
        "Excel command numbers": ("macro-command-id", None),
    }

    def add(name, kind, start, end, **detail):
        entry = dict(symbol=name, kind=kind, line_start=start, line_end=end, **detail)
        if guards:
            entry["guard"] = list(guards)
        records.append(entry)

    while index < len(lines):
        line = lines[index].strip()
        source_line = original[index]
        number = index + 1
        for heading, category in headings.items():
            if heading in source_line and not line:
                family = category
        index += 1
        if not line:
            continue
        if line == "#pragma once":
            continue
        if line.startswith("#ifdef "):
            guards.append(line.split()[1])
            continue
        if line.startswith("#ifndef "):
            guards.append("!" + line.split()[1])
            continue
        if line == "#else":
            if not guards:
                raise ReviewError("unmatched else")
            guards[-1] = guards[-1][1:] if guards[-1].startswith("!") else "!" + guards[-1]
            continue
        if line == "#endif":
            if not guards:
                raise ReviewError("unmatched endif")
            guards.pop()
            continue
        if line == 'extern "C" {':
            if guards != ["__cplusplus"] or linkage:
                raise ReviewError("unreviewed linkage opening")
            linkage = True
            continue
        if line == "}":
            if guards != ["__cplusplus"] or not linkage:
                raise ReviewError("unreviewed linkage closing")
            linkage = False
            continue
        if line.startswith("#define "):
            match = re.fullmatch(r"#define\s+(\w+)\s+(.+)", line)
            if not match or not family:
                raise ReviewError(f"unclassified macro at {number}")
            name, expr = match.groups()
            if name in values:
                raise ReviewError("duplicate macro")
            values[name] = integer(expr, values)
            kind, subfamily = family
            if kind == "type-flag" and name.startswith("xlbit"):
                kind = "ownership-flag"
            detail = dict(expression=expr, value=values[name])
            if subfamily:
                detail["subfamily"] = subfamily
            alias = re.fullmatch(r"\(*([A-Za-z_]\w*)\)*", normalized(expr))
            if alias and alias[1] in values:
                detail["alias_of"] = alias[1]
            add(name, kind, number, number, **detail)
            categories[subfamily or kind] += 1
            continue
        if line.startswith("typedef struct "):
            block = [line]
            depth = line.count("{") - line.count("}")
            while not (depth == 0 and block[-1].endswith(";")):
                if index == len(lines):
                    raise ReviewError("unterminated aggregate")
                item = lines[index].strip()
                block.append(item)
                depth += item.count("{") - item.count("}")
                index += 1
            tag = re.fullmatch(r"typedef struct (\w+)", block[0])
            tail = re.fullmatch(r"}\s*(\w+)(?:\s*,\s*\*\s*(\w+))?;", block[-1])
            if not tag or not tail:
                raise ReviewError("unsupported aggregate typedef")
            name, pointer = tail.groups()
            conditional = []
            inner_guards = []
            for item in block[1:-1]:
                if item.startswith("#ifdef "):
                    inner_guards.append(item.split()[1])
                elif item == "#else" and inner_guards:
                    inner_guards[-1] = "!" + inner_guards[-1]
                elif item == "#endif" and inner_guards:
                    inner_guards.pop()
                elif item.startswith("#"):
                    raise ReviewError("unsupported aggregate conditional")
                elif item.endswith(";"):
                    aggregate_fields[name] += 1
                    if inner_guards:
                        conditional.append(dict(guard=inner_guards[-1], declaration=item[:-1]))
            if inner_guards:
                raise ReviewError("unterminated aggregate conditional")
            fields = dict(tag=tag[1])
            if conditional:
                fields["conditional_fields"] = conditional
            add(name, "struct", number, index, **fields)
            add(tag[1], "struct-tag", number, index, target=name)
            if pointer:
                add(pointer, "pointer-typedef", number, index, target=name)
            continue
        if line.startswith("typedef "):
            if "CALLBACK" in line:
                match = re.fullmatch(r"typedef (int \(CALLBACK \*(\w+)\)\(.+\));", line)
                if not match:
                    raise ReviewError("unsupported callback typedef")
                add(match[2], "function-pointer-typedef", number, number, signature=match[1])
            else:
                match = re.fullmatch(r"typedef (\w+) (\w+);", line)
                if not match:
                    raise ReviewError("unsupported scalar typedef")
                add(match[2], "scalar-typedef", number, number, underlying=match[1])
            continue
        match = re.fullmatch(r"((?:int|long) (?:_cdecl|pascal) (\w+)\(.*\));", line)
        if match and linkage:
            add(match[2], "function-prototype", number, number,
                signature=match[1], linkage='extern "C" when __cplusplus')
            continue
        raise ReviewError(f"unaccounted source line {number}: {line}")
    if guards or linkage:
        raise ReviewError("unterminated source conditional/linkage")
    identities = [(r["kind"], r["symbol"]) for r in records]
    if len(set(identities)) != len(identities):
        raise ReviewError("duplicate source identity")
    return records, dict(categories), dict(aggregate_fields)


def compare(text, catalogue, require_subfamilies=False):
    expected, categories, members = expectations(text)
    actual = {}
    differences = []
    for row in catalogue:
        key = (row["kind"], row["symbol"])
        if key in actual:
            differences.append(f"duplicate catalogue identity: {key}")
        actual[key] = row
    for item in expected:
        key = (item["kind"], item["symbol"])
        row = actual.pop(key, None)
        if row is None:
            differences.append(f"missing: {key}")
            continue
        source = row["source"]
        for field, value in dict(repository="evnekdev/excel-api", revision=REVISION,
                                 path=SOURCE, blob_sha=BLOB,
                                 line_start=item["line_start"], line_end=item["line_end"]).items():
            if source.get(field) != value:
                differences.append(f"{key}: source.{field}")
        for field, value in item.items():
            if field in ("line_start", "line_end", "kind", "symbol"):
                continue
            if field == "subfamily" and not require_subfamilies:
                continue
            current = row.get(field)
            same = normalized(current) == normalized(value) if isinstance(value, str) and isinstance(current, str) else current == value
            if not same:
                differences.append(f"{key}: {field}")
        for field in ("alias_of", "guard", "conditional_fields"):
            if field not in item and field in row:
                differences.append(f"{key}: unexpected {field}")
    differences.extend(f"extra: {key}" for key in actual)
    return dict(review="independent agent technical review; not human sign-off",
                source_sha256=hashlib.sha256(text.encode("latin1")).hexdigest(),
                physical_lines=len(text.splitlines()), expected_records=len(expected),
                source_categories=categories, aggregate_member_declarations=members,
                differences=differences,
                limit="Aggregate members reviewed separately in anchored evidence; no ABI/runtime or exhaustive C-parser claim")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--catalogue-revision", help="compare a Git catalogue snapshot")
    parser.add_argument("--require-subfamilies", action="store_true")
    args = parser.parse_args()
    path = "knowledge/excel-c-api/symbols.jsonl"
    data = subprocess.check_output(["git", "show", f"{args.catalogue_revision}:{path}"], cwd=ROOT) if args.catalogue_revision else (ROOT / path).read_bytes()
    result = compare(read_source(), [json.loads(line) for line in data.splitlines()], args.require_subfamilies)
    print(json.dumps(result, indent=2))
    return bool(result["differences"])


if __name__ == "__main__":
    raise SystemExit(main())
