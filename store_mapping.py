import os
import re


def parse_mapping_text(text):
    bindings = []
    for raw_line in str(text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        parts = re.split(r"(?:\t+|\s{2,})", line, maxsplit=1)
        if len(parts) < 2:
            parts = line.split(None, 1)
        if len(parts) < 2:
            continue
        store_name = parts[0].strip()
        owned_store = parts[1].strip()
        if store_name and owned_store:
            bindings.append((store_name, owned_store))
    return bindings


def load_bindings(path):
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as file_obj:
        return parse_mapping_text(file_obj.read())


def dump_bindings(bindings):
    lines = []
    for store_name, owned_store in bindings:
        store = str(store_name or "").strip()
        owned = str(owned_store or "").strip()
        if store and owned:
            lines.append(f"{store}\t{owned}")
    return "\n".join(lines) + ("\n" if lines else "")


def save_bindings(path, bindings):
    with open(path, "w", encoding="utf-8") as file_obj:
        file_obj.write(dump_bindings(bindings))


def unique_store_names(bindings):
    names = []
    seen = set()
    for store_name, _ in bindings:
        if store_name not in seen:
            seen.add(store_name)
            names.append(store_name)
    return names


def owned_stores_for(bindings, store_name):
    store = str(store_name or "").strip()
    owned_names = []
    seen = set()
    for current_store, owned_store in bindings:
        if current_store != store:
            continue
        if owned_store in seen:
            continue
        seen.add(owned_store)
        owned_names.append(owned_store)
    return owned_names


def upsert_binding(bindings, store_name, owned_store):
    store = str(store_name or "").strip()
    owned = str(owned_store or "").strip()
    if not store or not owned:
        return list(bindings or [])
    normalized = list(bindings or [])
    if (store, owned) not in normalized:
        normalized.append((store, owned))
    return normalized
