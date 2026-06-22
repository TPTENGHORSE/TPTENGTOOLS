import io
import importlib
import json
import os
import re
from datetime import datetime

import pandas as pd
import requests
import streamlit as st


MEMORY_FILE = os.path.join(os.path.dirname(__file__), "horseluis_memory.json")
DEFAULT_MODEL = "llama3"


def _chunk_text(text: str, size: int = 900, overlap: int = 150) -> list[str]:
    cleaned = re.sub(r"\s+", " ", text).strip()
    if not cleaned:
        return []
    chunks = []
    i = 0
    step = max(size - overlap, 1)
    while i < len(cleaned):
        chunks.append(cleaned[i : i + size])
        i += step
    return chunks


def _tokenize(text: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]{3,}", text.lower()))


def _extract_text_and_tables(uploaded_file) -> tuple[str, dict[str, pd.DataFrame], str | None]:
    name = uploaded_file.name
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    data = uploaded_file.getvalue()

    try:
        if ext in {"txt", "md", "log"}:
            return data.decode("utf-8", errors="ignore"), {}, None

        if ext == "csv":
            df = pd.read_csv(io.BytesIO(data))
            return df.to_csv(index=False), {"data": df}, None

        if ext in {"xlsx", "xls"}:
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None)
            parts = []
            for sheet_name, df in sheets.items():
                parts.append(f"Sheet: {sheet_name}\n{df.to_csv(index=False)}")
            return "\n\n".join(parts), sheets, None

        if ext == "pdf":
            try:
                pypdf = importlib.import_module("pypdf")
                PdfReader = getattr(pypdf, "PdfReader")
            except Exception:
                return "", {}, "PDF support requires 'pypdf' package."

            reader = PdfReader(io.BytesIO(data))
            pages = []
            for page in reader.pages:
                pages.append(page.extract_text() or "")
            return "\n".join(pages), {}, None

        return "", {}, f"Unsupported file type: .{ext}"
    except Exception as exc:
        return "", {}, f"Failed to read {name}: {exc}"


def _build_kb(files) -> tuple[list[dict], dict[str, dict[str, pd.DataFrame]], list[str]]:
    kb_chunks: list[dict] = []
    tabular_data: dict[str, dict[str, pd.DataFrame]] = {}
    warnings: list[str] = []

    for up_file in files:
        text, tables, err = _extract_text_and_tables(up_file)
        if err:
            warnings.append(f"{up_file.name}: {err}")
            continue

        if tables:
            tabular_data[up_file.name] = tables

        chunks = _chunk_text(text)
        for chunk in chunks:
            kb_chunks.append(
                {
                    "source": up_file.name,
                    "text": chunk,
                    "tokens": _tokenize(chunk),
                }
            )

    return kb_chunks, tabular_data, warnings


def _retrieve(kb_chunks: list[dict], question: str, top_k: int = 3) -> list[dict]:
    q_tokens = _tokenize(question)
    if not q_tokens:
        return []

    scored = []
    for chunk in kb_chunks:
        overlap = len(q_tokens & chunk["tokens"])
        if overlap > 0:
            scored.append((overlap, chunk))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:top_k]]


def _fallback_retrieve(kb_chunks: list[dict], top_k: int = 3) -> list[dict]:
    # Provide representative chunks when keyword overlap fails.
    if not kb_chunks:
        return []
    selected = []
    seen_sources = set()
    for chunk in kb_chunks:
        source = chunk.get("source", "")
        if source not in seen_sources:
            selected.append(chunk)
            seen_sources.add(source)
        if len(selected) >= top_k:
            return selected
    return kb_chunks[:top_k]


def _looks_like_doc_request(text: str) -> bool:
    t = text.lower()
    markers = [
        "documento adjunto",
        "adjunto",
        "archivo adjunto",
        "resumen del pdf",
        "resume el pdf",
        "attached document",
        "attached file",
        "summarize the pdf",
    ]
    return any(m in t for m in markers)


def _normalize_memories(raw_data) -> list[dict]:
    normalized = []
    next_id = 1
    if not isinstance(raw_data, list):
        return normalized

    for item in raw_data:
        if not isinstance(item, dict):
            continue
        text = str(item.get("text", "")).strip()
        if not text:
            continue
        normalized.append(
            {
                "id": int(item.get("id", next_id)),
                "text": text,
                "category": str(item.get("category", "general")),
                "priority": int(item.get("priority", 3)),
                "confidence": float(item.get("confidence", 0.8)),
                "source": str(item.get("source", "manual")),
                "uses": int(item.get("uses", 0)),
                "created_at": str(
                    item.get(
                        "created_at",
                        datetime.utcnow().isoformat(timespec="seconds") + "Z",
                    )
                ),
                "last_used": str(item.get("last_used", "")),
            }
        )
        next_id = max(next_id, normalized[-1]["id"] + 1)

    normalized.sort(key=lambda m: m["id"])
    return normalized


def _load_memory() -> list[dict]:
    if not os.path.exists(MEMORY_FILE):
        return []
    try:
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return _normalize_memories(data)
    except Exception:
        return []


def _save_memory(memories: list[dict]) -> None:
    try:
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(memories, f, ensure_ascii=False, indent=2)
    except Exception:
        # Keep the app usable even if writing the memory file fails.
        pass


def _add_memory(
    memories: list[dict],
    text: str,
    source: str,
    category: str = "general",
    priority: int = 3,
    confidence: float = 0.9,
) -> list[dict]:
    clean = re.sub(r"\s+", " ", text).strip()
    if not clean:
        return memories
    lowered = clean.lower()
    existing = {m.get("text", "").strip().lower() for m in memories}
    if lowered in existing:
        return memories

    next_id = max([int(m.get("id", 0)) for m in memories], default=0) + 1
    memories.append(
        {
            "id": next_id,
            "text": clean,
            "category": category,
            "priority": int(max(min(priority, 5), 1)),
            "confidence": float(max(min(confidence, 1.0), 0.0)),
            "source": source,
            "uses": 0,
            "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "last_used": "",
        }
    )
    return memories


def _auto_extract_memory(user_text: str) -> str | None:
    patterns = [
        r"^\s*recuerda que\s+(.+)$",
        r"^\s*aprende que\s+(.+)$",
        r"^\s*remember that\s+(.+)$",
        r"^\s*my name is\s+(.+)$",
        r"^\s*mi nombre es\s+(.+)$",
    ]
    for pattern in patterns:
        match = re.search(pattern, user_text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip(" .")
    return None


def _retrieve_memory(memories: list[dict], question: str, top_k: int = 4) -> list[dict]:
    q_tokens = _tokenize(question)
    if not q_tokens:
        return []

    scored = []
    for mem in memories:
        text = str(mem.get("text", ""))
        overlap = len(q_tokens & _tokenize(text))
        if overlap > 0:
            score = (
                (overlap * 2.0)
                + (float(mem.get("priority", 3)) * 0.6)
                + (float(mem.get("confidence", 0.8)) * 0.8)
            )
            scored.append((score, mem))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:top_k]]


def _mark_memories_used(memories: list[dict], used_ids: set[int]) -> list[dict]:
    now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    updated = []
    for mem in memories:
        if int(mem.get("id", 0)) in used_ids:
            mem["uses"] = int(mem.get("uses", 0)) + 1
            mem["last_used"] = now
        updated.append(mem)
    return updated


def _apply_memory_edits(memories: list[dict], edited_df: pd.DataFrame) -> list[dict]:
    by_id = {int(m["id"]): m for m in memories}
    for _, row in edited_df.iterrows():
        mem_id = int(row["id"])
        if mem_id not in by_id:
            continue
        by_id[mem_id]["text"] = str(row["text"]).strip()
        by_id[mem_id]["category"] = str(row["category"]).strip() or "general"
        by_id[mem_id]["priority"] = int(max(min(int(row["priority"]), 5), 1))
        by_id[mem_id]["confidence"] = float(max(min(float(row["confidence"]), 1.0), 0.0))
    cleaned = [m for m in by_id.values() if m.get("text")]
    cleaned.sort(key=lambda m: int(m.get("id", 0)))
    return cleaned


def _delete_memories(memories: list[dict], ids_to_delete: set[int]) -> list[dict]:
    return [m for m in memories if int(m.get("id", 0)) not in ids_to_delete]


def _compute_quote(distance_km: float, rate_per_km: float, fixed_cost: float, fuel_pct: float) -> dict:
    linehaul = distance_km * rate_per_km
    subtotal = linehaul + fixed_cost
    fuel_surcharge = subtotal * (fuel_pct / 100.0)
    total = subtotal + fuel_surcharge
    return {
        "linehaul": round(linehaul, 2),
        "subtotal": round(subtotal, 2),
        "fuel_surcharge": round(fuel_surcharge, 2),
        "total": round(total, 2),
    }


def run():
    st.markdown("## HorseLuis")
    st.caption("General assistant with long-term memory, PDF/doc grounding, and optional calculation tools")

    OLLAMA_URL = "http://localhost:11434/api/chat"

    if "messages" not in st.session_state:
        st.session_state["messages"] = []
    if "kb_chunks" not in st.session_state:
        st.session_state["kb_chunks"] = []
    if "kb_sources" not in st.session_state:
        st.session_state["kb_sources"] = []
    if "tabular_data" not in st.session_state:
        st.session_state["tabular_data"] = {}
    if "memory_entries" not in st.session_state:
        st.session_state["memory_entries"] = _load_memory()
    if "auto_learn" not in st.session_state:
        st.session_state["auto_learn"] = True
    if "model_name" not in st.session_state:
        st.session_state["model_name"] = DEFAULT_MODEL
    if "temperature" not in st.session_state:
        st.session_state["temperature"] = 0.2

    with st.sidebar:
        st.header("Assistant")
        st.text_input("Model", key="model_name")
        st.slider("Temperature", 0.0, 1.5, key="temperature", step=0.1)
        st.checkbox("Strict document mode", key="strict_docs_mode", value=False)
        colsb1, colsb2 = st.columns(2)
        with colsb1:
            if st.button("New chat"):
                st.session_state["messages"] = []
                st.rerun()
        with colsb2:
            if st.button("Clear memory"):
                st.session_state["memory_entries"] = []
                _save_memory([])
                st.rerun()

        st.divider()
        st.subheader("Knowledge Base")
        files = st.file_uploader(
            "Upload docs",
            type=["txt", "csv", "xlsx", "xls", "pdf", "md", "log"],
            accept_multiple_files=True,
        )
        if st.button("Index documents"):
            if not files:
                st.warning("Upload at least one document.")
            else:
                kb_chunks, tabular_data, warnings = _build_kb(files)
                st.session_state["kb_chunks"] = kb_chunks
                st.session_state["tabular_data"] = tabular_data
                st.session_state["kb_sources"] = sorted({c["source"] for c in kb_chunks})
                st.success(
                    f"Indexed {len(kb_chunks)} chunks from {len(st.session_state['kb_sources'])} files."
                )
                for warning in warnings:
                    st.warning(warning)

            st.caption("Index documents procesa los archivos subidos y los vuelve consultables en el chat.")

        st.caption(
            "Indexed files: " + ", ".join(st.session_state["kb_sources"])
            if st.session_state["kb_sources"]
            else "No indexed files"
        )

        st.divider()
        st.subheader("Calculation Tools")
        distance_km = st.number_input("Distance (km)", min_value=0.0, value=100.0)
        rate_per_km = st.number_input("Rate per km", min_value=0.0, value=2.0)
        fixed_cost = st.number_input("Fixed cost", min_value=0.0, value=35.0)
        fuel_pct = st.number_input("Fuel surcharge %", min_value=0.0, value=8.0)
        if st.button("Calculate quote"):
            quote = _compute_quote(distance_km, rate_per_km, fixed_cost, fuel_pct)
            tool_msg = (
                "Quote calculation\n"
                f"Linehaul: {quote['linehaul']}\n"
                f"Subtotal: {quote['subtotal']}\n"
                f"Fuel surcharge: {quote['fuel_surcharge']}\n"
                f"Total: {quote['total']}"
            )
            st.session_state["messages"].append(
                {"role": "assistant", "content": tool_msg, "sources": ["quotation_tool"]}
            )
            st.success(f"Total: {quote['total']}")

        st.divider()
        st.subheader("Excel Tool")
        tabular_data = st.session_state["tabular_data"]
        if tabular_data:
            source_name = st.selectbox("File", options=list(tabular_data.keys()))
            sheet_name = st.selectbox("Sheet", options=list(tabular_data[source_name].keys()))
            df_sheet = tabular_data[source_name][sheet_name]
            numeric_columns = [
                col
                for col in df_sheet.columns
                if pd.api.types.is_numeric_dtype(df_sheet[col])
            ]
            if numeric_columns:
                col_name = st.selectbox("Numeric column", options=numeric_columns)
                op_name = st.selectbox("Operation", options=["sum", "avg", "min", "max"])
                if st.button("Run Excel tool"):
                    series = df_sheet[col_name].dropna()
                    if op_name == "sum":
                        result = float(series.sum())
                    elif op_name == "avg":
                        result = float(series.mean())
                    elif op_name == "min":
                        result = float(series.min())
                    else:
                        result = float(series.max())

                    excel_msg = (
                        f"Excel tool result\nFile: {source_name}\nSheet: {sheet_name}\n"
                        f"Column: {col_name}\nOperation: {op_name}\nResult: {round(result, 4)}"
                    )
                    st.session_state["messages"].append(
                        {"role": "assistant", "content": excel_msg, "sources": [source_name]}
                    )
                    st.success(f"Result: {round(result, 4)}")
            else:
                st.info("No numeric columns available in this sheet.")
        else:
            st.caption("Index CSV/XLSX files to enable Excel tools.")

    with st.expander("Long-term memory", expanded=False):
        st.checkbox("Auto-learn from 'recuerda que' / 'remember that'", key="auto_learn")
        mem_col1, mem_col2, mem_col3, mem_col4 = st.columns([2, 1, 1, 1])
        with mem_col1:
            teach_text = st.text_input("Teach a memory", key="teach_memory_input")
        with mem_col2:
            teach_category = st.selectbox(
                "Category",
                options=["general", "customer", "pricing", "route", "personal"],
                key="teach_memory_category",
            )
        with mem_col3:
            teach_priority = st.slider("Priority", min_value=1, max_value=5, value=3, key="teach_priority")
        with mem_col4:
            teach_conf = st.slider("Confidence", min_value=0.0, max_value=1.0, value=0.9, step=0.1, key="teach_conf")

        mem_actions1, mem_actions2 = st.columns(2)
        with mem_actions1:
            if st.button("Save memory") and teach_text.strip():
                st.session_state["memory_entries"] = _add_memory(
                    st.session_state["memory_entries"],
                    teach_text,
                    source="manual",
                    category=teach_category,
                    priority=teach_priority,
                    confidence=teach_conf,
                )
                _save_memory(st.session_state["memory_entries"])
                st.success("Memory saved")
        with mem_actions2:
            if st.button("Clear all memory"):
                st.session_state["memory_entries"] = []
                _save_memory([])
                st.warning("Memory cleared")

        mem_df = pd.DataFrame(st.session_state["memory_entries"])
        if not mem_df.empty:
            editable_cols = ["id", "text", "category", "priority", "confidence", "source", "uses"]
            editor_df = mem_df[editable_cols].copy()
            edited_df = st.data_editor(
                editor_df,
                use_container_width=True,
                num_rows="fixed",
                disabled=["id", "source", "uses"],
                key="memory_editor",
            )
            apply_col1, apply_col2 = st.columns(2)
            with apply_col1:
                if st.button("Apply memory edits"):
                    st.session_state["memory_entries"] = _apply_memory_edits(
                        st.session_state["memory_entries"], edited_df
                    )
                    _save_memory(st.session_state["memory_entries"])
                    st.success("Memory updated")
            with apply_col2:
                delete_ids = st.multiselect("Delete by id", options=editor_df["id"].tolist())
                if st.button("Delete selected") and delete_ids:
                    st.session_state["memory_entries"] = _delete_memories(
                        st.session_state["memory_entries"], set(int(i) for i in delete_ids)
                    )
                    _save_memory(st.session_state["memory_entries"])
                    st.success("Selected memories deleted")

    for msg in st.session_state["messages"]:
        role = "assistant" if msg["role"] == "assistant" else "user"
        with st.chat_message(role):
            st.markdown(msg["content"])
            if msg.get("sources"):
                st.caption("Sources: " + ", ".join(msg["sources"]))
            if msg.get("memories"):
                st.caption("Memory used: " + "; ".join(msg["memories"]))

    user_input = st.chat_input("Send a message to HorseLuis")

    if user_input:
        st.session_state["messages"].append({"role": "user", "content": user_input})

        if _looks_like_doc_request(user_input) and not st.session_state["kb_chunks"]:
            bot_reply = (
                "No hay documentos indexados todavía. "
                "Sube tu archivo en 'Upload docs' y luego pulsa 'Index documents'."
            )
            st.session_state["messages"].append(
                {"role": "assistant", "content": bot_reply, "sources": []}
            )
            st.rerun()

        if st.session_state.get("auto_learn", True):
            extracted = _auto_extract_memory(user_input)
            if extracted:
                st.session_state["memory_entries"] = _add_memory(
                    st.session_state["memory_entries"],
                    extracted,
                    source="auto",
                    category="personal",
                    priority=4,
                    confidence=0.8,
                )
                _save_memory(st.session_state["memory_entries"])

        retrieved = _retrieve(st.session_state["kb_chunks"], user_input, top_k=3)
        if not retrieved and st.session_state["kb_chunks"]:
            retrieved = _fallback_retrieve(st.session_state["kb_chunks"], top_k=3)
        context_blocks = [f"Source: {c['source']}\n{c['text']}" for c in retrieved]
        context_text = "\n\n---\n\n".join(context_blocks)
        retrieved_memory = _retrieve_memory(st.session_state["memory_entries"], user_input, top_k=4)
        memory_text = "\n".join(f"- {m['text']}" for m in retrieved_memory)

        strict_docs_mode = bool(st.session_state.get("strict_docs_mode", False))
        if strict_docs_mode:
            system_prompt = (
                "You are HorseLuis, a helpful assistant. "
                "Use uploaded document context and long-term memory as primary evidence. "
                "If the answer is not present in provided context or memory, say clearly that it was not found there."
            )
        else:
            system_prompt = (
                "You are HorseLuis, a helpful general-purpose assistant. "
                "Use uploaded document context and long-term memory when relevant, "
                "but you can also answer general questions with your own model knowledge. "
                "When documents are used, prefer them for specific factual details."
            )

        user_prompt = user_input
        prompt_parts = []
        if memory_text:
            prompt_parts.append("Long-term memory facts:\n" + memory_text)
        if context_text:
            prompt_parts.append("Context from uploaded documents:\n" + context_text)
        if prompt_parts:
            user_prompt = "\n\n".join(prompt_parts) + f"\n\nQuestion: {user_input}"

        ollama_messages = [{"role": "system", "content": system_prompt}]
        for m in st.session_state["messages"][:-1]:
            ollama_messages.append({"role": m["role"], "content": m["content"]})
        ollama_messages.append({"role": "user", "content": user_prompt})

        payload = {
            "model": st.session_state.get("model_name", DEFAULT_MODEL),
            "messages": ollama_messages,
            "options": {"temperature": float(st.session_state.get("temperature", 0.2))},
        }

        full_reply = ""
        with st.chat_message("assistant"):
            response_placeholder = st.empty()
        try:
            response = requests.post(OLLAMA_URL, json=payload, timeout=120, stream=True)
            response.raise_for_status()
            for line in response.iter_lines():
                if not line:
                    continue
                try:
                    data = json.loads(line.decode("utf-8"))
                    content = data.get("message", {}).get("content")
                    if content:
                        full_reply += content
                        response_placeholder.markdown(full_reply + "▌")
                except Exception:
                    continue
            bot_reply = full_reply if full_reply else "[No reply]"
            response_placeholder.markdown(bot_reply)
        except Exception as exc:
            bot_reply = f"Error: {exc}"
            response_placeholder.markdown(bot_reply)

        sources = sorted({c["source"] for c in retrieved})
        memory_ids_used = {int(m["id"]) for m in retrieved_memory}
        memories_used = [m["text"] for m in retrieved_memory]
        st.session_state["memory_entries"] = _mark_memories_used(
            st.session_state["memory_entries"], memory_ids_used
        )
        _save_memory(st.session_state["memory_entries"])
        st.session_state["messages"].append(
            {
                "role": "assistant",
                "content": bot_reply,
                "sources": sources,
                "memories": memories_used,
            }
        )

if __name__ == "__main__":
    run()
