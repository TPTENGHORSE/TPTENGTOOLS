import streamlit as st
import requests
import os
import pandas as pd

# HorseLuis integration for app.py

def run():
    st.markdown("""
    <h2 style='text-align: center;'>
        <span style='font-size:2.5em;'>
            🐴
        </span><br>HorseLuis
    </h2>
    """, unsafe_allow_html=True)

    OLLAMA_URL = "http://localhost:11434/api/chat"
    OLLAMA_MODEL = "llama3.2:1b"  # Smaller model that uses less RAM

    # Session state for chat history
    if "messages" not in st.session_state:
        st.session_state["messages"] = []

    # User input only (no file upload)
    user_input = st.text_area("Type your message:", key="user_input")

    if st.button("Send") and user_input:
        st.session_state["messages"].append({"role": "user", "content": user_input})

        # Prepara el historial para Ollama (solo rol y contenido)
        ollama_messages = [
            {"role": m["role"], "content": m["content"]}
            for m in st.session_state["messages"]
        ]
        payload = {
            "model": OLLAMA_MODEL,
            "messages": ollama_messages
        }
        try:
            response = requests.post(OLLAMA_URL, json=payload, timeout=120, stream=True)
            response.raise_for_status()
            # Ollama puede devolver varias líneas JSON (streaming)
            import json
            full_reply = ""
            for line in response.iter_lines():
                if line:
                    try:
                        data = json.loads(line.decode('utf-8'))
                        content = data.get("message", {}).get("content")
                        if content:
                            full_reply += content
                    except Exception:
                        continue
            bot_reply = full_reply if full_reply else "[No reply]"
        except Exception as e:
            bot_reply = f"Error: {e}"
        # Add bot reply to history
        st.session_state["messages"].append({"role": "assistant", "content": bot_reply})

    # Display chat history
    for msg in st.session_state["messages"]:
        if msg["role"] == "user":
            st.markdown(f"<div style='text-align: right; color: #1a73e8;'><b>You:</b> {msg['content']}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div style='text-align: left; color: #34a853;'><b>HorseLuis:</b> {msg['content']}</div>", unsafe_allow_html=True)

# Allow running this file independently
if __name__ == "__main__":
    st.set_page_config(page_title="HorseLuis Chatbot", layout="centered")
    run()
