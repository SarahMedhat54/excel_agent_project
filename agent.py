"""Agent: Groq LLM + the Excel MCP server. Used by streamlit_app.py."""
import json
import os
import re
import sys

from dotenv import load_dotenv
from groq import AsyncGroq
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

load_dotenv()
MODEL = os.getenv("MODEL", "openai/gpt-oss-120b")

# Chinese / Japanese / Korean characters: the model sometimes drifts into them
CJK = re.compile(r"[\u1100-\u11ff\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]")

SYSTEM_PROMPT = """You are an Excel assistant. With your tools you can create Excel files,
edit them, analyze them (EDA), and add charts.

Rules:
- Every file has a file_id. If the user message starts with [Current file_id: xxx], that is
  the file they are working on: use it for edits and analysis unless they ask for a new file.
- If you create a new file, use the file_id returned by create_excel for later steps.
- Before editing a file you did not just create, call describe_excel to see its structure.
- Do all edits first and add charts LAST (editing after charts can damage them).
- Never invent data the user did not ask for, except sample data when they ask for it.
- If a column the user mentions does not exist, say so instead of guessing.
- Tool arguments must be real JSON: numbers as numbers, lists as lists (not strings).
- LANGUAGE: reply in the language of the user's latest message (English or Arabic).
  If it is English, reply in English only. Never use Chinese, Japanese or Korean characters.
- When done, reply briefly: what you did, and key findings if you ran EDA.
  Do not mention file paths or file_id: the app shows a download button by itself."""

SERVER = StdioServerParameters(command=sys.executable, args=["server.py"])


async def run_agent(messages: list) -> tuple[str, list]:
    """One user turn. Returns (final_text, file_ids_touched)."""
    client = AsyncGroq()  # reads GROQ_API_KEY from .env
    file_ids: list = []
    extra: list = []  # temporary correction messages (not saved in the history)
    retried = False

    async with stdio_client(SERVER) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            listed = await session.list_tools()
            tools = [
                {
                    "type": "function",
                    "function": {
                        "name": t.name,
                        "description": t.description or "",
                        "parameters": t.inputSchema,
                    },
                }
                for t in listed.tools
            ]

            for _ in range(14):  # safety limit on tool rounds
                response = await client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role": "system", "content": SYSTEM_PROMPT}] + messages + extra,
                    tools=tools,
                    tool_choice="auto",
                    temperature=0.2,
                    max_tokens=2000,
                )
                msg = response.choices[0].message

                if not msg.tool_calls:
                    text = (msg.content or "").strip()
                    if CJK.search(text) and not retried:
                        # wrong language: ask for a rewrite once, without saving the bad reply
                        retried = True
                        extra = [
                            {"role": "assistant", "content": text},
                            {
                                "role": "user",
                                "content": "Rewrite your last reply in the same language as my last "
                                "message (English or Arabic). Do not use Chinese, Japanese or Korean.",
                            },
                        ]
                        continue
                    messages.append({"role": "assistant", "content": text})
                    return text, list(dict.fromkeys(file_ids))

                messages.append(
                    {
                        "role": "assistant",
                        "content": msg.content or "",
                        "tool_calls": [
                            {
                                "id": tc.id,
                                "type": "function",
                                "function": {"name": tc.function.name, "arguments": tc.function.arguments},
                            }
                            for tc in msg.tool_calls
                        ],
                    }
                )

                for tc in msg.tool_calls:
                    try:
                        args = json.loads(tc.function.arguments or "{}")
                    except json.JSONDecodeError:
                        out = "Error: tool arguments were not valid JSON. Try again."
                    else:
                        r = await session.call_tool(tc.function.name, args)
                        out = "\n".join(c.text for c in r.content if hasattr(c, "text"))
                        file_ids += re.findall(r'"file_id":\s*"(\w+)"', out)
                        if isinstance(args, dict) and args.get("file_id"):
                            file_ids.append(args["file_id"])
                    messages.append({"role": "tool", "tool_call_id": tc.id, "content": out})

    return "Stopped: too many steps. Try a simpler request.", list(dict.fromkeys(file_ids))