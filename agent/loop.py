"""الـ agent loop: ابعت -> لو tool_use نفّذ وارجّع النتيجة -> كرّر لحد end_turn."""
from __future__ import annotations

import json
import os
from typing import Callable

import anthropic

from .prompts import SYSTEM_PROMPT
from .tools import TOOLS, run_tool


def _text(content) -> str:
    return "".join(b.text for b in content if b.type == "text")


class AuditAgent:
    def __init__(
        self,
        client: anthropic.Anthropic | None = None,
        model: str | None = None,
        max_turns: int = 12,
        approver: Callable[[str, dict], bool] | None = None,
        log: Callable[[str], None] = print,
    ):
        self.client = client or anthropic.Anthropic()  # بيقرا ANTHROPIC_API_KEY
        self.model = model or os.environ.get("AUDIT_MODEL", "claude-sonnet-5-5")
        self.max_turns = max_turns
        self.approver = approver
        self.log = log
        self.messages: list[dict] = []  # ذاكرة المحادثة (قصيرة المدى)

    def ask(self, user_text: str) -> str:
        self.messages.append({"role": "user", "content": user_text})

        for _ in range(self.max_turns):
            resp = self.client.messages.create(
                model=self.model,
                max_tokens=4096,
                system=SYSTEM_PROMPT,
                tools=TOOLS,
                messages=self.messages,
            )
            self.messages.append({"role": "assistant", "content": resp.content})

            if resp.stop_reason == "tool_use":
                results = []
                for block in resp.content:
                    if block.type != "tool_use":
                        continue
                    self.log(f"  ⚙ {block.name}({json.dumps(block.input, ensure_ascii=False)})")
                    out, is_error = run_tool(block.name, block.input, self.approver)
                    results.append(
                        {
                            "type": "tool_result",
                            "tool_use_id": block.id,
                            "content": json.dumps(out, ensure_ascii=False, default=str),
                            "is_error": is_error,
                        }
                    )
                self.messages.append({"role": "user", "content": results})
                continue

            if resp.stop_reason == "end_turn":
                return _text(resp.content)

            # max_tokens أو refusal أو غيره
            return f"[توقف الـ agent: {resp.stop_reason}]\n{_text(resp.content)}"

        return "[وصل للحد الأقصى من الدورات قبل ما يخلّص. جرّب تقسّم المهمة.]"
