"""تشغيل الـ agent من الـ terminal.

أمثلة:
    python main.py "افحص ملف sample_trial_balance.xlsx وجهّز ورقة عمل"
    python main.py            # وضع محادثة تفاعلي
    python main.py --mock "افحص ملف sample_trial_balance.xlsx"   # من غير API key
"""
import json
import sys

from dotenv import load_dotenv

load_dotenv()

from agent.loop import AuditAgent  # noqa: E402


def approve(tool_name: str, args: dict) -> bool:
    """موافقة بشرية قبل أي أداة بتكتب ملفات."""
    print(f"\n⚠ الـ agent عايز ينفّذ: {tool_name}")
    print(json.dumps(args, ensure_ascii=False, indent=2)[:1500])
    return input("موافق؟ (y/n) ").strip().lower() == "y"


def main() -> None:
    args = sys.argv[1:]
    mock = "--mock" in args
    args = [a for a in args if a != "--mock"]

    if mock:
        from agent.mock import MockClient

        print("⚠ وضع تجريبي: كلود متحاكى (من غير API key)، والأدوات بتشتغل فعليًا.")
        agent = AuditAgent(client=MockClient(), approver=approve)
    else:
        agent = AuditAgent(approver=approve)

    if args:
        print(agent.ask(" ".join(args)))
        return

    print("مساعد المراجعة جاهز. اكتب 'خروج' للإنهاء.")
    while True:
        text = input("\nإنت: ").strip()
        if text in ("خروج", "exit", "quit"):
            break
        if text:
            print("\nالمساعد:", agent.ask(text))


if __name__ == "__main__":
    main()
