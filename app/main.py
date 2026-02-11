import json
import re
import secrets
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import END, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
import tkinter as tk

import pyzipper
import win32com.client


APP_NAME = "伊蓮娜的煩惱"
APP_DIR = Path(__file__).resolve().parent
PROJECT_DIR = APP_DIR.parent
RUNTIME_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else PROJECT_DIR
LOG_DIR = RUNTIME_DIR / "logs"
OUTBOX_DIR = RUNTIME_DIR / "outbox"
CONFIG_PATH = RUNTIME_DIR / "config.json"
AUDIT_LOG_PATH = LOG_DIR / "audit.jsonl"

SPLASH_IMAGE_REL = Path("photo") / "elena.png"
SPLASH_MS = 1800

DEFAULT_SUBJECT_1 = "履歷附件（加密壓縮檔）"
DEFAULT_BODY_1 = "您好，\n\n附件為加密壓縮後的履歷檔案。\n\n謝謝。"
DEFAULT_SUBJECT_2 = "履歷壓縮檔密碼"
DEFAULT_BODY_2 = "您好，\n\n履歷壓縮檔密碼如下：{password}\n\n謝謝。"


@dataclass
class Config:
    default_sender: str = ""
    subject_1: str = DEFAULT_SUBJECT_1
    body_1: str = DEFAULT_BODY_1
    subject_2: str = DEFAULT_SUBJECT_2
    body_2: str = DEFAULT_BODY_2


def resource_path(relative_path: Path) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return PROJECT_DIR / relative_path


def ensure_dirs() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    OUTBOX_DIR.mkdir(parents=True, exist_ok=True)


def load_config() -> Config:
    ensure_dirs()
    if not CONFIG_PATH.exists():
        return Config()
    try:
        raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return Config(
            default_sender=raw.get("default_sender", ""),
            subject_1=raw.get("subject_1", DEFAULT_SUBJECT_1),
            body_1=raw.get("body_1", DEFAULT_BODY_1),
            subject_2=raw.get("subject_2", DEFAULT_SUBJECT_2),
            body_2=raw.get("body_2", DEFAULT_BODY_2),
        )
    except Exception:
        return Config()


def save_config(cfg: Config) -> None:
    ensure_dirs()
    CONFIG_PATH.write_text(
        json.dumps(
            {
                "default_sender": cfg.default_sender,
                "subject_1": cfg.subject_1,
                "body_1": cfg.body_1,
                "subject_2": cfg.subject_2,
                "body_2": cfg.body_2,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def append_audit(record: dict) -> None:
    ensure_dirs()
    with AUDIT_LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def is_valid_email(value: str) -> bool:
    pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    return bool(re.match(pattern, value.strip()))


def generate_password_8_digits() -> str:
    return "".join(str(secrets.randbelow(10)) for _ in range(8))


def create_protected_zip(input_file: Path, password: str) -> Path:
    ensure_dirs()
    if not input_file.exists():
        raise FileNotFoundError(f"File not found: {input_file}")
    if not re.fullmatch(r"\d{8}", password):
        raise ValueError("Password must be exactly 8 digits.")

    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", input_file.stem)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = OUTBOX_DIR / f"{safe_name}_{timestamp}.zip"

    with pyzipper.AESZipFile(
        output,
        "w",
        compression=pyzipper.ZIP_DEFLATED,
        encryption=pyzipper.WZ_AES,
    ) as zf:
        zf.setpassword(password.encode("utf-8"))
        zf.write(str(input_file), arcname=input_file.name)
    return output


def verify_protected_zip(zip_path: Path, password: str, expected_member: str) -> None:
    if not zip_path.exists():
        raise FileNotFoundError(f"ZIP not found: {zip_path}")
    with pyzipper.AESZipFile(zip_path, "r") as zf:
        zf.setpassword(password.encode("utf-8"))
        members = zf.namelist()
        if expected_member not in members:
            raise RuntimeError("ZIP content mismatch.")
        _ = zf.read(expected_member)
        if zf.testzip() is not None:
            raise RuntimeError("ZIP integrity check failed.")


def cleanup_old_outbox(days: int = 7) -> None:
    ensure_dirs()
    now = datetime.now().timestamp()
    max_age_seconds = days * 24 * 60 * 60
    for path in OUTBOX_DIR.glob("*.zip"):
        try:
            if now - path.stat().st_mtime > max_age_seconds:
                path.unlink(missing_ok=True)
        except Exception:
            pass


def show_splash(root: Tk, on_done) -> None:
    image_path = resource_path(SPLASH_IMAGE_REL)
    if not image_path.exists():
        on_done()
        return
    try:
        img = tk.PhotoImage(file=str(image_path))
    except Exception:
        on_done()
        return

    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.attributes("-topmost", True)
    splash.configure(bg="black")

    label = tk.Label(splash, image=img, bd=0, highlightthickness=0)
    label.pack()

    w = img.width()
    h = img.height()
    x = max((splash.winfo_screenwidth() - w) // 2, 0)
    y = max((splash.winfo_screenheight() - h) // 2, 0)
    splash.geometry(f"{w}x{h}+{x}+{y}")

    splash._img_ref = img  # type: ignore[attr-defined]
    splash.after(SPLASH_MS, lambda: (splash.destroy(), on_done()))


class OutlookMailer:
    def __init__(self) -> None:
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self._account_map: dict[str, object] = {}

    def list_accounts(self) -> list[str]:
        results: list[str] = []
        self._account_map = {}
        for i in range(1, self.ns.Accounts.Count + 1):
            account = self.ns.Accounts.Item(i)
            smtp = getattr(account, "SmtpAddress", None)
            display_name = str(account.DisplayName or "").strip()
            smtp_str = str(smtp or "").strip()
            base_label = smtp_str if smtp_str else display_name if display_name else f"Account {i}"
            label = base_label
            suffix = 2
            while label in self._account_map:
                label = f"{base_label} ({suffix})"
                suffix += 1
            self._account_map[label] = account
            results.append(label)
        return results

    def get_account_by_label(self, account_label: str):
        return self._account_map.get(account_label)

    def _apply_send_account(self, mail, account) -> None:
        # Primary path.
        mail.SendUsingAccount = account
        # Fallback for Outlook builds where direct property set is flaky.
        try:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
        except Exception:
            pass

        smtp = str(getattr(account, "SmtpAddress", "") or "").strip()
        if smtp:
            try:
                mail.SentOnBehalfOfName = smtp
            except Exception:
                pass

    def send_mail(
        self,
        sender_label: str,
        to: str,
        subject: str,
        body: str,
        attachment: Path | None = None,
    ) -> None:
        mail = self.app.CreateItem(0)
        account = self.get_account_by_label(sender_label)
        if account is None:
            raise RuntimeError(f"Sender account not found: {sender_label}")
        self._apply_send_account(mail, account)
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if attachment is not None:
            mail.Attachments.Add(str(attachment))
        mail.Send()


class ResumeMailerApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("760x700")
        self.root.minsize(680, 620)
        self.config = load_config()

        self.resume_var = StringVar()
        self.to_var = StringVar()
        self.sender_var = StringVar()
        self.status_var = StringVar(value="Ready")

        self.mailer: OutlookMailer | None = None
        self.accounts: list[str] = []

        self._build_ui()
        self._init_outlook()
        self._load_defaults()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)
        frame.rowconfigure(8, weight=1)

        ttk.Label(frame, text="履歷檔案").grid(row=0, column=0, sticky=tk.W, pady=4)
        ttk.Entry(frame, textvariable=self.resume_var).grid(row=0, column=1, sticky=tk.EW, pady=4)
        ttk.Button(frame, text="選擇檔案", command=self.on_pick_file).grid(row=0, column=2, padx=6, pady=4)

        ttk.Label(frame, text="收件人 (To)").grid(row=1, column=0, sticky=tk.W, pady=4)
        ttk.Entry(frame, textvariable=self.to_var).grid(row=1, column=1, sticky=tk.EW, pady=4)

        ttk.Label(frame, text="寄件帳號").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.sender_combo = ttk.Combobox(frame, textvariable=self.sender_var, state="readonly")
        self.sender_combo.grid(row=2, column=1, sticky=tk.EW, pady=4)
        ttk.Button(frame, text="重新整理帳號", command=self.refresh_accounts).grid(
            row=2, column=2, padx=6, pady=4
        )

        ttk.Label(frame, text="第一封主旨").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.subject1_entry = ttk.Entry(frame)
        self.subject1_entry.grid(row=3, column=1, columnspan=2, sticky=tk.EW, pady=4)

        ttk.Label(frame, text="第一封內容").grid(row=4, column=0, sticky=tk.NW, pady=4)
        self.body1_text = tk.Text(frame, height=8)
        self.body1_text.grid(row=4, column=1, columnspan=2, sticky=tk.EW, pady=4)

        ttk.Label(frame, text="第二封主旨").grid(row=5, column=0, sticky=tk.W, pady=4)
        self.subject2_entry = ttk.Entry(frame)
        self.subject2_entry.grid(row=5, column=1, columnspan=2, sticky=tk.EW, pady=4)

        ttk.Label(frame, text="第二封內容 ({password})").grid(row=6, column=0, sticky=tk.NW, pady=4)
        self.body2_text = tk.Text(frame, height=8)
        self.body2_text.grid(row=6, column=1, columnspan=2, sticky=tk.EW, pady=4)

        buttons = ttk.Frame(frame)
        buttons.grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=10)
        ttk.Button(buttons, text="寄送兩封信", command=self.on_send).pack(side=tk.LEFT, padx=4)
        ttk.Button(buttons, text="儲存模板設定", command=self.on_save_templates).pack(side=tk.LEFT, padx=4)

        ttk.Label(frame, textvariable=self.status_var).grid(row=8, column=0, columnspan=3, sticky=tk.W, pady=6)

    def _init_outlook(self) -> None:
        try:
            self.mailer = OutlookMailer()
            self.refresh_accounts()
        except Exception as e:
            self.status_var.set("Outlook 初始化失敗")
            messagebox.showerror("Outlook 錯誤", f"無法連線 Outlook。\n{e}")

    def _load_defaults(self) -> None:
        self.subject1_entry.insert(0, self.config.subject_1)
        self.body1_text.insert("1.0", self.config.body_1)
        self.subject2_entry.insert(0, self.config.subject_2)
        self.body2_text.insert("1.0", self.config.body_2)

    def refresh_accounts(self) -> None:
        if self.mailer is None:
            return
        self.accounts = self.mailer.list_accounts()
        self.sender_combo["values"] = self.accounts
        if self.config.default_sender in self.accounts:
            self.sender_var.set(self.config.default_sender)
        elif self.accounts:
            self.sender_var.set(self.accounts[0])

    def on_pick_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="選擇履歷檔案",
            filetypes=[("Document files", "*.pdf *.doc *.docx *.txt"), ("All files", "*.*")],
        )
        if selected:
            self.resume_var.set(selected)

    def on_save_templates(self) -> None:
        self.config.subject_1 = self.subject1_entry.get().strip()
        self.config.body_1 = self.body1_text.get("1.0", END).strip()
        self.config.subject_2 = self.subject2_entry.get().strip()
        self.config.body_2 = self.body2_text.get("1.0", END).strip()
        self.config.default_sender = self.sender_var.get().strip()
        save_config(self.config)
        self.status_var.set("模板與預設寄件帳號已儲存")
        messagebox.showinfo("完成", "設定已儲存。")

    def _validate(self) -> tuple[Path, str, str]:
        if self.mailer is None:
            raise RuntimeError("Outlook 未初始化成功。")
        resume_path = Path(self.resume_var.get().strip())
        recipient = self.to_var.get().strip()
        sender = self.sender_var.get().strip()
        if not resume_path.exists():
            raise ValueError("請先選擇有效的履歷檔案。")
        if not is_valid_email(recipient):
            raise ValueError("收件人格式不合法。")
        if sender not in self.accounts:
            raise ValueError("請選擇有效的寄件帳號。")
        return resume_path, recipient, sender

    def on_send(self) -> None:
        started_at = datetime.now().isoformat(timespec="seconds")
        try:
            resume_path, recipient, sender = self._validate()
            self.status_var.set("處理中：建立加密壓縮檔...")
            self.root.update_idletasks()

            password = generate_password_8_digits()
            zip_path = create_protected_zip(resume_path, password)
            verify_protected_zip(zip_path, password, resume_path.name)

            subject_1 = self.subject1_entry.get().strip()
            body_1 = self.body1_text.get("1.0", END).strip()
            subject_2 = self.subject2_entry.get().strip()
            body_2_template = self.body2_text.get("1.0", END).strip()
            body_2 = body_2_template.replace("{password}", password)

            self.status_var.set("處理中：寄送第一封（附件信）...")
            self.root.update_idletasks()
            self.mailer.send_mail(sender, recipient, subject_1, body_1, zip_path)

            self.status_var.set("處理中：寄送第二封（密碼信）...")
            self.root.update_idletasks()
            try:
                self.mailer.send_mail(sender, recipient, subject_2, body_2, None)
            except Exception as e2:
                append_audit(
                    {
                        "timestamp": started_at,
                        "sender": sender,
                        "recipient": recipient,
                        "status": "mail_2_failed",
                        "error": str(e2),
                    }
                )
                self.status_var.set("第二封寄送失敗")
                retry = messagebox.askyesno("第二封失敗", f"第二封寄送失敗：\n{e2}\n\n是否重送第二封？")
                if retry:
                    self.mailer.send_mail(sender, recipient, subject_2, body_2, None)
                else:
                    return

            self.config.default_sender = sender
            self.config.subject_1 = subject_1
            self.config.body_1 = body_1
            self.config.subject_2 = subject_2
            self.config.body_2 = body_2_template
            save_config(self.config)

            append_audit(
                {
                    "timestamp": started_at,
                    "sender": sender,
                    "recipient": recipient,
                    "status": "success",
                    "error": "",
                }
            )
            self.status_var.set("完成：兩封信已送出")
            messagebox.showinfo("完成", "兩封信已送出。")
        except Exception as e:
            append_audit(
                {
                    "timestamp": started_at,
                    "sender": self.sender_var.get().strip(),
                    "recipient": self.to_var.get().strip(),
                    "status": "failed",
                    "error": str(e),
                }
            )
            self.status_var.set("失敗")
            messagebox.showerror("錯誤", str(e))


def main() -> None:
    ensure_dirs()
    cleanup_old_outbox()

    root = Tk()
    root.withdraw()

    def launch_main() -> None:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        root.deiconify()
        ResumeMailerApp(root)

    show_splash(root, launch_main)
    root.mainloop()


if __name__ == "__main__":
    main()
