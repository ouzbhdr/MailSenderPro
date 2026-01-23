import customtkinter as ctk
from tkinter import filedialog, messagebox
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import threading
import os
import mimetypes

# Configuration
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class BulkMailApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("MailSenderPro - Bulk Email Sender")
        self.geometry("900x700")
        
        self.attachment_paths = []

        # --- GRID LAYOUT CONFIGURATION ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # --- LEFT PANEL (Settings) ---
        self.frame_left = ctk.CTkFrame(self, corner_radius=10)
        self.frame_left.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.lbl_title = ctk.CTkLabel(self.frame_left, text="Configuration", font=("Roboto", 20, "bold"))
        self.lbl_title.pack(pady=20)

        # Gmail Inputs
        self.entry_email = ctk.CTkEntry(self.frame_left, placeholder_text="Your Gmail Address", width=250)
        self.entry_email.pack(pady=10)

        self.entry_password = ctk.CTkEntry(self.frame_left, placeholder_text="App Password", show="*", width=250)
        self.entry_password.pack(pady=10)
        
        self.lbl_info = ctk.CTkLabel(self.frame_left, text="* Use App Password, not login password.", text_color="gray", font=("Arial", 10))
        self.lbl_info.pack(pady=0)

        # Subject
        self.entry_subject = ctk.CTkEntry(self.frame_left, placeholder_text="Email Subject", width=250)
        self.entry_subject.pack(pady=20)

        # Recipient List Input
        self.lbl_recipients = ctk.CTkLabel(self.frame_left, text="Recipients (One per line):", anchor="w")
        self.lbl_recipients.pack(pady=(10, 0), padx=20, fill="x")
        
        self.txt_recipients = ctk.CTkTextbox(self.frame_left, width=250, height=200)
        self.txt_recipients.pack(pady=5)

        # --- RIGHT PANEL (Content) ---
        self.frame_right = ctk.CTkFrame(self, corner_radius=10)
        self.frame_right.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        self.lbl_content = ctk.CTkLabel(self.frame_right, text="Message Content", font=("Roboto", 20, "bold"))
        self.lbl_content.pack(pady=20)

        self.lbl_template_hint = ctk.CTkLabel(self.frame_right, text="Tip: The system automatically adds 'Hello {username}' at the start.", text_color="gray", font=("Arial", 12))
        self.lbl_template_hint.pack(pady=5)

        # Message Body
        self.txt_message = ctk.CTkTextbox(self.frame_right, width=500, height=250)
        self.txt_message.pack(pady=10)

        # Attachments Area
        self.btn_attach = ctk.CTkButton(self.frame_right, text="Attach Files", command=self.add_attachment, fg_color="#D35400", hover_color="#A04000")
        self.btn_attach.pack(pady=10)

        self.lbl_attachments = ctk.CTkLabel(self.frame_right, text="No files selected", text_color="gray")
        self.lbl_attachments.pack(pady=5)

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self.frame_right, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=20)

        # Send Button
        self.btn_send = ctk.CTkButton(self.frame_right, text="SEND EMAILS", command=self.start_sending_thread, height=40, font=("Arial", 14, "bold"), fg_color="green", hover_color="darkgreen")
        self.btn_send.pack(pady=10)

        # Status Log
        self.txt_log = ctk.CTkTextbox(self.frame_right, height=100)
        self.txt_log.pack(pady=10, fill="x", padx=20)
        self.log_message("System Ready...")

    def log_message(self, message):
        self.txt_log.insert("end", message + "\n")
        self.txt_log.see("end")

    def add_attachment(self):
        files = filedialog.askopenfilenames()
        if files:
            for f in files:
                if f not in self.attachment_paths:
                    self.attachment_paths.append(f)
            
            self.lbl_attachments.configure(text=f"{len(self.attachment_paths)} files selected")
            self.log_message(f"Files added: {len(files)}")

    def start_sending_thread(self):
        threading.Thread(target=self.send_emails, daemon=True).start()

    def send_emails(self):
        email_user = self.entry_email.get()
        email_pass = self.entry_password.get()
        subject = self.entry_subject.get()
        recipients = self.txt_recipients.get("1.0", "end").strip().split("\n")
        body_template = self.txt_message.get("1.0", "end")

        if not email_user or not email_pass or not recipients:
            messagebox.showerror("Error", "Please fill in all fields.")
            return

        self.btn_send.configure(state="disabled", text="Sending...")
        self.progress_bar.set(0)
        
        total_emails = len(recipients)
        success_count = 0
        fail_count = 0

        try:
            # Connect to Gmail SMTP Server
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_user, email_pass)
            self.log_message("Connected to SMTP Server successfully.")

            for i, recipient in enumerate(recipients):
                recipient = recipient.strip()
                if not recipient: continue

                try:
                    username = recipient.split("@")[0]
                    
                    msg = MIMEMultipart()
                    msg["From"] = email_user
                    msg["To"] = recipient
                    msg["Subject"] = subject

                    # Personalize the message
                    full_body = f"Hello {username},\n\n{body_template}"
                    msg.attach(MIMEText(full_body, "plain"))

                    # Attach files
                    for filepath in self.attachment_paths:
                        ctype, encoding = mimetypes.guess_type(filepath)
                        if ctype is None or encoding is not None:
                            ctype = "application/octet-stream"
                        
                        maintype, subtype = ctype.split("/", 1)
                        
                        with open(filepath, "rb") as f:
                            part = MIMEBase(maintype, subtype)
                            part.set_payload(f.read())
                        
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(filepath)}")
                        msg.attach(part)

                    server.send_message(msg)
                    success_count += 1
                    self.log_message(f"Sent to: {recipient}")
                
                except Exception as e:
                    fail_count += 1
                    self.log_message(f"Failed ({recipient}): {str(e)}")
                
                # Update Progress
                self.progress_bar.set((i + 1) / total_emails)

            server.quit()
            messagebox.showinfo("Completed", f"Process Finished!\nSuccess: {success_count}\nFailed: {fail_count}")

        except Exception as e:
            messagebox.showerror("Connection Error", str(e))
        
        finally:
            self.btn_send.configure(state="normal", text="SEND EMAILS")

if __name__ == "__main__":
    app = BulkMailApp()
    app.mainloop()
