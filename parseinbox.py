import re
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip  # For copying text to clipboard

class EmailParserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ParseInbox - Focuses on parsing inbox messages for relevant details")

        # Maximize the window on startup
        self.root.state('zoomed')  # Works on Windows
        # For cross-platform compatibility, use:
        # self.root.attributes('-zoomed', True)  # Works on Windows and some Linux environments

        # Configure grid to enforce 50/50 split
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1, uniform="column")  # Left column
        self.root.grid_columnconfigure(1, weight=1, uniform="column")  # Right column

        # Configure styles
        self.style = ttk.Style()
        self.style.configure("TFrame")
        self.style.configure("TLabel", font=("Helvetica", 10))
        self.style.configure("TButton", font=("Helvetica", 10, "bold"), padding=2)
        self.style.configure("Treeview", font=("Helvetica", 10), rowheight=25)
        
        # Configure Treeview Heading with light gray background
        self.style.configure("Treeview.Heading", font=("Helvetica", 11, "bold"), background="lightgray")  # Set background to light gray

        # Left Frame (Input and Parse Button)
        self.left_frame = ttk.Frame(root, padding="12")
        self.left_frame.grid(row=0, column=0, sticky="nsew")

        # Email Input Box
        tk.Label(self.left_frame, text="Paste Email Content:", font=("Helvetica", 11)).pack(pady=(10, 10))
        self.email_text = tk.Text(self.left_frame, height=15, width=60, font=("Helvetica", 10), wrap=tk.WORD)
        self.email_text.pack(fill="both", expand=True, pady=(10, 10))

        # Parse Button
        self.parse_button = ttk.Button(self.left_frame, text="Extract", command=self.parse_email)
        self.parse_button.pack()

        # Right Frame (General Info and Users)
        self.right_frame = ttk.Frame(root, padding="10")
        self.right_frame.grid(row=0, column=1, sticky="nsew")

        # General Information Section
        general_info_frame = ttk.LabelFrame(self.right_frame, text="General Information", padding="10")
        general_info_frame.pack(fill="x", pady=(10, 10))

        # Configure columns for General Information
        general_info_frame.grid_columnconfigure(1, weight=1)  # Allow the middle column to expand

        # Requestor Email
        ttk.Label(general_info_frame, text="Requestor Email ").grid(row=0, column=0, sticky="w", pady=2)
        self.requestor_email_label = ttk.Label(general_info_frame, text="", font=("Helvetica", 10, "bold"))
        self.requestor_email_label.grid(row=0, column=1, sticky="w", pady=2)
        ttk.Button(general_info_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.requestor_email_label.cget("text"))).grid(row=0, column=2, sticky="e", padx=5)

        # Client Name
        ttk.Label(general_info_frame, text="Client Name      ").grid(row=1, column=0, sticky="w", pady=2)
        self.client_name_label = ttk.Label(general_info_frame, text="", font=("Helvetica", 10, "bold"))
        self.client_name_label.grid(row=1, column=1, sticky="w", pady=2)
        ttk.Button(general_info_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.client_name_label.cget("text"))).grid(row=1, column=2, sticky="e", padx=5)

        # Contract Name
        ttk.Label(general_info_frame, text="Contract Name  ").grid(row=2, column=0, sticky="w", pady=2)
        self.contract_name_label = ttk.Label(general_info_frame, text="", font=("Helvetica", 10, "bold"))
        self.contract_name_label.grid(row=2, column=1, sticky="w", pady=2)
        ttk.Button(general_info_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.contract_name_label.cget("text"))).grid(row=2, column=2, sticky="e", padx=5)

        # Contract ID
        ttk.Label(general_info_frame, text="Contract ID       ").grid(row=3, column=0, sticky="w", pady=2)
        self.contract_id_label = ttk.Label(general_info_frame, text="", font=("Helvetica", 10, "bold"))
        self.contract_id_label.grid(row=3, column=1, sticky="w", pady=2)
        ttk.Button(general_info_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.contract_id_label.cget("text"))).grid(row=3, column=2, sticky="e", padx=5)

        # Access Level
        ttk.Label(general_info_frame, text="Access Level     ").grid(row=4, column=0, sticky="w", pady=2)
        self.access_level_label = ttk.Label(general_info_frame, text="", font=("Helvetica", 10, "bold"), wraplength=300, justify="left")
        self.access_level_label.grid(row=4, column=1, sticky="w", pady=2)
        ttk.Button(general_info_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.access_level_label.cget("text"))).grid(row=4, column=2, sticky="e", padx=5)

        # Users Section
        users_frame = ttk.LabelFrame(self.right_frame, text="Users", padding="10")
        users_frame.pack(fill="both", expand=True)

        # Treeview for Users (Table)
        self.user_tree = ttk.Treeview(
            users_frame,
            columns=("First Name", "Last Name", "Email", "Phone"),
            show="headings",
            selectmode="browse",
            height=5  # Set the height of the Treeview to 5 rows
        )
        self.user_tree.heading("First Name", text="First Name")
        self.user_tree.heading("Last Name", text="Last Name")
        self.user_tree.heading("Email", text="Email")
        self.user_tree.heading("Phone", text="Phone")
        self.user_tree.column("First Name", width=120, anchor="w")
        self.user_tree.column("Last Name", width=120, anchor="w")
        self.user_tree.column("Email", width=200, anchor="w")
        self.user_tree.column("Phone", width=120, anchor="w")

        # Add a vertical scrollbar to the Treeview
        scrollbar = ttk.Scrollbar(users_frame, orient="vertical", command=self.user_tree.yview)
        self.user_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.user_tree.pack(fill="both", expand=True)

        # Bind selection event to display user details
        self.user_tree.bind("<<TreeviewSelect>>", self.display_user_details)

        # User Details Section (Separate from Users section)
        user_details_frame = ttk.LabelFrame(self.right_frame, text="User Details", padding="10")
        user_details_frame.pack(fill="x", pady=(10, 0))  # Add padding to separate from the Treeview

        # Configure columns for User Details
        user_details_frame.grid_columnconfigure(1, weight=1)  # Allow the middle column to expand

        # Full Name
        ttk.Label(user_details_frame, text="Full Name  ").grid(row=0, column=0, sticky="w", pady=2)
        self.user_full_name_label = ttk.Label(user_details_frame, text="", font=("Helvetica", 10, "bold"))
        self.user_full_name_label.grid(row=0, column=1, sticky="w", pady=2)
        ttk.Button(user_details_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.user_full_name_label.cget("text"))).grid(row=0, column=2, sticky="e", padx=5)

        # First Name
        ttk.Label(user_details_frame, text="First Name ").grid(row=1, column=0, sticky="w", pady=2)
        self.user_first_name_label = ttk.Label(user_details_frame, text="", font=("Helvetica", 10, "bold"))
        self.user_first_name_label.grid(row=1, column=1, sticky="w", pady=2)
        ttk.Button(user_details_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.user_first_name_label.cget("text"))).grid(row=1, column=2, sticky="e", padx=5)

        # Last Name
        ttk.Label(user_details_frame, text="Last Name  ").grid(row=2, column=0, sticky="w", pady=2)
        self.user_last_name_label = ttk.Label(user_details_frame, text="", font=("Helvetica", 10, "bold"))
        self.user_last_name_label.grid(row=2, column=1, sticky="w", pady=2)
        ttk.Button(user_details_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.user_last_name_label.cget("text"))).grid(row=2, column=2, sticky="e", padx=5)

        # Email
        ttk.Label(user_details_frame, text="Email         ").grid(row=3, column=0, sticky="w", pady=2)
        self.user_email_label = ttk.Label(user_details_frame, text="", font=("Helvetica", 10, "bold"))
        self.user_email_label.grid(row=3, column=1, sticky="w", pady=2)
        ttk.Button(user_details_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.user_email_label.cget("text"))).grid(row=3, column=2, sticky="e", padx=5)

        # Phone
        ttk.Label(user_details_frame, text="Phone        ").grid(row=4, column=0, sticky="w", pady=2)
        self.user_phone_label = ttk.Label(user_details_frame, text="", font=("Helvetica", 10, "bold"))
        self.user_phone_label.grid(row=4, column=1, sticky="w", pady=2)
        ttk.Button(user_details_frame, text="Copy", command=lambda: self.copy_to_clipboard(self.user_phone_label.cget("text"))).grid(row=4, column=2, sticky="e", padx=5)

        # Developer
        tk.Label(self.right_frame, text="msantosuk", font=("Helvetica", 9)).pack(pady=(10, 10), anchor="e", side="right")
        
        self.user_data = []  # Store parsed user details

    def parse_email(self):
        """Parses the email and extracts required data"""
        email_content = self.email_text.get("1.0", tk.END).strip()

        # Extract General Information
        requestor_email = re.search(r"Requestor Email:\s*(.+)", email_content)
        client_name = re.search(r"Client Name\s*:\s*(.+)", email_content)
        contract_name = re.search(r"Contract Name\s*:\s*(.+)", email_content)
        contract_id = re.search(r"Contract ID\s*:\s*(.+)", email_content)
        access_level = re.search(r"Which level does the user need access to\?:\s*(.+)", email_content)

        self.requestor_email_label.config(text=requestor_email.group(1) if requestor_email else "N/A")
        self.client_name_label.config(text=client_name.group(1) if client_name else "N/A")
        self.contract_name_label.config(text=contract_name.group(1) if contract_name else "N/A")
        self.contract_id_label.config(text=contract_id.group(1) if contract_id else "N/A")
        self.access_level_label.config(text=access_level.group(1) if access_level else "N/A")

        # Extract Users
        user_names_section = re.search(r"User Full Name\s*:\s*([\s\S]+?)(?:\nUser Email Address:|$)", email_content)
        user_emails_section = re.search(r"User Email Address\s*:\s*([\s\S]+?)(?:\nUser Phone Number:|$)", email_content)
        user_phones_section = re.search(r"User Phone Number\s*:\s*([\s\S]+?)(?:\nCompany Domain:|$)", email_content)

        if not user_names_section or not user_emails_section:
            messagebox.showerror("Error", "Could not find users in the email!")
            return

        # Split names, emails, and phones into lists
        names = [name.strip() for name in re.split(r"[\n/]+", user_names_section.group(1)) if name.strip()]
        emails = [email.strip() for email in re.split(r"[\n/]+", user_emails_section.group(1)) if email.strip() and "@" in email]  # Filter valid emails

        # Handle phone numbers (optional)
        if user_phones_section:
            phones = [phone.strip() for phone in re.split(r"[\n/]+", user_phones_section.group(1)) if phone.strip()]
        else:
            phones = ["N/A"] * len(names)  # Fill with "N/A" if phone numbers are missing

        # Check if the number of names and emails match
        if len(names) != len(emails):
            messagebox.showerror("Error", "Mismatch in user names and emails count!")
            return

        # Clear old data
        self.user_tree.delete(*self.user_tree.get_children())
        self.user_data = []

        # Add each user as a separate row in the Treeview
        for i in range(len(names)):
            # Split full name into first name and last name
            full_name = names[i]
            name_parts = full_name.split()
            first_name = name_parts[0] if len(name_parts) > 0 else "N/A"
            last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else "N/A"

            # Store user details
            self.user_data.append({
                "full_name": full_name,
                "first_name": first_name,
                "last_name": last_name,
                "email": emails[i],
                "phone": phones[i] if i < len(phones) else "N/A"
            })
            self.user_tree.insert("", "end", values=(first_name, last_name, emails[i], phones[i]))

        # Adjust window size based on content
        self.root.update_idletasks()
        self.root.geometry(f"{self.root.winfo_reqwidth()}x{self.root.winfo_reqheight()}")

    def display_user_details(self, event):
        """Displays the selected user's details"""
        selected_item = self.user_tree.selection()
        if not selected_item:
            return
        
        # Get the index of the selected item in the Treeview
        selected_index = self.user_tree.index(selected_item[0])
        
        # Ensure the index is within the bounds of the user_data list
        if 0 <= selected_index < len(self.user_data):
            user = self.user_data[selected_index]  # Get the selected user's data
            self.user_full_name_label.config(text=f" {user['full_name']}")
            self.user_first_name_label.config(text=f" {user['first_name']}")
            self.user_last_name_label.config(text=f" {user['last_name']}")
            self.user_email_label.config(text=f" {user['email']}")
            self.user_phone_label.config(text=f" {user['phone']}")

    def copy_to_clipboard(self, text):
        """Copies the given text to the clipboard"""
        try:
            pyperclip.copy(text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy text: {e}")

# Run the Tkinter application
if __name__ == "__main__":
    root = tk.Tk()
    app = EmailParserApp(root)
    root.mainloop()