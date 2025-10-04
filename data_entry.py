import tkinter as tk
from tkinter import messagebox
import pandas as pd
from pathlib import Path
import re
import os

# File for saving submissions
FILE = Path("submissions.xlsx")

# --------------------------
# Validation Functions
# --------------------------
def validate_email(email):
    """Validate email format."""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def validate_phone(phone):
    """Validate phone number (allows various formats)."""
    if not phone:
        return False
    cleaned = re.sub(r'[\s\-$$$$\.]', '', phone)
    return cleaned.isdigit() and 10 <= len(cleaned) <= 15

def validate_name(name):
    """Validate name (letters, spaces, hyphens only)."""
    if not name or len(name) < 2:
        return False
    return bool(re.match(r'^[a-zA-Z\s\-]+$', name))

def check_duplicate_email(email):
    """Check if email already exists in database."""
    if not FILE.exists():
        return False
    try:
        df = pd.read_excel(FILE)
        return email.lower() in df['email'].str.lower().values
    except Exception as e:
        print(f"Error checking duplicates: {e}")
        return False

# --------------------------
# Excel handling
# --------------------------
def ensure_file():
    """Create Excel file with headers if not exists."""
    if not FILE.exists():
        try:
            df = pd.DataFrame(columns=["name", "email", "phone", "address", "notes"])
            df.to_excel(FILE, index=False, engine='openpyxl')
            print(f"✓ File created at: {FILE.absolute()}")
        except ImportError:
            messagebox.showerror(
                "Missing Library",
                "Please install openpyxl:\n\npip install openpyxl\n\nThen restart the application."
            )
            raise
        except Exception as e:
            messagebox.showerror("Error", f"Could not create file:\n{str(e)}")
            raise

def save_submission():
    """Save user input to Excel file with validation."""
    name = e_name.get().strip()
    email = e_email.get().strip()
    phone = e_phone.get().strip()
    address = txt_address.get("1.0", "end").strip()
    notes = txt_notes.get("1.0", "end").strip()
    
    # Validation checks
    errors = []
    
    if not name:
        errors.append("• Name is required")
    elif not validate_name(name):
        errors.append("• Name must be at least 2 characters and contain only letters")
    elif len(name) > 100:
        errors.append("• Name must be less than 100 characters")
    
    if not email:
        errors.append("• Email is required")
    elif not validate_email(email):
        errors.append("• Email format is invalid (example: user@domain.com)")
    elif check_duplicate_email(email):
        errors.append("• This email is already registered")
    
    if not phone:
        errors.append("• Phone number is required")
    elif not validate_phone(phone):
        errors.append("• Phone number must be 10-15 digits")
    
    if not address:
        errors.append("• Address is required")
    elif len(address) < 10:
        errors.append("• Address must be at least 10 characters")
    elif len(address) > 500:
        errors.append("• Address must be less than 500 characters")
    
    if len(notes) > 1000:
        errors.append("• Notes must be less than 1000 characters")
    
    if errors:
        messagebox.showerror("Validation Error", "\n".join(errors))
        return
    
    # Save to Excel
    try:
        row = {
            "name": name.title(),
            "email": email.lower(),
            "phone": phone,
            "address": address,
            "notes": notes
        }
        df = pd.read_excel(FILE, engine='openpyxl')
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        df.to_excel(FILE, index=False, engine='openpyxl')
        
        file_location = FILE.absolute()
        messagebox.showinfo(
            "Success", 
            f"Submission saved successfully!\n\nFile location:\n{file_location}"
        )
        clear_fields()
    except ImportError:
        messagebox.showerror(
            "Missing Library",
            "Please install openpyxl:\n\npip install openpyxl"
        )
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

def clear_fields():
    """Clear all form fields."""
    e_name.delete(0, tk.END)
    e_email.delete(0, tk.END)
    e_phone.delete(0, tk.END)
    txt_address.delete("1.0", "end")
    txt_notes.delete("1.0", "end")

def show_file_location():
    """Show where the Excel file is saved."""
    location = FILE.absolute()
    exists = "EXISTS" if FILE.exists() else "NOT FOUND"
    messagebox.showinfo(
        "File Location",
        f"File: {location}\n\nStatus: {exists}"
    )

# --------------------------
# GUI Setup
# --------------------------
try:
    ensure_file()
except:
    pass  # Error already shown

root = tk.Tk()
root.title("Data Entry Form")
root.geometry("500x450")
root.resizable(False, False)

padding = {'padx': 10, 'pady': 5}

tk.Label(root, text="Name *", font=('Arial', 10)).grid(row=0, column=0, sticky="w", **padding)
e_name = tk.Entry(root, width=40, font=('Arial', 10))
e_name.grid(row=0, column=1, **padding)

tk.Label(root, text="Email *", font=('Arial', 10)).grid(row=1, column=0, sticky="w", **padding)
e_email = tk.Entry(root, width=40, font=('Arial', 10))
e_email.grid(row=1, column=1, **padding)

tk.Label(root, text="Phone *", font=('Arial', 10)).grid(row=2, column=0, sticky="w", **padding)
e_phone = tk.Entry(root, width=40, font=('Arial', 10))
e_phone.grid(row=2, column=1, **padding)

tk.Label(root, text="Address *", font=('Arial', 10)).grid(row=3, column=0, sticky="nw", **padding)
txt_address = tk.Text(root, height=4, width=30, font=('Arial', 10))
txt_address.grid(row=3, column=1, **padding)

tk.Label(root, text="Notes", font=('Arial', 10)).grid(row=4, column=0, sticky="nw", **padding)
txt_notes = tk.Text(root, height=4, width=30, font=('Arial', 10))
txt_notes.grid(row=4, column=1, **padding)

# Buttons
btn_frame = tk.Frame(root)
btn_frame.grid(row=5, column=0, columnspan=2, pady=15)

tk.Button(btn_frame, text="Save", command=save_submission, width=12, 
          bg="#4CAF50", fg="white", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
tk.Button(btn_frame, text="Clear", command=clear_fields, width=12,
          bg="#f44336", fg="white", font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)

# File location button
tk.Button(root, text="Show File Location", command=show_file_location, 
          font=('Arial', 9)).grid(row=6, column=0, columnspan=2, pady=5)

# Required fields note
tk.Label(root, text="* Required fields", font=('Arial', 8), fg="red").grid(
    row=7, column=0, columnspan=2, sticky="w", padx=10)

root.mainloop()