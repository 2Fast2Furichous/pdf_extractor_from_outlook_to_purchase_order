"""
PDF Extractor from Outlook Emails
Standalone application to extract PO data from PDFs in Outlook emails
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import os
import csv
import tempfile
import re
import sys
import json
import hashlib

try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    from ttkbootstrap.scrolled import ScrolledText
except ImportError as e:
    print("ERROR: ttkbootstrap not installed. Run: pip install ttkbootstrap")
    print(f"Details: {e}")
    sys.exit(1)

try:
    import win32com.client
except ImportError as e:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    print(f"Details: {e}")
    sys.exit(1)

try:
    import pdfplumber
except ImportError as e:
    print("ERROR: pdfplumber not installed. Run: pip install pdfplumber")
    print(f"Details: {e}")
    sys.exit(1)

try:
    import pandas as pd
except ImportError as e:
    print("ERROR: pandas not installed. Run: pip install pandas openpyxl")
    print(f"Details: {e}")
    sys.exit(1)


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Data Extractor - Outlook")
        self.root.geometry("700x600")

        # Set window icon
        self.set_window_icon()

        # Variables
        self.email_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.output_path_var = tk.StringVar()

        # Set default output path (Excel file now)
        default_output = os.path.join(os.path.expanduser("."), "PO_Data.xlsx")
        self.output_path_var.set(default_output)

        # Settings file path
        self.settings_file = os.path.join(os.path.expanduser("~"), ".pdf_extractor_settings.json")

        # Load saved settings
        self.load_settings()

        self.create_widgets()

    def set_window_icon(self):
        """Set the window icon for the application"""
        try:
            # Determine the base path (works for both script and frozen exe)
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                base_path = sys._MEIPASS
            else:
                # Running as script
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

            icon_path = os.path.join(base_path, 'icon.ico')

            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                # Fallback: try current directory
                icon_path = os.path.join(os.getcwd(), 'icon.ico')
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
        except Exception as e:
            # Silently fail - don't interrupt user if icon can't be loaded
            pass

    def create_widgets(self):
        # Title
        title_label = ttk.Label(self.root, text="PDF Data Extractor",
                               font=("Segoe UI", 18, "bold"))
        title_label.pack(pady=15)

        # Input Frame
        input_frame = ttk.Labelframe(self.root, text="Email Filter Criteria", padding=15)
        input_frame.pack(fill="x", padx=15, pady=10)

        # Email Address
        ttk.Label(input_frame, text="Email Address:").grid(row=0, column=0, sticky="w", pady=8, padx=5)
        ttk.Entry(input_frame, textvariable=self.email_var, width=50).grid(row=0, column=1, pady=8, padx=5)

        # Folder Contains
        ttk.Label(input_frame, text="Folder Contains:").grid(row=1, column=0, sticky="w", pady=8, padx=5)
        ttk.Entry(input_frame, textvariable=self.folder_var, width=50).grid(row=1, column=1, pady=8, padx=5)

        # Subject Contains
        ttk.Label(input_frame, text="Subject Contains:").grid(row=2, column=0, sticky="w", pady=8, padx=5)
        ttk.Entry(input_frame, textvariable=self.subject_var, width=50).grid(row=2, column=1, pady=8, padx=5)

        # Start Date
        ttk.Label(input_frame, text="Start Date (MM/DD/YYYY):").grid(row=3, column=0, sticky="w", pady=8, padx=5)
        ttk.Entry(input_frame, textvariable=self.start_date_var, width=50).grid(row=3, column=1, pady=8, padx=5)

        # Output Frame
        output_frame = ttk.Labelframe(self.root, text="Output Settings", padding=15)
        output_frame.pack(fill="x", padx=15, pady=10)

        ttk.Label(output_frame, text="Output Excel File:").grid(row=0, column=0, sticky="w", pady=8, padx=5)
        ttk.Entry(output_frame, textvariable=self.output_path_var, width=40).grid(row=0, column=1, pady=8, padx=5)
        ttk.Button(output_frame, text="Browse...", command=self.browse_output, bootstyle="secondary").grid(row=0, column=2, padx=5)

        # Extract Button
        extract_btn = ttk.Button(self.root, text="Extract PDFs from Outlook",
                                 command=self.extract_pdfs, bootstyle="success", width=25)
        extract_btn.pack(pady=15)

        # Progress Frame
        progress_frame = ttk.Labelframe(self.root, text="Progress", padding=15)
        progress_frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.progress_text = ScrolledText(progress_frame, height=15, width=80, autohide=True)
        self.progress_text.pack(fill="both", expand=True)

        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, bootstyle="inverse-secondary")
        status_bar.pack(fill="x", side=tk.BOTTOM, pady=5)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="PO_Data.xlsx"
        )
        if filename:
            self.output_path_var.set(filename)

    def log(self, message):
        """Add message to progress text box"""
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.see(tk.END)
        self.root.update_idletasks()

    def extract_pdfs(self):
        """Main extraction logic"""
        self.progress_text.delete(1.0, tk.END)
        self.log("Starting PDF extraction...")

        try:
            # Validate inputs
            email_addr = self.email_var.get().strip()
            folder_text = self.folder_var.get().strip()
            subject_text = self.subject_var.get().strip()
            start_date_str = self.start_date_var.get().strip()
            output_path = self.output_path_var.get().strip()

            # Save settings for next run
            self.save_settings()

            # Parse start date
            start_date = None
            if start_date_str:
                try:
                    start_date = datetime.strptime(start_date_str, "%m/%d/%Y")
                    self.log(f"Filter: Emails after {start_date.strftime('%Y-%m-%d')}")
                except ValueError:
                    messagebox.showerror("Error", "Invalid date format. Use MM/DD/YYYY")
                    return

            # Connect to Outlook
            self.log("Connecting to Outlook...")
            self.status_var.set("Connecting to Outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

            # Find matching folder
            self.log(f"Searching for folder containing: '{folder_text}'")
            target_folder = self.find_folder(outlook, email_addr, folder_text)

            if not target_folder:
                messagebox.showerror("Error", f"Could not find folder containing '{folder_text}' in {email_addr}")
                return

            self.log(f"Found folder: {target_folder.Name}")

            # Filter emails
            self.log(f"Filtering emails with subject containing: '{subject_text}'")
            emails = self.filter_emails(target_folder, subject_text, start_date)
            self.log(f"Found {len(emails)} matching emails")

            if not emails:
                messagebox.showwarning("No Emails", "No emails match the filter criteria")
                return

            # Setup PDF storage location
            output_dir = os.path.dirname(output_path)
            pdf_base_folder = os.path.join(output_dir, "PDFs")

            # Process PDFs with deduplication
            all_data = []
            pdf_count = 0
            processed_pdfs = set()  # Track unique PDFs by hash

            for idx, email in enumerate(emails, 1):
                self.log(f"\n[{idx}/{len(emails)}] Processing: {email.Subject}")
                self.status_var.set(f"Processing email {idx}/{len(emails)}...")

                # Get email date for folder organization
                email_date = email.ReceivedTime
                date_folder = datetime(email_date.year, email_date.month, email_date.day).strftime("%Y-%m-%d")
                pdf_save_folder = os.path.join(pdf_base_folder, date_folder)

                # Create folder if it doesn't exist
                os.makedirs(pdf_save_folder, exist_ok=True)

                # Extract PDFs from attachments
                for attachment in email.Attachments:
                    if attachment.FileName.lower().endswith('.pdf'):
                        # Save to temp file first to calculate hash
                        temp_pdf = os.path.join(tempfile.gettempdir(), attachment.FileName)
                        attachment.SaveAsFile(temp_pdf)

                        # Calculate PDF hash for deduplication
                        try:
                            with open(temp_pdf, 'rb') as f:
                                pdf_hash = hashlib.md5(f.read()).hexdigest()

                            # Check if already processed
                            if pdf_hash in processed_pdfs:
                                self.log(f"  Skipping duplicate: {attachment.FileName}")
                                try:
                                    os.remove(temp_pdf)
                                except:
                                    pass
                                continue

                            # Mark as processed
                            processed_pdfs.add(pdf_hash)
                            pdf_count += 1
                            self.log(f"  Found PDF: {attachment.FileName}")

                            # Parse PDF
                            data = self.parse_pdf(temp_pdf, attachment.FileName)
                            all_data.extend(data)

                            # Save PDF to permanent location with retry on permission error
                            permanent_pdf_path = os.path.join(pdf_save_folder, attachment.FileName)
                            for pdf_attempt in range(3):  # Try up to 3 times for PDF saves
                                try:
                                    import shutil
                                    shutil.copy2(temp_pdf, permanent_pdf_path)
                                    self.log(f"  Saved to: {date_folder}/{attachment.FileName}")
                                    break
                                except PermissionError:
                                    if pdf_attempt < 2:
                                        retry = messagebox.askretrycancel(
                                            "PDF Save Error",
                                            f"Cannot save PDF:\n{attachment.FileName}\n\n"
                                            "The file may be open or the folder may be locked.\n"
                                            "Please close it and click 'Retry'."
                                        )
                                        if not retry:
                                            self.log(f"  User cancelled PDF save for: {attachment.FileName}")
                                            break
                                    else:
                                        self.log(f"  ERROR: Could not save PDF after 3 attempts: {attachment.FileName}")
                                        break
                                except Exception as e:
                                    self.log(f"  Warning: Could not save PDF: {e}")
                                    break

                        except Exception as e:
                            self.log(f"  Error processing {attachment.FileName}: {e}")

                        # Clean up temp file
                        try:
                            os.remove(temp_pdf)
                        except:
                            pass

            # Write to file (Excel or CSV) with retry on permission error
            self.log(f"\nWriting {len(all_data)} line items to file...")
            max_retries = 5
            for attempt in range(max_retries):
                try:
                    self.write_output(output_path, all_data)
                    break  # Success, exit retry loop
                except PermissionError:
                    if attempt < max_retries - 1:
                        # Ask user to retry
                        retry = messagebox.askretrycancel(
                            "File Access Error",
                            f"Cannot write to file:\n{output_path}\n\n"
                            "The file may be open in Excel or another program.\n"
                            "Please close the file and click 'Retry'."
                        )
                        if not retry:
                            self.log("User cancelled file write operation")
                            messagebox.showinfo("Cancelled", "Data extraction completed but file was not saved.")
                            return
                        else:
                            self.log(f"Retrying file write (attempt {attempt + 2}/{max_retries})...")
                    else:
                        # Max retries reached
                        self.log("ERROR: Max retries reached, file could not be written")
                        messagebox.showerror("Error", "Could not write file after multiple attempts.")
                        return

            self.log(f"\n{'='*60}")
            self.log(f"SUCCESS! Extracted {pdf_count} PDFs with {len(all_data)} line items")
            self.log(f"Output saved to: {output_path}")
            self.log(f"PDFs saved to: {pdf_base_folder}")
            self.log(f"{'='*60}")

            self.status_var.set(f"Complete! {len(all_data)} items extracted")

            # Ask to open file
            if messagebox.askyesno("Success", f"Extracted {len(all_data)} line items.\n\nOpen output file?"):
                os.startfile(output_path)

        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            self.status_var.set("Error occurred")

    def find_folder(self, outlook, email_addr, folder_text):
        """Find folder in Outlook"""
        try:
            # Try to get account by email
            for account in outlook.Folders:
                if email_addr.lower() in account.Name.lower():
                    return self.search_subfolder(account, folder_text)

            # If not found, search all folders
            for account in outlook.Folders:
                folder = self.search_subfolder(account, folder_text)
                if folder:
                    return folder

            return None
        except Exception as e:
            self.log(f"Error finding folder: {e}")
            return None

    def search_subfolder(self, parent_folder, search_text):
        """Recursively search for subfolder"""
        try:
            if search_text.lower() in parent_folder.Name.lower():
                return parent_folder

            for subfolder in parent_folder.Folders:
                result = self.search_subfolder(subfolder, search_text)
                if result:
                    return result

            return None
        except:
            return None

    def filter_emails(self, folder, subject_text, start_date):
        """Filter emails by criteria"""
        matching_emails = []

        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # Sort by received time, descending

            for item in items:
                try:
                    # Check if it's a mail item
                    if item.Class != 43:  # 43 = olMail
                        continue

                    # Check subject
                    if subject_text and subject_text.lower() not in item.Subject.lower():
                        continue

                    # Check date
                    if start_date:
                        try:
                            received_date = item.ReceivedTime
                            # Compare just the date parts (year, month, day) to avoid type issues
                            received_date_only = datetime(received_date.year, received_date.month, received_date.day)

                            if received_date_only < start_date:
                                continue
                        except Exception as date_err:
                            # If date conversion fails, log and skip date check for this email
                            self.log(f"    Warning: Date conversion failed: {date_err}")
                            pass

                    matching_emails.append(item)
                except:
                    continue

        except Exception as e:
            self.log(f"Error filtering emails: {e}")

        return matching_emails

    def parse_pdf(self, pdf_path, pdf_name):
        """Parse PDF using pdfplumber - much more reliable than Word"""
        data = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Extract text from all pages
                full_text = ""
                for page in pdf.pages:
                    full_text += page.extract_text() + "\n"

                # Check for tables
                has_tables = False
                for page in pdf.pages:
                    if page.extract_tables():
                        has_tables = True
                        break

                self.log(f"    Tables detected: {has_tables}")

                # Use table-based parsing if available
                if has_tables:
                    data = self.parse_pdf_tables(pdf, pdf_name)
                else:
                    data = self.parse_pdf_text(pdf, pdf_name)

                self.log(f"    Extracted {len(data)} line items")

        except Exception as e:
            self.log(f"    ERROR parsing PDF: {e}")

        return data

    def parse_pdf_tables(self, pdf, pdf_name):
        """Parse PDF using table extraction - handles multi-page line items"""
        data = []

        try:
            # Extract order info from text and coordinates
            full_text = ""
            for page in pdf.pages:
                full_text += page.extract_text() + "\n"

            order_number = self.extract_order_number(full_text)
            order_date = self.extract_order_date(full_text)

            # Use first page for coordinate-based extraction of addresses
            first_page = pdf.pages[0] if pdf.pages else None
            ship_to = self.extract_ship_to_coordinates(first_page) if first_page else ""
            ordering_office = self.extract_ordering_office_coordinates(first_page) if first_page else ""

            # Track column mappings for continuation pages
            col_line = 0  # Line is always first
            col_part = None
            col_date = None
            col_qty = None
            col_unit_price = None
            col_amount = None
            found_header = False

            # Look for line items table - check all pages
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()

                # Check all tables on this page
                for table_idx, table in enumerate(tables):
                    if not table or len(table) < 1:
                        continue

                    # Try to detect if this is a line items table
                    first_row = table[0]
                    first_cell = str(first_row[0]).strip() if first_row and first_row[0] else ""

                    # Check if first row is a header (contains "Line" and "Part")
                    is_header = False
                    if len(table) >= 2:
                        header = ' '.join([str(cell) for cell in first_row if cell])
                        if 'Line' in header and 'Part' in header:
                            is_header = True
                            found_header = True

                            # Parse column indices from header
                            col_line = 0  # Line is always first
                            for idx, cell in enumerate(first_row):
                                if cell:
                                    cell_lower = str(cell).lower()
                                    if 'part' in cell_lower:
                                        col_part = idx
                                    elif 'delivery' in cell_lower or 'date' in cell_lower:
                                        col_date = idx
                                    elif 'quantity' in cell_lower:
                                        col_qty = idx
                                    elif 'unit price' in cell_lower or 'price' in cell_lower:
                                        col_unit_price = idx
                                    elif 'amount' in cell_lower or 'total' in cell_lower:
                                        col_amount = idx

                    # Check if this is a continuation table (starts with line number like "4.1")
                    is_continuation = False
                    if not is_header and found_header and re.match(r'^\d+\.\d+$', first_cell):
                        is_continuation = True

                        # For continuation tables, detect columns from first data row
                        # because page breaks can shift column positions
                        if len(first_row) > 1:
                            # Reset price/amount for this table's detection
                            table_unit_price = None
                            table_amount = None

                            # Examine first row to find column positions
                            for idx, cell in enumerate(first_row):
                                cell_str = str(cell).strip() if cell else ""
                                # Date column - look for date pattern
                                if re.search(r'\d{1,2}-[A-Za-z]{3}-\d{4}', cell_str):
                                    col_date = idx
                                # Quantity - numeric value (but not price with decimal)
                                elif cell_str.isdigit():
                                    col_qty = idx
                                # UOM - typically "Each"
                                elif cell_str.lower() == 'each':
                                    pass  # We don't need UOM column index
                                # Unit Price - has decimal point, larger than quantity
                                elif re.match(r'^\d+\.\d{2,}$', cell_str):
                                    if not table_unit_price:
                                        col_unit_price = idx
                                        table_unit_price = idx
                                    else:
                                        col_amount = idx
                                        table_amount = idx

                    # Skip if this is neither a header table nor a continuation
                    if not is_header and not is_continuation:
                        continue

                    # Process data rows (skip header if present)
                    start_row = 1 if is_header else 0
                    for row in table[start_row:]:
                        if not row:
                            continue

                        line_num = str(row[col_line]).strip() if row[col_line] else ""

                        # Check if valid line number (e.g., "1.1", "2.1", "4.1")
                        if re.match(r'^\d+\.\d+$', line_num):
                            part_num = str(row[col_part]).strip() if col_part and len(row) > col_part and row[col_part] else ""
                            delivery_date = str(row[col_date]).strip() if col_date and len(row) > col_date and row[col_date] else ""
                            quantity = str(row[col_qty]).strip() if col_qty and len(row) > col_qty and row[col_qty] else ""
                            unit_price = str(row[col_unit_price]).strip() if col_unit_price and len(row) > col_unit_price and row[col_unit_price] else ""
                            amount = str(row[col_amount]).strip() if col_amount and len(row) > col_amount and row[col_amount] else ""

                            # Clean part number (remove /REV: and newlines)
                            if '/' in part_num:
                                part_num = part_num.split('/')[0].strip()
                            part_num = part_num.replace('\n', ' ').strip()

                            data.append({
                                'pdf_file': pdf_name,
                                'order_number': order_number,
                                'order_date': order_date,
                                'line': line_num,
                                'part_number': part_num,
                                'quantity': quantity,
                                'unit_price': unit_price,
                                'amount': amount,
                                'delivery_date': delivery_date,
                                'ship_to': ship_to,
                                'ordering_office': ordering_office
                            })

        except Exception as e:
            self.log(f"    Error in table parsing: {e}")

        return data

    def parse_pdf_text(self, pdf, pdf_name):
        """Parse PDF using text extraction (for vertical format)"""
        data = []

        try:
            # Extract full text
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

            lines = text.split('\n')

            order_number = self.extract_order_number(text)
            order_date = self.extract_order_date(text)

            # Use coordinate-based extraction for addresses
            first_page = pdf.pages[0] if pdf.pages else None
            ship_to = self.extract_ship_to_coordinates(first_page) if first_page else ""
            ordering_office = self.extract_ordering_office_coordinates(first_page) if first_page else ""

            # Find line items section
            in_line_items = False
            i = 0

            while i < len(lines):
                line = lines[i].strip()

                # Detect start of line items
                if not in_line_items and 'line' in line.lower() and len(line) < 50:
                    in_line_items = True
                    i += 1
                    continue

                if in_line_items:
                    # Check for line number (e.g., "1.1", "2.1")
                    if re.match(r'^\d+\.\d+$', line):
                        line_num = line
                        part_num = ""
                        delivery_date = ""
                        quantity = ""
                        unit_price = ""
                        amount = ""

                        # Get part number from next line
                        if i + 1 < len(lines):
                            part_line = lines[i + 1].strip()
                            if '/' in part_line:
                                part_num = part_line.split('/')[0].strip()
                            else:
                                part_num = part_line

                        # Look for delivery date, quantity, price, and amount in next 10 lines
                        for j in range(i + 1, min(i + 11, len(lines))):
                            check_line = lines[j].strip()

                            # Check for date pattern
                            if not delivery_date:
                                date_match = re.search(r'\d{1,2}-[A-Za-z]{3}-\d{4}', check_line)
                                if date_match:
                                    delivery_date = date_match.group()

                            # Check for quantity (numeric, after date)
                            if delivery_date and not quantity:
                                if check_line.isdigit() and len(check_line) < 10:
                                    quantity = check_line

                            # Check for price patterns (e.g., "$12.34" or "12.34")
                            if not unit_price:
                                price_match = re.search(r'\$?\d+\.\d{2}', check_line)
                                if price_match and not amount:
                                    unit_price = price_match.group()
                                elif price_match:
                                    amount = price_match.group()

                        if part_num:
                            data.append({
                                'pdf_file': pdf_name,
                                'order_number': order_number,
                                'order_date': order_date,
                                'line': line_num,
                                'part_number': part_num,
                                'quantity': quantity,
                                'unit_price': unit_price,
                                'amount': amount,
                                'delivery_date': delivery_date,
                                'ship_to': ship_to,
                                'ordering_office': ordering_office
                            })

                i += 1

        except Exception as e:
            self.log(f"    Error in text parsing: {e}")

        return data

    def extract_order_number(self, text):
        """Extract 10-digit order number"""
        match = re.search(r'\b(\d{10})\b', text)
        return match.group(1) if match else ""

    def extract_order_date(self, text):
        """Extract order date"""
        match = re.search(r'\d{1,2}-[A-Za-z]{3}-\d{4}', text)
        return match.group() if match else ""

    def extract_ship_to_coordinates(self, page):
        """Extract ship to address using pdfplumber coordinate-based extraction"""
        try:
            words = page.extract_words()

            # Find "Ship To Address" label position
            ship_label_y = None
            for word in words:
                if 'Ship' in word['text'] and word.get('top', 0) > 100:  # Skip header
                    # Look for words near this one to confirm "Ship To Address"
                    ship_label_y = word['top']
                    break

            if not ship_label_y:
                return ""

            # Find "Payment Terms" label to know where to stop
            payment_terms_y = None
            for word in words:
                if 'Payment' in word['text'] and word['top'] > ship_label_y:
                    payment_terms_y = word['top']
                    break

            # Extract words in LEFT column (x < 300) between ship_label and payment_terms
            ship_words = []
            for word in words:
                word_y = word['top']
                word_x = word['x0']

                # Must be below the "Ship To Address" label and above "Payment Terms"
                if word_y > ship_label_y + 10:  # Start below the label
                    if payment_terms_y and word_y >= payment_terms_y:
                        break
                    # Left column only (ship-to, not invoice)
                    if word_x < 300:  # Adjust threshold as needed
                        ship_words.append((word_y, word_x, word['text']))

            # Sort by y-coordinate first, then x-coordinate (to get correct reading order)
            ship_words.sort(key=lambda w: (w[0], w[1]))

            # Group words by line (same Y coordinate within tolerance)
            lines = []
            current_line = []
            last_y = None

            for word_y, word_x, text in ship_words:
                if last_y is None or abs(word_y - last_y) < 2:  # Same line
                    current_line.append(text)
                    last_y = word_y
                else:  # New line
                    if current_line:
                        lines.append(' '.join(current_line))
                    current_line = [text]
                    last_y = word_y

            if current_line:  # Don't forget the last line
                lines.append(' '.join(current_line))

            # Join all lines and clean up
            address = ', '.join(lines[:5])  # Limit to first 5 lines
            address = address.replace(', ,', ',').strip(' ,')
            return address[:300] if address else ""

        except Exception as e:
            self.log(f"Error in coordinate extraction for ship_to: {e}")
            return ""

    def extract_ordering_office_coordinates(self, page):
        """Extract ordering office using pdfplumber coordinate-based extraction"""
        try:
            words = page.extract_words()

            # Find "Ordering Office" label position
            ordering_label_y = None
            for word in words:
                if 'Ordering' in word['text']:
                    ordering_label_y = word['top']
                    break

            if not ordering_label_y:
                return ""

            # Find "Supplier Contact" or "Buyer" to know where to stop
            stop_y = None
            for word in words:
                if word['top'] > ordering_label_y:
                    if 'Supplier' in word['text'] and 'Contact' in [w['text'] for w in words if abs(w['top'] - word['top']) < 5]:
                        stop_y = word['top']
                        break
                    if 'Buyer' in word['text']:
                        stop_y = word['top']
                        break

            # Extract words in RIGHT column (x > 300) between ordering_label and stop
            ordering_words = []
            for word in words:
                word_y = word['top']
                word_x = word['x0']

                # Must be below the "Ordering Office" label
                if word_y > ordering_label_y + 10:
                    if stop_y and word_y >= stop_y:
                        break
                    # Right column only (x > 300)
                    if word_x > 300:
                        ordering_words.append((word_y, word_x, word['text']))

            # Sort by y-coordinate first, then x-coordinate (to get correct reading order)
            ordering_words.sort(key=lambda w: (w[0], w[1]))

            # Group words by line (same Y coordinate within tolerance)
            lines = []
            current_line = []
            last_y = None

            for word_y, word_x, text in ordering_words:
                if last_y is None or abs(word_y - last_y) < 2:  # Same line
                    current_line.append(text)
                    last_y = word_y
                else:  # New line
                    if current_line:
                        lines.append(' '.join(current_line))
                    current_line = [text]
                    last_y = word_y

            if current_line:  # Don't forget the last line
                lines.append(' '.join(current_line))

            # Join all lines and clean up
            office = ', '.join(lines[:6])  # Limit to first 6 lines
            office = office.replace(', ,', ',').strip(' ,')
            return office[:300] if office else ""

        except Exception as e:
            self.log(f"Error in coordinate extraction for ordering_office: {e}")
            return ""

    def write_output(self, output_path, data):
        """Write data to Excel or CSV file (append mode if exists)"""
        try:
            if not data:
                self.log("Warning: No data to write")
                return

            fieldnames = ['pdf_file', 'order_number', 'order_date', 'line',
                         'part_number', 'quantity', 'unit_price', 'amount',
                         'delivery_date', 'ship_to', 'ordering_office']

            # Create DataFrame from new data
            new_df = pd.DataFrame(data, columns=fieldnames)

            # Rename columns to Title Case for better readability
            new_df.columns = ['PDF File', 'Order Number', 'Order Date', 'Line',
                             'Part Number', 'Quantity', 'Unit Price', 'Amount',
                             'Delivery Date', 'Ship To', 'Ordering Office']

            # Convert numeric columns to proper numeric types
            numeric_columns = ['Quantity', 'Unit Price', 'Amount']
            for col in numeric_columns:
                if col in new_df.columns:
                    # Remove commas and convert to numeric
                    new_df[col] = new_df[col].astype(str).str.replace(',', '')
                    new_df[col] = pd.to_numeric(new_df[col], errors='coerce')

            # Convert identifier columns to numeric to prevent Excel warnings
            # Order Number: convert to integer
            if 'Order Number' in new_df.columns:
                new_df['Order Number'] = pd.to_numeric(new_df['Order Number'], errors='coerce').astype('Int64')

            # Line: convert to float (e.g., 1.1, 2.1)
            if 'Line' in new_df.columns:
                new_df['Line'] = pd.to_numeric(new_df['Line'], errors='coerce')

            # Check if file exists and append if it does
            if os.path.exists(output_path):
                self.log(f"File exists, attempting to append to: {output_path}")

                # Read existing file
                if output_path.endswith('.xlsx'):
                    existing_df = pd.read_excel(output_path)
                else:
                    existing_df = pd.read_csv(output_path)

                self.log(f"Successfully read existing file with {len(existing_df)} rows")

                # Combine old and new data
                combined_df = pd.concat([existing_df, new_df], ignore_index=True)

                # Remove duplicates based on key columns (PDF file, order number, and line number)
                # This prevents removing legitimate duplicate part numbers across different orders
                before_dedup = len(combined_df)
                combined_df = combined_df.drop_duplicates(subset=['PDF File', 'Order Number', 'Line'], keep='last')
                after_dedup = len(combined_df)

                duplicates_removed = before_dedup - after_dedup
                if duplicates_removed > 0:
                    self.log(f"Removed {duplicates_removed} duplicate line items")

                self.log(f"Appended {len(new_df)} new rows, total rows: {after_dedup}")
            else:
                self.log(f"Creating new file: {output_path}")
                combined_df = new_df

            # Write to file
            if output_path.endswith('.xlsx'):
                # Write Excel with proper formatting and auto-fit columns
                self.write_excel_with_formatting(output_path, combined_df)
                self.log(f"Excel file written successfully: {output_path}")
            else:
                combined_df.to_csv(output_path, index=False, encoding='utf-8')
                self.log(f"CSV file written successfully: {output_path}")

        except PermissionError:
            # Excel file is open - re-raise to be handled by retry logic
            self.log(f"ERROR: Cannot write to file (file is open or locked)")
            raise
        except Exception as e:
            self.log(f"Error writing output file: {e}")
            raise

    def write_excel_with_formatting(self, output_path, df):
        """Write DataFrame to Excel with formatting and auto-fit columns"""
        try:
            from openpyxl import load_workbook
            from openpyxl.utils import get_column_letter

            # Write to Excel using pandas
            df.to_excel(output_path, index=False, sheet_name='PO Data', engine='openpyxl')

            # Open workbook to apply formatting
            wb = load_workbook(output_path)
            ws = wb['PO Data']

            # Auto-fit column widths
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)

                for cell in column:
                    try:
                        if cell.value:
                            cell_length = len(str(cell.value))
                            if cell_length > max_length:
                                max_length = cell_length
                    except:
                        pass

                # Set column width (add a little padding)
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                ws.column_dimensions[column_letter].width = adjusted_width

            # Format Order Number as integer (no decimals, with thousand separators)
            try:
                if 'Order Number' in df.columns:
                    col_idx = list(df.columns).index('Order Number') + 1
                    col_letter = get_column_letter(col_idx)

                    for row in range(2, ws.max_row + 1):
                        cell = ws[f'{col_letter}{row}']
                        if cell.value is not None:
                            cell.number_format = '0'  # Integer with no decimals, no separators
            except:
                pass

            # Format Line as number with conditional decimals
            try:
                if 'Line' in df.columns:
                    col_idx = list(df.columns).index('Line') + 1
                    col_letter = get_column_letter(col_idx)

                    for row in range(2, ws.max_row + 1):
                        cell = ws[f'{col_letter}{row}']
                        if cell.value is not None:
                            cell.number_format = '0.0#'  # Show up to 2 decimals only if needed (1, 1.1, 1.23)
            except:
                pass

            # Format number columns with thousand separators
            from openpyxl.styles import numbers
            numeric_cols = ['Quantity', 'Unit Price', 'Amount']

            for col_name in numeric_cols:
                try:
                    # Find column index
                    col_idx = list(df.columns).index(col_name) + 1
                    col_letter = get_column_letter(col_idx)

                    # Apply number format to data rows (skip header)
                    for row in range(2, ws.max_row + 1):
                        cell = ws[f'{col_letter}{row}']
                        if cell.value is not None:
                            if col_name == 'Quantity':
                                # No decimals for quantity
                                cell.number_format = '#,##0'
                            else:
                                # 2 decimals for prices
                                cell.number_format = '#,##0.00'
                except:
                    pass

            # Save the formatted workbook
            wb.save(output_path)
            wb.close()

        except PermissionError:
            # File is locked or open - re-raise to trigger retry logic
            self.log(f"ERROR: Cannot write Excel file (file is open or locked)")
            raise
        except Exception as e:
            # If formatting fails, at least we have the basic file
            self.log(f"Warning: Could not apply Excel formatting: {e}")
            pass

    def save_settings(self):
        """Save user settings to file"""
        try:
            settings = {
                'email_address': self.email_var.get(),
                'folder_contains': self.folder_var.get(),
                'subject_contains': self.subject_var.get(),
                'start_date': self.start_date_var.get(),
                'output_path': self.output_path_var.get()
            }
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception as e:
            # Silently fail - don't interrupt user
            pass

    def load_settings(self):
        """Load user settings from file"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)

                self.email_var.set(settings.get('email_address', ''))
                self.folder_var.set(settings.get('folder_contains', ''))
                self.subject_var.set(settings.get('subject_contains', ''))
                self.start_date_var.set(settings.get('start_date', ''))

                # Only load output path if it's not the default
                saved_output = settings.get('output_path', '')
                if saved_output:
                    self.output_path_var.set(saved_output)
        except Exception as e:
            # Silently fail - don't interrupt user
            pass


def main():
    root = ttk.Window(themename="superhero")
    app = PDFExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
