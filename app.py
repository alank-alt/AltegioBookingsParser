import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import math
from thefuzz import process
from difflib import SequenceMatcher

class MatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Parser and Fuzzy Matcher")
        self.root.geometry("650x450")

        # Variables
        self.bookings_path = tk.StringVar(value=os.path.abspath("Bookings.xlsx"))
        self.staff_path = tk.StringVar(value=os.path.abspath("Staff.xlsx"))
        self.services_path = tk.StringVar(value=os.path.abspath("services.xls"))
        self.threshold = tk.DoubleVar(value=70.0)

        # UI Elements
        self.create_file_selector("Bookings Excel:", self.bookings_path)
        self.create_file_selector("Staff Excel:", self.staff_path)
        self.create_file_selector("Services Excel:", self.services_path)

        # Threshold
        frame_thresh = tk.Frame(root)
        frame_thresh.pack(pady=10, padx=20, fill=tk.X)
        tk.Label(frame_thresh, text="Min Match %:").pack(side=tk.LEFT)
        tk.Spinbox(frame_thresh, from_=0, to=100, increment=0.1, textvariable=self.threshold, width=10).pack(side=tk.LEFT, padx=5)

        # Progress UI
        self.status_label = tk.Label(root, text="Ready", fg="blue", font=("Arial", 10))
        self.status_label.pack(pady=10)
        self.progress = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=5)

        # Run Button
        self.btn_run = tk.Button(root, text="Start Processing", command=self.run_matching, bg="green", fg="white", font=("Arial", 12, "bold"))
        self.btn_run.pack(pady=20)

    def create_file_selector(self, label_text, string_var):
        frame = tk.Frame(self.root)
        frame.pack(pady=5, padx=20, fill=tk.X)
        tk.Label(frame, text=label_text, width=15, anchor="w").pack(side=tk.LEFT)
        tk.Entry(frame, textvariable=string_var, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="Browse...", command=lambda: self.browse_file(string_var)).pack(side=tk.LEFT)

    def load_excel_safe(self, file_path):
        if str(file_path).lower().endswith('.xls'):
            try:
                return pd.read_excel(file_path)
            except Exception:
                return pd.read_excel(file_path, engine_kwargs={'ignore_workbook_corruption': True})
        return pd.read_excel(file_path)

    def browse_file(self, string_var):
        filename = filedialog.askopenfilename(
            title="Select File",
            filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        )
        if filename:
            string_var.set(filename)

    def run_matching(self):
        bookings_file = self.bookings_path.get()
        staff_file = self.staff_path.get()
        services_file = self.services_path.get()

        if not (os.path.exists(bookings_file) and os.path.exists(staff_file) and os.path.exists(services_file)):
            messagebox.showerror("Error", "Please select all three valid input files.")
            return

        self.btn_run.config(state=tk.DISABLED)
        self.root.update()

        try:
            self.status_label.config(text=f"Loading data from {os.path.basename(bookings_file)}...")
            self.root.update()

            # 1. Load Data
            df_bookings = self.load_excel_safe(bookings_file)
            df_staff = self.load_excel_safe(staff_file)
            df_services = self.load_excel_safe(services_file)

            # 2. Delete columns
            cols_to_delete = ['booking_id', 'customer_id', 'customer_card_id', 'added_by', 'appointment_type', 'booking_finished_at', 'source_name', 'price']
            df_bookings.drop(columns=[c for c in cols_to_delete if c in df_bookings.columns], inplace=True)

            # 3. Create new columns
            new_cols = ['staffer_ID', 'service_ID', 'duration', 'paid', 'method', 'match']
            for col in new_cols:
                if col not in df_bookings.columns:
                    df_bookings[col] = None

            # 4. Copy final_price to paid
            if 'final_price' in df_bookings.columns:
                df_bookings['paid'] = df_bookings['final_price']
            
            # 5. Datetimes and duration
            self.status_label.config(text="Calculating durations and formatting dates...")
            self.root.update()

            if 'booked_from' in df_bookings.columns and 'booked_till' in df_bookings.columns:
                df_bookings['booked_from_dt'] = pd.to_datetime(df_bookings['booked_from'], errors='coerce')
                df_bookings['booked_till_dt'] = pd.to_datetime(df_bookings['booked_till'], errors='coerce')

                # Calculate duration in seconds
                durations = (df_bookings['booked_till_dt'] - df_bookings['booked_from_dt']).dt.total_seconds()
                df_bookings['duration'] = durations

                # Convert columns to format dd-mm-yyyy HH:MM
                df_bookings['booked_from'] = df_bookings['booked_from_dt'].dt.strftime('%d-%m-%Y %H:%M')
                df_bookings['booked_till'] = df_bookings['booked_till_dt'].dt.strftime('%d-%m-%Y %H:%M')

                # Drop temp columns
                df_bookings.drop(columns=['booked_from_dt', 'booked_till_dt'], inplace=True)

            # 6. Fuzzy Matching
            self.status_label.config(text="Preparing dictionaries for fuzzy matching...")
            self.root.update()

            # Prepare staff dict
            staff_dict = {}
            staff_names = []
            if 'Name' in df_staff.columns and 'ID' in df_staff.columns:
                for _, row in df_staff.iterrows():
                    name = str(row['Name']).strip()
                    if name and name.lower() != 'nan':
                        staff_dict[name] = row['ID']
                        staff_names.append(name)

            # Prepare services dict
            services_dict = {}
            services_names = []
            if 'Имя' in df_services.columns and 'ID' in df_services.columns:
                for _, row in df_services.iterrows():
                    name = str(row['Имя']).strip()
                    if name and name.lower() != 'nan':
                        services_dict[name] = row['ID']
                        services_names.append(name)

            total_rows = len(df_bookings)
            self.progress["maximum"] = total_rows

            self.status_label.config(text="Fuzzy matching in progress...")
            self.root.update()

            threshold_val = self.threshold.get()

            # Iterate over rows for matching
            for i in range(total_rows):
                staffer_score = 0
                service_score = 0
                
                # Staff match
                if 'staffer' in df_bookings.columns:
                    staffer_name = str(df_bookings.at[i, 'staffer']).strip()
                    if staffer_name and staffer_name.lower() != 'nan' and staff_names:
                        best_match, score = process.extractOne(staffer_name, staff_names)
                        if best_match:
                            exact_score = SequenceMatcher(None, staffer_name.lower(), best_match.lower()).ratio() * 100
                            staffer_score = math.floor(exact_score * 10) / 10.0
                            if staffer_score >= threshold_val:
                                df_bookings.at[i, 'staffer_ID'] = staff_dict[best_match]
                
                # Service match
                if 'service_name' in df_bookings.columns:
                    service_name = str(df_bookings.at[i, 'service_name']).strip()
                    if service_name and service_name.lower() != 'nan' and services_names:
                        best_match, score = process.extractOne(service_name, services_names)
                        if best_match:
                            exact_score = SequenceMatcher(None, service_name.lower(), best_match.lower()).ratio() * 100
                            service_score = math.floor(exact_score * 10) / 10.0
                            if service_score >= threshold_val:
                                df_bookings.at[i, 'service_ID'] = services_dict[best_match]

                # Combine scores into match column as requested
                match_texts = []
                if staffer_score > 0:
                    match_texts.append(f"Staff: {staffer_score}%")
                if service_score > 0:
                    match_texts.append(f"Service: {service_score}%")
                
                if match_texts:
                    df_bookings.at[i, 'match'] = " | ".join(match_texts)

                self.progress["value"] = i + 1
                if i % 10 == 0:
                    self.root.update_idletasks()

            # 7. Export
            self.status_label.config(text="Saving output file...")
            self.root.update()

            output_dir = os.path.dirname(bookings_file)
            output_path = os.path.join(output_dir, "Actual Upload Visits.xlsx")
            
            df_bookings.to_excel(output_path, index=False)

            self.status_label.config(text=f"Done! Saved to {os.path.basename(output_path)}")
            messagebox.showinfo("Success", f"Processing complete!\nSaved to:\n{output_path}")

        except Exception as e:
            self.status_label.config(text="Error occurred.")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            self.btn_run.config(state=tk.NORMAL)
            self.progress["value"] = 0
            self.root.update()

if __name__ == "__main__":
    root = tk.Tk()
    app = MatcherApp(root)
    root.mainloop()
