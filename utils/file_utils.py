import os
import pandas as pd
from datetime import datetime

def read_excel(file_path):
    """
    Membaca file Excel dan mengembalikan data sebagai list of dictionaries.
    Jika file tidak ada, buat file kosong dengan header yang sesuai.
    """
    if not os.path.isfile(file_path):
        # Tentukan header berdasarkan file
        if "user_accounts.xlsx" in file_path:
            headers = ["Username", "PasswordHash"]
        elif "inventory_items.xlsx" in file_path:
            headers = ["product_id", "date", "product_name", "category", "quantity", "purchase_price", "selling_price",
                       "total", "total_price"]
        elif "purchases.xlsx" in file_path:
            headers = ["purchase_id", "date", "product_name", "purchase_price", "quantity", "total", "total_price"]
        elif "sales.xlsx" in file_path:
            headers = ["sale_id", "date", "product_name", "quantity_sold", "selling_price", "total_income", "cogs",
                       "gross_profit"]
        elif "journal_entries.xlsx" in file_path:
            headers = ["entry_id", "date", "transaction_type", "product_name", "quantity", "unit_price", "total_amount"]
        else:
            headers = []

        df = pd.DataFrame(columns=headers)
        df.to_excel(file_path, index=False)
        return df.to_dict(orient='records')
    else:
        try:
            df = pd.read_excel(file_path)

            # Tambahkan kolom yang hilang dengan nilai NaN
            if "journal_entries.xlsx" in file_path:
                required_columns = ["entry_id", "date", "transaction_type", "product_name", "quantity", "unit_price",
                                    "total_amount"]
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = pd.NA
                # Pastikan urutan kolom
                df = df[required_columns]

            return df.to_dict(orient='records')
        except Exception as e:
            print(f"Error occurred while reading the Excel file: {e}")
            return []


def write_excel(file_path, data, headers):
    """
    Menulis seluruh data ke file Excel, menggantikan konten yang ada.
    """
    try:
        df = pd.DataFrame(data, columns=headers)
        with pd.ExcelWriter(file_path) as writer:
            df.to_excel(writer, index=False)
    except Exception as e:
        print(f"Error occurred while writing to the Excel file: {e}")


def append_excel(file_path, row, headers):
    """
    Menambahkan satu baris data ke file Excel.
    Menghapus kolom yang semua entri-nya adalah NaN sebelum concatenation.
    """
    try:
        if not os.path.exists(file_path):
            write_excel(file_path, [row], headers)
        else:
            existing_df = pd.read_excel(file_path)

            # Buat DataFrame untuk baris baru dan hapus kolom yang semua NaN
            new_row = pd.DataFrame([row], columns=headers).dropna(axis=1, how='all')

            if new_row.empty:
                # Jika new_row tidak memiliki kolom yang valid, hentikan proses
                print("Warning: New row is empty after dropping all-NA columns. Skipping append.")
                return

            # Pastikan new_row memiliki kolom yang sama dengan existing_df
            # Jika tidak, tambahkan kolom yang hilang di new_row dengan nilai NaN
            for col in existing_df.columns:
                if col not in new_row.columns:
                    new_row[col] = pd.NA

            # Reorder new_row untuk mengikuti urutan kolom existing_df
            new_row = new_row[existing_df.columns]

            # Drop columns in new_row that are all NA (untuk menghindari FutureWarning)
            new_row = new_row.dropna(axis=1, how='all')

            # Gabungkan existing_df dan new_row
            combined_df = pd.concat([existing_df, new_row], ignore_index=True)

            # Tulis kembali ke file Excel
            write_excel(file_path, combined_df.to_dict(orient='records'), existing_df.columns.tolist())
    except Exception as e:
        print(f"Error occurred while appending to the Excel file: {e}")


def log_journal(entry):
    """
    Mencatat transaksi ke journal_entries.xlsx
    """
    JOURNAL_FILE = "database/journal_entries.xlsx"
    journal_entries = read_excel(JOURNAL_FILE)

    new_entry_id = max([e.get("entry_id", 0) for e in journal_entries if "entry_id" in e and pd.notna(e.get("entry_id"))],
                       default=0) + 1

    # Pastikan semua kunci ada, gunakan nilai default jika tidak
    entry_data = {
        "entry_id": new_entry_id,
        "date": entry.get("date", datetime.now().strftime("%Y-%m-%d")),
        "transaction_type": entry.get("transaction_type", "Unknown"),
        "product_name": entry.get("product_name", "Unknown"),
        "quantity": entry.get("quantity", 0),
        "unit_price": entry.get("unit_price", 0.0),
        "total_amount": entry.get("total_amount", 0.0)
    }

    headers = ["entry_id", "date", "transaction_type", "product_name", "quantity", "unit_price", "total_amount"]

    append_excel(JOURNAL_FILE, entry_data, headers)