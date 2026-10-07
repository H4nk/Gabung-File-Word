from pathlib import Path
import re
import sys
import time

import win32com.client


# ============================================================
# KONFIGURASI
# ============================================================

FOLDER_INPUT = Path(".")
OUTPUT_FILE = "HASIL_GABUNGAN.docx"

# Word constants
WD_FORMAT_DOCUMENT_DEFAULT = 16   # .docx
WD_COLLAPSE_END = 0
WD_PAGE_BREAK = 7


# ============================================================
# URUTAN FILE
# ============================================================

def numeric_key(path):
    """
    Urutkan berdasarkan angka pertama pada nama file.

    1.doc
    2.doc
    ...
    9.docx
    10.doc
    """

    match = re.search(r"\d+", path.stem)

    if match:
        return (
            0,
            int(match.group()),
            path.name.lower()
        )

    return (
        1,
        999999999,
        path.name.lower()
    )


def get_word_files(folder):
    files = []

    for path in folder.iterdir():

        if not path.is_file():
            continue

        if path.name.lower() == OUTPUT_FILE.lower():
            continue

        if path.suffix.lower() in (".doc", ".docx"):
            files.append(path)

    return sorted(
        files,
        key=numeric_key
    )


# ============================================================
# MAIN
# ============================================================

def main():

    folder = FOLDER_INPUT.resolve()
    output = folder / OUTPUT_FILE

    print("=" * 72)
    print("GABUNG FILE WORD MENGGUNAKAN MICROSOFT WORD")
    print("=" * 72)
    print(f"Folder : {folder}")
    print(f"Output : {output}")
    print()

    files = get_word_files(folder)

    if not files:
        print("Tidak ditemukan file .doc atau .docx.")
        sys.exit(1)

    print("Urutan dokumen:")
    print("-" * 72)

    for i, file in enumerate(files, start=1):

        match = re.search(
            r"\d+",
            file.stem
        )

        nomor = (
            match.group()
            if match
            else "-"
        )

        print(
            f"{i:03d}. "
            f"{nomor:>5}  "
            f"{file.name}"
        )

    print()

    # --------------------------------------------------------
    # BUKA MICROSOFT WORD
    # --------------------------------------------------------

    try:

        word = win32com.client.DispatchEx(
            "Word.Application"
        )

    except Exception as error:

        print("=" * 72)
        print("MICROSOFT WORD TIDAK DITEMUKAN")
        print("=" * 72)
        print(
            "Program ini membutuhkan Microsoft Word "
            "desktop yang terinstall di Windows."
        )
        print()
        print("Install dependency:")
        print(
            r".\.venv\Scripts\python.exe -m pip install pywin32"
        )
        print()
        print(f"Detail error: {error}")

        sys.exit(1)

    word.Visible = False
    word.DisplayAlerts = 0

    master = None

    try:

        # ----------------------------------------------------
        # BUKA DOKUMEN PERTAMA
        # ----------------------------------------------------

        first_file = str(
            files[0].resolve()
        )

        print(
            f"[1/{len(files)}] "
            f"Membuka: {files[0].name}"
        )

        master = word.Documents.Open(
            first_file,
            ReadOnly=False,
            AddToRecentFiles=False
        )

        # ----------------------------------------------------
        # GABUNG DOKUMEN BERIKUTNYA
        # ----------------------------------------------------

        for index, file in enumerate(
            files[1:],
            start=2
        ):

            print(
                f"[{index}/{len(files)}] "
                f"Menggabungkan: {file.name}"
            )

            # Posisi akhir dokumen master.
            selection = master.Range(
                master.Content.End - 1,
                master.Content.End - 1
            )

            # Page break sebelum dokumen berikutnya.
            selection.InsertBreak(
                WD_PAGE_BREAK
            )

            # Range setelah page break.
            insertion_range = master.Range(
                master.Content.End - 1,
                master.Content.End - 1
            )

            # InsertFile menggunakan Microsoft Word sendiri.
            # Ini mempertahankan format Word jauh lebih baik
            # dibanding menyalin paragraph dengan python-docx.
            insertion_range.InsertFile(
                str(file.resolve())
            )

        # ----------------------------------------------------
        # SIMPAN SEBAGAI DOCX
        # ----------------------------------------------------

        if output.exists():
            output.unlink()

        print()
        print("Menyimpan hasil...")

        master.SaveAs2(
            str(output.resolve()),
            FileFormat=WD_FORMAT_DOCUMENT_DEFAULT
        )

        # Pastikan file selesai ditulis.
        time.sleep(1)

        print()
        print("=" * 72)
        print("SELESAI")
        print("=" * 72)
        print(f"Jumlah file : {len(files)}")
        print(f"Hasil       : {output}")
        print()
        print(
            "File asli tidak diubah."
        )
        print(
            "Penggabungan dilakukan langsung oleh Microsoft Word."
        )
        print("=" * 72)

    except Exception as error:

        print()
        print("=" * 72)
        print("ERROR")
        print("=" * 72)
        print(error)
        print("=" * 72)

        sys.exit(1)

    finally:

        if master is not None:
            try:
                master.Close(
                    SaveChanges=False
                )
            except Exception:
                pass

        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
