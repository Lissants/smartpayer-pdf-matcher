"""
convert_missing_pdfs.py
-----------------------
Scans the Generated_Letters directory for .docx files that have no
matching .pdf counterpart, then converts them using Microsoft Word COM
(pywin32).  A timestamped log is written to the same directory.

Naming convention in this folder:
  <base>_with_tables.docx   ←→   <base>_letter.pdf

The script strips known suffixes (_with_tables, _letter) to obtain the
base name, then checks whether *any* .pdf with that base exists.
"""

import os
import sys
import logging
import datetime
import pathlib

# ── Configuration ────────────────────────────────────────────────────────────
SCRIPT_DIR      = pathlib.Path(__file__).parent
LETTERS_DIR     = SCRIPT_DIR / "Generated_Letters"
LOG_FILE        = LETTERS_DIR / f"conversion_log_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt"

DOCX_SUFFIX     = "_with_tables"   # suffix on the docx side
PDF_SUFFIX      = "_letter"        # suffix on the pdf side
# ─────────────────────────────────────────────────────────────────────────────


def setup_logger() -> logging.Logger:
    """Configure a logger that writes to both console and a log file."""
    logger = logging.getLogger("pdf_converter")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    # Console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # File handler
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


def get_base_name(stem: str) -> str:
    """
    Strip the known docx/pdf suffixes to get the shared base name.
    e.g.  'Smartpayer April 2026 FOOBAR_with_tables' → 'Smartpayer April 2026 FOOBAR'
          'Smartpayer April 2026 FOOBAR_letter'       → 'Smartpayer April 2026 FOOBAR'
    Falls back to the original stem if neither suffix is found.
    """
    for sfx in (DOCX_SUFFIX, PDF_SUFFIX):
        if stem.endswith(sfx):
            return stem[: -len(sfx)]
    return stem


def scan_directory(letters_dir: pathlib.Path, logger: logging.Logger):
    """
    Scan letters_dir and return two sets:
      docx_bases  – base names that have a .docx file
      pdf_bases   – base names that have a .pdf file
    Also returns a mapping  base_name → full docx path  for conversion.
    Temporary Word lock files (~$...) are ignored.
    """
    docx_map: dict[str, pathlib.Path] = {}   # base → Path
    pdf_bases: set[str]               = set()

    for f in letters_dir.iterdir():
        if not f.is_file():
            continue
        if f.name.startswith("~$"):          # skip temp/lock files
            logger.debug("Skipping temp file: %s", f.name)
            continue

        base = get_base_name(f.stem)

        if f.suffix.lower() == ".docx":
            docx_map[base] = f
        elif f.suffix.lower() == ".pdf":
            pdf_bases.add(base)

    return docx_map, pdf_bases


def convert_docx_to_pdf(docx_path: pathlib.Path,
                        pdf_path: pathlib.Path,
                        logger: logging.Logger) -> bool:
    """
    Convert a single .docx to .pdf using Microsoft Word COM (pywin32).
    Returns True on success, False on failure.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError:
        logger.error("pywin32 is not installed.  Run: pip install pywin32")
        return False

    word = None
    doc  = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0          # wdAlertsNone

        doc = word.Documents.Open(str(docx_path.resolve()))
        # wdFormatPDF = 17
        doc.SaveAs2(str(pdf_path.resolve()), FileFormat=17)
        logger.info("  ✔  Converted → %s", pdf_path.name)
        return True

    except Exception as exc:
        logger.error("  ✗  Failed to convert %s: %s", docx_path.name, exc)
        return False

    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def main():
    logger = setup_logger()

    logger.info("=" * 70)
    logger.info("SmartPayer PDF Converter")
    logger.info("Scanning: %s", LETTERS_DIR)
    logger.info("=" * 70)

    if not LETTERS_DIR.is_dir():
        logger.error(
            "Generated_Letters folder not found!\n"
            "  Expected location : %s\n"
            "  Make sure this script sits in the same folder as Generated_Letters\\ "
            "and that the folder name is spelled exactly 'Generated_Letters'.",
            LETTERS_DIR,
        )
        sys.exit(1)

    # ── Step 1: Scan ──────────────────────────────────────────────────────────
    docx_map, pdf_bases = scan_directory(LETTERS_DIR, logger)

    logger.info("\n[SCAN] %d .docx files found (excl. temp files)", len(docx_map))
    logger.info("[SCAN] %d .pdf  files found", len(pdf_bases))

    # ── Step 2: Find docx-only files ──────────────────────────────────────────
    missing = {base: path
               for base, path in docx_map.items()
               if base not in pdf_bases}

    if not missing:
        logger.info("\n✔  All .docx files already have a matching .pdf. Nothing to do.")
        return

    logger.info("\n[MISSING PDFs]  %d file(s) have .docx but no .pdf:\n", len(missing))
    for base, docx_path in sorted(missing.items()):
        logger.info("  • %s", docx_path.name)

    # ── Step 3: Convert ───────────────────────────────────────────────────────
    logger.info("\n[CONVERTING]  Starting Word COM conversion …\n")

    success_list: list[str] = []
    failure_list: list[str] = []

    for base, docx_path in sorted(missing.items()):
        pdf_name = base + PDF_SUFFIX + ".pdf"   # e.g. Smartpayer April 2026 FOOBAR_letter.pdf
        pdf_path = LETTERS_DIR / pdf_name
        logger.info("Converting: %s", docx_path.name)

        ok = convert_docx_to_pdf(docx_path, pdf_path, logger)
        if ok:
            success_list.append(pdf_name)
        else:
            failure_list.append(docx_path.name)

    # ── Step 4: Summary ───────────────────────────────────────────────────────
    logger.info("\n" + "=" * 70)
    logger.info("CONVERSION SUMMARY")
    logger.info("=" * 70)
    logger.info("Total missing PDFs found : %d", len(missing))
    logger.info("Successfully converted   : %d", len(success_list))
    logger.info("Failed                   : %d", len(failure_list))

    if success_list:
        logger.info("\n[SUCCESS]  The following PDFs were created:")
        for name in success_list:
            logger.info("  ✔  %s", name)

    if failure_list:
        logger.info("\n[FAILED]  The following files could NOT be converted:")
        for name in failure_list:
            logger.info("  ✗  %s", name)

    logger.info("\nLog saved to: %s", LOG_FILE)
    logger.info("=" * 70)


if __name__ == "__main__":
    main()
