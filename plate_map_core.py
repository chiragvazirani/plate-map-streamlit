import os
import subprocess
import sys
import shutil

def generate_plate_maps(input_csv, output_dir):
    """
    Calls generate_plate_map.py using an absolute path
    and safely collects the outputs (Windows-safe).
    """

    base_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(base_dir, "generate_plate_map.py")

    if not os.path.exists(script_path):
        raise FileNotFoundError(f"Cannot find generate_plate_map.py at {script_path}")

    subprocess.check_call(
        [sys.executable, script_path, input_csv],
        cwd=base_dir
    )

    base = os.path.splitext(os.path.basename(input_csv))[0]

    pdf_name = f"{base}_source_plate_map.pdf"
    xlsx_name = f"{base}_source_plate_map.xlsx"

    pdf_src = os.path.join(base_dir, pdf_name)
    xlsx_src = os.path.join(base_dir, xlsx_name)

    if not os.path.exists(pdf_src) or not os.path.exists(xlsx_src):
        raise FileNotFoundError("Expected output files were not generated")

    os.makedirs(output_dir, exist_ok=True)

    pdf_out = os.path.join(output_dir, pdf_name)
    xlsx_out = os.path.join(output_dir, xlsx_name)

    # Windows-safe copy
    shutil.copy2(pdf_src, pdf_out)
    shutil.copy2(xlsx_src, xlsx_out)

    # Best-effort cleanup
    try:
        os.remove(pdf_src)
    except PermissionError:
        pass

    try:
        os.remove(xlsx_src)
    except PermissionError:
        pass

    return pdf_out, xlsx_out
