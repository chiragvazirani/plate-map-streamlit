import sys
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side

import matplotlib.cm as cm
import matplotlib.colors as mcolors

# ============================================================
# INPUT / OUTPUT
# ============================================================
INPUT_FILE = sys.argv[1] if len(sys.argv) > 1 else "dubunit_file.csv"
run_tag = os.path.splitext(os.path.basename(INPUT_FILE))[0]

OUTPUT_PDF = f"{run_tag}_source_plate_map.pdf"
OUTPUT_XLSX = f"{run_tag}_source_plate_map.xlsx"

# ============================================================
# PLATE DEFINITION
# ============================================================
PLATE_ROWS = ["A", "B"]
PLATE_COLS = list(range(1, 25))

GLOBAL_WELLS = []
for c in PLATE_COLS:
    GLOBAL_WELLS.append(f"A{c:02d}")
    GLOBAL_WELLS.append(f"B{c:02d}")

# ============================================================
# LOAD DATA
# ============================================================
df = pd.read_csv(INPUT_FILE)
df.columns = df.columns.str.strip()

required = ["Source Plate Name", "Source Well", "Sequence Name", "Transfer Volume"]
missing = [c for c in required if c not in df.columns]
if missing:
    raise ValueError(f"Missing required columns: {missing}")

plates = list(df["Source Plate Name"].unique())

# ============================================================
# GLOBAL GENE LIST + IDs (first-seen order)
# ============================================================
ALL_GENES = []
for g in df["Sequence Name"]:
    if g not in ALL_GENES:
        ALL_GENES.append(g)

GENE_ID = {gene: f"G{i+1:02d}" for i, gene in enumerate(ALL_GENES)}

# ============================================================
# PROFESSIONAL DISTINCT COLORS (tab20 + tab20b + tab20c, softened)
# ============================================================
def soften_rgba(rgba, mix_with_white=0.18, desat=0.12):
    r, g, b, a = rgba
    h, s, v = mcolors.rgb_to_hsv((r, g, b))
    s = max(0.0, s * (1.0 - desat))
    r2, g2, b2 = mcolors.hsv_to_rgb((h, s, v))
    r3 = r2 * (1.0 - mix_with_white) + 1.0 * mix_with_white
    g3 = g2 * (1.0 - mix_with_white) + 1.0 * mix_with_white
    b3 = b2 * (1.0 - mix_with_white) + 1.0 * mix_with_white
    return (r3, g3, b3, a)

def build_palette():
    palettes = []
    for name in ["tab20", "tab20b", "tab20c"]:
        cmap = cm.get_cmap(name)
        palettes.extend([cmap(i) for i in range(cmap.N)])
    palettes = [soften_rgba(rgba) for rgba in palettes]
    return palettes

PALETTE = build_palette()
GENE_COLOR = {gene: PALETTE[i % len(PALETTE)] for i, gene in enumerate(ALL_GENES)}

# ============================================================
# EXCEL SETUP
# ============================================================
wb = Workbook()
wb.remove(wb.active)
thin = Side(style="thin")
legend_rows = []

# ============================================================
# PDF GENERATION
# ============================================================
with PdfPages(OUTPUT_PDF) as pdf:

    for plate in plates:
        plate_df = df[df["Source Plate Name"] == plate]

        gene_order = []
        for g in plate_df["Sequence Name"]:
            if g not in gene_order:
                gene_order.append(g)

        gene_counts = plate_df.groupby("Sequence Name").size().to_dict()
        well_volume = dict(zip(plate_df["Source Well"], plate_df["Transfer Volume"]))

        assignments = {}
        mer_index = {}
        gene_blocks = {}
        ptr = 0

        for gene in gene_order:
            n = gene_counts[gene]
            wells = GLOBAL_WELLS[ptr:ptr + n]
            gene_blocks[gene] = wells

            for i, w in enumerate(wells, start=1):
                assignments[w] = gene
                mer_index[w] = i

            ptr += n

            vols = plate_df[plate_df["Sequence Name"] == gene]["Transfer Volume"].unique()
            vol_label = vols[0] if len(vols) == 1 else "varies"
            legend_rows.append([plate, GENE_ID[gene], gene, n, vol_label])

        fig, ax = plt.subplots(figsize=(28, 3.3))

        for r, row in enumerate(PLATE_ROWS):
            for c, col in enumerate(PLATE_COLS):
                well = f"{row}{col:02d}"
                y = 1 - r

                if well in assignments:
                    ax.add_patch(Rectangle(
                        (c, y), 1, 1,
                        facecolor=GENE_COLOR[assignments[well]],
                        edgecolor="black",
                        linewidth=0.6
                    ))
                else:
                    ax.add_patch(Rectangle(
                        (c, y), 1, 1,
                        fill=False,
                        edgecolor="black",
                        linewidth=0.6
                    ))

                if well in mer_index:
                    ax.text(
                        c + 0.5, y + 0.38,
                        f"#{mer_index[well]}",
                        fontsize=9.5,
                        fontweight="bold",
                        ha="center", va="center"
                    )
                    ax.text(
                        c + 0.5, y + 0.18,
                        f"{well_volume.get(well, '')} µL",
                        fontsize=8.5,
                        ha="center", va="center"
                    )

        for c, col in enumerate(PLATE_COLS):
            ax.add_patch(Rectangle((c, 2), 1, 0.28, fill=False, linewidth=0.8))
            ax.text(c + 0.5, 2.14, str(col),
                    ha="center", va="center",
                    fontsize=9, fontweight="bold")

        for r, row in enumerate(PLATE_ROWS):
            y = 1 - r
            ax.add_patch(Rectangle((-0.7, y), 0.7, 1, fill=False, linewidth=0.8))
            ax.text(-0.35, y + 0.5, row,
                    ha="center", va="center",
                    fontsize=11, fontweight="bold")

        for gene, wells in gene_blocks.items():
            idxs = [GLOBAL_WELLS.index(w) for w in wells]
            mid = (min(idxs) + max(idxs)) / 2
            col_mid = int(mid // 2)
            row_mid = 1 if int(mid) % 2 == 0 else 0
            ax.text(col_mid + 0.5, row_mid + 0.82,
                    GENE_ID[gene],
                    fontsize=12,
                    fontweight="bold",
                    ha="center")

        legend_x0 = len(PLATE_COLS) + 0.6
        legend_y_top = 2.25
        row_step = 0.30
        genes_per_col = 8
        box_size = 0.16
        col_width = 4.6

        for i, gene in enumerate(gene_order):
            col_idx = i // genes_per_col
            row_idx = i % genes_per_col
            lx = legend_x0 + col_idx * col_width
            ly = legend_y_top - row_idx * row_step

            ax.add_patch(Rectangle(
                (lx, ly - box_size),
                box_size, box_size,
                facecolor=GENE_COLOR[gene],
                edgecolor="black",
                linewidth=0.5
            ))
            ax.text(
                lx + box_size + 0.12,
                ly - box_size / 2,
                f"{GENE_ID[gene]}  {gene}  ({gene_counts[gene]}×100mers)",
                fontsize=9.0,
                ha="left",
                va="center"
            )

        ax.text(
            len(PLATE_COLS) / 2,
            2.48,
            f"Source Plate Map – {plate}",
            ha="center",
            va="bottom",
            fontsize=13,
            fontweight="bold"
        )

        ax.set_xlim(-0.9, len(PLATE_COLS) + 10)
        ax.set_ylim(0, 2.6)
        ax.axis("off")

        fig.tight_layout(pad=0.3)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

# ============================================================
# EXCEL OUTPUT (Plate map sheets + Legend with volume)
# ============================================================
thin = Side(style="thin")

for plate in plates:
    ws = wb.create_sheet(title=plate[:31])
    ws.append([""] + [str(c) for c in PLATE_COLS])
    for row in PLATE_ROWS:
        ws.append([row] + [""] * len(PLATE_COLS))

    plate_df = df[df["Source Plate Name"] == plate]
    gene_order = []
    for g in plate_df["Sequence Name"]:
        if g not in gene_order:
            gene_order.append(g)

    gene_counts = plate_df.groupby("Sequence Name").size().to_dict()

    ptr = 0
    for gene in gene_order:
        n = gene_counts[gene]
        wells = GLOBAL_WELLS[ptr:ptr + n]

        rgba = GENE_COLOR[gene]
        hex_color = "%02X%02X%02X" % tuple(int(255 * x) for x in rgba[:3])
        fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

        for w in wells:
            r = PLATE_ROWS.index(w[0]) + 2
            c = int(w[1:]) + 1
            cell = ws.cell(row=r, column=c)
            cell.fill = fill
            cell.value = GENE_ID[gene]
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ptr += n

ws_leg = wb.create_sheet(title="Legend")
ws_leg.append(["Source Plate", "Gene ID", "Gene Name", "#100mers", "Volume (µL)"])
for row in legend_rows:
    ws_leg.append(row)

wb.save(OUTPUT_XLSX)

print("✅ Outputs generated:")
print(f" - {OUTPUT_PDF}")
print(f" - {OUTPUT_XLSX}")
