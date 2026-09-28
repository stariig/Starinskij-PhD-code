"""Create an ELISA heatmap from a user-selected Excel workbook."""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

"""
03_heatmap.py

Purpose:
    Create an ELISA heatmap from a user-selected Excel workbook.

Input:
    Excel workbook containing ELISA data.

Output:
    1. Excel workbook containing summary of ELISA readings for each sample/condition combination.
	2. PNG file containing the ELISA heatmap for each sample/condition combination.

Thesis:
    Chapter: ERVEV biosurveillance survey

Author:
    Igor Starinskij

AI assistance:
    Generative artificial intelligence tools, primarily GitHub Copilot 0.63.0 (GPT-5.6 Luna), were used during the development of software used in this thesis.
	AI was used to assist with Python code generation, debugging, code refinement and explanation of programming concepts.
	Generated code was reviewed, modified, tested and validated by the author.
	The analytical methodology, statistical decisions, interpretation of results and final responsibility for the analyses remain with the author.
"""

def choose_workbook() -> Path | None:
	"""Return the workbook selected in the Windows file picker."""
	root = tk.Tk()
	root.withdraw()
	root.attributes("-topmost", True)
	filename = filedialog.askopenfilename(
		title="Select ELISA workbook",
		filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
	)
	root.destroy()
	return Path(filename) if filename else None


def create_heatmap(workbook_path: Path) -> Path:
	"""Average replicate readings and save a condition-by-sample heatmap."""
	import pandas as pd
	import seaborn as sns
	import matplotlib.pyplot as plt

	data = pd.read_excel(workbook_path, sheet_name="heatmap")
	if data.shape[1] < 6:
		raise ValueError(
			'The "heatmap" sheet must contain at least six columns: '
			"four conditions, sample name, and ELISA reading."
		)

	data = data.iloc[:, :6].copy()
	column_names = ["condition_1", "condition_2", "condition_3", "condition_4", "sample", "ODn"]
	data.columns = column_names
	data["ODn"] = pd.to_numeric(data["ODn"], errors="coerce")
	data = data.dropna(subset=["ODn"])
	if data.empty:
		raise ValueError("The sixth column contains no numeric ELISA readings.")

	condition_columns = column_names[:4]
	group_columns = condition_columns + ["sample"]
	replicate_stats = data.groupby(group_columns, dropna=False, sort=False, as_index=False)["ODn"].agg(
		mean="mean",
		sd="std",
	)
	replicate_stats["cv"] = (
		replicate_stats["sd"] / replicate_stats["mean"].abs()
	).where(replicate_stats["mean"] > 0.5)
	heatmap_data = replicate_stats.pivot(
		index=condition_columns,
		columns="sample",
		values="mean",
	)
	cv_data = replicate_stats.pivot(index=condition_columns, columns="sample", values="cv")
	layout = pd.read_excel(workbook_path, sheet_name="layout", usecols=[0, 1])
	layout.columns = ["order", "sample"]
	layout["order"] = pd.to_numeric(layout["order"], errors="coerce")
	layout = layout.dropna(subset=["order", "sample"]).sort_values("order", kind="stable")
	listed_samples = layout["sample"].astype(str).drop_duplicates().tolist()
	present_samples = [str(sample) for sample in heatmap_data.columns]
	ordered_samples = [sample for sample in listed_samples if sample in present_samples]
	ordered_samples.extend(
		sample for sample in sorted(present_samples, key=str.casefold) if sample not in ordered_samples
	)
	heatmap_data.columns = present_samples
	heatmap_data = heatmap_data.reindex(columns=ordered_samples)
	cv_data.columns = present_samples
	cv_data = cv_data.reindex(index=heatmap_data.index, columns=ordered_samples)
	heatmap_data.index = heatmap_data.index.map(
		lambda values: " | ".join("" if pd.isna(value) else str(value) for value in values)
	)
	cv_data.index = heatmap_data.index
	transposed = False
	if len(heatmap_data.index) < len(heatmap_data.columns):
		heatmap_data = heatmap_data.T
		cv_data = cv_data.T
		transposed = True

	figure_width = max(8, 0.65 * len(heatmap_data.columns) + 3)
	figure_height = max(5, 0.38 * len(heatmap_data.index) + 2)
	figure, axis = plt.subplots(figsize=(figure_width, figure_height), constrained_layout=True)
	sns.heatmap(
		heatmap_data,
		ax=axis,
		cmap="YlOrRd",
		annot=True,
		fmt=".3f",
		linewidths=0.5,
		linecolor="white",
		cbar_kws={"label": "Mean ODn"},
	)
	for row_index, row in enumerate(cv_data.itertuples(index=False)):
		for column_index, cv in enumerate(row):
			if pd.notna(cv) and cv > 0.20:
				axis.text(
					column_index + 0.95,
					row_index + 0.5,
					"!",
					color="black",
					fontweight="bold",
					ha="right",
					va="center",
				)
	axis.set_xlabel("Conditions" if transposed else "Sample")
	axis.set_ylabel("Sample" if transposed else "Conditions")
	axis.set_title("ELISA mean ODn")
	if len(heatmap_data.columns) <= 3:
		axis.tick_params(axis="x", rotation=0)
		figure.canvas.draw()
		x_label_boxes = [label.get_window_extent() for label in axis.get_xticklabels()]
		x_labels_overlap = any(
			left_box.overlaps(right_box)
			for left_box, right_box in zip(x_label_boxes, x_label_boxes[1:])
		)
		if x_labels_overlap:
			axis.tick_params(axis="x", rotation=45)
	else:
		axis.tick_params(axis="x", rotation=45)
	axis.tick_params(axis="y", rotation=0)

	output_path = workbook_path.with_name(f"{workbook_path.stem}_heatmap.png")
	figure.savefig(output_path, dpi=300, bbox_inches="tight")
	plt.close(figure)
	return output_path


def main() -> None:
	workbook_path = choose_workbook()
	if workbook_path is None:
		return

	try:
		output_path = create_heatmap(workbook_path)
	except Exception as error:
		messagebox.showerror("Heatmap generation failed", str(error))
		raise
	else:
		messagebox.showinfo("Heatmap created", f"Saved heatmap to:\n{output_path}")


if __name__ == "__main__":
	main()
