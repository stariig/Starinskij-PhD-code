#!/usr/bin/env python3
"""Create overall and subgroup histograms from a user-selected Excel sheet."""

from itertools import product
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

"""
05_histograms.py

Purpose:
    Create overall and subgroup histograms from a user-selected Excel sheet.

Input:
    Excel workbook containing continuos and discrete data arranged in columns (1 paticipant per row).


Output:
    PNG files containing histograms of the continuous response variable for all data, each level of each grouping variable, and each combination of levels of the grouping variables.

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

def choose_file():
	root = tk.Tk()
	root.withdraw()
	selected = filedialog.askopenfilename(
		title="Select Excel workbook",
		filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
	)
	root.destroy()
	return Path(selected) if selected else None


def choose_sheet(sheets):
	"""Return the selected worksheet name, or None if the dialog is cancelled."""
	result = []
	root = tk.Tk()
	root.title("Choose worksheet")
	root.resizable(False, False)

	tk.Label(root, text="Which worksheet should be used?").pack(padx=12, pady=(12, 4))
	sheet_box = tk.Listbox(root, width=48, height=min(20, max(5, len(sheets))), exportselection=False)
	for sheet in sheets:
		sheet_box.insert(tk.END, sheet)
	sheet_box.selection_set(0)
	sheet_box.pack(padx=12, pady=4)

	def accept():
		choices = sheet_box.curselection()
		if choices:
			result.append(sheets[choices[0]])
			root.destroy()

	tk.Button(root, text="Use worksheet", command=accept, width=16).pack(pady=(4, 12))
	root.protocol("WM_DELETE_WINDOW", root.destroy)
	root.mainloop()
	return result[0] if result else None


def choose_columns(columns):
	"""Return (response column, grouping columns), or None if cancelled."""
	result = []
	root = tk.Tk()
	root.title("Choose columns")
	root.resizable(False, False)

	tk.Label(root, text="Continuous response variable (x axis):").grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")
	response = tk.StringVar(value=columns[0])
	ttk.Combobox(root, textvariable=response, values=columns, state="readonly", width=35).grid(
		row=1, column=0, padx=12, pady=(0, 10), sticky="ew"
	)
	tk.Label(root, text="Grouping columns (select one or more; optional):").grid(
		row=2, column=0, padx=12, pady=(0, 4), sticky="w"
	)
	group_box = tk.Listbox(root, selectmode=tk.MULTIPLE, height=min(12, max(5, len(columns))), width=40, exportselection=False)
	for column in columns:
		group_box.insert(tk.END, column)
	group_box.grid(row=3, column=0, padx=12, pady=(0, 10), sticky="ew")

	def accept():
		groups = [columns[index] for index in group_box.curselection()]
		if response.get() in groups:
			messagebox.showerror("Invalid selection", "The response column cannot also be a grouping column.", parent=root)
			return
		result.append((response.get(), groups))
		root.destroy()

	tk.Button(root, text="Create histograms", command=accept, width=16).grid(row=4, column=0, pady=(0, 12))
	root.protocol("WM_DELETE_WINDOW", root.destroy)
	root.mainloop()
	return result[0] if result else None


def safe_name(value):
	"""Make a value suitable for use in a Windows filename."""
	cleaned = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", str(value)).strip(" .")
	return cleaned or "blank"


def make_histogram(values, title, response_name, output_path, bins):
	fig, ax = plt.subplots(figsize=(8, 5.5))
	ax.hist(values, bins=bins, color="#2f6f8f", edgecolor="white", linewidth=0.6)
	ax.set_xlabel(response_name)
	ax.set_ylabel("Frequency")
	ax.set_title(title)
	fig.tight_layout()
	fig.savefig(output_path, dpi=300)
	plt.close(fig)


def create_histograms(data, response_name, group_names, output_dir):
	response = pd.to_numeric(data[response_name], errors="coerce")
	valid = data.loc[response.notna()].copy()
	valid[response_name] = response[response.notna()]
	if valid.empty:
		raise ValueError(f'No numeric values were found in response column "{response_name}".')

	values = valid[response_name].to_numpy(dtype=float)
	if np.nanmin(values) == np.nanmax(values):
		centre = values[0]
		bins = np.linspace(centre - 0.5, centre + 0.5, 11)
	else:
		bins = np.histogram_bin_edges(values, bins="auto")

	plot_count = 0
	make_histogram(values, "All data", response_name, output_dir / "00_all_data.png", bins)
	plot_count += 1

	if not group_names:
		return plot_count

	levels = []
	for group_name in group_names:
		group_values = valid[group_name].dropna().unique().tolist()
		if not group_values:
			continue
		levels.append((group_name, group_values))
		for level in group_values:
			subset = valid[valid[group_name] == level]
			if subset.empty:
				continue
			title = f"{group_name} = {level}"
			filename = f"{plot_count:02d}_{safe_name(group_name)}_{safe_name(level)}.png"
			make_histogram(subset[response_name], title, response_name, output_dir / filename, bins)
			plot_count += 1

	for combination in product(*[items[1] for items in levels]):
		mask = pd.Series(True, index=valid.index)
		labels = []
		for group_name, level in zip([items[0] for items in levels], combination):
			mask &= valid[group_name] == level
			labels.append(f"{group_name} = {level}")
		subset = valid.loc[mask]
		if subset.empty:
			continue
		filename = f"{plot_count:02d}_combination_" + "_".join(safe_name(level) for level in combination) + ".png"
		make_histogram(subset[response_name], "; ".join(labels), response_name, output_dir / filename, bins)
		plot_count += 1

	return plot_count


def main():
	try:
		workbook = choose_file()
		if workbook is None:
			return
		print(f"Reading workbook: {workbook}")
		sheets = pd.ExcelFile(workbook).sheet_names
		sheet = choose_sheet(sheets)
		if sheet is None:
			return
		print(f"Reading worksheet: {sheet}")
		data = pd.read_excel(workbook, sheet_name=sheet)
		if data.empty or len(data.columns) < 1:
			raise ValueError("The selected worksheet has no usable columns.")
		data.columns = [str(column) for column in data.columns]
		choices = choose_columns(list(data.columns))
		if choices is None:
			return
		response_name, group_names = choices
		output_dir = workbook.parent / f"{workbook.stem}_histograms"
		output_dir.mkdir(exist_ok=True)
		print("Creating histograms...")
		count = create_histograms(data, response_name, group_names, output_dir)
		messagebox.showinfo("Histograms complete", f"Created {count} histogram(s) in:\n{output_dir}")
	except Exception as error:
		messagebox.showerror("Could not create histograms", str(error))


if __name__ == "__main__":
	main()
