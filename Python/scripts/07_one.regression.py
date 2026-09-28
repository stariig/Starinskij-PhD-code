"""Run a regression using columns selected from an Excel workbook."""

from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path
from typing import Optional

"""
06_regression.py

Purpose:
    Run a regression using columns selected from an Excel workbook.

Input:
    Excel workbook containing continuous and/or two-level response data, and one or more explanatory variables.


Output:
    1. CSV file containing the regression coefficients, standard errors, p-values, and confidence intervals.
	2. PNG file containing a scatter plot of observed versus fitted response values.
	3. TXT file containing a summary of the regression analysis, including model type, transformation, and fit statistics.

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

def choose_input_file() -> Optional[Path]:
	from tkinter import Tk, filedialog

	root = Tk()
	root.withdraw()
	selected = filedialog.askopenfilename(
		title="Select ELISA and neutralisation workbook",
		filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
	)
	root.destroy()
	return Path(selected) if selected else None


def ask_for_sheet(sheet_names: list[str]) -> str:
	print("Available sheets:")
	for index, sheet_name in enumerate(sheet_names, start=1):
		print(f"{index}. {sheet_name}")
	while True:
		selection = input("Enter the sheet number or name: ").strip()
		if selection.isdigit() and 1 <= int(selection) <= len(sheet_names):
			return sheet_names[int(selection) - 1]
		if selection in sheet_names:
			return selection
		print("Please enter one of the listed sheet numbers or names.")


def ask_for_columns(columns: list[str], prompt: str, allow_multiple: bool) -> list[str]:
	print("Available columns:")
	for index, column in enumerate(columns, start=1):
		print(f"{index}. {column}")
	while True:
		selection = input(prompt).strip()
		parts = [part.strip() for part in selection.split(",") if part.strip()]
		chosen: list[str] = []
		for part in parts:
			if part.isdigit() and 1 <= int(part) <= len(columns):
				column = columns[int(part) - 1]
			elif part in columns:
				column = part
			else:
				break
			if column not in chosen:
				chosen.append(column)
		else:
			if chosen and (allow_multiple or len(chosen) == 1):
				return chosen
		print("Please select valid column numbers or names separated by commas.")


def ask_for_random_effect(columns: list[str]) -> Optional[str]:
	print("\nChoose a random-effects grouping variable, or press Enter for none:")
	for index, column in enumerate(columns, start=1):
		print(f"{index}. {column}")
	while True:
		selection = input("Enter the random-effects column number or name: ").strip()
		if not selection:
			return None
		if selection.isdigit() and 1 <= int(selection) <= len(columns):
			return columns[int(selection) - 1]
		if selection in columns:
			return selection
		print("Please enter one of the listed column numbers or names, or press Enter for none.")


def ask_for_response_reference(data, response_column: str):
	values = data[response_column].dropna().drop_duplicates().tolist()
	if len(values) != 2:
		return None
	print(f"\nSelect the reference (0) level for '{response_column}':")
	for index, value in enumerate(values, start=1):
		print(f"{index}. {value}")
	while True:
		selection = input("Enter the reference response number or value: ").strip()
		if selection.isdigit() and 1 <= int(selection) <= len(values):
			return values[int(selection) - 1]
		for value in values:
			if str(value) == selection:
				return value
		print("Please enter one of the listed reference numbers or values.")


def ask_for_references(data, predictor_columns: list[str]) -> dict[str, object]:
	import pandas as pd

	references: dict[str, object] = {}
	for column in predictor_columns:
		if pd.api.types.is_numeric_dtype(data[column]):
			continue
		levels = list(data[column].dropna().drop_duplicates())
		if len(levels) < 2:
			continue
		print(f"\nSelect the reference level for '{column}':")
		for index, level in enumerate(levels, start=1):
			print(f"{index}. {level}")
		while True:
			selection = input("Enter the reference number or value: ").strip()
			if selection.isdigit() and 1 <= int(selection) <= len(levels):
				references[column] = levels[int(selection) - 1]
				break
			if selection in {str(level) for level in levels}:
				references[column] = next(level for level in levels if str(level) == selection)
				break
			print("Please enter one of the listed reference numbers or values.")
	return references


def normalise_column_names(columns) -> dict[str, str]:
	return {str(column).strip().strip('"'): column for column in columns}


def neutralisation_to_binary(value: object) -> Optional[int]:
	if value is None:
		return None
	text = str(value).strip().lower()
	if text in {"pos", "positive", "+", "1", "true", "yes"}:
		return 1
	if text in {"neg", "negative", "-", "0", "false", "no"}:
		return 0
	return None


def safe_name(value: str) -> str:
	return re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._-") or "workbook"


def load_data(input_path: Path, sheet_name: str, response_column: str,
			  predictor_columns: list[str], response_reference=None):
	import pandas as pd

	data = pd.read_excel(input_path, sheet_name=sheet_name)
	column_lookup = normalise_column_names(data.columns)
	selected = [response_column, *predictor_columns]
	missing = [column for column in selected if column not in column_lookup]
	if missing:
		raise ValueError("Selected columns were not found: " + ", ".join(missing))
	data = data.rename(columns={original: clean for clean, original in column_lookup.items()})
	data = data[selected].copy()
	data = data.rename(columns={response_column: "_response"})
	response_values = data["_response"]
	numeric_response = pd.to_numeric(response_values, errors="coerce")
	if numeric_response.notna().any():
		data["_response"] = numeric_response
	else:
		categories = response_values.astype("string").str.strip()
		if categories.dropna().nunique() != 2:
			raise ValueError("The response column must be numeric or contain exactly two categories.")
		codes = pd.Categorical(categories).codes.astype(float)
		data["_response"] = pd.Series(codes, index=data.index).replace(-1, float("nan"))
	if response_reference is not None:
		data["_response"] = (response_values != response_reference).astype(float)
		data.loc[response_values.isna(), "_response"] = float("nan")
	for column in predictor_columns:
		values = pd.to_numeric(data[column], errors="coerce")
		if values.notna().sum() == data[column].notna().sum():
			data[column] = values
		else:
			data[column] = data[column].astype("string").str.strip()
	return data, predictor_columns


def fit_model(data, predictor_columns, references=None, random_effect_column=None):
	import pandas as pd
	import numpy as np
	import statsmodels.api as sm
	import statsmodels.formula.api as smf

	references = references or {}
	model_columns = ["_response", *predictor_columns]
	if random_effect_column is not None:
		model_columns.append(random_effect_column)
	complete = data.dropna(subset=model_columns).copy()
	if len(complete) < 5:
		raise ValueError("At least five complete rows are needed for the model.")
	if complete["_response"].nunique() < 2:
		raise ValueError("The response column must contain at least two distinct values.")
	response_is_binary = complete["_response"].nunique() == 2
	if random_effect_column is not None and response_is_binary:
		raise ValueError("Random effects are currently supported only for continuous responses.")
	if response_is_binary:
		response_transform = "none (binary response)"
	else:
		if (complete["_response"] <= 0).any():
			raise ValueError(
				"Continuous response values must all be greater than zero for log transformation."
			)
		complete["_response"] = np.log(complete["_response"])
		response_transform = "natural log"
	predictor_terms = []
	for column in predictor_columns:
		term = f"Q({column!r})"
		if not pd.api.types.is_numeric_dtype(complete[column]):
			reference = references.get(column)
			if reference is None:
				term = f"C({term})"
			else:
				term = f"C({term}, Treatment(reference={reference!r}))"
		predictor_terms.append(term)
	formula = "_response ~ " + (" + ".join(predictor_terms) or "1")
	if response_is_binary:
		model = smf.glm(formula, data=complete, family=sm.families.Binomial()).fit()
		model_type = "binary"
	elif random_effect_column is not None:
		if complete[random_effect_column].nunique() < 2:
			raise ValueError("The random-effects variable must contain at least two groups.")
		model = smf.mixedlm(
			formula, data=complete, groups=complete[random_effect_column]
		).fit()
		model_type = "continuous mixed-effects"
	else:
		model = smf.ols(formula, data=complete).fit()
		model_type = "continuous"
	return complete, model, model_type, response_transform


def model_effect_table(model, odds_ratios: bool):
	import numpy as np
	import pandas as pd

	table = pd.DataFrame(
		{
			"term": model.params.index,
			"estimate": model.params.values,
			"standard_error": model.bse.values,
			"p_value": model.pvalues.values,
			"ci_low": model.conf_int()[0].values,
			"ci_high": model.conf_int()[1].values,
		}
	)
	if odds_ratios:
		def safe_exp(value):
			with np.errstate(over="ignore", invalid="ignore"):
				return np.exp(value)

		for column in ["estimate", "ci_low", "ci_high"]:
			table[column] = table[column].map(safe_exp)
		table = table.rename(columns={"estimate": "odds_ratio"})
	return table


def write_plot(data, model, model_type: str, output_dir: Path):
	import matplotlib.pyplot as plt
	fig, ax = plt.subplots(figsize=(8, 5.5))
	predictions = model.predict(data)
	ax.scatter(data["_response"], predictions, alpha=0.8, edgecolor="white")
	minimum = min(data["_response"].min(), predictions.min())
	maximum = max(data["_response"].max(), predictions.max())
	ax.plot([minimum, maximum], [minimum, maximum], linestyle="--", color="black")
	ax.set(title="Observed versus fitted response", xlabel="Observed response",
	       ylabel="Fitted response")
	if model_type == "binary":
		ax.set_ylim(-0.05, 1.05)
	fig.tight_layout()
	fig.savefig(output_dir / "model_fit.png", dpi=180)
	plt.close(fig)


def write_report(input_path, sheet_name, response_column, predictor_columns,
					 references, response_reference, random_effect_column, output_dir,
					 data, model, model_type,
				 response_transform):
	table = model_effect_table(model, odds_ratios=model_type == "binary")
	table.to_csv(output_dir / "model_coefficients.csv", index=False)
	if model_type == "binary":
		table.to_csv(output_dir / "model_odds_ratios.csv", index=False)

	def effect_sentence(table, term_column):
		effects = table[table["term"].str.contains(term_column, regex=False)]
		if effects.empty:
			return f"No estimable {term_column} effect was available."
		strongest = effects.loc[effects["p_value"].idxmin()]
		estimate_column = "odds_ratio" if "odds_ratio" in effects else "estimate"
		estimate_label = "odds ratio" if estimate_column == "odds_ratio" else "estimate"
		return (f"The strongest {term_column} term was {strongest['term']} "
			f"({estimate_label} {strongest[estimate_column]:.3g}, "
				f"p={strongest['p_value']:.3g}).")

	report = f"""Regression analysis
Input: {input_path}
Sheet: {sheet_name}
Response: {response_column}
Explanatory variables: {', '.join(predictor_columns)}
Categorical reference levels: {', '.join(f'{column}={value}' for column, value in references.items()) or 'default levels'}
Binary response reference (0): {response_reference if response_reference is not None else 'not applicable'}
Random-effects grouping variable: {random_effect_column or 'none'}

Analysis
The model includes all selected explanatory variables simultaneously.
Numeric responses use OLS, or a random-intercept mixed model when a grouping variable is selected.
Two-level responses use binomial logistic regression.
A smaller p-value indicates stronger evidence of an association, but effect sizes and
confidence intervals should also be considered, especially with small samples.

Model type: {model_type}
Response transformation: {response_transform}
Rows used: {len(data)}

Model fit
"""
	if model_type == "binary":
		report += f"Pseudo R-squared (CS): {model.pseudo_rsquared(kind='cs'):.3f}\n"
	elif model_type == "continuous mixed-effects":
		report += f"Log-likelihood: {model.llf:.3f}\nAIC: {model.aic:.3f}\n"
	else:
		report += (f"R-squared: {model.rsquared:.3f}\n"
				   f"Adjusted R-squared: {model.rsquared_adj:.3f}\n")
	report += f"""

Model summaries
{model.summary()}
"""
	(output_dir / "regression_report.txt").write_text(report, encoding="utf-8")


def parse_args() -> argparse.Namespace:
	parser = argparse.ArgumentParser(description=__doc__)
	parser.add_argument("--input", type=Path, help="Path to the Excel workbook")
	parser.add_argument("--output", type=Path, help="Parent folder for the new results folder")
	parser.add_argument("--sheet", help="Workbook sheet name")
	parser.add_argument("--response", help="Response column name")
	parser.add_argument("--predictors", nargs="+", help="Explanatory column names")
	return parser.parse_args()


def main() -> None:
	args = parse_args()
	input_path = args.input or choose_input_file()
	if input_path is None:
		print("No workbook selected.")
		return
	input_path = input_path.expanduser().resolve()
	try:
		import pandas as pd
		sheet_names = pd.ExcelFile(input_path).sheet_names
		sheet_name = args.sheet or ask_for_sheet(sheet_names)
		raw_data = pd.read_excel(input_path, sheet_name=sheet_name)
		columns = list(normalise_column_names(raw_data.columns))
		response_column = args.response or ask_for_columns(
			columns, "Enter the response column number or name: ", allow_multiple=False
		)[0]
		response_reference = ask_for_response_reference(raw_data, response_column)
		predictor_columns = args.predictors or ask_for_columns(
			columns, "Enter explanatory columns, separated by commas: ", allow_multiple=True
		)
		if response_column in predictor_columns:
			raise ValueError("The response column cannot also be an explanatory variable.")
		random_effect_column = ask_for_random_effect(predictor_columns)
		fixed_predictor_columns = [
			column for column in predictor_columns if column != random_effect_column
		]
		data, predictor_columns = load_data(
			input_path, sheet_name, response_column, predictor_columns,
			response_reference
		)
		references = ask_for_references(data, fixed_predictor_columns)
		data, model, model_type, response_transform = fit_model(
			data, fixed_predictor_columns, references, random_effect_column
		)
		parent = args.output or input_path.parent
		output_dir = Path(parent) / f"{safe_name(input_path.stem)}_regression_{datetime.now():%Y%m%d_%H%M%S_%f}"
		output_dir.mkdir(parents=True, exist_ok=False)
		write_plot(data, model, model_type, output_dir)
		write_report(input_path, sheet_name, response_column, fixed_predictor_columns,
					 references, response_reference, random_effect_column, output_dir,
					 data, model, model_type,
					 response_transform)
	except ImportError as exc:
		raise SystemExit(
			"Missing analysis dependency. Install the required packages with: "
			"python -m pip install pandas numpy scipy statsmodels matplotlib openpyxl"
		) from exc
	except (OSError, ValueError) as exc:
		raise SystemExit(f"Analysis could not be completed: {exc}") from exc
	print(f"Results written to: {output_dir}")


if __name__ == "__main__":
	main()
