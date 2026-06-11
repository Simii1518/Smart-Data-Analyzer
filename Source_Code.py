import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import matplotlib.pyplot as plt

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class DataAnalysisTool:
    """
    Professional Data Analysis Desktop Application
    ------------------------------------------------
    Features:
    - Read CSV / Excel files
    - Dataset Summary
    - Dynamic Column Detection
    - Interactive GroupBy Report Builder
    - Report Preview in GUI
    - Export Report (Excel / CSV)
    - Interactive Chart Builder
    - Chart Preview in GUI
    - Export Chart as PNG
    """

    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Data Analysis Tool")
        self.root.geometry("1400x850")
        self.root.minsize(1200, 750)

        # Data holders
        self.df = None
        self.report_df = None
        self.input_file = ""
        self.input_folder = ""

        self.text_columns = []
        self.numeric_columns = []

        self.chart_canvas = None
        self.current_figure = None

        self.create_widgets()

    # ==================================================
    # GUI DESIGN
    # ==================================================
    def create_widgets(self):

        # -------------------------
        # Title
        # -------------------------
        title_label = tk.Label(
            self.root,
            text="Corporate Data Analysis Tool",
            font=("Segoe UI", 18, "bold"),
            fg="navy"
        )
        title_label.pack(pady=10)

        # =========================
        # FILE FRAME
        # =========================
        file_frame = ttk.LabelFrame(self.root, text="File Selection")
        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_var = tk.StringVar()

        file_entry = ttk.Entry(
            file_frame,
            textvariable=self.file_var,
            width=100
        )
        file_entry.grid(row=0, column=0, padx=5, pady=10)

        browse_btn = ttk.Button(
            file_frame,
            text="Browse",
            command=self.browse_file
        )
        browse_btn.grid(row=0, column=1, padx=5)

        read_btn = ttk.Button(
            file_frame,
            text="Read",
            command=self.read_file
        )
        read_btn.grid(row=0, column=2, padx=5)

        # =========================
        # DATASET INFO FRAME
        # =========================
        info_frame = ttk.LabelFrame(self.root, text="Dataset Information")
        info_frame.pack(fill="x", padx=10, pady=5)

        self.info_text = tk.Text(
            info_frame,
            height=6,
            wrap="word"
        )
        self.info_text.pack(fill="x", padx=5, pady=5)

        # =========================
        # REPORT BUILDER FRAME
        # =========================
        report_frame = ttk.LabelFrame(self.root, text="Report Builder")
        report_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(report_frame, text="Group By Column").grid(
            row=0, column=0, padx=5, pady=5
        )

        self.groupby_combo = ttk.Combobox(
            report_frame,
            width=25,
            state="readonly"
        )
        self.groupby_combo.grid(row=0, column=1, padx=5)

        ttk.Label(report_frame, text="Aggregation").grid(
            row=0, column=2, padx=5
        )

        self.agg_combo = ttk.Combobox(
            report_frame,
            width=20,
            state="readonly",
            values=[
                "sum",
                "mean",
                "average",
                "max",
                "min",
                "count",
                "median"
            ]
        )
        self.agg_combo.grid(row=0, column=3, padx=5)

        ttk.Label(report_frame, text="Value Column").grid(
            row=0, column=4, padx=5
        )

        self.value_combo = ttk.Combobox(
            report_frame,
            width=25,
            state="readonly"
        )
        self.value_combo.grid(row=0, column=5, padx=5)

        preview_btn = ttk.Button(
            report_frame,
            text="Preview Report",
            command=self.preview_report
        )
        preview_btn.grid(row=0, column=6, padx=10)

        export_report_btn = ttk.Button(
            report_frame,
            text="Export Report",
            command=self.export_report
        )
        export_report_btn.grid(row=0, column=7, padx=5)

        # =========================
        # CHART BUILDER FRAME
        # =========================
        chart_frame = ttk.LabelFrame(self.root, text="Chart Builder")
        chart_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(chart_frame, text="Chart Type").grid(
            row=0,
            column=0,
            padx=5,
            pady=5
        )

        self.chart_combo = ttk.Combobox(
            chart_frame,
            width=25,
            state="readonly",
            values=[
                "Bar Chart",
                "Column Chart",
                "Line Chart",
                "Pie Chart"
            ]
        )
        self.chart_combo.grid(row=0, column=1, padx=5)

        preview_chart_btn = ttk.Button(
            chart_frame,
            text="Preview Chart",
            command=self.preview_chart
        )
        preview_chart_btn.grid(row=0, column=2, padx=10)

        export_chart_btn = ttk.Button(
            chart_frame,
            text="Export Chart",
            command=self.export_chart
        )
        export_chart_btn.grid(row=0, column=3, padx=5)

        # =========================
        # MAIN DISPLAY AREA
        # =========================
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # -------------------------
        # Report Area
        # -------------------------
        report_display_frame = ttk.LabelFrame(
            main_frame,
            text="Report Preview"
        )
        report_display_frame.pack(
            side="left",
            fill="both",
            expand=True,
            padx=5
        )

        self.report_tree = ttk.Treeview(report_display_frame)

        tree_scroll_y = ttk.Scrollbar(
            report_display_frame,
            orient="vertical",
            command=self.report_tree.yview
        )

        tree_scroll_x = ttk.Scrollbar(
            report_display_frame,
            orient="horizontal",
            command=self.report_tree.xview
        )

        self.report_tree.configure(
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set
        )

        self.report_tree.pack(fill="both", expand=True)

        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x.pack(side="bottom", fill="x")

        # -------------------------
        # Chart Area
        # -------------------------
        self.chart_display_frame = ttk.LabelFrame(
            main_frame,
            text="Chart Preview"
        )

        self.chart_display_frame.pack(
            side="right",
            fill="both",
            expand=True,
            padx=5
        )

    # ==================================================
    # FILE SELECTION
    # ==================================================
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if file_path:
            self.file_var.set(file_path)

    # ==================================================
    # READ FILE
    # ==================================================
    def read_file(self):

        file_path = self.file_var.get().strip()

        if not file_path:
            messagebox.showerror(
                "Error",
                "Please select a file first."
            )
            return

        try:
            ext = os.path.splitext(file_path)[1].lower()

            if ext == ".csv":
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)

            self.input_file = file_path
            self.input_folder = os.path.dirname(file_path)

            self.detect_columns()
            self.display_dataset_info()

            messagebox.showinfo(
                "Success",
                "File loaded successfully."
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Unable to read file.\n\n{e}"
            )

    # ==================================================
    # DETECT COLUMNS
    # ==================================================
    def detect_columns(self):

        self.text_columns = []

        self.numeric_columns = []

        for col in self.df.columns:

            if pd.api.types.is_numeric_dtype(self.df[col]):
                self.numeric_columns.append(col)

            else:
                converted = pd.to_numeric(
                    self.df[col],
                    errors="coerce"
                )

                ratio = converted.notna().mean()

                if ratio > 0.8:
                    self.numeric_columns.append(col)
                else:
                    self.text_columns.append(col)

        self.groupby_combo["values"] = self.text_columns
        self.value_combo["values"] = self.numeric_columns

    # ==================================================
    # DATASET INFO
    # ==================================================
    def display_dataset_info(self):

        self.info_text.delete("1.0", tk.END)

        rows, cols = self.df.shape

        info = (
            f"Total Rows    : {rows}\n"
            f"Total Columns : {cols}\n\n"
            f"Column Names:\n"
        )

        for col in self.df.columns:
            info += f"• {col}\n"

        self.info_text.insert(tk.END, info)

    # ==================================================
    # REPORT PREVIEW
    # ==================================================
    def preview_report(self):

        if self.df is None:
            messagebox.showerror(
                "Error",
                "Please read a file first."
            )
            return

        group_col = self.groupby_combo.get()
        agg_method = self.agg_combo.get()
        value_col = self.value_combo.get()

        if not group_col:
            messagebox.showerror(
                "Error",
                "Select Group By Column."
            )
            return

        if not agg_method:
            messagebox.showerror(
                "Error",
                "Select Aggregation Method."
            )
            return

        if not value_col:
            messagebox.showerror(
                "Error",
                "Select Value Column."
            )
            return

        try:

            temp_df = self.df.copy()

            temp_df[value_col] = pd.to_numeric(
                temp_df[value_col],
                errors="coerce"
            )

            if agg_method == "average":
                agg_method = "mean"

            result = (
                temp_df
                .groupby(group_col)[value_col]
                .agg(agg_method)
                .reset_index()
            )

            result = result.sort_values(
                by=value_col,
                ascending=False
            )

            self.report_df = result

            self.display_report(result)

        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e)
            )

    # ==================================================
    # DISPLAY REPORT
    # ==================================================
    def display_report(self, df):

        self.report_tree.delete(
            *self.report_tree.get_children()
        )

        self.report_tree["columns"] = list(df.columns)
        self.report_tree["show"] = "headings"

        for col in df.columns:
            self.report_tree.heading(
                col,
                text=col
            )

            self.report_tree.column(
                col,
                width=180,
                anchor="center"
            )

        for _, row in df.iterrows():
            self.report_tree.insert(
                "",
                tk.END,
                values=list(row)
            )

    # ==================================================
    # EXPORT REPORT
    # ==================================================
    def export_report(self):

        if self.report_df is None:
            messagebox.showerror(
                "Error",
                "Generate report first."
            )
            return

        export_type = messagebox.askquestion(
            "Export",
            "Click YES for Excel\nClick NO for CSV"
        )

        try:

            base_name = os.path.splitext(
                os.path.basename(self.input_file)
            )[0]

            if export_type == "yes":

                output = os.path.join(
                    self.input_folder,
                    f"{base_name}_Report.xlsx"
                )

                self.report_df.to_excel(
                    output,
                    index=False
                )

            else:

                output = os.path.join(
                    self.input_folder,
                    f"{base_name}_Report.csv"
                )

                self.report_df.to_csv(
                    output,
                    index=False
                )

            messagebox.showinfo(
                "Success",
                f"Report exported:\n{output}"
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e)
            )

    # ==================================================
    # PREVIEW CHART
    # ==================================================
    def preview_chart(self):

        if self.report_df is None:
            messagebox.showerror(
                "Error",
                "Generate report first."
            )
            return

        chart_type = self.chart_combo.get()

        if not chart_type:
            messagebox.showerror(
                "Error",
                "Select chart type."
            )
            return

        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()

        x_col = self.report_df.columns[0]
        y_col = self.report_df.columns[1]

        fig, ax = plt.subplots(
            figsize=(6, 5)
        )

        if chart_type == "Bar Chart":

            ax.barh(
                self.report_df[x_col],
                self.report_df[y_col]
            )

        elif chart_type == "Column Chart":

            ax.bar(
                self.report_df[x_col],
                self.report_df[y_col]
            )

        elif chart_type == "Line Chart":

            ax.plot(
                self.report_df[x_col],
                self.report_df[y_col],
                marker="o"
            )

        elif chart_type == "Pie Chart":

            ax.pie(
                self.report_df[y_col],
                labels=self.report_df[x_col],
                autopct="%1.1f%%"
            )

        if chart_type != "Pie Chart":
            ax.set_xlabel(x_col)
            ax.set_ylabel(y_col)
            ax.tick_params(axis='x', rotation=45)

        ax.set_title(chart_type)

        fig.tight_layout()

        self.current_figure = fig

        self.chart_canvas = FigureCanvasTkAgg(
            fig,
            master=self.chart_display_frame
        )

        self.chart_canvas.draw()

        self.chart_canvas.get_tk_widget().pack(
            fill="both",
            expand=True
        )

    # ==================================================
    # EXPORT CHART
    # ==================================================
    def export_chart(self):

        if self.current_figure is None:
            messagebox.showerror(
                "Error",
                "Preview chart first."
            )
            return

        try:

            base_name = os.path.splitext(
                os.path.basename(self.input_file)
            )[0]

            output = os.path.join(
                self.input_folder,
                f"{base_name}_Chart.png"
            )

            self.current_figure.savefig(
                output,
                dpi=300,
                bbox_inches="tight"
            )

            messagebox.showinfo(
                "Success",
                f"Chart exported:\n{output}"
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e)
            )


# ======================================================
# MAIN
# ======================================================
if __name__ == "__main__":

    root = tk.Tk()

    style = ttk.Style()
    style.theme_use("clam")

    app = DataAnalysisTool(root)

    root.mainloop()
