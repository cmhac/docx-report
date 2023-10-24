"""Generates a docx report."""

from datetime import datetime
import os
from typing import Optional
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH  # pylint: disable=no-name-in-module
import pandas as pd


class DocxReport:
    """Generates a docx report."""

    def __init__(self, title: str) -> None:
        self.doc = docx.Document()
        self.title = title
        self._add_heading()

    def _add_heading(self) -> None:
        """draws the heading for the report"""
        # add subtitle above the header
        subtitle_p = self.doc.add_paragraph()
        subtitle_p.add_run(
            f"Generated on {datetime.now().strftime('%B %-d, %-I:%-M %p')}"
        )
        subtitle_p.style = "Subtitle"
        # add the title
        self.doc.add_heading(self.title, 0)

    @staticmethod
    def _cleanup_dataframe(
        df: pd.DataFrame,
        round_numeric: bool = True,
        round_decimals: int = 1,
        auto_format_dates: bool = True,
        rename_cols: Optional[dict] = None,
        strftime_format: str = "%Y-%m-%d",
    ) -> pd.DataFrame:
        """cleans up a dataframe to be ready for plotting

        Args:
            df: the dataframe to clean up
            round_numeric: whether to round numeric columns to 2 decimal places
            round_decimals: how many decimal places to round to
            auto_format_dates: whether to automatically format dates
            rename_cols: a dictionary of columns to rename
            strftime_format: the format to use when converting dates to strings

        Returns:
            the cleaned up dataframe
        """
        # rename the columns
        if rename_cols:
            df = df.rename(columns=rename_cols)

        # clean up the names
        df = df.clean_names().rename(columns=lambda x: x.replace("_", " "))

        # round all floats to 2 decimal places
        if round_numeric:
            df = df.applymap(
                lambda x: round(x, round_decimals) if isinstance(x, float) else x
            )

        # automatically clean up dates
        if auto_format_dates:
            # find all columns that are dates
            date_cols = df.select_dtypes(include="datetime").columns
            for date_col in date_cols:
                # convert to strings
                df[date_col] = df[date_col].dt.strftime(strftime_format)

        return df

    def _center_last_paragraph(self) -> None:
        """centers the last paragraph in the doc"""
        last_paragraph = self.doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def add_plot(
        self,
        df: pd.DataFrame,
        title: str,
        x_label: str,
        y_label: str,
        rename_cols: Optional[dict] = None,
        **kwargs,
    ) -> None:
        """uses matplotlib to plot a dataframe, then adds it to the docx file

        Args:
            df: the dataframe to plot
            title: the title of the plot
            x_label: the label for the x axis
            y_label: the label for the y axis
            rename_cols: a dictionary of columns to rename
            **kwargs: keyword arguments to pass to the plot function

        Returns:
            None
        """
        df = self._cleanup_dataframe(
            df, round_numeric=False, auto_format_dates=False, rename_cols=rename_cols
        )
        # create the plot
        ax = df.plot(**kwargs)
        ax.set_title(title)
        ax.set_xlabel(x_label)
        ax.set_ylabel(y_label)
        # save the plot as a png
        ax.get_figure().savefig("temp.png")
        # add the plot to the docx file
        self.doc.add_picture("temp.png", width=docx.shared.Inches(5))
        # center the image
        self._center_last_paragraph()
        # delete the temp png
        os.remove("temp.png")

    def add_table(
        self,
        df: pd.DataFrame,
        include_index: bool = True,
        rename_cols: dict = None,
        pct_cols: list = None,
    ) -> None:
        """turns a dataframe into a table in the document

        Args:
            df: the dataframe to turn into a table
            include_index: whether to include the index as a column
            rename_cols: a dictionary of columns to rename
            pct_cols: a list of columns to format as percentages

        Returns:
            None
        """
        # if include_index, reset the index so it's a column
        if include_index:
            df = df.reset_index()

        # save the original column names for later to check against pct_cols
        original_cols = df.columns.tolist()

        # do the cleanup
        df = self._cleanup_dataframe(df, rename_cols=rename_cols)

        # create the table based on the size of the dataframe
        rows = df.shape[0] + 1  # add 1 for the header
        cols = df.shape[1]
        table = self.doc.add_table(rows=rows, cols=cols)
        # set the style
        table.style = "TableGrid"

        # add the header
        header_cells = table.rows[0].cells
        for col_index, col in enumerate(df.columns):
            header_cells[col_index].text = col

        # add the data
        for row_index, row in df.iterrows():
            row_cells = table.rows[row_index + 1].cells
            for value_index, value in enumerate(row):
                # if value is numeric, add commas
                if isinstance(value, (int, float)):
                    # get the key from the original column names
                    # to check against pct_cols
                    key = original_cols[value_index]
                    # format percentages if needed
                    if pct_cols and key in pct_cols:
                        value = f"{value:.1%}"
                    else:
                        value = f"{value:,}"
                else:
                    value = str(value)
                row_cells[value_index].text = value
        return df

    def add_list_bullet(self, text: str) -> None:
        """adds a bullet point to the document"""
        return self.doc.add_paragraph(text, style="List Bullet")

    def save(self, filename: str) -> None:
        """saves the docx file"""
        self.doc.save(filename)
