/*Copyright (c) 2016 Flynet Ltd.

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.
*/

// Represents a single column in the multi row.
function Column(name, start, length, numberFormat, columnWidth, formatFn) {
    // Excel column width in characters.
    this.ColumnWidth = columnWidth;

    // Fn called by TextForRow() to format text.
    this.FormatFunction = formatFn;

    // Max length of an entry in this column on the screen.
    this.Length = length;

    // The name of the column.
    this.Name = name;

    // Excel number format for this column.
    this.NumberFormat = numberFormat;

    // Index of this column in the worksheet.
    this.SheetColumn = 0;

    // Index of 1st character of this column on the screen.
    this.Start = start;

    // Returns a flag indicating if this column has either a number format or column width set.
    this.HasColumnSettings = function () { return NumberFormat || ColumnWidth; };

    // Gets the text for this column on the given row. Text is trimmed and modified by the format function if set.
    this.TextForRow = function (row) { var text = GetTrimmedText(row, this.Start, this.Length); if (this.FormatFunction) { text = this.FormatFunction(text); } return text; };
}

// Represents the more indicator on a multi row.
function MoreIndicator(row, column, text) {
    // The index of the column where the more indicator text starts.
    this.Column = column;

    // The length of the more indicator text (calculated).
    this.Length = text.length;

    // The row on which the more indicator appears.
    this.Row = row;

    // The text of the more indicator.
    this.Text = text;

    // Returns a flag indicating if there are more rows beyond the current screen.
    this.HasMoreRows = function () { return GetTrimmedText(this.Row, this.Column, this.Length) == this.Text; };
}

// Represents a multi row on the screen.
function MultiRow(headerRow, firstRow, lastRow, pageDownKey, pageUpKey) {
    // The columns.
    this.Columns = [];

    // The index of the 1st row of DATA on the screen.
    this.FirstRow = firstRow;

    // The index of the header row on the screen.
    this.HeaderRow = headerRow;

    // The index of the last row of DATA on the screen.
    this.LastRow = lastRow;

    // The key used to page down to get the next page of data.
    this.PageDownKey = pageDownKey;

    // the key used to page up to get the previous page of data.
    this.PageUpKey = pageUpKey;

    // Adds a column to the columns collection.
    this.AddColumn = function (name, start, length, numberFormat, columnWidth, formatFn) { this.Columns.push(new Column(name, start, length, numberFormat, columnWidth, formatFn)); };

    // Gets a column with the given name or null if no such column exists.
    this.GetColumn = function (name) {
        for (var index = 0; index < this.Columns.length; ++index) {
            var column = this.Columns[index];

            if (column.Name == name) {
                return column;
            }
        }

        return null;
    };

    // Returns a flag indicating if there are more rows beyond the current screen.
    this.HasMoreRows = function () { return this.MoreIndicator.HasMoreRows(); };

    // Pages down to get the next page of data.
    this.PageDown = function () { fvt.SendKeys(this.PageDownKey); };

    // Sets the more indicator.
    this.SetMoreIndicator = function (row, column, text) { this.MoreIndicator = new MoreIndicator(row, column, text); };
}