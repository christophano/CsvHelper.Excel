
namespace CsvHelper.Excel
{
    using System;
    using System.Linq;
    using ClosedXML.Excel;
    using CsvHelper.Configuration;

    /// <summary>
    /// Parses an Excel file.
    /// </summary>
    public class ExcelParser : ICsvParser
    {
        private readonly bool shouldDisposeWorkbook;
        private readonly IXLRangeBase range;
        private bool isDisposed;

        /// <summary>
        /// Creates a new parser using a new <see cref="XLWorkbook"/> from the given <paramref name="path"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, CsvConfiguration configuration = null)
            : this(new XLWorkbook(path, XLEventTracking.Disabled), configuration)
        {
            shouldDisposeWorkbook = true;
        }

        /// <summary>
        /// Creates a new parser using a new <see cref="XLWorkbook"/> from the given <paramref name="path"/> and uses the given <paramref name="configuration"/>.
        /// </summary>
        /// <param name="path">The path to the workbook.</param>
        /// <param name="sheetName">The name of the sheet to import data from.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, string sheetName, CsvConfiguration configuration = null)
            : this(new XLWorkbook(path, XLEventTracking.Disabled), sheetName, configuration)
        {
            shouldDisposeWorkbook = true;
        }

        /// <summary>
        /// Creates a new parser using the given <see cref="XLWorkbook"/> and <see cref="CsvConfiguration"/>.
        /// <remarks>
        /// Will attempt to read the data from the first worksheet in the workbook.
        /// </remarks>
        /// </summary>
        /// <param name="workbook">The <see cref="XLWorkbook"/> with the data.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(XLWorkbook workbook, CsvConfiguration configuration = null) : this(workbook.Worksheets.First(), configuration) { }

        /// <summary>
        /// Creates a new parser using the given <see cref="XLWorkbook"/> and <see cref="CsvConfiguration"/>.
        /// <remarks>
        /// Will attempt to read the data from the first worksheet in the workbook.
        /// </remarks>
        /// </summary>
        /// <param name="workbook">The <see cref="XLWorkbook"/> with the data.</param>
        /// <param name="sheetName">The name of the sheet to import from.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(XLWorkbook workbook, string sheetName, CsvConfiguration configuration = null) : this(workbook.Worksheet(sheetName), configuration) { }

        /// <summary>
        /// Creates a new parser using the given <see cref="IXLWorksheet"/> and <see cref="CsvConfiguration"/>.
        /// </summary>
        /// <param name="worksheet">The <see cref="IXLWorksheet"/> with the data.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(IXLWorksheet worksheet, CsvConfiguration configuration = null) : this((IXLRangeBase)worksheet, configuration) { }

        /// <summary>
        /// Creates a new parser using the given <see cref="IXLRange"/> and <see cref="CsvConfiguration"/>.
        /// </summary>
        /// <param name="range">The <see cref="IXLRange"/> with the data.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(IXLRange range, CsvConfiguration configuration = null) : this((IXLRangeBase)range, configuration) { }

        private ExcelParser(IXLRangeBase range, CsvConfiguration configuration)
        {
            Workbook = range.Worksheet.Workbook;
            this.range = range;
            Configuration = configuration ?? new CsvConfiguration();
            FieldCount = range.CellsUsed().Max(cell => cell.Address.ColumnNumber) - range.CellsUsed().Min(cell => cell.Address.ColumnNumber) + 1;
        }

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public CsvConfiguration Configuration { get; }

        /// <summary>
        /// Gets the workbook from which we are reading data.
        /// </summary>
        /// <value>
        /// The workbook.
        /// </value>
        public XLWorkbook Workbook { get; }

        /// <summary>
        /// Gets the field count.
        /// </summary>
        public int FieldCount { get; }

        /// <summary>
        /// Gets the character position that the parser is currently on.
        /// <remarks>This feature is unused.</remarks>
        /// </summary>
        public long CharPosition => -1;

        /// <summary>
        /// Gets the byte position that the parser is currently on.
        /// <remarks>This feature is unused.</remarks>
        /// </summary>
        public long BytePosition => -1;

        /// <summary>
        /// Gets the row of the Excel file that the parser is currently on.
        /// </summary>
        public int Row { get; private set; } = 1;

        /// <summary>
        /// Gets and sets the number of rows to offset the start position from.
        /// </summary>
        public int RowOffset { get; set; } = 0;

        /// <summary>
        /// Gets and sets the number of columns to offset the start position from.
        /// </summary>
        public int ColumnOffset { get; set; } = 0;

        /// <summary>
        /// Gets the raw row for the current record that was parsed.
        /// </summary>
        public virtual string RawRecord => range.AsRange().Row(Row).Cells(1, FieldCount).ToString();

        /// <summary>
        /// Reads a record from the Excel file.
        /// </summary>
        /// <returns>
        /// A <see cref="T:String[]" /> of fields for the record read.
        /// </returns>
        /// <exception cref="ObjectDisposedException">Thrown if the parser has been disposed.</exception>
        public virtual string[] Read()
        {
            CheckDisposed();
            var row = range.AsRange().Row(Row + RowOffset);
            if (row.CellsUsed().Any())
            {
                var result = row.Cells(1 + ColumnOffset, FieldCount + ColumnOffset)
                    .Select(cell => cell.Value.ToString())
                    .ToArray();
                Row++;
                return result;
            }
            return null;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        ~ExcelParser()
        {
            Dispose(false);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (isDisposed) return;
            if (disposing)
            {
                if (shouldDisposeWorkbook) Workbook.Dispose();
            }
            isDisposed = true;
        }

        /// <summary>
        /// Checks if the instance has been disposed of.
        /// </summary>
        /// <exception cref="ObjectDisposedException" />
        protected virtual void CheckDisposed()
        {
            if (isDisposed)
            {
                throw new ObjectDisposedException(GetType().ToString());
            }
        }
    }
}
