
namespace CsvHelper.Excel
{
    using System;
    using System.Text.RegularExpressions;
    using ClosedXML.Excel;
    using CsvHelper.Configuration;

    /// <summary>
    /// Defines methods used to serialize data into an Excel (2007+) file.
    /// </summary>
    public class ExcelSerializer : ICsvSerializer
    {
        private readonly string path;
        private readonly bool disposeWorkbook;
        private readonly IXLRangeBase range;
        private readonly bool disposeWorksheet;
        private bool disposed;
        private int currentRow = 1;

        /// <summary>
        /// Creates a new serializer using a new <see cref="XLWorkbook"/> saved to the given <paramref name="path"/>.
        /// <remarks>
        /// The workbook will not be saved until the serializer is disposed.
        /// </remarks>
        /// </summary>
        /// <param name="path">The path to which to save the workbook.</param>
        /// <param name="configuration">The configuration</param>
        public ExcelSerializer(string path, CsvConfiguration configuration = null) 
            : this(new XLWorkbook(XLEventTracking.Disabled), configuration)
        {
            this.path = path;
            disposeWorkbook = true;
        }

        /// <summary>
        /// Creates a new serializer using a new <see cref="XLWorkbook"/> saved to the given <paramref name="path"/>.
        /// <remarks>
        /// The workbook will not be saved until the serializer is disposed.
        /// </remarks>
        /// </summary>
        /// <param name="path">The path to which to save the workbook.</param>
        /// <param name="sheetName">The name of the sheet to which to save</param>
        public ExcelSerializer(string path, string sheetName) : this(new XLWorkbook(XLEventTracking.Disabled).AddWorksheet(sheetName))
        {
            this.path = path;
            disposeWorkbook = true;
        }
        
        /// <summary>
        /// Creates a new serializer using the given <see cref="XLWorkbook"/> and <see cref="CsvConfiguration"/>.
        /// <remarks>
        /// The <paramref name="workbook"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The workbook will <b><i>not</i></b> be saved by the serializer.
        /// A new worksheet will be added to the workbook.
        /// </remarks>
        /// </summary>
        /// <param name="workbook">The workbook to write the data to.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelSerializer(XLWorkbook workbook, CsvConfiguration configuration = null)
            : this(workbook, "Export", configuration)
        {
            disposeWorksheet = true;
        }

        /// <summary>
        /// Creates a new serializer using the given <see cref="XLWorkbook"/> and <see cref="CsvConfiguration"/>.
        /// <remarks>
        /// The <paramref name="workbook"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The workbook will <b><i>not</i></b> be saved by the serializer.
        /// A new worksheet will be added to the workbook.
        /// </remarks>
        /// </summary>
        /// <param name="workbook">The workbook to write the data to.</param>
        /// <param name="sheetName">The name of the sheet to write to.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelSerializer(XLWorkbook workbook, string sheetName, CsvConfiguration configuration = null)
            : this(workbook.GetOrAddWorksheet(sheetName), configuration)
        {
            disposeWorksheet = true;
        }
        
        /// <summary>
        /// Creates a new serializer using the given <see cref="IXLWorksheet"/>.
        /// <remarks>
        /// The <paramref name="worksheet"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The workbook will <b><i>not</i></b> be saved by the serializer.
        /// </remarks>
        /// </summary>
        /// <param name="worksheet">The worksheet to write the data to.</param>
        /// <param name="configuration">The configuration</param>
        public ExcelSerializer(IXLWorksheet worksheet, CsvConfiguration configuration = null) : this((IXLRangeBase)worksheet, configuration) { }
    
        /// <summary>
        /// Creates a new serializer using the given <see cref="IXLWorksheet"/>.
        /// </summary>
        /// <param name="range">The range to write the data to.</param>
        /// <param name="configuration">The configuration</param>
        public ExcelSerializer(IXLRange range, CsvConfiguration configuration = null) : this((IXLRangeBase)range, configuration) { }

        private ExcelSerializer(IXLRangeBase range, CsvConfiguration configuration)
        {
            Workbook = range.Worksheet.Workbook;
            this.range = range;
            Configuration = configuration ?? new CsvConfiguration();
            Configuration.QuoteNoFields = true;
        }

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public CsvConfiguration Configuration { get; }

        /// <summary>
        /// Gets the workbook to which the data is being written.
        /// </summary>
        /// <value>
        /// The workbook.
        /// </value>
        public XLWorkbook Workbook { get; }

        /// <summary>
        /// Gets and sets the number of rows to offset the start position from.
        /// </summary>
        public int RowOffset { get; set; } = 0;

        /// <summary>
        /// Gets and sets the number of columns to offset the start position from.
        /// </summary>
        public int ColumnOffset { get; set; } = 0;

        /// <summary>
        /// Writes a record to the Excel file.
        /// </summary>
        /// <param name="record">The record to write.</param>
        /// <exception cref="ObjectDisposedException">
        /// Thrown is the serializer has been disposed.
        /// </exception>
        public virtual void Write(string[] record)
        {
            CheckDisposed();
            for (var i = 0; i < record.Length; i++)
            {
                range.AsRange().Cell(currentRow + RowOffset, i + 1 + ColumnOffset).Value = ReplaceHexadecimalSymbols(record[i]);
            }
            currentRow++;
        }

        /// <summary>
        /// Replaces the hexadecimal symbols.
        /// </summary>
        /// <param name="text">The text to replace.</param>
        /// <returns>The input</returns>
        protected static string ReplaceHexadecimalSymbols(string text)
        {
            if (String.IsNullOrEmpty(text)) return text;
            return Regex.Replace(text, "[\x00-\x08\x0B\x0C\x0E-\x1F]", String.Empty, RegexOptions.Compiled);
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
        /// Finalizes an instance of the <see cref="ExcelSerializer"/> class.
        /// </summary>
        ~ExcelSerializer()
        {
            Dispose(false);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) return;
            if (disposing)
            {
                if (disposeWorksheet) range.Worksheet.Dispose();
                if (disposeWorkbook)
                {
                    Workbook.SaveAs(path);
                    Workbook.Dispose();
                }
            }
            disposed = true;
        }

        /// <summary>
        /// Checks if the instance has been disposed of.
        /// </summary>
        /// <exception cref="ObjectDisposedException">
        /// Thrown is the serializer has been disposed.
        /// </exception>
        protected virtual void CheckDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(GetType().ToString());
            }
        }
    }
}
