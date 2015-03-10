
namespace CsvHelper.Excel
{
    using System;
    using ClosedXML.Excel;
    using CsvHelper.Configuration;

    /// <summary>
    /// Defines methods used to serialize data into an Excel (2007+) file.
    /// </summary>
    public class ExcelSerializer : ICsvSerializer
    {
        private readonly string path;
        private readonly XLWorkbook workbook;
        private readonly bool disposeWorkbook;
        private readonly IXLWorksheet worksheet;
        private readonly bool disposeWorksheet;
        private readonly CsvConfiguration configuration;
        private bool disposed;
        private int currentRow = 1;

        /// <summary>
        /// Creates a new serializer using a new <see cref="XLWorkbook"/> saved to the given <paramref name="path"/>.
        /// <remarks>
        /// The workbook will not be saved until the serializer is disposed.
        /// </remarks>
        /// </summary>
        /// <param name="path">The path to which to save the workbook.</param>
        public ExcelSerializer(string path) : this(new XLWorkbook(XLEventTracking.Disabled))
        {
            this.path = path;
            disposeWorkbook = true;
        }

        /// <summary>
        /// Creates a new serializer using the given <see cref="XLWorkbook"/>.
        /// <remarks>
        /// The <paramref name="workbook"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The workbook will <b><i>not</i></b> be saved by the serializer.
        /// A new worksheet will be added to the workbook.
        /// </remarks>
        /// </summary>
        /// <param name="workbook">The workbook to write the data to.</param>
        public ExcelSerializer(XLWorkbook workbook) : this(workbook, new CsvConfiguration()) { }

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
        public ExcelSerializer(XLWorkbook workbook, CsvConfiguration configuration)
            : this(workbook.AddWorksheet("Export"), configuration)
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
        public ExcelSerializer(IXLWorksheet worksheet) : this(worksheet, new CsvConfiguration()) { }

        /// <summary>
        /// Creates a new serializer using the given <see cref="IXLWorksheet"/>.
        /// <remarks>
        /// The <paramref name="worksheet"/> will <b><i>not</i></b> be disposed of when the serializer is disposed.
        /// The workbook will <b><i>not</i></b> be saved by the serializer.
        /// </remarks>
        /// </summary>
        /// <param name="worksheet">The worksheet to write the data to.</param>
        /// <param name="configuration">The configuration</param>
        public ExcelSerializer(IXLWorksheet worksheet, CsvConfiguration configuration)
        {
            workbook = worksheet.Workbook;
            this.worksheet = worksheet;
            this.configuration = configuration;
        }

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public CsvConfiguration Configuration { get { return configuration; } }

        /// <summary>
        /// Gets the workbook to which the data is being written.
        /// </summary>
        /// <value>
        /// The workbook.
        /// </value>
        public XLWorkbook Workbook { get { return workbook; } }

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
                worksheet.Cell(currentRow, i + 1).Value = record[i];
            }
            currentRow++;
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
                if (disposeWorksheet) worksheet.Dispose();
                if (disposeWorkbook)
                {
                    workbook.SaveAs(path);
                    workbook.Dispose();
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