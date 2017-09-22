
namespace CsvHelper.Excel.Specs
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using ClosedXML.Excel;
    using Xunit;

    public class ExcelParserSpecs
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values =
            {
                new Person { Name = "Bill", Age = 40 },
                new Person { Name = "Ben", Age = 30 },
                new Person { Name = "Weed", Age = 40 }
            };
            private XLWorkbook workbook;
            private IXLWorksheet worksheet;
            protected Person[] Results;

            protected Spec()
            {
                var workbook = Helpers.GetOrCreateWorkbook(Path, WorksheetName);
                var worksheet = workbook.GetOrAddWorksheet(WorksheetName);
                var headerRow = worksheet.Row(StartRow);
                headerRow.Cell(StartColumn).Value = nameof(Person.Name);
                headerRow.Cell(StartColumn + 1).Value = nameof(Person.Age);
                for (int i = 0; i < Values.Length; i++)
                {
                    var row = worksheet.Row(StartRow + i + 1);
                    row.Cell(StartColumn).Value = Values[i].Name;
                    row.Cell(StartColumn + 1).Value = Values[i].Age;
                }
                workbook.SaveAs(Path);
            }

            protected abstract string Path { get; }

            protected virtual string WorksheetName => "Export";

            protected virtual int StartRow => 1;

            protected virtual int StartColumn => 1;

            protected XLWorkbook Workbook => workbook ?? (workbook = Helpers.GetOrCreateWorkbook(Path, WorksheetName));

            protected IXLWorksheet Worksheet => worksheet ?? (worksheet = Workbook.GetOrAddWorksheet(WorksheetName));

            protected void Run(ExcelParser parser)
            {
                using (var reader = new CsvReader(parser))
                {
                    reader.Configuration.AutoMap<Person>();
                    Results = reader.GetRecords<Person>().ToArray();
                }
            }

            [Fact]
            public void TheResultsAreNotNull()
            {
                Assert.NotNull(Results);
            }

            [Fact]
            public void TheResultsAreCorrect()
            {
                Assert.Equal(Values, Results, EqualityComparer<Person>.Default);
            }

            public void Dispose()
            {
                Workbook?.Dispose();
                File.Delete(Path);
            }
        }

        public class ParseUsingPathSpec : Spec
        {
            public ParseUsingPathSpec()
            {
                using (var parser = new ExcelParser(Path))
                {
                    Run(parser);
                }
            }

            protected override string Path => "parse_by_path.xlsx";
        }

        public class ParseUsingPathWithOffsetsSpec : Spec
        {
            public ParseUsingPathWithOffsetsSpec()
            {
                using (var parser = new ExcelParser(Path) { ColumnOffset = StartColumn - 1, RowOffset = StartRow - 1})
                {
                    Run(parser);
                }
            }

            protected override int StartColumn => 5;

            protected override int StartRow => 5;

            protected override string Path => "parse_by_path_with_offset.xlsx";
        }

        public class ParseUsingPathAndSheetNameSpec : Spec
        {
            public ParseUsingPathAndSheetNameSpec()
            {
                using (var parser = new ExcelParser(Path, WorksheetName))
                {
                    Run(parser);
                }
            }

            protected override string Path => "parse_by_path_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }

        public class ParseUsingWorkbookSpec : Spec
        {
            public ParseUsingWorkbookSpec()
            {
                using (var parser = new ExcelParser(Workbook))
                {
                    Run(parser);
                }
            }

            protected override string Path => "parse_by_workbook.xlsx";
        }

        public class ParseUsingWorkbookAndSheetNameSpec : Spec
        {
            public ParseUsingWorkbookAndSheetNameSpec()
            {
                using (var parser = new ExcelParser(Workbook, WorksheetName))
                {
                    Run(parser);
                }
            }

            protected override string Path => "parse_by_workbook_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }

        public class ParseUsingWorksheetSpec : Spec
        {
            public ParseUsingWorksheetSpec()
            {
                using (var parser = new ExcelParser(Worksheet))
                {
                    Run(parser);
                }
            }

            protected override string Path => "parse_by_worksheet.xlsx";
        }

        public class ParseUsingRangeSpec : Spec
        {
            public ParseUsingRangeSpec()
            {
                var range = Worksheet.Range(StartRow, StartColumn, StartRow + Values.Length, StartColumn + 1);
                using (var parser = new ExcelParser(range))
                {
                    Run(parser);
                }
            }

            protected override int StartColumn => 5;

            protected override int StartRow => 4;

            protected override string Path => "parse_with_range.xlsx";
        }

        public class ParseWithFormulaSpec : Spec
        {
            public ParseWithFormulaSpec()
            {
                for (int i = 0; i < Values.Length; i++)
                {
                    var row = Worksheet.Row(2 + i);
                    row.Cell(2).FormulaA1 = $"=LEN({row.Cell(1).Address.ToStringFixed()})*10";
                }
                Workbook.SaveAs(Path);
                using (var parser = new ExcelParser(Path))
                {
                    Run(parser);
                }
            }

            protected override string Path => "parser_with_formula.xlsx";
        }
    }
}
