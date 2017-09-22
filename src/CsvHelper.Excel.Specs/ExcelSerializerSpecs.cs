
namespace CsvHelper.Excel.Specs
{
    using System;
    using System.IO;
    using ClosedXML.Excel;
    using Xunit;

    public class ExcelSerializerSpecs
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values = 
            {
                new Person { Name = "Bill", Age = 20 },
                new Person { Name = "Ben", Age = 20 },
                new Person { Name = "Weed", Age = 30 }
            };
            private XLWorkbook workbook;
            private IXLWorksheet worksheet;

            protected abstract string Path { get; }

            protected virtual string WorksheetName => "Export";
            
            protected virtual int StartRow => 1;

            protected virtual int StartColumn => 1;

            protected XLWorkbook Workbook => workbook ?? (workbook = Helpers.GetOrCreateWorkbook(Path, WorksheetName));

            protected IXLWorksheet Worksheet => worksheet ?? (worksheet = Workbook.GetOrAddWorksheet(WorksheetName));

            protected void Run(ExcelSerializer serialiser)
            {
                using (var writer = new CsvWriter(serialiser))
                {
                    writer.Configuration.AutoMap<Person>();
                    writer.WriteRecords(Values);
                }
            }

            [Fact]
            public void TheFileIsAValidExcelFile()
            {
                Assert.NotNull(Workbook);
            }

            [Fact]
            public void TheExcelWorkbookHeadersAreCorrect()
            {
                Assert.Equal(nameof(Person.Name), Worksheet.Row(StartRow).Cell(StartColumn).Value);
                Assert.Equal(nameof(Person.Age), Worksheet.Row(StartRow).Cell(StartColumn+1).Value);
            }

            [Fact]
            public void TheExcelWorkbookValuesAreCorrect()
            {
                for (int i = 0; i < Values.Length; i++)
                {
                    Assert.Equal(Values[i].Name, Worksheet.Row(StartRow+i+1).Cell(StartColumn).Value);
                    Assert.Equal((double)Values[i].Age, Worksheet.Row(StartRow+i+1).Cell(StartColumn+1).Value);
                }
            }
            
            public void Dispose()
            {
                 Workbook?.Dispose();
                 File.Delete(Path);
            }
        }
        
        public class SerialiseUsingPathSpec : Spec
        {
            public SerialiseUsingPathSpec()
            {
                using (var serialiser = new ExcelSerializer(Path))
                {
                    Run(serialiser);
                }
            }

            protected sealed override string Path => "serialise_by_path.xlsx";
        }

        public class SerialiseUsingPathWithOffsetsSpec : Spec
        {
            public SerialiseUsingPathWithOffsetsSpec()
            {
                using (var serialiser = new ExcelSerializer(Path) { ColumnOffset = StartColumn - 1, RowOffset = StartRow - 1})
                {
                    Run(serialiser);
                }
            }

            protected override int StartColumn => 5;

            protected override int StartRow => 5;

            protected sealed override string Path => "serialise_by_path_with_offsets.xlsx";
        }

        public class SerialiseUsingPathAndSheetnameSpec : Spec
        {
            public SerialiseUsingPathAndSheetnameSpec()
            {
                using (var serialiser = new ExcelSerializer(Path, WorksheetName))
                {
                    Run(serialiser);
                }
            }

            protected sealed override string Path => "serialise_by_path_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }

        public class SerialiseUsingWorkbookSpec : Spec
        {
            public SerialiseUsingWorkbookSpec()
            {
                using (var serialiser = new ExcelSerializer(Workbook))
                {
                    Run(serialiser);
                }
            }

            protected override string Path => "serialise_by_workbook.xlsx";
        }

        public class SerialiseUsingWorkbookAndSheetnameSpec : Spec
        {
            public SerialiseUsingWorkbookAndSheetnameSpec()
            {
                using (var serialiser = new ExcelSerializer(Workbook, WorksheetName))
                {
                    Run(serialiser);
                }
            }

            protected override string Path => "serialise_by_workbook_and_sheetname.xlsx";

            protected override string WorksheetName => "a_different_sheet_name";
        }

        public class SerialiseUsingWorksheetSpec : Spec
        {
            public SerialiseUsingWorksheetSpec()
            {
                using (var serialiser = new ExcelSerializer(Worksheet))
                {
                    Run(serialiser);
                }
            }

            protected override string Path => "serialise_by_worksheet.xlsx";

            protected override string WorksheetName => "a_different_sheetname";
        }

        public class SerialiseUsingRangeSpec : Spec
        {
            public SerialiseUsingRangeSpec()
            {
                var range = Worksheet.Range(StartRow, StartColumn, StartRow + Values.Length, StartColumn + 1);
                using (var serialiser = new ExcelSerializer(range))
                {
                    Run(serialiser);
                }
            }

            protected override int StartRow => 4;

            protected override int StartColumn => 8;

            protected override string Path => "serialise_by_range.xlsx";
        }
    }
}