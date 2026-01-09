namespace ExcelReader.Test.Implementation
{
    public class ReaderExcelSimple
    {

        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileAsync()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simpletwo.xlsx");
            // Act
            var rows = implementation.GetNextRow("sheet1").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(2));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.A, Is.EqualTo("Hello"));
            Assert.That(firstRow.B, Is.EqualTo("World"));
            dynamic secondRow = rows[1];
            Assert.That(secondRow.A, Is.EqualTo("Foo"));
            Assert.That(secondRow.B, Is.EqualTo("Bar"));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileMissingColumnsAsync()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/missingcolumns.xlsx");
            // Act
            var rows = implementation.GetNextRow("sheet1").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(3));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.A, Is.EqualTo("Hello"));
            Assert.That(firstRow.B, Is.EqualTo("World"));
            Assert.That(firstRow.C1, Is.EqualTo("Missing"));
            Assert.That(firstRow.D, Is.EqualTo(12));
            dynamic secondRow = rows[1];
            Assert.That(secondRow.A, Is.EqualTo("Foo"));
            Assert.That(secondRow.B, Is.EqualTo("Bar"));
            Assert.That(secondRow.D, Is.EqualTo(123));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileSecondPage()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simple.xlsx");
            // Act
            var rows = implementation.GetNextRow("sheet2").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(3));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.A, Is.EqualTo("Hello"));
            Assert.That(firstRow.B, Is.EqualTo("World"));
            Assert.That(firstRow.C1, Is.EqualTo("Missing"));
            Assert.That(firstRow.D, Is.EqualTo(12));
            dynamic secondRow = rows[1];
            Assert.That(secondRow.A, Is.EqualTo("Foo"));
            Assert.That(secondRow.B, Is.EqualTo("Bar"));
            Assert.That(secondRow.D, Is.EqualTo(123));
        }

        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileWithNoColumsn()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simple.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = false,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet4").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(3));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.AC1, Is.EqualTo("Hello"));
            dynamic secondRow = rows[1];
            Assert.That(secondRow.AC1, Is.EqualTo("Foo"));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileWithNumbers()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/numbers.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = false,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet1").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(4));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.A1, Is.EqualTo(12));
            dynamic secondRow = rows[1];
            Assert.That(secondRow.A1, Is.EqualTo(13));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileWithDate()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simple.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet5").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(4));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.Name, Is.EqualTo("Jhon"));
            Assert.That(firstRow.Employee_Date, Is.EqualTo(new DateTime(1984, 06, 13)));
            Assert.That(firstRow.Salary, Is.EqualTo(1200));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileWithDateAsTable()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simple.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet6").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(4));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.Name, Is.EqualTo("Jhon"));
            Assert.That(firstRow.Employee_Date, Is.EqualTo(new DateTime(1984, 06, 13)));
            Assert.That(firstRow.Salary, Is.EqualTo(1200));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileWithDateAsFormula()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simple.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet7").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(4));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.Name, Is.EqualTo("Jhon"));
            Assert.That(firstRow.Employee_Date, Is.EqualTo(new DateTime(1984, 06, 13)));
            Assert.That(firstRow.Salary, Is.EqualTo(1200));
            Assert.That(firstRow.Tax, Is.EqualTo(1380));

            var lastRow = rows[3];
            Assert.That(lastRow.Name, Is.EqualTo("Peter"));
            Assert.That(lastRow.Salary, Is.EqualTo(12044));
            Assert.That(lastRow.Tax, Is.EqualTo(13850.599999999999d));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadBigFile()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/bigfile.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = false
            });
            // Act
            var rows = implementation.GetNextRow("sheet1").ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(1000020));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.ip_address, Is.EqualTo("86.129.5.175"));
        }

        [Test]
        public async Task ReaderExcelSimple_ReadBigFileWithLinq()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/bigfile.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet1").
                Where(r => r.gender == "Male")
                .ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(503007));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.ip_address, Is.EqualTo("59.148.21.186"));
        }

        [Test]
        public async Task ReaderExcelSimple_ReadBigFileWithLinqOnly20()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/bigfile.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet1")
                .Take(20)
                .ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(20));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadBigFileWithLinqGroupBy()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/bigfile.xlsx", new src.Config.ConfigurationReader
            {
                HasHeaders = true,
                UseMemoryForStrings = true
            });
            // Act
            var rows = implementation.GetNextRow("sheet1")
                 .GroupBy(it => it.gender)
                 .Select(it => new
                 {
                     it.Key,
                     Count = it.Count()
                 })
                .ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(2));
         
        }

    }
}
