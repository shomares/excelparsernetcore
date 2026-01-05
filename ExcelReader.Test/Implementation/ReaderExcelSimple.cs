namespace ExcelReader.Test.Implementation
{
    public class ReaderExcelSimple
    {

        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileAsync()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simpletwo.xlsx", "sheet1");
            // Act
            var rows = implementation.GetNextRow().ToList();
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
            await implementation.ReadFileAsync("Files/missingcolumns.xlsx", "sheet1");
            // Act
            var rows = implementation.GetNextRow().ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(3));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.A, Is.EqualTo("Hello"));
            Assert.That(firstRow.B, Is.EqualTo("World"));
            Assert.That(firstRow.C1, Is.EqualTo("Missing"));
            Assert.That(firstRow.D, Is.EqualTo("12"));
            dynamic secondRow = rows[1];
            Assert.That(secondRow.A, Is.EqualTo("Foo"));
            Assert.That(secondRow.B, Is.EqualTo("Bar"));
            Assert.That(secondRow.D, Is.EqualTo("123"));
        }


        [Test]
        public async Task ReaderExcelSimple_ReadSimpleFileSecondPage()
        {
            // Arrange
            var implementation = new src.Implementation.ReaderExcelSimple();
            await implementation.ReadFileAsync("Files/simple.xlsx", "sheet2");
            // Act
            var rows = implementation.GetNextRow().ToList();
            // Assert
            Assert.That(rows.Count, Is.EqualTo(3));
            dynamic firstRow = rows[0];
            Assert.That(firstRow.A, Is.EqualTo("Hello"));
            Assert.That(firstRow.B, Is.EqualTo("World"));
            Assert.That(firstRow.C1, Is.EqualTo("Missing"));
            Assert.That(firstRow.D, Is.EqualTo("12"));
            dynamic secondRow = rows[1];
            Assert.That(secondRow.A, Is.EqualTo("Foo"));
            Assert.That(secondRow.B, Is.EqualTo("Bar"));
            Assert.That(secondRow.D, Is.EqualTo("123"));
        }
    }
}
