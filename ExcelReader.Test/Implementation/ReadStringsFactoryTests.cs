using ExcelReader.src.Config;
using ExcelReader.src.Implementation;

namespace ExcelReader.Test.Implementation
{
    internal class ReadStringsFactoryTests
    {
        [Test]
        public void TestReaderStringConfigurationBigFileNull()
        {
            var implementation = new ReadStringsFactory();

            var result = implementation.CreateReadStrings("Files/bigfile.xlsx");

            Assert.That(result, Is.InstanceOf(typeof(StreamReadStrings)));
        }

        [Test]
        public void TestReaderStringConfigurationNull()
        {
            var implementation = new ReadStringsFactory();

            var result = implementation.CreateReadStrings("Files/empty.xlsx");

            Assert.That(result, Is.InstanceOf(typeof(MemoryReadStrings)));
        }


        [Test]
        public void TestReaderStringConfigurationTrue()
        {
            var implementation = new ReadStringsFactory();

            var result = implementation.CreateReadStrings("Files/empty.xlsx", new ConfigurationReader
            {
                UseMemoryForStrings = true
            });

            Assert.That(result, Is.InstanceOf(typeof(MemoryReadStrings)));
        }

        [Test]
        public void TestReaderStringConfigurationFalse()
        {
            var implementation = new ReadStringsFactory();

            var result = implementation.CreateReadStrings("Files/empty.xlsx", new ConfigurationReader
            {
                UseMemoryForStrings = false
            });

            Assert.That(result, Is.InstanceOf(typeof(StreamReadStrings)));
        }
    }
}
