using JsonToExcel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace JsonToExcel.Test
{
    [TestClass]
    public class JsonConverterTests
    {
        [TestMethod]
        public void ToExcel_NoListMode()
        {
            // Arrange
            var jsonConverter = new JsonConverter(new ConvertOptions { ListMode = false });
            string path = @"";

            // Act
            jsonConverter.ToExcelFile(
               path, "test.xlsx");

            // Assert
            //Assert.Fail();
        }

        [TestMethod]
        public void ToExcel_ListMode()
        {
            // Arrange
            var jsonConverter = new JsonConverter(new ConvertOptions { ListMode = true });
            string path = @"";

            // Act
            jsonConverter.ToExcelFile(
               path, "test.xlsx");

            // Assert
            //Assert.Fail();
        }
    }
}
