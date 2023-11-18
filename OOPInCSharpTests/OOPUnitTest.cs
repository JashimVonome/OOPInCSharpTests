using OfficeOpenXml;
using OOPInCSharp;

namespace OOPInCSharpTests
{
    public class OOPUnitTest
    {
        [Fact]
        public void CalculateSalary_shouldReturnCorrectSalary_ForManager()
        {
            // Arrange
            var manager = new Manager("Jashim Uddin");

            // Act
            decimal salary = manager.CalculateSalary();

            // Assert
            Assert.Equal(5000, salary);
        }
        [Fact]
        public void CalculateSalary_shouldReturnCorrectSalary_ForDeveloper()
        {
            // Arrange
            var developer = new Developer("Jamal Uddin");

            // Act
            decimal salary = developer.CalculateSalary();

            // Assert
            Assert.Equal(4000, salary);
        }

        [Fact]
        public void CalculateSalary_shouldReturnCorrectSalary_ForTester()
        {
            // Arrange
            var tester = new Tester("Jamal Uddin");

            // Act
            decimal salary = tester.CalculateSalary();

            // Assert
            Assert.Equal(3500, salary);
        }

        [Fact]
        public void CalculateSalary_shouldReturnCorrectSalary_ForAnalyst()
        {
            // Arrange
            var analyst = new Analyst("Kamal Uddin");

            // Act
            decimal salary = analyst.CalculateSalary();

            // Assert
            Assert.Equal(4500, salary);
        }

        [Fact]
        public void ExportSalarySheetToExcel_ShouldCreateExcelFileWithCorrectData()
        {
            // Arrange
             var employees= new List<Employee>();

            new Manager("Jashim Uddin");
            new Developer("Jamal Uddin");
            new Tester("Jamal Uddin");
            new Analyst("Kamal Uddin");

            string fileName = "SalarySheet.xlsx";

            // Act
            SalarySheetExporter.ExportSalarySheetToExcel(employees, fileName);

            // Assert
            using (var package = new ExcelPackage(new FileInfo(fileName)))
            {
                var worksheet = package.Workbook.Worksheets["SalarySheet"];

                // Verify the column headers
                Assert.Equal("Name", worksheet.Cells[1,1].Value.ToString());
                Assert.Equal("Designation", worksheet.Cells[1,2].Value.ToString());
                Assert.Equal("Salary", worksheet.Cells[1,3].Value.ToString());

                // Verify the employee data
                for (int row = 2; row <= employees.Count + 1  ; row++)
                {
                    var employee = employees[row - 2];
                    Assert.Equal(employee.Name, worksheet.Cells[row, 1].Value.ToString());
                    Assert.Equal(employee.Designation.ToString(), worksheet.Cells[row, 2].Value.ToString());
                    Assert.Equal(employee.CalculateSalary(), decimal.Parse(worksheet.Cells[row, 3].Value.ToString()));
                }
            }
        }
    }
}