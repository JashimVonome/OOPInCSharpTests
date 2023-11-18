using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;

namespace OOPInCSharp
{
    public enum Designation
    {
        Manager,
        Developer,
        Tester,
        Analyst
    }

    public interface ISalaryCalculator
    {
        decimal CalculateSalary();
    }

    public abstract class Employee
    {
        public string Name { get; set; }
        public Designation Designation { get; set; }

        public Employee(string name, Designation designation)
        {
            Name = name;
            Designation = designation;
        }

        public virtual decimal CalculateSalary()
        {
            return 0;
        }
    }

    public class Manager : Employee, ISalaryCalculator
    {
        public Manager(string name) : base(name, Designation.Manager)
        {
        }

        public override decimal CalculateSalary()
        {
            return 5000;
        }
    }

    public class Developer : Employee, ISalaryCalculator
    {
        public Developer(string name) : base(name, Designation.Developer)
        {
        }

        public override decimal CalculateSalary()
        {
            return 4000;
        }
    }

    public class Tester : Employee, ISalaryCalculator
    {
        public Tester(string name) : base(name, Designation.Tester)
        {
        }

        public override decimal CalculateSalary()
        {
            return 3500;
        }
    }

    public class Analyst : Employee, ISalaryCalculator
    {
        public Analyst(string name) : base(name, Designation.Analyst)
        {
        }

        public override decimal CalculateSalary()
        {
            return 4500;
        }
    }

    public static class SalarySheetExporter
    {
        public static void ExportSalarySheetToExcel(List<Employee> employees, string fileName)
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            string fullFilePath = Path.Combine(currentDirectory, fileName);

            try
            {
                using (var package = new ExcelPackage(new FileInfo(fullFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.Add("SalarySheet");

                    worksheet.Cells[1, 1].Value = "Name";
                    worksheet.Cells[1, 2].Value = "Designation";
                    worksheet.Cells[1, 3].Value = "Salary";

                    int row = 2;
                    foreach (var employee in employees)
                    {
                        worksheet.Cells[row, 1].Value = employee.Name;
                        worksheet.Cells[row, 2].Value = employee.Designation.ToString();
                        worksheet.Cells[row, 3].Value = employee.CalculateSalary();
                        row++;
                    }

                    package.Save();
                }

                Console.WriteLine($"Salary sheet exported successfully! Saved in: {fullFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while exporting the salary sheet: {ex.Message}");
            }
        }
    }

    public static class DatabaseHelper
    {
        private const string ConnectionString = "Data Source=VONOME-PC-03\\MSSQLSERVER01\r\n\r\n;Initial Catalog=OOPInCSharpDB;Integrated Security=True";

        public static void SaveEmployeeData(Employee employee)
        {
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                string query = "INSERT INTO Employees (Name, Designation) VALUES (@Name, @Designation)";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Name", employee.Name);
                command.Parameters.AddWithValue("@Designation", employee.Designation.ToString());

                connection.Open();
                command.ExecuteNonQuery();
            }

            Console.WriteLine("Employee data saved to the database!");
        }

        public static List<Employee> LoadEmployeeData()
        {
            List<Employee> employees = new List<Employee>();

            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                string query = "SELECT Name, Designation FROM Employees";
                SqlCommand command = new SqlCommand(query, connection);

                connection.Open();

                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    string designation = reader.GetString(1);

                    Designation empDesignation;
                    Enum.TryParse(designation, out empDesignation);

                    Employee employee = CreateEmployee(name, empDesignation);
                    employees.Add(employee);
                }

                reader.Close();
            }

            return employees;
        }

        private static Employee CreateEmployee(string name, Designation designation)
        {
            switch (designation)
            {
                case Designation.Manager:
                    return new Manager(name);
                case Designation.Developer:
                    return new Developer(name);
                case Designation.Tester:
                    return new Tester(name);
                case Designation.Analyst:
                    return new Analyst(name);
                default:
                    return null;
            }
        }
    }

    public partial class Program
    {
        private static List<Employee> employees = new List<Employee>();

        static void Main(string[] args)
        {
            UI();
            LoadEmployeesFromDatabase();

            Console.WriteLine("Welcome to the Salary Application!");

            bool exit = false;
            while (!exit)
            {
                Console.WriteLine("Please select an option:");
                Console.WriteLine("\t\t\t1. Add Employee");
                Console.WriteLine("\t\t\t2. Calculate Salary");
                Console.WriteLine("\t\t\t3. Export Salary Sheet");
                Console.WriteLine("\t\t\t4. Modify Employee Data");
                Console.WriteLine("\t\t\t5. Exit");

                int option = Convert.ToInt32(Console.ReadLine());

                switch (option)
                {
                    case 1:
                        AddEmployee();
                        break;
                    case 2:
                        CalculateSalary();
                        break;
                    case 3:
                        ExportSalarySheet();
                        break;
                    case 4:
                        ModifyEmployeeData();
                        break;
                    case 5:
                        exit = true;
                        break;
                    default:
                        Console.WriteLine("Invalid option!");
                        break;
                }
            }
        }

        private static void UI()
        {
            Console.BackgroundColor = ConsoleColor.Blue;
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.White;
        }

        private static void AddEmployee()
        {
            Console.WriteLine("Enter the employee name:");
            string name = Console.ReadLine();

            Console.WriteLine("Select the employee designation:");
            for (int i = 0; i < Enum.GetNames(typeof(Designation)).Length; i++)
            {
                Console.WriteLine($"{i + 1}. {Enum.GetNames(typeof(Designation))[i]}");
            }
            int designationIndex = Convert.ToInt32(Console.ReadLine()) - 1;
            if (designationIndex < 0 || designationIndex >= Enum.GetNames(typeof(Designation)).Length)
            {
                Console.WriteLine("Invalid designation!");
                return;
            }
            Designation designation = (Designation)designationIndex;

            Employee employee = CreateEmployee(name, designation);
            if (employee != null)
            {
                employees.Add(employee);
                DatabaseHelper.SaveEmployeeData(employee);
                Console.WriteLine("Employee added successfully!");
            }
            else
            {
                Console.WriteLine("Invalid employee designation!");
            }
        }

        private static Employee CreateEmployee(string name, Designation designation)
        {
            switch (designation)
            {
                case Designation.Manager:
                    return new Manager(name);
                case Designation.Developer:
                    return new Developer(name);
                case Designation.Tester:
                    return new Tester(name);
                case Designation.Analyst:
                    return new Analyst(name);
                default:
                    throw new ArgumentException("Invalid employee designation");
            }
        }

        private static void CalculateSalary()
        {
            Console.WriteLine("Enter the employee name:");
            string name = Console.ReadLine();

            Employee employee = employees.Find(emp => emp.Name == name);
            if (employee != null)
            {
                Console.WriteLine($"Salary for {employee.Name}: {employee.CalculateSalary()}");
            }
            else
            {
                Console.WriteLine("Employee not found!");
            }
        }

        private static void ExportSalarySheet()
        {
            Console.WriteLine("Enter the file name:");
            string fileName = Console.ReadLine();

            SalarySheetExporter.ExportSalarySheetToExcel(employees, fileName);
        }

        private static void ModifyEmployeeData()
        {
            Console.WriteLine("Enter the employee name:");
            string name = Console.ReadLine();

            Employee employee = employees.Find(emp => emp.Name == name);
            if (employee != null)
            {
                Console.WriteLine("Select the employee designation:");
                for (int i = 0; i < Enum.GetNames(typeof(Designation)).Length; i++)
                {
                    Console.WriteLine($"{i + 1}. {Enum.GetNames(typeof(Designation))[i]}");
                }
                int designationIndex = Convert.ToInt32(Console.ReadLine()) - 1;
                if (designationIndex < 0 || designationIndex >= Enum.GetNames(typeof(Designation)).Length)
                {
                    Console.WriteLine("Invalid designation!");
                    return;
                }
                Designation designation = (Designation)designationIndex;

                employee.Designation = designation;
                DatabaseHelper.SaveEmployeeData(employee);
                Console.WriteLine("Employee data modified successfully!");
            }
            else
            {
                Console.WriteLine("Employee not found!");
            }
        }

        private static void LoadEmployeesFromDatabase()
        {
            employees = DatabaseHelper.LoadEmployeeData();
        }
    }
}
