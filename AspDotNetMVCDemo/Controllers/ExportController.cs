using AspDotNetMVCDemo.Models;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Web;
using System.Web.Mvc;
using System.Collections;
using System.Collections.Generic;
using System.Web.UI.WebControls;
using System.Web.UI;
using System.Reflection;

namespace AspDotNetMVCDemo.Controllers
{
    public class ExportController : Controller
    {
        // GET: Import
        public ActionResult Index()
        {
            return View(GetEmployeesList());
        }

        
        public ActionResult ExportToExcel()
        {
            var gv = new GridView();
            gv.DataSource = ToDataTable<Employee>(GetEmployeesList());
            gv.DataBind();

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=DemoExcel.xls");
            Response.ContentType = "application/ms-excel";

            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);

            gv.RenderControl(objHtmlTextWriter);

            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();

            return View("Index");
        }

        /// <summary>
        /// Returns Employee List
        /// </summary>
        /// <returns></returns>
        private List<Employee> GetEmployeesList()
        {
            var employees = new List<Employee>();
            employees.Add(new Employee()
            {
                EmployeeId = 1001,
                EmployeeName = "Kannadasan",
                Designation = "Tech Lead",
                Salary = 60000
            });
            employees.Add(new Employee()
            {
                EmployeeId = 1002,
                EmployeeName = "Manju",
                Designation = "UX Lead",
                Salary = 70000
            });
            employees.Add(new Employee()
            {
                EmployeeId = 1003,
                EmployeeName = "Ranadhir",
                Designation = "Manager",
                Salary = 80000
            });
            
            return employees;
        }

        //Generic method to convert List to DataTable
        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Defining type of data column gives proper data table 
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
    }
}