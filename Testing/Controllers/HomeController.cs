
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Testing.Models
{
    public class HomeController : Controller
    {
        private EA_Testing_1Entities1 db = new EA_Testing_1Entities1();
        public ActionResult Index()
        {
            return View();

        }

        [HttpPost]
        public FileResult ExportToExcel()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[7] {
                new DataColumn("Sno"),
                new DataColumn("Firstname"),
                new DataColumn("Lastname"),
                new DataColumn("DateOfBirth"),
                new DataColumn("Age"),
                new DataColumn("Gender"),
                new DataColumn("Mobile")
            });


            var EA_Testing_1Entities = from Employee1 in db.Employee1 select Employee1;

            foreach (var e in EA_Testing_1Entities)
            {
                dt.Rows.Add(e.Sno, e.Firstname, e.Lastname, e.DateOfBirth,
                    e.Age, e.Gender, e.Mobile);
            }

            using (XLWorkbook wb = new XLWorkbook()) //Install ClosedXml from Nuget for XLWorkbook  
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream()) //using System.IO;  
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelFile.xlsx");
                }
            }
        }

        [HttpPost]
        public ActionResult ImportFromExcel(HttpPostedFileBase postedFile)
        {
            string filePath = string.Empty;
            if (postedFile != null)
            {
                string path = Server.MapPath("~/file/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                filePath = path + Path.GetFileName(postedFile.FileName);
                string extension = Path.GetExtension(postedFile.FileName);
                postedFile.SaveAs(filePath);

                string conString = string.Empty;
                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        conString = ConfigurationManager.ConnectionStrings["ExcelConString03"].ConnectionString;
                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = ConfigurationManager.ConnectionStrings["ExcelConString07"].ConnectionString;
                        break;
                }
                DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                using (OleDbConnection oleconexcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdexcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter oleexcel = new OleDbDataAdapter())
                        {
                            cmdexcel.Connection = oleconexcel;

                            //Get the name of First Sheet.
                            oleconexcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = oleconexcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            oleconexcel.Close();

                            //Read Data from First Sheet.
                            oleconexcel.Open();
                            cmdexcel.CommandText = "SELECT * From [" + sheetName + "]";
                            oleexcel.SelectCommand = cmdexcel;
                            oleexcel.Fill(dt);
                            oleconexcel.Close();
                        }
                    }
                }
                conString = ConfigurationManager.ConnectionStrings["Constring"].ConnectionString;
                try
                {
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name.
                            sqlBulkCopy.DestinationTableName = "dbo.Employee1";

                            //[OPTIONAL]: Map the Excel columns with that of the database table
                            sqlBulkCopy.ColumnMappings.Add("Sno", "Sno");
                            sqlBulkCopy.ColumnMappings.Add("Firstname", "Firstname");
                            sqlBulkCopy.ColumnMappings.Add("Lastname", "Lastname");
                            sqlBulkCopy.ColumnMappings.Add("DateOfBirth", "DateOfBirth");
                            sqlBulkCopy.ColumnMappings.Add("Age", "Age");
                            sqlBulkCopy.ColumnMappings.Add("Gender", "Gender");
                            sqlBulkCopy.ColumnMappings.Add("Mobile", "Mobile");

                            con.Open();
                            sqlBulkCopy.WriteToServer(dt);
                            con.Close();
                        }
                    }

                    return Json(new { success = true, message = "Saved successfully" }, JsonRequestBehavior.AllowGet);
                }

                catch (Exception ex)

                {

                    Console.WriteLine(ex);
                    return View();
                }  
            }
            return View("Index");
        }

         
        public ActionResult Getlist()

        {

            List<Employee1> empResult = null;

            try
            {

                using (EA_Testing_1Entities1 db = new EA_Testing_1Entities1())

                {

                    List<Employee1> emplist = db.Employee1.ToList<Employee1>();

                    empResult = emplist;

                }

            }

            catch (Exception ex)

            {

                Console.WriteLine(ex);

                empResult = null;

            }

            return Json(new { data = empResult }, JsonRequestBehavior.AllowGet);

        }

        [HttpGet]
        public ActionResult AddorEdit(int Id = 0)

        {

            if (Id == 0)

                return View(new Employee1());

            else

            {

                using (EA_Testing_1Entities1 db = new EA_Testing_1Entities1())

                {

                    return View(db.Employee1.Where(x => x.Id == Id).FirstOrDefault<Employee1>());

                }

            }

        }

        [HttpPost]
        public ActionResult AddorEdit(Employee1 Employee1)

        {

            using (EA_Testing_1Entities1 db = new EA_Testing_1Entities1())

            {

                if (Employee1.Id == 0)

                {

                    db.Employee1.Add(Employee1);

                    db.SaveChanges();

                    return Json(new { success = true, message = "Saved successfully" }, JsonRequestBehavior.AllowGet);

                }

                else

                {

                    db.Entry(Employee1).State = EntityState.Modified;

                    db.SaveChanges();

                    return Json(new { success = true, message = "Updated successfully" }, JsonRequestBehavior.AllowGet);

                }

            }

        }

       
        [HttpPost]

        public ActionResult Delete(int Id)

        {

            using (EA_Testing_1Entities1 db = new EA_Testing_1Entities1())

            {

                Employee1 Employee1 = db.Employee1.Where(x => x.Id == Id).FirstOrDefault<Employee1>();

                db.Employee1.Remove(Employee1);

                db.SaveChanges();

                return Json(new { success = true, message = "Deleted successfylly" }, JsonRequestBehavior.AllowGet);

            }

        }

        public ActionResult Edit(Employee1 Employee1)

        {

            using (EA_Testing_1Entities1 db = new EA_Testing_1Entities1())

            {

                if (Employee1.Id == 0)

                {

                    db.Employee1.Add(Employee1);

                    db.SaveChanges();

                    return Json(new { success = true, message = "Saved successfully" }, JsonRequestBehavior.AllowGet);

                }

                else

                {

                    db.Entry(Employee1).State = EntityState.Modified;

                    db.SaveChanges();

                    return Json(new { success = true, message = "Updated successfully" }, JsonRequestBehavior.AllowGet);

                }

            }
        }
    }
}