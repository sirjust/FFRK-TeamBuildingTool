using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using FFRK_Element_Info.Models;
using System.Data.OleDb;
using System.Data.SqlClient;
using FFRK_Element_Info.ViewModels;

namespace FFRK_Element_Info.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Boost()
        {
            return View();
        }

        public ActionResult SeeAll()
        {
            using (FFRK_DatabaseEntities db = new FFRK_DatabaseEntities())
            {
                var AllCharacters = db.FFRK_Data;
                var AllCharactersVms = new List<AllCharactersVm>();
                //for (int i = 0; i < AllCharactersVms.Count; i++)
                //{
                //    var charVm = new AllCharactersVm();
                //    charVm.Id = AllCharactersVms[i].Id.ToString();
                //    charVm.Name = AllCharactersVms[i].Name;
                //    charVm.Iteration = AllCharactersVms[i].Iteration;
                //    charVm.Type = AllCharactersVms[i].Type;
                //    charVm.ElementType = AllCharactersVms[i].ElementType;
                //    charVm.EnElement = AllCharactersVms[i].EnElement;
                //    charVm.SBType = AllCharactersVms[i].SBType;
                //    charVm.Level99 = AllCharactersVms[i].Level99;
                //    charVm.Dived = AllCharactersVms[i].Dived;
                //    charVm.ExtraInfo = AllCharactersVms[i].ExtraInfo;
                //    AllCharactersVms.Add(charVm);
                //}
                foreach(var character in AllCharacters)
                {
                    var charVm = new AllCharactersVm();
                    charVm.Id = character.Id;
                    charVm.Name = character.Name;
                    charVm.Iteration = character.Iteration;
                    charVm.Type = character.Type;
                    charVm.ElementType = character.ElementType;
                    charVm.EnElement = character.EnElement;
                    charVm.SBType = character.SBType;
                    charVm.Level99 = character.Level99;
                    charVm.Dived = character.Dived;
                    charVm.ExtraInfo = character.ExtraInfo;
                    AllCharactersVms.Add(charVm);
                }
                return View(AllCharactersVms);
            }

        }

        /*[HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an excel file.";
                return View("Index");
            }
            else
            {
                if(excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith ("xlsx"))
                {
                    string path = Server.MapPath("~/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    // read data from excel file ~14:00
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<Character> characters = new List<Character>();
                    for(int row = 2; row < range.Rows.Count; row++)
                    {
                        Character c = new Character();
                        c.Id = ((Excel.Range)range.Cells[row, 1]).Text;
                        c.Name = ((Excel.Range)range.Cells[row, 2]).Text;
                        c.Iteration = ((Excel.Range)range.Cells[row, 3]).Text;
                        c.ElementType = ((Excel.Range)range.Cells[row, 4]).Text;
                        c.EnElement = ((Excel.Range)range.Cells[row, 5]).Text;
                        c.SBType = ((Excel.Range)range.Cells[row, 6]).Text;
                        c.Level99 = ((Excel.Range)range.Cells[row, 7]).Text;
                        c.Dived = ((Excel.Range)range.Cells[row, 8]).Text;
                        c.ExtraInfo = ((Excel.Range)range.Cells[row, 9]).Text;
                        characters.Add(c);
                    }
                    ViewBag.Characters = characters;
                    return View("SeeAll");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect.";
                    return View("Index");
                }
            }

        }*/

        public ActionResult ToDatabase(HttpPostedFileBase excelfile)
        {
            string path = Server.MapPath("~/" + excelfile.FileName);
            OleDbConnection OleDbcon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;");
            OleDbCommand command = new OleDbCommand("SELECT * FROM [Sheet1$]", OleDbcon);
            OleDbDataAdapter objAdapter1 = new OleDbDataAdapter(command);

            OleDbcon.Open();
            OleDbDataReader dr = command.ExecuteReader();

            string constr = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\SirJUST\source\repos\FFRK_Element_Info\FFRK_Element_Info\App_Data\FFRK_Database.mdf;Integrated Security=True;MultipleActiveResultSets=True;Application Name=EntityFramework";

            SqlBulkCopy bulkInsert = new SqlBulkCopy(constr);
            bulkInsert.DestinationTableName = "FFRK_Data";
            bulkInsert.WriteToServer(dr);

            return View("Success");
        }
    }
}