using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using FFRK_Element_Info.Models;
using System.Data;
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

        [HttpPost]
        public ActionResult ToDatabase(HttpPostedFileBase excelfile)
        {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an excel file.";
                return View("Index");
            }
            else
            {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    // delete data currently in database
                    SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\SirJUST\source\repos\FFRK_Element_Info\FFRK_Element_Info\App_Data\FFRK_Database.mdf;Integrated Security=True;MultipleActiveResultSets=True;Application Name=EntityFramework");
                    string sqlStatement = "DELETE FROM FFRK_Data";
                    try
                    {
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(sqlStatement, connection);
                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();
                    }
                    finally
                    {
                        connection.Close();
                    }
                    // insert spreadsheet data into database
                    OleDbConnection OleDbcon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;");
                    OleDbCommand command = new OleDbCommand("SELECT * FROM [Sheet1$]", OleDbcon);
                    OleDbDataAdapter objAdapter1 = new OleDbDataAdapter(command);

                    OleDbcon.Open();
                    OleDbDataReader dr = command.ExecuteReader();

                    string constr = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\SirJUST\source\repos\FFRK_Element_Info\FFRK_Element_Info\App_Data\FFRK_Database.mdf;Integrated Security=True;MultipleActiveResultSets=True;Application Name=EntityFramework";

                    SqlBulkCopy bulkInsert = new SqlBulkCopy(constr);
                    bulkInsert.DestinationTableName = "FFRK_Data";
                    bulkInsert.WriteToServer(dr);
                    // do i now need to close to prevent sql injection?

                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect.";
                    return View("Index");
                }
            }
        }

        public ActionResult ElementView()
        {
            return View();
        }

        public ActionResult BoostView()
        {
            return View();
        }

        public ActionResult ElementChoice(string Element)
        {
            using (FFRK_DatabaseEntities db = new FFRK_DatabaseEntities())
            {
                var AllCharacters = db.FFRK_Data;
                var ElementDisplayVms = new List<ElementDisplayVm>();
                foreach(var character in AllCharacters)
                {
                    if(character.ElementType.Contains (Element))
                    {
                        var charVm = new ElementDisplayVm();
                        charVm.Name = character.Name;
                        charVm.ElementType = character.ElementType;
                        charVm.EnElement = character.EnElement;
                        ElementDisplayVms.Add(charVm);
                    }
                }
                return View("ElementView", ElementDisplayVms);
            }
        }

        public ActionResult BoostChoice(string Boost)
        {
            using (FFRK_DatabaseEntities db = new FFRK_DatabaseEntities())
            {
                var AllCharacters = db.FFRK_Data;
                var BoostDisplayVms = new List<BoostDisplayVm>();
                foreach (var character in AllCharacters)
                {
                    if (character.ExtraInfo.Contains(Boost))
                    {
                        var charVm = new BoostDisplayVm();
                        charVm.Name = character.Name;
                        charVm.Iteration = character.Iteration;
                        charVm.ExtraInfo = character.ExtraInfo;
                        BoostDisplayVms.Add(charVm);
                    }
                }
                return View("BoostView", BoostDisplayVms);
            }
        }
    }
}