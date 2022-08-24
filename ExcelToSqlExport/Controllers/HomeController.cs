using ExcelToSqlExport.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelToSqlExport.Controllers
{
    public class HomeController : Controller
    {
        string conString = string.Empty;
       public HomeController()
        {
             conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
        }
        DataAccessLayer _dataAccessLayer = new DataAccessLayer();
        public DataTable dt = null;
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            dt = _dataAccessLayer.GetDataTable();
            _dataAccessLayer.sqlBulkCopy(dt);

            return View();
        }
    
    public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}