using System.Web.Mvc;
using Aspose.Cells;
using System.IO;
using ExcelChartPdf.Models;
using RazorPDF;

namespace ExcelChartPdf.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Message = "Welcome to Charting Solution!";

            var rootPath = Server.MapPath("~/Content");

            //Open the existing excel file which contains the pie chart.
            var excelPath = Path.Combine(rootPath, "Chart.xlsx");
            var workbook = new Workbook(excelPath);

            //Get the designer chart (first chart) in the first worksheet
            //of the workbook.
            var chart = workbook.Worksheets[0].Charts[0];

            // Set virtual image path. Example - http://localhost/..
            //ImageModel.ImagePath = Request.Url + @"Content/PieChart.png";

            //Convert the chart to an image file.
            var imgFilePath = Path.Combine(rootPath, "PieChart.png");
            chart.ToImage(imgFilePath, System.Drawing.Imaging.ImageFormat.Png);

            return View();
        }

        public ActionResult About()
        {
            return View();
        }

        public PdfResult Pdf()
        {
            ImageModel.ImagePath = Server.MapPath(@"~/Content/PieChart.png");
            return new PdfResult();
        }
    }
}
