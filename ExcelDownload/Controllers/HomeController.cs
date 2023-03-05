using ExcelDownload.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace ExcelDownload.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public void ExcellCreate()
        {
            // Veritabanı bağlantı dizesi
            string connString = "Data Source=DESKTOP-MBGVKF7;Initial Catalog=Northwind;Integrated Security=True;Trusted_Connection=True;";

            // Veritabanı sorgusu
            string query = "SELECT * FROM Employees";

            // Veritabanı bağlantısı oluşturun
            System.Data.SqlClient.SqlConnection conn = new System.Data.SqlClient.SqlConnection(connString);

            // Komut nesnesi oluşturun ve sorguyu belirtin
            SqlCommand cmd = new SqlCommand(query, conn);

            // Verileri DataTable nesnesine doldurun
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
    
            // Excel dosyası oluşturun
            ExcelPackage excel = new ExcelPackage();

            // Yeni bir Excel çalışma sayfası ekleyin
            ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Employees");

            // DataTable nesnesindeki verileri Excel sayfasına doldurun
            worksheet.Cells["A1"].LoadFromDataTable(dt, true);

            // Excel dosyasını bellekte oluşturun
            byte[] bytes = excel.GetAsByteArray();

            // Dosya adını ve uzantısını belirleyin
            string fileName = "Employees.xlsx";

            // Yanıt nesnesini alın
            HttpResponse response = HttpContext.Response;

            // Yanıt başlıklarını belirleyin
            response.Clear();
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            response.Headers.Add("content-disposition", "attachment;  filename=" + fileName);

            // Excel dosyasını yanıt nesnesine yazın
             response.Body.WriteAsync(bytes, 0, bytes.Length);
             response.Body.FlushAsync();
        }


        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}