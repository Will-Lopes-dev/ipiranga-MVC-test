using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace ipiranga_mais_1
{
    public class HomeController : Controller
    {
        private readonly List<Usuario> usuarios = new List<Usuario>();
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public ActionResult Index()
        {
            return View(usuarios);
        }

        [HttpPost]
        public ActionResult AdicionarUsuario(Usuario usuario, IFormFile arquivo)
        {
            if (arquivo != null && arquivo.Length > 0)
            {
                ViewBag.ArquivoExcelSelecionado = true;
            }
            else
            {
                ViewBag.ArquivoExcelSelecionado = false;
            }

            usuarios.Add(usuario);
            ModelState.Clear();

            return RedirectToAction("Index");
        }

        public IActionResult SalvarEmExcel()
        {
            var fileName = "usuarios.xlsx";
            var filePath = Path.Combine(_hostingEnvironment.WebRootPath, "App_Data", fileName);

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Count > 0
                    ? package.Workbook.Worksheets[1]
                    : package.Workbook.Worksheets.Add("Sheet1");

                var rowCount = worksheet.Dimension.Rows;

                foreach (var usuario in usuarios)
                {
                    rowCount++;
                    worksheet.Cells[rowCount, 1].Value = usuario.Nome;
                    worksheet.Cells[rowCount, 2].Value = usuario.Email;
                    worksheet.Cells[rowCount, 3].Value = usuario.CPF;
                    worksheet.Cells[rowCount, 4].Value = usuario.NumeroCelular;
                }

                package.Save();
            }

            return RedirectToAction("Index");
        }
    }
}
