using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using practice2.Data;
using practice2.Models;

namespace practice2.Controllers
{
    public class StudentController : Controller
    {
        private readonly SchoolContext _context;

        public StudentController(SchoolContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> Index()
        {
            var estudiantes = await _context.Students.Include(e => e.Course).ToListAsync();
            return View(estudiantes);
        }

        [HttpGet]
        public IActionResult Create()
        {
            ViewBag.Cursos = _context.Courses.ToList();
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CreateStudent(Student estudiante)
        {


            _context.Add(estudiante);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));


        }

        [HttpGet]
        public async Task<IActionResult> DownloadPdf(int id)
        {
            var estudiante = await _context.Students.Include(e => e.Course).FirstOrDefaultAsync(e => e.Id == id);
            if (estudiante == null) return NotFound();

            var pdf = GenerarPdf(estudiante);
            var pdfStream = new MemoryStream();
            pdf.Save(pdfStream, false);

            pdfStream.Position = 0;
            return File(pdfStream, "application/pdf", "DatosEstudiante.pdf");
        }

        private PdfDocument GenerarPdf(Student estudiante)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Datos del Estudiante";

            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XFont font = new XFont("Verdana", 12);

            gfx.DrawString($"Nombre: {estudiante.Name}", font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString($"Edad: {estudiante.Age}", font, XBrushes.Black, new XRect(0, 20, page.Width, page.Height), XStringFormats.TopLeft);
            gfx.DrawString($"Curso: {estudiante.Course.Name}", font, XBrushes.Black, new XRect(0, 40, page.Width, page.Height), XStringFormats.TopLeft);

            return document;
        }

        [HttpGet]
        public async Task<IActionResult> ExportToExcel()
        {
            var estudiantes = await _context.Students.Include(e => e.Course).ToListAsync();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Estudiantes");

                // Agregar encabezados
                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "Nombre";
                worksheet.Cells[1, 3].Value = "Edad";
                worksheet.Cells[1, 4].Value = "Curso";

                // Agregar datos
                for (int i = 0; i < estudiantes.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = estudiantes[i].Id;
                    worksheet.Cells[i + 2, 2].Value = estudiantes[i].Name;
                    worksheet.Cells[i + 2, 3].Value = estudiantes[i].Age;
                    worksheet.Cells[i + 2, 4].Value = estudiantes[i].Course.Name;
                }

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Estudiantes.xlsx");
            }

        }
    }
}