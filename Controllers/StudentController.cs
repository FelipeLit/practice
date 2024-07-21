using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using practice2.Data;
using practice2.Models;
using X.PagedList.Extensions;
using X.PagedList;
using X.PagedList.Mvc.Core;

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
        public async Task<IActionResult> Index(int? page)
        {
            var pageNumber = page ?? 1;
            var pageSize = 10;

            // Convertir la consulta a una lista antes de aplicar la paginación
            var studentsList = await _context.Students.Include(s => s.Course).ToListAsync();
            var studentsPagedList = studentsList.ToPagedList(pageNumber, pageSize);

            return View(studentsPagedList);
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


        [HttpPost]
        public async Task<IActionResult> ImportFromExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            try
            {
                using (var package = new ExcelPackage(file.OpenReadStream()))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var nameCell = worksheet.Cells[row, 1].Text.Trim();
                        var ageCell = worksheet.Cells[row, 2].Text.Trim();
                        var courseIdCell = worksheet.Cells[row, 3].Text.Trim();

                        // Debug: Check the values read from the Excel
                        Console.WriteLine($"Row {row} - Name: '{nameCell}', Age: '{ageCell}', CourseId: '{courseIdCell}'");

                        if (string.IsNullOrWhiteSpace(nameCell) || string.IsNullOrWhiteSpace(ageCell) || string.IsNullOrWhiteSpace(courseIdCell))
                        {
                            // Skipping row due to missing data
                            Console.WriteLine($"Skipping row {row} due to missing data: Name='{nameCell}', Age='{ageCell}', CourseId='{courseIdCell}'");
                            continue;
                        }

                        if (!int.TryParse(ageCell, out int age) || !int.TryParse(courseIdCell, out int courseId))
                        {
                            // Skipping row due to invalid data
                            Console.WriteLine($"Skipping row {row} due to invalid data: Name='{nameCell}', Age='{ageCell}', CourseId='{courseIdCell}'");
                            continue;
                        }

                        var estudiante = new Student
                        {
                            Name = nameCell,
                            Age = age,
                            CourseId = courseId
                        };

                        _context.Students.Add(estudiante);
                        Console.WriteLine($"Added student: {nameCell}, Age: {age}, CourseId: {courseId}");
                    }

                    await _context.SaveChangesAsync();
                    Console.WriteLine("Database updated successfully");
                }

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                // Log the exception (you can use a logging framework for this)
                Console.WriteLine($"Error importing Excel file: {ex.Message}");
                return StatusCode(500, "Internal server error");
            }
        }
    }
}