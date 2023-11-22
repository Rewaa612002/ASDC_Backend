using Microsoft.AspNetCore.Mvc;
using ExcelUpload.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelUpload.Data;
using OfficeOpenXml;

namespace ExcelUpload.Controllers
{
    [ApiController]
    [Route("api/excel")]
    public class ExcelController : ControllerBase
    {
        private readonly ApplicationDbContext _context;

        public ExcelController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpPost("upload")]
        [Consumes("multipart/form-data")]
        [ProducesResponseType(typeof(List<ExcelData>), StatusCodes.Status200OK)]

        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File is not selected.");

            // Read Excel file using EPPlus
            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    // Assuming that the ID is in the first column
                    var data = new List<ExcelData>();
                    for (int row = 2; row <= rowCount; row++)
                    {
                        data.Add(new ExcelData
                        {
                            Name = worksheet.Cells[row, 2].GetValue<string>(),
                            Description = worksheet.Cells[row, 3].GetValue<string>(),
                            Location = worksheet.Cells[row, 4].GetValue<string>(),
                            Price = worksheet.Cells[row, 5].GetValue<double>(),
                            Color = worksheet.Cells[row, 6].GetValue<string>(),

                            // Add other properties based on your Excel columns
                        });
                    }

                    // Filter data based on the given ID
                    await _context.ExcelData.AddRangeAsync(data);
                    await _context.SaveChangesAsync();

                    return Ok(data);
                }
            }
        }

        [HttpGet("{id} GetById")]
        public async Task<IActionResult> GetById(int id)
        {
            var data = await _context.ExcelData.FirstOrDefaultAsync(d => d.ID == id);
            if (data == null)
                return NotFound("Data not found");

            return Ok(data);
        }

        [HttpGet("GetAll")]
        public async Task<IActionResult> GetAll()
        {
            var data = await _context.ExcelData.ToListAsync();
            return Ok(data);
        }

        [HttpPost("Add")]
        public async Task<IActionResult> Add([FromBody] ExcelData newData)
        {
            _context.ExcelData.Add(newData);
            await _context.SaveChangesAsync();

            return CreatedAtAction(nameof(GetById), new { id = newData.ID }, newData);
        }

        [HttpPut("{id} Update")]
        public async Task<IActionResult> Update(int id, [FromBody] ExcelData updatedData)
        {
            var existingData = await _context.ExcelData.FirstOrDefaultAsync(d => d.ID == id);
            if (existingData == null)
                return NotFound("Data not found");

            // Update properties based on your Excel columns
            existingData.Name = updatedData.Name;
            existingData.Description = updatedData.Description;
            existingData.Location = updatedData.Location;
            existingData.Price = updatedData.Price;
            existingData.Color = updatedData.Color;


            // Update other properties

            await _context.SaveChangesAsync();

            return Ok(existingData);
        }

        [HttpDelete("{id} Delete")]
        public async Task<IActionResult> Delete(int id)
        {
            var existingData = await _context.ExcelData.FirstOrDefaultAsync(d => d.ID == id);
            if (existingData == null)
                return NotFound("Data not found");

            _context.ExcelData.Remove(existingData);
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }

}
