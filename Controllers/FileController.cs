using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using RabbitMQ.Client;
using System.Text;

namespace RabbitPracticeProject.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {

        private string _getRowAsString(ExcelWorksheet worksheet, int row)
        {

            int colCount = worksheet.Dimension.Columns;
            StringBuilder res = new StringBuilder();
            for (int col = 1; col <= colCount; col++)
            {
                // Access cell value using worksheet.Cells[row, col].Value
                var cellValue = worksheet.Cells[row, col].Value;
                // Do something with the cell value
                res.Append(cellValue + ", ");
            }
            return res.ToString();
        }

        private async Task _uploadWorksheet(ExcelWorksheet worksheet)
        {
            ConnectionFactory connectionFactory = new ConnectionFactory();
            connectionFactory.Uri = new Uri("amqp://guest:guest@localhost:5672");

            using (var connection = connectionFactory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    //channel.QueueDelete("excelQueue", true, false);
                    channel.QueueDeclare(queue: "excelQueue", true, false, false);
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        string rowToString = _getRowAsString(worksheet, row);
                        var body = Encoding.UTF8.GetBytes(rowToString);
                        channel.BasicPublish(exchange: "", routingKey: "excelQueue", body: body);
                    }

                    var end = Encoding.UTF8.GetBytes("morcha");
                    channel.BasicPublish(exchange: "", routingKey: "excelQueue",body: end);
                }
            }

        }

        private async Task<IActionResult> _uploadExcel(IFormFile file)
        {
            try
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    
                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        await _uploadWorksheet(worksheet);
                    }
                }

                return Ok("File processed successfully");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }   

        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile excelFile)
        {
            if (excelFile is not { Length: > 0 })
            {
                return BadRequest("Invalid file");
            }

            return await _uploadExcel(excelFile);

        }
    }
}
