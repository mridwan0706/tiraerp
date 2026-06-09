using Dapper;
using JobNefo.Contracts;
using JobNefo.Controllers;
using JobNefo.Model.JobEWACNDN;
using JobNefo.Models;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NefoCore.Extensions;
using NefoCore.Mails;
using NefoCore.Models;
using NefoCore.Reports;
using NefoModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobNefo.Repositories
{
    

    public class JobEWACNDNRepository : IJobEWACNDNRepository
    {
        private readonly IConfiguration _config;
        private readonly ConfigDBModel _config_db;        
        private readonly ILogger<JobEWACNDNController> log;
        private readonly BackupConfigModel _backup;
        //private readonly IEnumerable<SellpoConfigModel> _sellpo_config;
        private readonly string id;

        public JobEWACNDNRepository(
            IConfiguration config,          
            ConfigDBModel config_db,
            ILogger<JobEWACNDNController> ilogger,           
            BackupConfigModel backup)
        {
            _config = config;
            _config_db = config_db;           
            _backup = backup;
            log = ilogger;
            id = Path.GetRandomFileName();
        }
        public IDbConnection Connection { get; private set; }
        private IDbConnection ConnectionKAM => new SqlConnection(_config_db.DefaultConnection);        
        public string ViewName { get; private set; }
        public string StoreProcedurreName { get; private set; }
        public string TabelName { get; set; }

        public async Task<bool> Post()
        {
            var stopwatch = new Stopwatch();

            try
            {
                stopwatch.Start();
                log.LogInformation($"[{id}] Job Start.");

                var dataExcel = (await GetDataExcel())?.ToList() ?? new List<ReportEWACNDNModel>();
                if (!dataExcel.Any())
                {
                    log.LogInformation($"[{id}] No data found. Job finished without sending email.");
                    stopwatch.Stop();
                    return true;
                }

                string filename = CreateExcel(dataExcel, _backup);
                if (string.IsNullOrWhiteSpace(filename) || !File.Exists(filename))
                {
                    log.LogWarning($"[{id}] Excel file gagal dibuat.");
                    stopwatch.Stop();
                    return false;
                }

                log.LogInformation($"[{id}] Send Email Start.");

                try
                {
                    EmailAddressCollection emails = await GetPenerimaEmail();
                    await SendEmailAsync(emails, stopwatch.Elapsed, dataExcel, filename);
                }
                catch (Exception exc)
                {
                    log.LogError($"[{id}] Send email error: {exc.GetMessageInnerOrDefault()}");
                    return false;
                }

                log.LogInformation($"[{id}] Send Email End.");

                stopwatch.Stop();
                log.LogInformation($"[{id}] Job End.");

                return true;
            }
            catch (Exception exc)
            {
                log.LogError($"[{id}] {exc.GetMessageInnerOrDefault()}");
                return false;
            }
        }

        public async Task<IEnumerable<ReportEWACNDNModel>> GetDataExcel()
        {
            try
            {            
                ViewName = "VW_EWA_CNDN";
                using var conn = new SqlConnection(_config_db.DefaultConnection);
                await conn.OpenAsync();

                string strSQL = $@"
                    SELECT *
                    FROM {ViewName}";

                var models = await conn.QueryAsync<ReportEWACNDNModel>(
                    sql: strSQL,
                    commandTimeout: 0);

                return models.ToList();
            }
            catch (Exception ex)
            {
                log.LogError($"[{id}] GetDataExcel error: {ex.GetMessageInnerOrDefault()}");
                return Enumerable.Empty<ReportEWACNDNModel>();
            }
        }

        public string CreateExcel(IEnumerable<ReportEWACNDNModel> models, BackupConfigModel folder)
        {
            string filename = string.Empty;

            try
            {
                string targetFolder = Path.Combine(folder.Directory, "files");
                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                filename = Path.Combine(
                    targetFolder,
                    $"Report CNDN_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                using var excel = new ExcelPackage();
                var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

                workSheet.TabColor = Color.Black;
                workSheet.DefaultRowHeight = 12;

                workSheet.Row(1).Height = 20;
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.Row(1).Style.Font.Bold = true;
                workSheet.Cells[1, 1].Value = $"Report CNDN";

                Color colFromHexHeader = ColorTranslator.FromHtml("#ffe599");
                workSheet.Row(3).Height = 20;
                workSheet.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(3).Style.Font.Bold = true;
                workSheet.Row(3).Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Row(3).Style.Fill.BackgroundColor.SetColor(colFromHexHeader);

                workSheet.Cells[3, 1].Value = "BRANCH";
                workSheet.Cells[3, 2].Value = "NIK";
                workSheet.Cells[3, 3].Value = "FLAG VERIFIKASI";
                workSheet.Cells[3, 4].Value = "TANGGAL INPUT NIK";
                workSheet.Cells[3, 5].Value = "TANGGAL VERIFIKASI";
                workSheet.Cells[3, 6].Value = "KODE OUTLET";
                workSheet.Cells[3, 7].Value = "NAMA OUTLET";
                workSheet.Cells[3, 8].Value = "REASON";

                int recordIndex = 4;
                foreach (var data in models)
                {
                    workSheet.Cells[recordIndex, 1].Value = data.branch;
                    workSheet.Cells[recordIndex, 2].Value = data.nik;
                    workSheet.Cells[recordIndex, 3].Value = data.flag_verifikasi;

                    //workSheet.Cells[recordIndex, 4].Value = data.nik_input_date;
                    //workSheet.Cells[recordIndex, 5].Value = data.tanggal_verifikasi;

                    if (data.nik_input_date.HasValue)
                        workSheet.Cells[recordIndex, 4].Value = data.nik_input_date.Value;
                    else
                        workSheet.Cells[recordIndex, 4].Value = string.Empty;

                    if (data.tanggal_verifikasi.HasValue)
                        workSheet.Cells[recordIndex, 5].Value = data.tanggal_verifikasi.Value;
                    else
                        workSheet.Cells[recordIndex, 5].Value = string.Empty;

                    workSheet.Cells[recordIndex, 6].Value = data.kode_outlet;
                    workSheet.Cells[recordIndex, 7].Value = data.nama_outlet;
                    workSheet.Cells[recordIndex, 8].Value = data.reason;
                    recordIndex++;
                }

                workSheet.Column(4).Style.Numberformat.Format = "dd-MMM-yyyy";
                workSheet.Column(5).Style.Numberformat.Format = "dd-MMM-yyyy";

                if (recordIndex > 4)
                {
                    workSheet.Cells[3, 1, recordIndex - 1, 8].AutoFitColumns();
                    workSheet.Cells[3, 1, recordIndex - 1, 8].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[3, 1, recordIndex - 1, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[3, 1, recordIndex - 1, 8].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[3, 1, recordIndex - 1, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                File.WriteAllBytes(filename, excel.GetAsByteArray());
                return filename;
            }
            catch (Exception ex)
            {
                log.LogError($"[{id}] CreateExcel error: {ex.GetMessageInnerOrDefault()}");
                return string.Empty;
            }
        }

        public async Task<EmailAddressCollection> GetPenerimaEmail()
        {
            TabelName = "GS_GEN_HARDCODED";

            var model = new EmailAddressCollection
            {
                Subject = $"EWA CNDN",
            };

            try
            {
                //using var conn = new SqlConnection(_config_db.ConnectionKAM);
                using var conn = new SqlConnection(_config_db.DefaultConnection);
                await conn.OpenAsync();

                string sql = $@"
                    SELECT gh_function_desc AS [Address]
                    FROM {TabelName} WITH (NOLOCK)
                    WHERE gh_function_name = 'EWA_CNDN'
                      AND gh_function_code = 'TO'";

                var to = (await conn.QueryAsync<EmailAddressModel>(sql)).ToList();
                if (to.Any())
                    model.TO = to;
            }
            catch (Exception exc)
            {
                log.LogError($"[{id}] GetPenerimaEmail error: {exc.GetMessageInnerOrDefault()}");
                throw;
            }

            return model;
        }

        public async Task SendEmailAsync(
            EmailAddressCollection collection,
            TimeSpan time,
            IEnumerable<ReportEWACNDNModel> models,
            string path)
        {
            using var mail = new TigaraksaMail();

            if (collection.TO?.Any() == true)
            {
                mail.To = new List<EmailAddressCoreModel>();
                mail.To.AddRange(collection.TO.Select(s =>
                    new EmailAddressCoreModel(code: "TO", desc: s.Address, display: s.Display)));
            }

            if (collection.CC?.Any() == true)
            {
                mail.CC = new List<EmailAddressCoreModel>();
                mail.CC.AddRange(collection.CC.Select(s =>
                    new EmailAddressCoreModel(code: "CC", desc: s.Address, display: s.Display)));
            }

            if (collection.BCC?.Any() == true)
            {
                mail.BCC = new List<EmailAddressCoreModel>();
                mail.BCC.AddRange(collection.BCC.Select(s =>
                    new EmailAddressCoreModel(code: "BCC", desc: s.Address, display: s.Display)));
            }

            var bodyEmail = models
                .GroupBy(x => x.branch)
                .Select((g, index) => new HeaderEmailDisplayModel
                {
                    NO = index + 1,
                    CABANG = g.Key,
                    TOTAL = g.Count()
                })
                .ToList();

            var html = new StringBuilder();
            var excel = new ExcelHtml();

            html.Append(excel.CreateParagraph("Dear All,"));
            html.Append(excel.CreateParagraph(
                $"Berikut adalah ewa cndn, mohon diperhatikan outlet yang belum verified per {DateTime.Now:dd MMM yyyy - HH:mm}."));
            html.Append(CreateHeader());
            html.AppendLine("<table class='table'>");
            html.AppendLine("<tr>");
            html.AppendLine("<th class='header'>CABANG</th>");
            html.AppendLine("<th class='header'>TOTAL</th>");
            html.AppendLine("</tr>");

            foreach (var item in bodyEmail)
            {
                html.AppendLine("<tr>");
                html.AppendLine($"<td>{item.CABANG}</td>");
                html.AppendLine($"<td>{item.TOTAL:#,##0}</td>");
                html.AppendLine("</tr>");
            }

            html.AppendLine("</table>");
            html.Append(excel.Enter());
            html.Append($"<b>Time</b><br/>{time}");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            await mail.SendAsync(
                subject: collection.Subject,
                isBodyHtml: true,
                body: html.ToString(),
                filename: path);
        }

        public StringBuilder CreateHeader()
        {
            var sb = new StringBuilder();
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("<style>");
            sb.AppendLine(".table {");
            sb.AppendLine("font-size: 15px;");
            sb.AppendLine("border-collapse: collapse;");
            sb.AppendLine("}");
            sb.AppendLine(".header {");
            sb.AppendLine("background-color: #008080;");
            sb.AppendLine("color: white;");
            sb.AppendLine("text-align: center;");
            sb.AppendLine("vertical-align: middle;");
            sb.AppendLine("height: 40px;");
            sb.AppendLine("}");
            sb.AppendLine(".total, .totalnumberic, .totaldatetime, .totalcenterrd {");
            sb.AppendLine("background-color: #008080;");
            sb.AppendLine("color: white;");
            sb.AppendLine("font-weight: bold;");
            sb.AppendLine("}");
            sb.AppendLine(".numberic {");
            sb.AppendLine("text-align: right;");
            sb.AppendLine(@"mso-number-format: ""\#\,\#\#0"";");
            sb.AppendLine("}");
            sb.AppendLine(".datetime {");
            sb.AppendLine("text-align: center;");
            sb.AppendLine(@"mso-number-format: ""dd-mmm-yyyy"";");
            sb.AppendLine("}");
            sb.AppendLine(".centerrd {");
            sb.AppendLine("text-align: center;");
            sb.AppendLine("}");
            sb.AppendLine("td, th {");
            sb.AppendLine("border: 1px solid black;");
            sb.AppendLine("padding-left: 9px;");
            sb.AppendLine("padding-right: 9px;");
            sb.AppendLine("padding-top: 4px;");
            sb.AppendLine("padding-bottom: 4px;");
            sb.AppendLine("}");
            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            return sb;
        }
    }
}
