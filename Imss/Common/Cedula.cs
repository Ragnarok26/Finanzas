using GemBox.Spreadsheet;
using Imss.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Imss.Common
{
    public class Cedula
    {
        public static Tools.ProcessResult Procesar(Stream Parametros, Stream Cedula, HttpPostedFileBase Pdf, string Para, string usuario, string app, string mailSubject, string mailBody, long logId)
        {
            OfficeOpenXml.ExcelPackage pkg = null;
            OfficeOpenXml.ExcelPackage pkgCed = null;
            OfficeOpenXml.ExcelWorksheet sheet = null;
            OfficeOpenXml.ExcelWorksheet sheetCed = null;
            Empleado empleado = null;
            List<Empleado> empleados = new List<Empleado>();
            MemoryStream ms = null;
            MemoryStream pdf = null;
            List<System.Net.Mail.Attachment> attachments = null;
            int? ascIni = null;
            int? ascFin = null;
            int? descIni = null;
            int? descFin = null;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile xls = null;
            ExcelWorksheet worksheet = null;
            Tools.ProcessResult result = new ProcessResult();
            try
            {
                pkg = new OfficeOpenXml.ExcelPackage(Parametros);
                sheet = pkg.Workbook.Worksheets[1];
                for (int x = sheet.Dimension.Start.Row + 1; x <= sheet.Dimension.End.Row; x++)
                {
                    empleado = new Empleado()
                    {
                        Nombre = (string)sheet.Cells[x, 1].Value,
                        Correo = (string)sheet.Cells[x, 2].Value
                    };
                    if (!string.IsNullOrEmpty(empleado.Nombre))
                    {
                        pkgCed = new OfficeOpenXml.ExcelPackage(Cedula);
                        sheetCed = pkgCed.Workbook.Worksheets[1];
                        try
                        {
                            ascFin = (from cell in sheetCed.Cells where cell.Value is string && ((string)cell.Value).Equals(empleado.Nombre) select cell).FirstOrDefault().Start.Row;
                        }
                        catch
                        {
                            ascFin = null;
                        }
                        if (ascFin.HasValue)
                        {
                            ascIni = (from cell in sheetCed.Cells where cell.Value is string && ((string)cell.Value).Equals(@"SDI") select cell).FirstOrDefault().Start.Row;
                            descFin = (from cell in sheetCed.Cells where cell.Value is string && ((string)cell.Value).Equals(@"Total de Días cotizados para el calculo de trabajadores promedio expuestos al riesgo") select cell).FirstOrDefault().Start.Row;
                            try
                            {
                                descIni = (from cell in sheetCed.Cells where cell.Start.Row > ascFin.Value + 1 && cell.Value is string && ((string)cell.Value).Contains("-") && ((string)cell.Value).Replace("-", "").Length == 11 select cell).FirstOrDefault().Start.Row;
                            }
                            catch
                            {
                                descIni = null;
                            }
                            if (descIni.HasValue && descFin.HasValue)
                            {
                                //descFin -= 2;
                                if (descFin.Value - descIni.Value > 0)
                                {
                                    sheetCed.DeleteRow(descIni.Value, (descFin.Value - descIni.Value));
                                }
                            }
                            if (ascIni.HasValue && ascFin.HasValue)
                            {
                                ascIni += 2;
                                ascFin -= 1;
                                if (ascFin.Value - ascIni.Value > 0)
                                {
                                    sheetCed.DeleteRow(ascIni.Value, (ascFin.Value - ascIni.Value));
                                }
                            }
                            ms = new MemoryStream();
                            pkgCed.SaveAs(ms);
                            ms.Position = 0;
                            pdf = new MemoryStream();
                            xls = ExcelFile.Load(ms, LoadOptions.XlsxDefault);
                            worksheet = xls.Worksheets.ActiveWorksheet;
                            worksheet.PrintOptions.FitWorksheetWidthToPages = 1;
                            worksheet.PrintOptions.FitWorksheetHeightToPages = 1;
                            xls.Save(pdf, SaveOptions.PdfDefault);
                            attachments = new List<System.Net.Mail.Attachment>();
                            attachments.Add(new System.Net.Mail.Attachment(pdf, "Cédula IMSS.pdf", "application/pdf"));
                            Pdf.InputStream.Position = 0;
                            String PdfFileName = Path.GetFileName(Pdf.FileName);
                            attachments.Add(new System.Net.Mail.Attachment(Pdf.InputStream, PdfFileName, Pdf.ContentType));
                            if (!string.IsNullOrEmpty(Para))
                            {
                                empleado.Correo = Para;
                            }
                            empleado.Procesado = Correo.Enviar(empleado, usuario, app, mailSubject, mailBody, attachments, logId);
                            empleado.Guardar();
                            empleados.Add(empleado);
                            if (pdf != null)
                            {
                                pdf.Dispose();
                                pdf = null;
                            }
                            if (ms != null)
                            {
                                ms.Dispose();
                                ms = null;
                            }
                        }
                        else
                        {
                            empleado.Procesado = false;
                            empleado.Guardar();
                            empleados.Add(empleado);
                        }
                    }
                    empleado = null;
                    attachments = null;
                    worksheet = null;
                    xls = null;
                    if (sheetCed != null)
                    {
                        sheetCed.Dispose();
                        sheetCed = null;
                    }
                    if (pkgCed != null)
                    {
                        pkgCed.Dispose();
                        pkgCed = null;
                    }
                    ascIni = null;
                    ascFin = null;
                    descFin = null;
                }
                if (sheet != null)
                {
                    sheet.Dispose();
                    sheet = null;
                }
                if (pkg != null)
                {
                    pkg.Dispose();
                    pkg = null;
                }
                if (Parametros != null)
                {
                    Parametros.Dispose();
                    Parametros = null;
                }
                if (Cedula != null)
                {
                    Cedula.Dispose();
                    Cedula = null;
                }
                result.Collection = empleados;
                return result;
            }
            catch (Exception ex)
            {
                LogTools.RegisterLog(logId, usuario, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "Error al procesar la información: " + ex.Message + ".", DateTime.Now);
                result.HasErrors = true;
                result.Message = "Error al procesar la información: " + ex.Message + ".";
                return result;
            }
            finally
            {
                if (sheet != null)
                {
                    sheet.Dispose();
                    sheet = null;
                }
                if (pkg != null)
                {
                    pkg.Dispose();
                    pkg = null;
                }
                if (Parametros != null)
                {
                    Parametros.Dispose();
                    Parametros = null;
                }
                empleado = null;
                attachments = null;
                xls = null;
                worksheet = null;
                ascIni = null;
                ascFin = null;
                descFin = null;
                if (sheetCed != null)
                {
                    sheetCed.Dispose();
                    sheetCed = null;
                }
                if (pkgCed != null)
                {
                    pkgCed.Dispose();
                    pkgCed = null;
                }
                if (Cedula != null)
                {
                    Cedula.Dispose();
                    Cedula = null;
                }
                if (ms != null)
                {
                    ms.Dispose();
                    ms = null;
                }
                if (pdf != null)
                {
                    pdf.Dispose();
                    pdf = null;
                }
                result = null;
            }
        }
    }
}
