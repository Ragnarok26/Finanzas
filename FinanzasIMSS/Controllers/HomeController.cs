using FinanzasIMSS.Models;
using Imss.Common;
using Imss.Tools;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace FinanzasIMSS.Controllers
{
    public class HomeController : Controller
    {
        private async Task<Usuario> GetUsuario(string app)
        {
            HttpCookie myCookie = null;
            JavaScriptSerializer js = null;
            Usuario usuario = new Usuario();
            bool flagPermisos = false;
            try
            {
                myCookie = Request.Cookies["UserSettingsFinanzasImss"];
                js = new JavaScriptSerializer();
                if (myCookie != null)
                {
                    if (!string.IsNullOrEmpty(myCookie["UserData"]))
                    {
                        try
                        {
                            usuario = (Usuario)js.Deserialize(myCookie["UserData"], typeof(Usuario));
                            flagPermisos = usuario.Permisos != null ? (usuario.Permisos.Count > 0) : false;
                            if (!flagPermisos)
                            {
                                usuario = null;
                            }
                        }
                        catch
                        {
                            usuario = null;
                        }
                    }
                    else
                    {
                        usuario = null;
                    }
                }
                else
                {
                    try
                    {
                        myCookie = new HttpCookie("UserSettingsFinanzasImss");
                        usuario.Nombre = User.Identity.Name.Substring(User.Identity.Name.LastIndexOf(@"\") + 1);
                        usuario.Permisos = usuario.ObtenerPermisos(app);
                        flagPermisos = usuario.Permisos != null ? (usuario.Permisos.Count > 0) : false;
                        if (flagPermisos)
                        {
                            myCookie["UserData"] = js.Serialize(usuario);
                            myCookie.Expires = DateTime.Now.AddYears(1000);
                        }
                        else
                        {
                            usuario = null;
                            myCookie = null;
                        }
                    }
                    catch
                    {
                        myCookie = null;
                        usuario = null;
                    }
                    if (myCookie != null)
                    {
                        Response.Cookies.Add(myCookie);
                    }
                }
                return usuario;
            }
            finally
            {
                js = null;
                usuario = null;
                myCookie = null;
            }
        }

        public FileResult Ayuda()
        {
            string NombreArchivo = Server.MapPath("~/Ayuda/Manual usuario.pdf");
            var cd = new System.Net.Mime.ContentDisposition
            {
                // for example foo.bak
                FileName = "Manual de usuario.pdf",
                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());
            return File(NombreArchivo, "application/pdf");
            //return null;
        }

        public FileResult ArchivoEjemploFormato()
        {
            string NombreArchivo = Server.MapPath("~/Ayuda/Ejemplo formato.xlsx");
            var cd = new System.Net.Mime.ContentDisposition
            {
                // for example foo.bak
                FileName = "Ejemplo formato.xlsx",
                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());
            return File(NombreArchivo, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }

        public FileResult ArchivoEjemploParametros()
        {
            string NombreArchivo = Server.MapPath("~/Ayuda/Ejemplo parametros.xlsx");
            var cd = new System.Net.Mime.ContentDisposition
            {
                // for example foo.bak
                FileName = "Ejemplo parametros.xlsx",
                // always prompt the user for downloading, set to true if you want 
                // the browser to try to show the file inline
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());
            return File(NombreArchivo, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }

        public async Task<ViewResult> Index()
        {
            Usuario usuario = null;
            string app = ConfigurationManager.AppSettings["App.Name"];
            try
            {
                usuario = await GetUsuario(app);
                if (usuario != null)
                {
                    ViewBag.UserName = usuario.Nombre;
                    return View();
                }
                else
                {
                    return View("ErrorPermisos");
                }
            }
            finally
            {
                usuario = null;
            }
        }

        [HttpPost]
        public async Task<ActionResult> Upload()
        {
            long logId = 0;
            ProcessResult result = null;
            string app = ConfigurationManager.AppSettings["App.Name"];
            Usuario usuario = await GetUsuario(app);
            logId = LogTools.RegisterLog(0, usuario.Nombre, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "Procesando Archivo.", DateTime.Now);
            if (usuario != null)
            {
                HttpPostedFileBase file = null;
                HttpPostedFileBase fileCed = null;
                HttpPostedFileBase filePdf = null;
                string to = ConfigurationManager.AppSettings["MailTo"].ToString();
                try
                {
                    if (Request.Files.Count == 3)
                    {
                        file = Request.Files["fileInputEnv"];
                        fileCed = Request.Files["fileInputCed"];
                        filePdf = Request.Files["fileInputPdf"];
                        result = Cedula.Procesar(file.InputStream, fileCed.InputStream, filePdf, to, usuario.Nombre, app, Request.Params["mailSubject"], Request.Params["mailBody"].Replace("\n","<br/>"), logId);
                        if (filePdf.InputStream != null)
                        {
                            filePdf.InputStream.Dispose();
                            filePdf = null;
                        }
                        if (fileCed.InputStream != null)
                        {
                            fileCed.InputStream.Dispose();
                            fileCed = null;
                        }
                        if (file.InputStream != null)
                        {
                            file.InputStream.Dispose();
                            file = null;
                        }
                        if (!result.HasErrors)
                        {
                            return PartialView("Grid", result.Collection);
                        }
                        else
                        {
                            return Json(result);
                        }
                    }
                    else
                    {
                        result = new ProcessResult()
                        {
                            HasErrors = true,
                            Message = "No hay suficientes archivos para iniciar el proceso."
                        };
                        return Json(result);
                    }
                }
                catch (Exception ex)
                {
                    LogTools.RegisterLog(logId, usuario.Nombre, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "Error al procesar la información: " + ex.Message + ".", DateTime.Now);
                    result = new ProcessResult()
                    {
                        HasErrors = true,
                        Message = "Error al procesar la información: " + ex.Message + "."
                    };
                    return Json(result);
                }
                finally
                {
                    if (file != null)
                    {
                        if (file.InputStream != null)
                        {
                            file.InputStream.Dispose();
                        }
                        file = null;
                    }
                    if (fileCed != null)
                    {
                        if (fileCed.InputStream != null)
                        {
                            fileCed.InputStream.Dispose();
                        }
                        fileCed = null;
                    }
                    if (filePdf != null)
                    {
                        if (filePdf.InputStream != null)
                        {
                            filePdf.InputStream.Dispose();
                        }
                        filePdf = null;
                    }
                }
            }
            else
            {
                LogTools.RegisterLog(logId, "", app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "No se han cargado los datos del usuario para poder procesar la carga de datos.", DateTime.Now);
                result = new ProcessResult()
                {
                    HasErrors = true,
                    Message = "No se han cargado los datos del usuario para poder procesar la carga de datos."
                };
                return Json(result);
            }
        }

        /*
        [HttpPost]
        public async Task<PartialViewResult> Upload()
        {
            long logId = 0;
            string app = ConfigurationManager.AppSettings["App.Name"];
            Usuario usuario = await GetUsuario(app);
            logId = LogTools.RegisterLog(0, usuario.Nombre, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "Procesando Archivo.", DateTime.Now);
            if (usuario != null)
            {
                HttpPostedFileBase file = null;
                Workbook workbook = null;
                Worksheet sheet = null;
                DataTable data = null;
                DataTable search = null;
                CellRange range = null;
                Empleado empleado = null;
                List<Empleado> empleados = new List<Empleado>();
                MemoryStream ms = null;
                string to = ConfigurationManager.AppSettings["MailTo"].ToString();
                System.Net.Mail.Attachment attachment = null;
                int? ascIni = null;
                int? ascFin = null;
                int? descIni = null;
                int? descFin = null;
                try
                {
                    if (Request.Files.Count == 2)
                    {
                        file = Request.Files["fileInputEnv"]; //Uploaded file
                        //Use the following properties to get file's name, size and MIMEType
                        workbook = new Workbook();
                        //Load a file and imports its data
                        workbook.LoadFromStream(file.InputStream);
                        //workbook.SaveToStream(, FileFormat.PDF);
                        //Initialize worksheet
                        sheet = workbook.Worksheets[0];
                        // get the data source that the grid is displaying data for
                        data = sheet.ExportDataTable();
                        sheet.Dispose();
                        workbook.Dispose();
                        file = null;
                        sheet = null;
                        workbook = null;
                        foreach (DataRow row in data.Rows)
                        {
                            empleado = new Empleado()
                            {
                                Nombre = (string)row[0],
                                Correo = (string)row[1]
                            };
                            file = Request.Files["fileInputCed"]; //Uploaded file
                            workbook = new Workbook();
                            workbook.LoadFromStream(file.InputStream);
                            sheet = workbook.Worksheets[0];
                            search = sheet.ExportDataTable();
                            try
                            {
                                range = sheet.FindString(@"SDI", false, false);
                                if (range != null)
                                {
                                    ascIni = range.Row + 2;
                                }
                            }
                            catch
                            {
                                ascIni = null;
                            }
                            finally
                            {
                                range = null;
                            }
                            try
                            {
                                range = sheet.FindString(empleado.Nombre, false, false);
                                if (range != null)
                                {
                                    ascFin = range.Row - 1;
                                }
                            }
                            catch
                            {
                                ascFin = null;
                            }
                            finally
                            {
                                range = null;
                            }
                            try
                            {
                                if (ascFin.HasValue)
                                {
                                    descIni = ascFin.Value;
                                    for (descIni = ascFin.Value; descIni < search.Rows.Count; descIni++)
                                    {
                                        if (search.Rows[descIni.Value][0] != DBNull.Value)
                                        {
                                            if (!string.IsNullOrEmpty((string)search.Rows[descIni.Value][0]))
                                            {
                                                if (((string)search.Rows[descIni.Value][0]).Replace(" ", "").Replace("-", "").Length == 11)
                                                {
                                                    descIni += 2;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    descIni = null;
                                }
                            }
                            catch
                            {
                                descIni = null;
                            }
                            finally
                            {
                                search = null;
                            }
                            try
                            {
                                range = sheet.FindString(@"Total de Días cotizados para el calculo de trabajadores promedio expuestos al riesgo", false, false);
                                if (range != null)
                                {
                                    descFin = range.Row - 4;
                                }
                            }
                            catch
                            {
                                descFin = null;
                            }
                            finally
                            {
                                range = null;
                            }
                            if (descIni.HasValue && descFin.HasValue)
                            {
                                if (descFin.Value - descIni.Value > 0)
                                {
                                    sheet.DeleteRow(descIni.Value, descFin.Value - descIni.Value);
                                }
                            }
                            if (ascIni.HasValue && ascFin.HasValue)
                            {
                                if (ascFin.Value - ascIni.Value > 0)
                                {
                                    sheet.DeleteRow(ascIni.Value, ascFin.Value - ascIni.Value);
                                }
                            }
                            //sheet.AllocatedRange.AutoFitColumns();
                            ms = new MemoryStream();
                            workbook.SaveToStream(ms, FileFormat.PDF);
                            sheet.Dispose();
                            workbook.Dispose();
                            file = null;
                            sheet = null;
                            workbook = null;
                            attachment = new System.Net.Mail.Attachment(ms, "Cédula IMSS.pdf", "application/pdf");
                            if (!string.IsNullOrEmpty(to))
                            {
                                empleado.Correo = to;
                            }
                            empleado.Procesado = Correo.Enviar(empleado, usuario.Nombre, app, "Cédula IMSS", "Se envía cédula IMSS.", attachment, logId);
                            empleados.Add(empleado);
                        }

                    }
                    return PartialView("Grid", empleados);
                }
                catch (Exception ex)
                {
                    LogTools.RegisterLog(logId, usuario.Nombre, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "Error al procesar la información: " + ex.Message + ".", DateTime.Now);
                    return PartialView("Grid");
                }
                finally
                {
                    file = null;
                    workbook = null;
                    sheet = null;
                    data = null;
                    empleado = null;
                }
            }
            else
            {
                LogTools.RegisterLog(logId, "", app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "No se han cargado los datos del usuario para poder procesar la carga de datos.", DateTime.Now);
                return PartialView("Grid");
            }
        }
         */
    }
}