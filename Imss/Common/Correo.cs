using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Mail;
using System.Reflection;
using System.Threading.Tasks;

namespace Imss.Common
{
    public class Correo
    {
        public static bool Enviar(Empleado empleado, string usuario, string app, string subject, string body, List<Attachment> attachments, long logId)
        {
            Tools.LogTools.RegisterLog(logId, usuario, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "sending mail...", DateTime.Now);

            try
            {
                //validamos datos
                if (empleado.Correo == null)
                {
                    throw new Exception("Email de destinatario no puede estar vacío.");
                }

                //creamos el objeto mail
                SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["MailSMTP"].ToString());
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["MailUser"].ToString(), ConfigurationManager.AppSettings["MailPass"].ToString());
                MailMessage message = new MailMessage();

                //llenamos el remitente y el destinatario
                message.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"].ToString());

                foreach (string mailDir in empleado.Correo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    message.To.Add(mailDir.Trim());
                }
                foreach (string mailDir in ConfigurationManager.AppSettings["MailBcc"].ToString().Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    message.Bcc.Add(mailDir.Trim());
                }

                //agregamos archivos adjuntos
                if (attachments != null)
                {
                    if (attachments.Count > 0)
                    {
                        foreach (Attachment attachment in attachments)
                        {
                            message.Attachments.Add(attachment);
                        }
                    }
                }

                //configuramos el mensaje
                message.Body = body;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.IsBodyHtml = true;
                message.Subject = subject;
                message.SubjectEncoding = System.Text.Encoding.UTF8;

                //enviamos el mail
                client.TargetName = "STARTTLS/smtp.office365.com";
                client.Send(message);
                Tools.LogTools.RegisterLog(logId, usuario, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, "Mail Sent...", DateTime.Now);
                return true;
            }
            catch (Exception ex)
            {
                Tools.LogTools.RegisterLog(logId, usuario, app, MethodBase.GetCurrentMethod().DeclaringType.Name, MethodBase.GetCurrentMethod().Name, ex.Message, DateTime.Now);
            }
            return false;
        }
    }
}
