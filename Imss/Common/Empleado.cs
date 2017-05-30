using System;
using System.ComponentModel.DataAnnotations;
using System.Data;

namespace Imss.Common
{
    public class Empleado
    {
        [Display(Name = "Empleado")]
        public string Nombre { get; set; }
        [Display(Name = "Correo Electrónico")]
        public string Correo { get; set; }
        [Display(Name = "Estatus")]
        public bool Procesado { get; set; }

        public bool Guardar()
        {
            int? rowsAffected = null;
            try
            {
                using (System.Data.SqlClient.SqlConnection conn = new System.Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["FinanzasDB"].ConnectionString))
                {
                    conn.Open();
                    using (System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand(
                            @"INSERT INTO imss_EmployeeToProcess(Empleado, Correo, Fecha, Procesado)
                            VALUES(@Employee, @MailAddress, GETDATE(), @Processed);", conn))
                    {
                        if (!string.IsNullOrEmpty(this.Nombre))
                        {
                            command.Parameters.Add("Employee", SqlDbType.NVarChar).Value = this.Nombre;
                        }
                        if (!string.IsNullOrEmpty(this.Correo))
                        {
                            command.Parameters.Add("MailAddress", SqlDbType.NVarChar).Value = this.Correo;
                        }
                        command.Parameters.Add("Processed", SqlDbType.Bit).Value = this.Procesado;
                        rowsAffected = command.ExecuteNonQuery();
                    }
                    conn.Close();
                }
                return rowsAffected.HasValue ? rowsAffected.Value > 0 : rowsAffected.HasValue;
            }
            catch
            {
                return false;
            }
            finally
            {
                rowsAffected = null;
            }
        }
    }
}
