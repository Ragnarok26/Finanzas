using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Web;

namespace FinanzasIMSS.Models
{
    public class Usuario
    {
        public string Nombre { get; set; }
        public List<Permiso> Permisos { get; set; }
        public List<Permiso> ObtenerPermisos(string app)
        {
            Type listType = null;
            dynamic dataObjects = null;
            List<string> columnName = null;
            dynamic element = null;
            PropertyInfo propertyInfo = null;
            columnName = new List<string>();
            dynamic value = null;
            Type dataType = null;
            try
            {
                using (System.Data.SqlClient.SqlConnection conn = new System.Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["FinanzasDB"].ConnectionString))
                {
                    conn.Open();
                    using (System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand("SELECT ISNULL(p.Id, 0) AS Id, ISNULL(p.Nombre, '') AS Nombre FROM Permiso p LEFT OUTER JOIN PermisoUsuario pu ON p.Id = pu.IdPermiso LEFT OUTER JOIN Sistema s ON p.IdSistema = s.Id WHERE pu.Usuario = @Usuario AND s.Nombre = @NombreSistema GROUP BY p.Id, p.Nombre", conn))
                    {
                        if (!string.IsNullOrEmpty(this.Nombre))
                        {
                            command.Parameters.Add("Usuario", SqlDbType.NVarChar).Value = this.Nombre;
                        }
                        if (!string.IsNullOrEmpty(app))
                        {
                            command.Parameters.Add("NombreSistema", SqlDbType.NVarChar).Value = app;
                        }
                        using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                        {
                            for (int x = 0; x < reader.FieldCount; x++)
                            {
                                columnName.Add(reader.GetName(x));
                            }
                            listType = typeof(List<>).MakeGenericType(new Type[] { typeof(Permiso) });
                            dataObjects = Activator.CreateInstance(listType);
                            while (reader.Read())
                            {
                                element = Activator.CreateInstance(typeof(Permiso));
                                for (int x = 0; x < columnName.Count; x++)
                                {
                                    propertyInfo = element.GetType().GetProperty(columnName[x]);
                                    dataType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
                                    value = (reader[columnName[x]] == null ? null : (reader[columnName[x]] == DBNull.Value ? null : Convert.ChangeType(reader[columnName[x]], dataType)));
                                    propertyInfo.SetValue(element, value, null);
                                    propertyInfo = null;
                                }
                                dataObjects.Add(element);
                                element = null;
                            }
                            reader.Close();
                        }
                    }
                    conn.Close();
                }
                return dataObjects;
            }
            catch
            {
                return new List<Permiso>();
            }
            finally
            {
                listType = null;
                dataObjects = null;
                columnName = null;
                element = null;
                propertyInfo = null;
                columnName = null;
                value = null;
                dataType = null;
            }
        }
    }
}