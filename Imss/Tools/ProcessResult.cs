using System;
using System.Collections.Generic;

namespace Imss.Tools
{
    public class ProcessResult
    {
        public bool HasErrors { get; set; }
        public string Message { get; set; }
        public List<Common.Empleado> Collection { get; set; } 

        public ProcessResult()
        {
            HasErrors = false;
            Message = string.Empty;
            Collection = new List<Common.Empleado>();
        }
    }
}
