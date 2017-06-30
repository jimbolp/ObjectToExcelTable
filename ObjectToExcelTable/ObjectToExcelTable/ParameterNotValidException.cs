using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace mvcTests.Models
{
    public class ParameterNotValidException : Exception
    {
        public override string Message { get; }
        public ParameterNotValidException(string message) :base(message)
        {
            Message = message;
        }
    }
}