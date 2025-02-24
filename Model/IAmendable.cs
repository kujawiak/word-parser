using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WordParser.Model
{
    public interface IAmendable
    {
        List<Amendment> Amendments { get; set; }
    }
}