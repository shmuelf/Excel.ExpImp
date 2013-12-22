using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelExpImp
{
    public interface IExpImp<T>
    {
        IEnumerable<T> Import(Stream file, string fileExtension);
        void Export(IEnumerable<T> objs, string name);
    }
}
