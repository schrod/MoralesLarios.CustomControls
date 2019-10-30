using System;
using System.Collections.Generic;

namespace MoralesLarios.Utilities.Helper.Excel
{
    internal interface IBuilderObjects
    {
        object BuildData(string line, Type type, List<string> columnNames);
        IEnumerable<object> BuildObject(string data, Type type, bool shwoErrorMessages, bool cancelWithErrors, List<string> columnNames);
    }
}