using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows;

namespace MoralesLarios.Utilities.Helper.Excel
{
    internal class BuilderObjects : IBuilderObjects
    {


        public IEnumerable<object> BuildObject(string data, Type type, bool showErrorMessages, bool cancelWithErrors, List<string> columnNames)
        {
            var result = new List<object>();

            var lines = data.Split('\n').Where(a => !string.IsNullOrWhiteSpace(a)).ToList();

            foreach (var line in lines)
            {
                var lineDef = line;

                if (line.Last() == '\r') lineDef = line.Remove(line.Count() - 1);

                object resultObj = null;

                try
                {
                    resultObj = BuildData(lineDef, type, columnNames);

                    result.Add(resultObj);
                }
                catch (Exception ex)
                {
                    if (showErrorMessages) MessageBox.Show($"Error Building an object.{Environment.NewLine} Source:{line}{Environment.NewLine}{Environment.NewLine}Error:{ex.Message} + {Environment.NewLine}{ex?.InnerException.Message}",
                                                            "Error",
                                                            MessageBoxButton.OK,
                                                            MessageBoxImage.Error);

                    if (cancelWithErrors) return new List<object>();
                }
            }

            return result;
        }


        public object BuildData(string line, Type type, List<string> columnNames)
        {
            var fields = line.Split('\t');

            var properties = type.GetProperties();

            var result = Activator.CreateInstance(type);

            for (int i = 0; i < fields.Count(); i++)
            {
                bool found = false;
                for(int j = 0; j < properties.Count(); j++)
                {
                    //Console.WriteLine("{0} {1}", properties[j].Name, columnNames[i]);
                    if (properties[j].Name == columnNames[i])
                    {
                        //Console.WriteLine("{0}={1}", properties[j].Name, fields[i]);
                        var value = Converters.ConverterTo(fields[i], properties[j]);

                        properties[j].SetValue(result, value);
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    Console.WriteLine("Couldn't find match for {0}", columnNames[i]);
                }
            }

            return result;
        }


    }
}
