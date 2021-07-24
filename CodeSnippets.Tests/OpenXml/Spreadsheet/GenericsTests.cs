using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CodeSnippets.Tests.OpenXml.Spreadsheet
{
    public class ModelBase
    {
        private readonly Dictionary<string, object> _values = new Dictionary<string, object>();

        public object GetValue(string fieldName)
        {
            return _values.TryGetValue(fieldName, out object value) ? value : null;
        }

        public void SetValue(string fieldName, object value)
        {
            _values[fieldName] = value;
        }
    }

    public class Foo : ModelBase
    {
        public string First
        {
            get => (string) GetValue(nameof(First));
            set => SetValue(nameof(First), value);
        }

        public int Second
        {
            get => (int) GetValue(nameof(Second));
            set => SetValue(nameof(Second), value);
        }
    }

    public class Bar : ModelBase
    {
        public string One
        {
            get => (string)GetValue(nameof(One));
            set => SetValue(nameof(One), value);
        }

        public DateTime Two
        {
            get => (DateTime) GetValue(nameof(Two));
            set => SetValue(nameof(Two), value);
        }
    }

    public class SheetConfig<T> where T : ModelBase
    {
        /// <summary>
        /// Name of the sheet (shown in the tab)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Objects to display in the sheet
        /// </summary>
        public IList<T> RowData { get; set; }

        /// <summary>
        /// Field names to extract from the RowData and use as header names
        /// </summary>
        public List<ColumnConfig> Columns { get; set; }
    }

    public class ColumnConfig
    {
        /// <summary>
        /// Header text
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Object property to use under header
        /// </summary>
        public string FieldName { get; set; }
    }

    public class Spreadsheet
    {
        public static void CreateDocument(IEnumerable<SheetConfig<ModelBase>> configs)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create("file.path.here", SpreadsheetDocumentType.Workbook))
            {
                // OpenXml stuff...
                var sheets = new Sheets();

                foreach (SheetConfig<ModelBase> s in configs)
                {
                    sheets.AppendChild(CreateSheet(s));
                    var y = sheets.Elements<Sheet>().Where(x => x.GetType().GetProperties().Any(p => p.GetValue(x) != null));
                }
            }
        }

        private static Sheet CreateSheet(SheetConfig<ModelBase> config)
        {
            // OpenXML stuff...
            return new Sheet { Name = config.Name };
        }
    }
}
