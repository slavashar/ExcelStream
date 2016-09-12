using System;

namespace ExcelStream
{
    public class Sheet
    {
        public Sheet(int id, string name)
        {
            this.Id = id;
            this.Url = new Uri("/xl/worksheets/sheet" + id + ".xml", UriKind.Relative);
            this.Name = name;
        }

        public int Id { get; private set; }

        public Uri Url { get; private set; }

        public string Name { get; private set; }
    }
}
