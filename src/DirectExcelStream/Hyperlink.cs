namespace ExcelStream
{
    public class Hyperlink
    {
        public Hyperlink(string @ref, string location, string display)
        {
            this.Ref = @ref;
            this.Location = location;
            this.Display = display;
        }

        public string Ref { get; private set; }

        public string Location { get; private set; }

        public string Display { get; private set; }
    }
}
