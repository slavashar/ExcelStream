namespace ExcelStream
{
    public class Font
    {
        public Font(int id)
        {
            this.Id = id;
        }

        public int Id { get; private set; }

        public bool Bold { get; set; }

        public bool Underline { get; set; }

        public string Color { get; set; }
    }
}
