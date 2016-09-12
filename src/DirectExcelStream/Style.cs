namespace ExcelStream
{
    public class Style
    {
        public Style(int id)
        {
            this.Id = id;
        }

        public int Id { get; private set; }

        public Font Font { get; set; }

        public Fill Fill { get; set; }

        public StyleFormat? NumberFormat { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; }
    }
}
