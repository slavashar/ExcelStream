namespace ExcelStream
{
    public class Fill
    {
        public Fill(int id)
        {
            this.Id = id;
        }

        public int Id { get; private set; }

        public string ForegroundColourRGB { get; set; }

        public int? ForegroundColourIndexed { get; set; }

        public string BackgroundColourRGB { get; set; }

        public int? BackgroundColourIndexed { get; set; }
    }
}
