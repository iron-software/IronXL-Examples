public class Country
    {
        [Key]
        public Guid Key { get; set; }
        public string Name { get; set; }
        public decimal GDP { get; set; }
    }
