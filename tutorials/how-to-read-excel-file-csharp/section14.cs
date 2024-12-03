using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section14
    {
        public void Run()
        {
            public class CountryContext : DbContext
                {
                    public DbSet<Country> Countries { get; set; }
                    public CountryContext()
                    {
                        //TODO: Make async
                        Database.EnsureCreated();
                    }
                    /// <summary>
                    /// Configure context to use Sqlite
                    /// </summary>
                    /// <param name="optionsBuilder"></param>
                    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
                    {
                        var connection = new SqliteConnection($"Data Source=Country.db");
                        connection.Open();
                        var command = connection.CreateCommand();
                        //Create the database if it doesn't already exist
                        command.CommandText = $"PRAGMA foreign_keys = ON;";
                        command.ExecuteNonQuery();
                        optionsBuilder.UseSqlite(connection);
                        base.OnConfiguring(optionsBuilder);
                    }
                }
        }
    }
}