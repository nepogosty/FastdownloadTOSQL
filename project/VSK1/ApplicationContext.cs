using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;

namespace VSK1
{
    class ApplicationContext: DbContext
    {
        //Класс-контекс, нужный для связи с MSSQL
        public DbSet<Person> Persons { get; set; }

        public ApplicationContext()
        {
            Database.EnsureCreated();

        }
        

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=(localdb)\\mssqllocaldb;Database=vskstrahovanie;Trusted_Connection=True;");
        }
    }
}
