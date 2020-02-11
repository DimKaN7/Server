using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace Server.Models
{
    public class OrdersContext : DbContext
    {
        public OrdersContext() : base("DefaultConnection")
        {
        }

        public DbSet<Order> Orders { get; set; }
    }
}