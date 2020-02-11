using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Server.Models
{
    public class Order
    {
        public int Id { get; set; }
        public string Type { get; set; }

        public string Executor { get; set; }
        public string ExecutorAddress { get; set; }
        public string ExecutorINN { get; set; }

        public string Customer { get; set; }
        public string CustomerAddress { get; set; }
        public string CustomerINN { get; set; }

        public string Materials { get; set; }
        public string Squares { get; set; }
        public string Prices { get; set; }
        public string Cost { get; set; }
        public int TileNum { get; set; }

    }
}