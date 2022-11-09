using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace WorkerEmail
{
    public class Model 
    {
        public class TradeProgress
        {
            [Key]
            public string buyerId { get; set; }
            public string sellerId { get; set; }
            public DateTime businessdate { get; set; }
        }
        public class ClearingMember
        {
            [Key]
            public string code { get; set; }
            public string StatusDomisiliFlag { get; set; }
        }
    }
}
