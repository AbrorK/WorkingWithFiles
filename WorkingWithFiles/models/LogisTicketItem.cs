using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WorkingWithFiles.models
{
    public class LogisTicketItem
    {
        public int deal_number { get; set; }
        public DateTime deal_date { get; set; }
        public decimal? deal_price { get; set; }
        public decimal? start_price { get; set; }
        public int term_delivering_days { get; set; }
        public int term_payment_days { get; set; }
        public int warranty_id { get; set; }
        public string warranty_name { get; set; }
        public string cargo_name { get; set; }
        public string unit { get; set; }
        public int weight { get; set; }
        public string description { get; set; }
        public string package_name { get; set; }
        public int pledge_percent { get; set; }
        public string length { get; set; } = "-";
        public string width { get; set; } = "-";
        public string height { get; set; } = "-";
        public string provider_name { get; set; }
        public string provider_address { get; set; }
        public string provider_zip { get; set; }
        public string provider_phone { get; set; }
        public string provider_tin { get; set; }
        public string provider_oked { get; set; }
        public string provider_director { get; set; }
        public string provider_bank_account { get; set; }
        public string provider_bank_name { get; set; }
        public string provider_mfo { get; set; }
        public string customer_name { get; set; }
        public string customer_address { get; set; }
        public string customer_zip { get; set; }
        public string customer_phone { get; set; }
        public string customer_tin { get; set; }
        public string customer_oked { get; set; }
        public string customer_director { get; set; }
        public string customer_bank_account { get; set; }
        public string customer_bank_name { get; set; }
        public string customer_mfo { get; set; }
        public string from_address { get; set; }
        public string to_address { get; set; }
        public decimal? customer_commission { get; set; } = 0;
        public decimal? provider_commission { get; set; } = 0;
    }
}
