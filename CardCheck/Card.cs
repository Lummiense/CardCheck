using System.Collections.Generic;
using System.Net.NetworkInformation;

namespace CardCheck
{
    public class Card
    {
        public int ID { get; set; }
        public int BIN { get; set; }
        public string Brand { get; set; }
        public string Bank { get; set; }
        public string BINType { get; set; }
        public string BINLevel { get; set; }
        public string Country { get; set; }
        public bool Withdrawal { get; set; }
        public bool Put { get; set; }
        }
}