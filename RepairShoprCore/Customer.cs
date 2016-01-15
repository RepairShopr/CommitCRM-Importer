using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairShoprCore
{
    public class Customer
    {
        [JsonProperty("id")]
        public string Id
        {
            get;
            set;
        }

        [JsonProperty("fullname")]
        public string FullName
        {
            get;
            set;
        }
    }   

    public class CustomerRoot
    {
        [JsonProperty("customer")]
        public Customer customer
        {
            get;
            set;
        }
    }

    public class CustomerListRoot
    {
        [JsonProperty("customers")]
        public List<Customer> customers
        {
            get;
            set;
        }
    }
}
