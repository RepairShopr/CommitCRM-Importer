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

        [JsonProperty("success")]
        public Boolean success { get; set; }

        [JsonProperty("message")]
        public String[] message { get; set; }

        [JsonProperty("params")]
        public Param params_ { get; set; }

        [JsonProperty("status")]
        public String status { get; set; }
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


    public class Param
    {
        [JsonProperty("email")]
        public String email { get; set; }

        [JsonProperty("firstname")]
        public String firstname { get; set; }

        [JsonProperty("lastname")]
        public String lastname { get; set; }
    }

}
