using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RepairShoprCore
{
   public class Ticket
    {
       [JsonProperty("id")]
       public string Id
       {
           get;
           set;
       }
    }

   public class TicketRoot
   {
       [JsonProperty("ticket")]
       public Ticket ticket
       {
           get;
           set;
       }
   }
}
