using Newtonsoft.Json;

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
