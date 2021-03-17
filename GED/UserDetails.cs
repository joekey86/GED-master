using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GED
{
    public class UserDetails
    {
        [JsonProperty("UF-Default-Label")]
        public string UFDefaultLabel { get; set; }

        [JsonProperty("UF-Metadata-ID")]
        public string UFMetadataID { get; set; }
    
        public string UserName {get;set;}

    }
    public class ListUserDetails
    {
        public List<UserDetails> usersDetails { get; set; }
    }
}
