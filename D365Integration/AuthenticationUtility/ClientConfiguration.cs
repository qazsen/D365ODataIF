using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AuthenticationUtility
{
    public partial class ClientConfiguration
    {
        public static ClientConfiguration Default { get { return ClientConfiguration.OneBox; } }

        public static ClientConfiguration OneBox = new ClientConfiguration()
        {
            UriString = "https://redac-sand1devaos.sandbox.ax.dynamics.com/",
            UserName = "crm_test@redacinc.com",            
            // Insert the correct password here for the actual test.
            Password = "Redac@!2016",

            ActiveDirectoryResource = "https://redac-sand1devaos.sandbox.ax.dynamics.com",
            ActiveDirectoryTenant = "https://sts.windows.net/redacinc.com",
            ActiveDirectoryClientAppId = "7139ca74-6c4a-4b53-8e9a-49c56fc5e652",
            // Insert here the application secret when authenticate with AAD by the application
            ActiveDirectoryClientAppSecret = "6InYCSI10oIXXCXzcJgvKcg0fneYQ/iUsyWM5OUHcmk=",

            // Change TLS version of HTTP request from the client here
            // Ex: TLSVersion = "1.2"
            // Leave it empty if want to use the default version
            TLSVersion = "",
        };

        public string TLSVersion { get; set; }
        public string UriString { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string ActiveDirectoryResource { get; set; }
        public String ActiveDirectoryTenant { get; set; }
        public String ActiveDirectoryClientAppId { get; set; }
        public string ActiveDirectoryClientAppSecret { get; set; }
    }
}
