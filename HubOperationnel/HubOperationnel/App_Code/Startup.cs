using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(HubOperationnel.Startup))]
namespace HubOperationnel
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
