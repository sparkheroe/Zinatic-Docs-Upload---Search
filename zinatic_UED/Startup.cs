using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(zinatic_UED.Startup))]
namespace zinatic_UED
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
