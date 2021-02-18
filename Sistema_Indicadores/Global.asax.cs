using FirebaseAdmin;
using Google.Apis.Auth.OAuth2;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Configuration;
using System.IO;
using System.Web.Mvc;
using System.Web.Routing;

namespace Sistema_Indicadores
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            RouteConfig.RegisterRoutes(RouteTable.Routes);

            //AreaRegistration.RegisterAllAreas();
            //FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            //RouteConfig.RegisterRoutes(RouteTable.Routes);
            //BundleConfig.RegisterBundles(BundleTable.Bundles);

            /* - - - - - - - - - - FCM Admin SDK - - - - - - - - - - - - - - - - */            
        }

        //public void ConfigureServices(IServiceCollection services)
        //{
        //    // Initialize the Firebase Admin SDK
        //    FirebaseApp.Create();
        //    services.AddMvc();
        //}

        //public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        //{
        //    if (env.IsDevelopment())
        //    {
        //        app.UseDeveloperExceptionPage();
        //    }

        //    app.UseMvc();
        //}

        //public IServiceProvider ConfigureServices(IServiceCollection services)
        //{
           
        //    var googleCredential = _hostingEnvironment.ContentRootPath;
        //    var filePath = Configuration.GetSection("GoogleFirebase")["fileName"];
        //    googleCredential = Path.Combine(googleCredential, filePath);
        //    var credential = GoogleCredential.FromFile(googleCredential);
        //    FirebaseApp.Create(new AppOptions()
        //    {

        //        Credential = credential
        //    });
        //}
    }
}
