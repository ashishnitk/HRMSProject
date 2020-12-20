using DinkToPdf;
using DinkToPdf.Contracts;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PDF_Generator.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace Jivi
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {


            //windows
            string filePath = $@"{System.IO.Directory.GetCurrentDirectory()}\libwkhtmltox.dll";

            ////linux
            //string filePath = @$"{(($@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}/libwkhtmltox.so").Replace(@"\", @"/"))}";

            CustomAssemblyLoadContext context = new CustomAssemblyLoadContext();
            context.LoadUnmanagedLibrary(filePath);
            // serviceCollection.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));


            services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
            services.AddControllers();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseHttpsRedirection();

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
