using Microsoft.AspNetCore.Mvc;
using WebAPISample;

bool devMode = false;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
        {
            policy.WithOrigins("https://apps1.daves.tips", "https://outlook.office.com")
            .AllowAnyHeader()
            .AllowAnyMethod();
        });
});

// Add services to the container.

if (devMode)
{
    builder.Services.AddControllers().AddMvcOptions(options =>
        options.Filters.Add(
            new ResponseCacheAttribute
            {
                NoStore = true,
                Location = ResponseCacheLocation.None
            }));

    // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
    builder.Services.AddEndpointsApiExplorer();
    builder.Services.AddSwaggerGen();

}

builder.Services.AddControllers(o => o.InputFormatters.Insert(o.InputFormatters.Count, new TextPlainInputFormatter()));

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseCors();

app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers.Append(
            "access-control-allow-origin", "https://outlook.office.com");
    }
});

app.UseAuthorization();

app.MapControllers();

app.Run();
