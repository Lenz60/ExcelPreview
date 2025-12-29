using Backend.Context;
using ExcelPreview.Repository;
using ExcelPreview.Repository.Interface;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Connection String 
builder.Services.AddDbContext<ExcelPreviewContext>(options => options.UseSqlServer(
                builder.Configuration.GetConnectionString("ExcelPreviewContext"))
            );

// Repo & Interface
builder.Services.AddScoped<IExcelRepository, ExcelRepository>();


builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
