var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowLocalhost", policy =>
    {
        policy.WithOrigins(
            "http://localhost:51831",
            "https://localhost:51831",
            "http://localhost:5173",
            "https://localhost:5173",
            "http://localhost:3000",
            "https://localhost:3000"
        )
              .AllowAnyMethod()
              .AllowAnyHeader()
              .AllowCredentials();
    });
});

builder.Services.AddControllers();
builder.Services.AddOpenApi();

var app = builder.Build();

app.UseCors("AllowLocalhost");

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
    // Enable SPA proxy in development
    app.UseDefaultFiles();
    app.MapStaticAssets();
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

// Serve SPA - redirect all non-API routes to index.html
app.MapFallbackToFile("index.html");

app.Run();
