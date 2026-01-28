using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using MonitoringSystem.Hubs;
using MonitoringSystem.Data;
using MonitoringSystem.Filters;
using static NuGet.Packaging.PackagingConstants;


var builder = WebApplication.CreateBuilder(args);

// Tambahkan konfigurasi koneksi database
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

builder.Services.AddDefaultIdentity<ApplicationUser>(
    options => options.SignIn.RequireConfirmedAccount = true
    )
    .AddEntityFrameworkStores<ApplicationDbContext>();

// Tambahkan RazorPages dan SignalR
builder.Services.AddRazorPages()
    .AddMvcOptions(options =>
    {
        options.Filters.Add<AuthorizeFilter>();
    });

builder.Services.AddHostedService<PlanUpdaterService>()
    .Configure<HostOptions>(options =>
    {
        // Jangan menghentikan aplikasi ketika ada unhandled exception
        options.BackgroundServiceExceptionBehavior = BackgroundServiceExceptionBehavior.Ignore;
    });

builder.Services.AddSignalR();

//builder.Services.AddHostedService<PlanUpdaterService>();


builder.Services.Configure<CookiePolicyOptions>(options =>
{
    options.CheckConsentNeeded = context => true;
    options.MinimumSameSitePolicy = SameSiteMode.None;
});

builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Identity/Account/Login";
    options.LogoutPath = "/Identity/Account/Logout";
    options.AccessDeniedPath = "/Identity/Account/AccessDenied";
    options.SlidingExpiration = true;
    options.ExpireTimeSpan = TimeSpan.FromHours(12); // Set session timeout
});

var app = builder.Build();

// Middleware
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

// Tambahkan authentication & authorization
app.UseAuthentication();
app.UseAuthorization();

app.MapHub<LossTimeHub>("/dataHub");
app.MapRazorPages();

app.Run();
