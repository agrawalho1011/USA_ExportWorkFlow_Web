using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using System.Text.Json.Serialization;
using USA_ExportWorkFlow_Web;
using USA_ExportWorkFlow_Web.DBContext;
using USA_ExportWorkFlow_Web.LDAP;
using USA_ExportWorkFlow_Web.Models;

var builder = WebApplication.CreateBuilder(args);
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");
builder.Services.AddDbContext<ApplicationDbContext>(options =>
	options.UseSqlServer(connectionString));

builder.Services.AddAntiforgery(o => o.HeaderName = "XSRF-TOKEN");


//builder.Services.AddCors(options =>
//{
//	options.AddPolicy(name: "MyAllowSpecificOrigins",
//					  builder =>
//					  {
//						  builder.WithOrigins("http://outlook.com",
//							  "http://www.outlook.com",
//							  "https://outlook.office365.com");
//					  });
//});

builder.Services.AddIdentity<UserMaster, IdentityRole>().
	AddRoles<IdentityRole>().AddUserManager<LdapUserManager<UserMaster>>()
	.AddEntityFrameworkStores<ApplicationDbContext>().AddDefaultTokenProviders();


// Add services to the container.
builder.Services.AddControllersWithViews().AddJsonOptions(options =>
	options.JsonSerializerOptions.ReferenceHandler = ReferenceHandler.IgnoreCycles
);
//builder.Services.AddScoped(typeof(IQueryHandlerAsync<ExportDataListQuery, ContainerTypeViewModel>), typeof(ExportDataListQueryAsync));

var app = builder.Build();

using (var scope = app.Services.CreateScope())
{
	var services = scope.ServiceProvider;
	SeedRoles.Initialize(services);
}

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
	app.UseExceptionHandler("/Home/Error");
	// The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
	app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapControllerRoute(
	name: "default",
	pattern: "{controller=Account}/{action=Login}/{id?}");

app.Run();
