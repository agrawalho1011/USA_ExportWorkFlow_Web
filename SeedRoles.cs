using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using USA_ExportWorkFlow_Web.DBContext;

namespace USA_ExportWorkFlow_Web
{
    public static class SeedRoles
    {
        public static void Initialize(IServiceProvider serviceProvider)
        {
            using (var context = new ApplicationDbContext(serviceProvider.GetRequiredService<DbContextOptions<ApplicationDbContext>>()))
            {
                string[] roles = new string[] { "Administrator", "Manager", "User" };

                var newrolelist = new List<IdentityRole>();
                foreach (string role in roles)
                {
                    if (!context.Roles.Any(r => r.Name == role))
                    {
                        newrolelist.Add(new IdentityRole(role));
                    }
                }
                context.Roles.AddRange(newrolelist);
                context.SaveChanges();



            }
        }

    }
}
