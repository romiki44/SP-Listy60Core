using ListyApp.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp.Repositories
{
    public class SdsConfigRepo: ISdsConfigRepo
    {
        private readonly ApplicationDbContext db;

        public SdsConfigRepo(ApplicationDbContext db)
        {
            this.db = db;
        }

        public async Task<List<SdsConfig>> GetConfigsAsync()
        {
            return await db.SdsConfig.ToListAsync();
        }

        public async Task<SdsConfigList> LoadConfigListAsync()
        {
            List<SdsConfig> configs = await db.SdsConfig.ToListAsync();
            return new SdsConfigList(configs);
        }

        public SdsConfigList LoadConfigList()
        {
            List<SdsConfig> configs = db.SdsConfig.ToList();
            return new SdsConfigList(configs);
        }

        public async Task<SdsConfig?> GetConfig(int id)
        {
            return await db.SdsConfig.FirstOrDefaultAsync(c => c.Id == id);
        }
    }
}
