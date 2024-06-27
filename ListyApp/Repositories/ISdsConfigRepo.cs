using ListyApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp.Repositories
{
    public interface ISdsConfigRepo
    {
        Task<List<SdsConfig>> GetConfigsAsync();
        Task<SdsConfigList> LoadConfigListAsync();
        SdsConfigList LoadConfigList();
        Task<SdsConfig?> GetConfig(int id);
    }
}
