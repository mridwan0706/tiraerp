using JobNefo.Model.JobEWACNDN;
using NefoModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobNefo.Contracts
{
    
    public interface IJobEWACNDNRepository
    {

        Task<bool> Post();
        Task<IEnumerable<ReportEWACNDNModel>> GetDataExcel();
        string CreateExcel(IEnumerable<ReportEWACNDNModel> models, BackupConfigModel folder);
        Task SendEmailAsync(EmailAddressCollection email, TimeSpan time, IEnumerable<ReportEWACNDNModel> models, string filename);

    }
}
