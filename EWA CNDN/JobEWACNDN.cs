using JobNefo.Controllers;
using JobNefo.Handler;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace JobNefo.Job
{
   

    public class JobEWACNDN : IJobHandler
    {
        public string Name => "EWA_CNDN";

        private readonly ILogger<JobEWACNDNController> _log;
        private readonly JobEWACNDNController _controller;

        public JobEWACNDN(ILogger<JobEWACNDNController> log, JobEWACNDNController controller)
        {
            _log = log;
            _controller = controller;
        }

        public async Task RunAsync(CancellationToken ct)
        {
            var id = Guid.NewGuid().ToString("N")[..8];
            _log.LogInformation("EWA CNDN starting (id={id})", id);
            await _controller.Execute(id, ct);
            _log.LogInformation("EWA CNDN finished (id={id})", id);
        }
    }
}
