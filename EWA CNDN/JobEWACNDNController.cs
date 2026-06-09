using JobNefo.Contracts;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace JobNefo.Controllers
{
   

    public class JobEWACNDNController
    {
        private readonly ILogger<JobEWACNDNController> _log;
        private readonly IJobEWACNDNRepository _repository;

        public JobEWACNDNController(
            ILogger<JobEWACNDNController> log,
            IJobEWACNDNRepository repository)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public async Task<bool> Execute(string id, CancellationToken ct)
        {
            try
            {
                _log.LogInformation("[{Id}] Controller Execute Start.", id);

                if (ct.IsCancellationRequested)
                {
                    _log.LogWarning("[{Id}] Proses dibatalkan sebelum mulai.", id);
                    return false;
                }
                var result = await _repository.Post();

                _log.LogInformation("[{Id}] Controller Execute End. Result: {Result}", id, result);
                return result;
            }
            catch (OperationCanceledException)
            {
                _log.LogWarning("[{Id}] Proses dibatalkan saat controller berjalan.", id);
                return false;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "[{Id}] Error pada controller.", id);
                return false;
            }
        }
    }
}
