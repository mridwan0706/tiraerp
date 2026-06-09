using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobNefo.Model.JobEWACNDN
{
    public class ReportEWACNDNModel
    {
        public string SalesOffice { get; set; }
        public string SoldToParty { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string DistributionChannel { get; set; }
        public string Division { get; set; }

        public string Delivery { get; set; }
        public DateTime? DeliveryDate { get; set; }
        public DateTime? GIDate { get; set; }
        public DateTime? PODDate { get; set; }

        public string Status { get; set; }
        public decimal? Amount { get; set; }

        public string ReceivingPoint { get; set; }
        public string Agent { get; set; }
        public string AgentName { get; set; }
        public string Remark { get; set; }

        public string AgingPendingPC { get; set; }
        public string AgingPendingInvoice { get; set; }
    }
}
