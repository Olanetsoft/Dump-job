using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSIPortalDumpJob.Helper
{
    public class CSIPortalDump : SPJobDefinition
    {
        public CSIPortalDump() : base()
        {

        }


        public CSIPortalDump(string jobName, SPService service) : base(jobName, service, null, SPJobLockType.None)
        {
            this.Title = "MTN CSI PORTAL Dump Job";
        }


        public CSIPortalDump(string jobName, SPWebApplication webapp) : base(jobName, webapp, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "MTN CSI PORTAL Dump Job";

        }

        public override void Execute(Guid targetInstanceId)
        {
            // Dump csi Initiative ratings
            AutoDumpListItems AutoDumpListItems = new AutoDumpListItems();

            // Dump csi request list
            AutoDumpCSIRequestList AutoDumpCSIRequestList = new AutoDumpCSIRequestList();
            
        }
    }
}
