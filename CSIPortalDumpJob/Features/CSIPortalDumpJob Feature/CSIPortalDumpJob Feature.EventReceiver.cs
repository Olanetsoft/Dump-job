using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using CSIPortalDumpJob.Helper;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using TaviaTech.SharePoint.Diagnostics;

namespace CSIPortalDumpJob.Features.CSIPortalDumpJob_Feature
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("df191f30-4618-4bfb-b473-f57924bd9542")]
    public class CSIPortalDumpJob_FeatureEventReceiver : SPFeatureReceiver
    {
        // Uncomment the method below to handle the event raised after a feature has been activated.
        const string JobName = "MTN CSI PORTAL Dump Job";

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    SPWebApplication parentWebApp = (SPWebApplication)properties.Feature.Parent;

                    CreateJob(parentWebApp);
                });
            }
            catch (Exception exp)
            {
                var msg = $"Exception: {exp.Message}. StackTrace: {exp.StackTrace}";
                LoggingService.LogError("CSIPortal Dump Job - Feature Activated", msg);
            }

        }


        // Uncomment the method below to handle the event raised before a feature is deactivated.

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            lock (this)
            {
                try
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate ()
                    {
                        SPWebApplication parentWebApp = (SPWebApplication)properties.Feature.Parent;
                        DeleteExistingJob(JobName, parentWebApp);
                    });
                }
                catch (Exception exp)
                {
                    var msg = $"Exception: {exp.Message}. StackTrace: {exp.StackTrace}";
                    LoggingService.LogError("CSIPortal Dump Job - Feature DeActivated", msg);
                }
            }
        }


        // Uncomment the method below to handle the event raised after a feature has been installed.

        //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        //{
        //}


        // Uncomment the method below to handle the event raised before a feature is uninstalled.

        //public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        //{
        //}

        // Uncomment the method below to handle the event raised when a feature is upgrading.

        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        //{
        //}

            // Create Job
        private bool CreateJob(SPWebApplication site)
        {
            bool jobCreated = false;
            try
            {
                CSIPortalDump job = new CSIPortalDump(JobName, site);

                SPDailySchedule schedule = new SPDailySchedule();
                schedule.BeginHour = 1;
                schedule.EndHour = 2;
                schedule.BeginMinute = 0;
                schedule.EndMinute = 0;
                schedule.BeginSecond = 0;
                schedule.EndSecond = 0;

                job.Schedule = schedule;
                job.IsDisabled = true;
                job.Update();
            }
            catch (Exception)
            {

                return jobCreated;
            }
            return jobCreated;
        }

        // Delete Existing Job
        public bool DeleteExistingJob(string jobName, SPWebApplication site)
        {
            bool jobDeleted = false;
            try
            {
                foreach (SPJobDefinition job in site.JobDefinitions)
                {
                    if (job.Name == jobName)
                    {
                        job.Delete();
                        jobDeleted = true;
                    }
                }
            }
            catch (Exception)
            {
                return jobDeleted;
            }
            return jobDeleted;
        }
    }
}
