using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaviaTech.SharePoint.Diagnostics;
using TaviaTech.SPExtensions;

namespace CSIPortalDumpJob.Helper
{
    public class AutoDumpListItems
    {
        public AutoDumpListItems()
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                string url = "http://sharepoint2016:6722/newportal/";

                try
                {
                    using (SPSite site = new SPSite(url))
                    {
                        using (SPWeb web = site.OpenWeb())
                        {
                            SPList list = web.Lists["CSIPortal Initiative Ratings"];

                            SPQuery qry = new SPQuery();
                            qry.Query =
                             @"<Where></Where>";
                            SPListItemCollection listItems = list.GetItems(qry);

                            //initiate the stringBuilder
                            var csv = new StringBuilder();

                            // Iterate through the listitems
                            foreach (SPListItem item in listItems)
                            {
                                var first = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.Title);
                                var second = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.VoteRating);
                                var third = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.opportunityNumber);
                                var fourth = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.Quarter);
                                var fifth = item.ReadItemDateTime(GlobalStrings.CSIVoteDetails.VoteDate);
                                var sixth = item.ReadItemFieldUser(GlobalStrings.CSIVoteDetails.Voter).Name;
                                var seven = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.RequesterEmail);
                                var eight = item.ReadItemFieldLookup(GlobalStrings.CSIVoteDetails.CSIInitiative).LookupId;

                                var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", first, second, third, fourth, fifth, sixth, seven, eight);
                                csv.AppendLine(newLine);
                            }

                            // Get today's date 
                            string datetimeString = string.Format("{0:yyyy-MM-dd_hh-mm-ss-tt}.csv", DateTime.Now);

                            //Initialise the file path
                            var filePath = @"C:\Temp\" + datetimeString;

                            // Write text to path
                            File.WriteAllText(filePath, csv.ToString());

                        }
                    }

                }
                catch (Exception exp)
                {
                    var msg = $"Exception: {exp.Message}. StackTrace: {exp.StackTrace}";
                    LoggingService.LogError("CSIPortal Dump Job - AutoDumpListItems", msg);
                }
            });
        }

    }
}
