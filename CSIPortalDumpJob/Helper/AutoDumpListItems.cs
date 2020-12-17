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
                // Dev Url
                //string url = "http://sharepoint2016:6722/newportal/";

                // Test Url
                string url = "https://testshare.mtnnigeria.net/csiportal/";

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

                            // Format the column titles
                            var titleData = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}", 
                                GlobalStrings.CSIVoteDetails.Title,
                                GlobalStrings.CSIVoteDetails.CSIInitiative, 
                                GlobalStrings.CSIVoteDetails.opportunityNumber, 
                                GlobalStrings.CSIVoteDetails.Quarter, 
                                GlobalStrings.CSIVoteDetails.VoteDate, 
                                GlobalStrings.CSIVoteDetails.Voter, 
                                GlobalStrings.CSIVoteDetails.Requester, 
                                GlobalStrings.CSIVoteDetails.RequesterEmail, 
                                GlobalStrings.CSIVoteDetails.VoteRating, 
                                GlobalStrings.CSIVoteDetails.ItemID);

                            //Append column Titles
                            csv.AppendLine(titleData);

                            // Iterate through the listitems
                            foreach (SPListItem item in listItems)
                            {

                                var Title = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.Title);
                                var CSIInitiative = item.ReadItemFieldLookup(GlobalStrings.CSIVoteDetails.CSIInitiative).LookupId;
                                var OppNum = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.opportunityNumber);
                                var Quarter = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.Quarter);
                                var voteDate = item.ReadItemDateTime(GlobalStrings.CSIVoteDetails.VoteDate);
                                var Voter = item.ReadItemFieldUser(GlobalStrings.CSIVoteDetails.Voter).Name;
                                var Requester = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.Requester);
                                var requesterEmail = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.RequesterEmail);
                                var voteRating = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.VoteRating);
                                var itemID = item.ReadItemFieldString(GlobalStrings.CSIVoteDetails.ItemID);

                                var Data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}", Title, CSIInitiative, OppNum, Quarter, voteDate, Voter, Requester, requesterEmail, voteRating, itemID);
                                csv.AppendLine(Data);
                            }

                            // Get today's date 
                            string datetimeString = string.Format("CSIRATINGS_LIST_{0:yyyy-MM-dd_hh-mm-ss-tt}.csv", DateTime.Now);

                            //Initialise the file path
                            //prod/test env
                            var filePath = @"E:\DUMP FILES\" + datetimeString;

                            // local env
                            //var filePath = @"C:\temp\" + datetimeString;

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
