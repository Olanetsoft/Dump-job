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
    public class AutoDumpCSIRequestList
    {
        public AutoDumpCSIRequestList()
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
                            SPList list = web.Lists["CSIPortal Request"];

                            SPQuery qry = new SPQuery();
                            qry.Query =
                             @"<Where></Where>";
                            SPListItemCollection listItems = list.GetItems(qry);

                            //initiate the stringBuilder
                            var csv = new StringBuilder();

                            // Iterate through the listitems
                            foreach (SPListItem item in listItems)
                            {
                                var one = item.ReadItemFieldString(GlobalStrings.CSIRequest.Title);
                                var two = item.ReadItemFieldUser(GlobalStrings.CSIRequest.Requester)?.Name;
                                var three = item.ReadItemDateTime(GlobalStrings.CSIRequest.Date);
                                var four = item.ReadItemFieldString(GlobalStrings.CSIRequest.Email);
                                var five = item.ReadItemFieldString(GlobalStrings.CSIRequest.opportunityNumber);
                                var six = item.ReadItemFieldString(GlobalStrings.CSIRequest.Unit);
                                var seven = item.ReadItemFieldUser(GlobalStrings.CSIRequest.lineManager)?.Name;
                                var eight = item.ReadItemFieldString(GlobalStrings.CSIRequest.Category);
                                var nine = item.ReadItemFieldString(GlobalStrings.CSIRequest.strategicObjective);
                                var ten = item.ReadItemDateTime(GlobalStrings.CSIRequest.ProposedStartDate);
                                var eleven = item.ReadItemDateTime(GlobalStrings.CSIRequest.ProposedEndDate);
                                var twelve = item.ReadItemFieldString(GlobalStrings.CSIRequest.Description);
                                var thirteen = item.ReadItemFieldString(GlobalStrings.CSIRequest.Department);
                                var fourteen = item.ReadItemFieldString(GlobalStrings.CSIRequest.BusinessJustification);
                                var fifteen = item.ReadItemFieldString(GlobalStrings.CSIRequest.Benefit);
                                var sixteen = item.ReadItemFieldString(GlobalStrings.CSIRequest.Sponsor);
                                var seventeen = item.ReadItemFieldString(GlobalStrings.CSIRequest.CostSave);
                                var eighteen = item.ReadItemFieldString(GlobalStrings.CSIRequest.TimeSave);
                                var nineteen = item.ReadItemFieldString(GlobalStrings.CSIRequest.AuditClosure);
                                var twenty = item.ReadItemFieldString(GlobalStrings.CSIRequest.Automation);
                                var twentyone = item.ReadItemFieldString(GlobalStrings.CSIRequest.Timeline);
                                var twentytwo = item.ReadItemFieldString(GlobalStrings.CSIRequest.BenefitIT);
                                var twentythree = item.ReadItemFieldString(GlobalStrings.CSIRequest.CSAT);
                                var twentyfour = item.ReadItemFieldString(GlobalStrings.CSIRequest.Oxygen);
                                var twentyfive = item.ReadItemFieldString(GlobalStrings.CSIRequest.CIOApproval);
                                var twentysix = item.ReadItemFieldString(GlobalStrings.CSIRequest.Priority);
                                var twentyseven = item.ReadItemFieldString(GlobalStrings.CSIRequest.MaturityUplift);
                                var twentyeight = item.ReadItemFieldString(GlobalStrings.CSIRequest.Status);
                                var twentynine = item.ReadItemFieldString(GlobalStrings.CSIRequest.OwnerComment);
                                var thirty = item.ReadItemFieldString(GlobalStrings.CSIRequest.AverageRating);
                                var thirtyone = item[GlobalStrings.CSIRequest.Member];
                                var thirtytwo = item.ReadItemDateTime(GlobalStrings.CSIRequest.ActualStartDate);
                                var thirtythree = item.ReadItemDateTime(GlobalStrings.CSIRequest.ActualEndDate);
                                var thirtyfour = item.ReadItemFieldString(GlobalStrings.CSIRequest.ClosureRemark);
                                var thirtyfive = item.ReadItemFieldString(GlobalStrings.CSIRequest.FinalOutcome);



                                var newLine =
                                    string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}," +
                                    "{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28},{29},{30},{31},{32},{33},{34}",
                                        one, two, three, four, five, six, seven, eight, nine, ten, eleven, twelve, thirteen, fourteen, fifteen,
                                        sixteen, seventeen, eighteen, nineteen, twenty, twentyone, twentytwo, twentythree, twentyfour,
                                        twentyfive, twentysix, twentyseven, twentyeight, twentynine, thirty, thirtyone, thirtytwo,
                                        thirtythree, thirtyfour, thirtyfive);

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
                    LoggingService.LogError("CSIPortal Dump Job - AutoDumpCSIRequestList", msg);
                }
            });
        }
    }
}
