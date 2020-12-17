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
                            SPList list = web.Lists["CSIPortal Request"];

                            SPQuery qry = new SPQuery();
                            qry.Query =
                             @"<Where></Where>";
                            SPListItemCollection listItems = list.GetItems(qry);

                            //initiate the stringBuilder
                            var csv = new StringBuilder();

                            // format the column titles

                            var dataTitles =
                                string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}," +
                                "{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28},{29},{30},{31},{32},{33},{34},{35},{36},{37},{38},{39},{40},{41}",
                                    GlobalStrings.CSIRequest.ID,
                                    GlobalStrings.CSIRequest.Title,
                                    GlobalStrings.CSIRequest.Requester,
                                    GlobalStrings.CSIRequest.Date,
                                    GlobalStrings.CSIRequest.Email,
                                    GlobalStrings.CSIRequest.opportunityNumber,
                                    GlobalStrings.CSIRequest.Unit,
                                    GlobalStrings.CSIRequest.lineManager,
                                    GlobalStrings.CSIRequest.Category,
                                    GlobalStrings.CSIRequest.strategicObjective,
                                    GlobalStrings.CSIRequest.ProposedEndDate,
                                    GlobalStrings.CSIRequest.ProposedEndDate,
                                    GlobalStrings.CSIRequest.Description,
                                    GlobalStrings.CSIRequest.Department,
                                    GlobalStrings.CSIRequest.BusinessJustification,
                                    GlobalStrings.CSIRequest.Benefit,
                                    GlobalStrings.CSIRequest.Sponsor,
                                    GlobalStrings.CSIRequest.CostSave,
                                    GlobalStrings.CSIRequest.TimeSave,
                                    GlobalStrings.CSIRequest.AuditClosure,
                                    GlobalStrings.CSIRequest.Automation,
                                    GlobalStrings.CSIRequest.Timeline,
                                    GlobalStrings.CSIRequest.BenefitIT,
                                    GlobalStrings.CSIRequest.CSAT,
                                    GlobalStrings.CSIRequest.Oxygen,
                                    GlobalStrings.CSIRequest.CIOApproval,
                                    GlobalStrings.CSIRequest.Priority,
                                    GlobalStrings.CSIRequest.MaturityUplift,
                                    GlobalStrings.CSIRequest.Status,
                                    GlobalStrings.CSIRequest.OwnerComment,
                                    GlobalStrings.CSIRequest.AverageRating,
                                    GlobalStrings.CSIRequest.Member,
                                    GlobalStrings.CSIRequest.ActualStartDate,
                                    GlobalStrings.CSIRequest.ActualEndDate,
                                    GlobalStrings.CSIRequest.ClosureRemark,
                                    GlobalStrings.CSIRequest.FinalOutcome,
                                    GlobalStrings.CSIRequest.Quarter,
                                    GlobalStrings.CSIRequest.Outcome,
                                    GlobalStrings.CSIRequest.OwnerLinemanger,
                                    GlobalStrings.CSIRequest.Owner,
                                    GlobalStrings.CSIRequest.Comment,
                                    GlobalStrings.CSIRequest.Stage);

                            csv.AppendLine(dataTitles);

                            // Iterate through the listitems
                            foreach (SPListItem item in listItems)
                            {
                                try
                                {
                                    var ID = item.ReadItemFieldInt(GlobalStrings.CSIRequest.ID);
                                    var Title = item.ReadItemFieldString(GlobalStrings.CSIRequest.Title);
                                    var Requester = item.ReadItemFieldUser(GlobalStrings.CSIRequest.Requester).Name;
                                    var Date = item.ReadItemDateTime(GlobalStrings.CSIRequest.Date);
                                    var Email = item.ReadItemFieldString(GlobalStrings.CSIRequest.Email);
                                    var oppNum = item.ReadItemFieldString(GlobalStrings.CSIRequest.opportunityNumber);
                                    var Unit = item.ReadItemFieldString(GlobalStrings.CSIRequest.Unit);
                                    var lineManager = item.ReadItemFieldUser(GlobalStrings.CSIRequest.lineManager).Name;
                                    var Category = item.ReadItemFieldString(GlobalStrings.CSIRequest.Category);
                                    var StrategicObj = item.ReadItemFieldString(GlobalStrings.CSIRequest.strategicObjective);
                                    var propStart = item.ReadItemDateTime(GlobalStrings.CSIRequest.ProposedStartDate);
                                    var propEnd = item.ReadItemDateTime(GlobalStrings.CSIRequest.ProposedEndDate);
                                    var Description = item.ReadItemFieldString(GlobalStrings.CSIRequest.Description);
                                    var Department = item.ReadItemFieldString(GlobalStrings.CSIRequest.Department);
                                    var businessJustification = item.ReadItemFieldString(GlobalStrings.CSIRequest.BusinessJustification);
                                    var Benefit = item.ReadItemFieldString(GlobalStrings.CSIRequest.Benefit);
                                    var Sponser = item.ReadItemFieldString(GlobalStrings.CSIRequest.Sponsor);
                                    var costSave = item.ReadItemFieldString(GlobalStrings.CSIRequest.CostSave);
                                    var timeSave = item.ReadItemFieldString(GlobalStrings.CSIRequest.TimeSave);
                                    var auditClosure = item.ReadItemFieldString(GlobalStrings.CSIRequest.AuditClosure);
                                    var Automation = item.ReadItemFieldString(GlobalStrings.CSIRequest.Automation);
                                    var Timeline = item.ReadItemFieldString(GlobalStrings.CSIRequest.Timeline);
                                    var benefitIT = item.ReadItemFieldString(GlobalStrings.CSIRequest.BenefitIT);
                                    var CSAT = item.ReadItemFieldString(GlobalStrings.CSIRequest.CSAT);
                                    var Oxygen = item.ReadItemFieldString(GlobalStrings.CSIRequest.Oxygen);
                                    var CIOApproval = item.ReadItemFieldString(GlobalStrings.CSIRequest.CIOApproval);
                                    var Priority = item.ReadItemFieldString(GlobalStrings.CSIRequest.Priority);
                                    var MaturityUplift = item.ReadItemFieldString(GlobalStrings.CSIRequest.MaturityUplift);
                                    var Status = item.ReadItemFieldString(GlobalStrings.CSIRequest.Status);
                                    var ownerComment = item.ReadItemFieldString(GlobalStrings.CSIRequest.OwnerComment);
                                    var averageRating = item.ReadItemFieldString(GlobalStrings.CSIRequest.AverageRating);
                                    var Member = item[GlobalStrings.CSIRequest.Member];
                                    var actualStart = item.ReadItemDateTime(GlobalStrings.CSIRequest.ActualStartDate);
                                    var actualEnd = item.ReadItemDateTime(GlobalStrings.CSIRequest.ActualEndDate);
                                    var closureRemark = item.ReadItemFieldString(GlobalStrings.CSIRequest.ClosureRemark);
                                    var finalOutcome = item.ReadItemFieldString(GlobalStrings.CSIRequest.FinalOutcome);
                                    var Quarter = item.ReadItemFieldString(GlobalStrings.CSIRequest.Quarter);
                                    var Outcome = item.ReadItemFieldString(GlobalStrings.CSIRequest.Outcome);
                                    var ownerLineManager = item.ReadItemFieldUser(GlobalStrings.CSIRequest.OwnerLinemanger).Name;
                                    var Owner = item.ReadItemFieldUser(GlobalStrings.CSIRequest.Owner).Name;
                                    var Comment = item.ReadItemFieldString(GlobalStrings.CSIRequest.Comment);
                                    var Stage = item.ReadItemFieldString(GlobalStrings.CSIRequest.Stage);


                                    // Data
                                    var newLine =
                                        string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}," +
                                    "{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28},{29},{30},{31},{32},{33},{34},{35},{36},{37},{38},{39},{40},{41}",
                                            ID, Title, Requester, Date, Email, oppNum, Unit, lineManager, Category, StrategicObj, propStart, propEnd, Description, Department, businessJustification,
                                            Benefit, Sponser, costSave, timeSave, auditClosure, Automation, Timeline, benefitIT, CSAT, Oxygen, CIOApproval, Priority, MaturityUplift, Status, ownerComment, averageRating, Member,
                                            actualStart, actualEnd, closureRemark, finalOutcome, Quarter, Outcome, ownerLineManager, Owner, Comment, Stage);

                                    // Append Data
                                    csv.AppendLine(newLine);
                                }
                                catch (Exception exp)
                                {

                                    var msg = $"Exception: {exp.Message}. StackTrace: {exp.StackTrace}";
                                    LoggingService.LogError("CSIPortal Request List- Job", msg);
                                }



                            }

                            // Get today's date 
                            string datetimeString = string.Format("CSIREQUEST_LIST_{0:yyyy-MM-dd_hh-mm-ss-tt}.csv", DateTime.Now);

                            //Initialise the file path
                            //Prod/test env
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
                    LoggingService.LogError("CSIPortal Dump Job - AutoDumpCSIRequestList", msg);
                }
            });
        }
    }
}
