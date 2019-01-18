using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Diagnostics;


namespace SPMigrationTool
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        //SP PnP
        private AuthenticationManager authManager = new AuthenticationManager();

        //Excel App
        Excel.Application xlApp;

        //Excel Workbook
        Excel.Workbook xlSettingsFile;

        //Excel sheets
        Excel.Worksheet xlSiteCollectionSheet;
        Excel.Worksheet xlListMappingsSheet;
        Excel.Worksheet xlMetadataMappingsSheet;
        Excel.Worksheet xlMigrationReportSheet;

        //Dictiobaries
        Dictionary<string, string> metadataMappingsDict = new Dictionary<string, string>();
        Dictionary<string, bool> migratedSitesDict = new Dictionary<string, bool>();

        //Lists
        List<MigrationSite> sitesForMig = new List<MigrationSite>();
        List<ListMapping> listMappings = new List<ListMapping>();

        public Form1()
        {
            InitializeComponent();
            btnStartMigration.Enabled = false;
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            //Update UI
            lbStatus.Items.Add("Opening Setting file at " + tbMappingFile.Text);


            //create instance of Excel and open settings file
            xlApp = new Excel.Application();
            xlApp.Visible = true;
            xlSettingsFile = xlApp.Workbooks.Open(tbMappingFile.Text);


            //Update UI
            lbStatus.Items.Add("Excel file opened");
            Application.DoEvents();

            //1: Create a mapping betweem project sites and site collections with the sheet NewSiteCollections
            //////////////////////////////////////////////////////////////////////////////////////////////////
            //Update UI
            lbStatus.Items.Add("Reading \"NewSiteCollections\" sheet");
            Application.DoEvents();
            if (CheckIfSheetExists(xlSettingsFile, "NewSiteCollections"))
            {
                xlSiteCollectionSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSettingsFile.Worksheets.Item["NewSiteCollections"];
                Excel.Range xlSiteCollectionRange = xlSiteCollectionSheet.UsedRange;
                for (int rCount = 2; rCount <= xlSiteCollectionRange.Rows.Count; rCount++)
                {
                    MigrationSite ms = new MigrationSite(
                        xlSiteCollectionSheet.Cells[rCount, 1].Value2.ToString(),
                        xlSiteCollectionSheet.Cells[rCount, 2].Value2.ToString(),
                        xlSiteCollectionSheet.Cells[rCount, 3].Value2.ToString(),
                        xlSiteCollectionSheet.Cells[rCount, 4].Value2.ToString(),
                        xlSiteCollectionSheet.Cells[rCount, 5].Value2.ToString()
                    );
                    sitesForMig.Add(ms);
                }
            }
            else
            {
                //"NewSiteCollections" sheet not found
            }




            //2: Create list mapping dict with the ListMappings sheet
            /////////////////////////////////////////////////////////
            //Update UI
            lbStatus.Items.Add("Reading \"ListMappings\" sheet");
            Application.DoEvents();
            if (CheckIfSheetExists(xlSettingsFile, "ListMappings"))
            {
                xlListMappingsSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSettingsFile.Worksheets.Item["ListMappings"];
                Excel.Range xlSiteListMappingsRange = xlListMappingsSheet.UsedRange;
                for (int rCount = 2; rCount <= xlSiteListMappingsRange.Rows.Count; rCount++)
                {
                    ListMapping ls = new ListMapping(
                        xlListMappingsSheet.Cells[rCount, 1].Value2.ToString(),
                        xlListMappingsSheet.Cells[rCount, 2].Value2.ToString(),
                        xlListMappingsSheet.Cells[rCount, 3].Value2.ToString()
                    );
                    listMappings.Add(ls);
                }
            }
            else
            {
                //"ListMappings" sheet not found
            }



            //3: Create Metadata Mappings with MetadataMappings sheet
            /////////////////////////////////////////////////////////
            //Update UI

            MetadataMappings mm = new MetadataMappings();

            metadataMappingsDict = mm.getDict();

            //lbStatus.Items.Add("Reading \"MetadataMappings\" sheet");
            //Application.DoEvents();
            //if (CheckIfSheetExists(xlSettingsFile, "MetadataMappings"))
            //{
            //    xlMetadataMappingsSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSettingsFile.Worksheets.Item["MetadataMappings"];
            //    Excel.Range xlMetadataMappingsRange = xlMetadataMappingsSheet.UsedRange;
            //    for (int rCount = 2; rCount <= xlMetadataMappingsRange.Rows.Count; rCount++)
            //    {
            //        if (xlMetadataMappingsSheet.Cells[rCount, 1].Value2 != null) {
            //            metadataMappingsDict.Add(xlMetadataMappingsSheet.Cells[rCount, 1].Value2.ToString(), xlMetadataMappingsSheet.Cells[rCount, 2].Value2.ToString());
            //        }
            //    }
            //}
            //else
            //{
            //    //"MetadataMappings" sheet not found
            //}



            //4.  Read Migration report sheet 
            /////////////////////////////////////////////////////////
            //Update UI
            lbStatus.Items.Add("Reading \"MigrationReport\" sheet");
            Application.DoEvents();
            if (CheckIfSheetExists(xlSettingsFile, "MigrationReport"))
            {
                xlMigrationReportSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSettingsFile.Worksheets.Item["MigrationReport"];
                Excel.Range xlMigrationReportRange = xlMigrationReportSheet.UsedRange;
                for (int rCount = 2; rCount <= xlMigrationReportRange.Rows.Count; rCount++)
                {
                    if (xlMigrationReportSheet.Cells[rCount, 1].Value2 != null)
                    {
                        migratedSitesDict.Add(xlMigrationReportSheet.Cells[rCount, 1].Value2.ToString(), Boolean.Parse(xlMigrationReportSheet.Cells[rCount, 2].Value2.ToString()));
                    }
                }
            }
            else
            {
                //"MigrationReport" sheet not found
            }

            //Update UI
            lbStatus.Items.Add("Sheets all read");
            Application.DoEvents();



            //save settings
            Properties.Settings.Default.MigrationSettings = tbMappingFile.Text;
            //Properties.Settings.Default.Username = tbUsername.Text;
            //Properties.Settings.Default.Password = tbPassword.Text;

            //Properties.Settings.Default.Save();

            //close excel

            Cursor = Cursors.Arrow;

            btnStartMigration.Enabled = true;
            //xlApp.Quit();


        }

        private void btnStartMigration_Click(object sender, EventArgs e)
        {
            //1.  Loop through sites being migrated
            foreach (MigrationSite ms in sitesForMig)
            {

                bool migrated = false;
                if (migratedSitesDict.TryGetValue(ms.OldSiteUrl, out migrated))
                {
                    if (migrated == true)
                    {
                        //Jumpt to next iteration in foreach loop if migrated is true
                        continue;
                    }
                }

                //Connect to SPO Site
                try
                {
                    using (var clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(ms.OldSiteUrl, tbUsername.Text, tbPassword.Text))
                    {
                        using (var DestClientContent = authManager.GetSharePointOnlineAuthenticatedContextTenant(ms.NewSiteUrl, tbUsername.Text, tbPassword.Text))
                        {
                            //Destination client content
                            Web destWeb = DestClientContent.Web;
                            DestClientContent.Load(destWeb);
                            DestClientContent.ExecuteQuery();

                            //1.  load lists and libraries
                            Web web = clientContext.Web;
                            ListCollection collectionLists = web.Lists;
                            clientContext.Load(
                                collectionLists,
                                lists => lists.Include(
                                    list => list.Title,
                                    list => list.Id,
                                    List => List.Hidden,
                                    List => List.IsCatalog,
                                    List => List.IsSiteAssetsLibrary,
                                    List => List.BaseTemplate
                                ));
                            clientContext.ExecuteQuery();

                            //2. Loop lists & libraries
                            foreach (List oList in collectionLists)
                            {
                                if (oList.Hidden == false && oList.IsSiteAssetsLibrary == false && oList.IsCatalog == false)
                                {

                                    try
                                    {
                                        //The find ListMapping can generate an exception and if it does then that list is not in scope for migration
                                        ListMapping lm = listMappings.Find(list => list.ListName == oList.Title);
                                        Debug.WriteLine("========================");
                                        Debug.WriteLine("List Name " + oList.Title);
                                        Debug.WriteLine("List Name " + lm.ListName + " " + lm.ListType + " " + lm.MappedListname);
                                        Debug.WriteLine("========================");

                                        //3. Load the list items
                                        ListItemCollection items = oList.GetItems(CamlQuery.CreateAllItemsQuery());
                                        clientContext.Load(items);
                                        clientContext.ExecuteQuery();

                                        //4. Loop through the lists items
                                        foreach (ListItem item in items)
                                        {
                                            clientContext.Load(item);
                                            clientContext.ExecuteQuery();

                                            //copy item - check if list or library
                                            if (oList.BaseTemplate == 100)
                                            {
                                                Debug.WriteLine("MIGRATING LIST");

                                            }
                                            else if (oList.BaseTemplate == 101)
                                            {
                                                Debug.WriteLine("MIGRATING LIBRARY");

                                                try
                                                {
                                                    Microsoft.SharePoint.Client.File file = item.File;
                                                    clientContext.Load(file);
                                                    clientContext.ExecuteQuery();
                                                    string destination = destWeb.ServerRelativeUrl.TrimEnd('/') + "/" + lm.MappedListname + "/" + file.Name;
                                                    Debug.WriteLine("DESTINATION " + destination);
                                                    FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, file.ServerRelativeUrl);
                                                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(DestClientContent, destination, fileInfo.Stream, true);


                                                    var uploadedFile = DestClientContent.Web.GetFileByServerRelativeUrl(destination);
                                                    var listItem = uploadedFile.ListItemAllFields;
                                                    //set properties
                                                    //5. Loop the fields
                                                    foreach (var obj in item.FieldValues)
                                                    {

                                                        Debug.WriteLine("   textField   " + obj.Key);
                                                        Debug.WriteLine("   Value   " + obj.Value);

                                                        //check if field in scope for migration
                                                        string fieldName;
                                                        if (metadataMappingsDict.TryGetValue(obj.Key, out fieldName))
                                                        {
                                                            Debug.WriteLine("========================");
                                                            Debug.WriteLine(obj.Key + " mapped to " + metadataMappingsDict[fieldName]);
                                                            Debug.WriteLine("========================");

                                        
                                                            listItem[metadataMappingsDict[fieldName]] = obj.Value;
                                          
                                                        }
                                                    }//end foreach (var obj in item.FieldValues)
                                                    listItem["Created"] = item["Created"];
                                                    listItem.Update();
                                                    DestClientContent.ExecuteQuery();


                                                    //update created date & metadata fields
                                                }
                                                catch (Exception ex)
                                                {
                                                    Debug.WriteLine(ex.Message);
                                                }


                                            }
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine("========================");
                                        Debug.WriteLine("NOT IN SCOPE " + oList.Title);
                                        Debug.WriteLine("========================");

                                    }

                                }
                            }
                            migrated = true;

                        }//End destination client context                       
                        
                    }//end using clientContent
                }
                catch (Exception ex)
                {

                }

                //End Connect to SPO Site

                //update Excel Report Sheet
                Excel.Range last = xlMigrationReportSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range range = xlMigrationReportSheet.get_Range("A1", last);
                int lastUsedRow = last.Row + 1;     
                xlMigrationReportSheet.Cells[lastUsedRow, 1] = ms.OldSiteUrl;
                xlMigrationReportSheet.Cells[lastUsedRow, 2] = migrated.ToString();
                xlSettingsFile.Save();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //load last values used
            tbMappingFile.Text = Properties.Settings.Default.MigrationSettings;
            tbUsername.Text = Properties.Settings.Default.Username;
            tbPassword.Text = Properties.Settings.Default.Password;
        }


        private bool CheckIfSheetExists(Excel.Workbook wb, string SheetName) {
            bool found = false;
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                // Check the name of the current sheet
                if (sheet.Name == SheetName)
                {
                    found = true;
                    break; // Exit the loop now
                }
            }
            return found;
        }


    }//partial class Form1
}
