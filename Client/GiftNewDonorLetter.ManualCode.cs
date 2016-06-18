//
// DO NOT REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
//
// @Authors:
//       timop
//
// Copyright 2004-2016 by OM International
//
// This file is part of OpenPetra.org.
//
// OpenPetra.org is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// OpenPetra.org is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with OpenPetra.org.  If not, see <http://www.gnu.org/licenses/>.
//
using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing.Printing;
using System.Xml;
using GNU.Gettext;
using Ict.Common;
using Ict.Common.Controls;
using Ict.Common.Printing;
using Ict.Common.Verification;
using Ict.Common.Data;
using Ict.Common.IO;
using Ict.Petra.Client.CommonDialogs;
using Ict.Petra.Client.App.Core.RemoteObjects;
using Ict.Petra.Client.CommonControls;
using Ict.Petra.Plugins.NewDonorSubscriptions.RemoteObjects;
using Ict.Petra.Shared;
using Ict.Petra.Shared.MPartner;
using Ict.Petra.Shared.MFinance.Gift.Data;
using Ict.Petra.Shared.MPartner.Partner.Data;
using Ict.Petra.Shared.Interfaces.MFinance;
using Ict.Petra.Plugins.NewDonorSubscriptions.Data;

namespace Ict.Petra.Plugins.NewDonorSubscriptions.Client
{
    partial class TFrmGiftNewDonorLetter
    {
        private void InitializeManualCode()
        {
            dtpStartDate.Date = DateTime.Now.AddMonths(-1);
            dtpEndDate.Date = DateTime.Now;

            cmbPublicationCode.ColumnWidthCol3 = 0;
            cmbPublicationCode.Text = TAppSettingsManager.GetValue("NewDonorSubscriptions.Publication", true);
            //txtExtract.Text = TAppSettingsManager.GetValue("NewDonorSubscriptions.Extract", true);
            txtPathHTMLTemplate.Text = TAppSettingsManager.GetValue("NewDonorSubscriptions.HtmlTemplate", true);
        }

        private void FilterChanged(System.Object sender, EventArgs e)
        {
        }

        private void GenerateLetters(System.Object sender, EventArgs e)
        {
            if ((!dtpStartDate.Date.HasValue) || (!dtpEndDate.Date.HasValue))
            {
                MessageBox.Show(Catalog.GetString("Please supply valid Start and End dates."));
                return;
            }

            TMNewDonorSubscriptionsNamespace TRemote = new TMNewDonorSubscriptionsNamespace();
            FMainDS = TRemote.Server.WebConnectors.GetNewDonorSubscriptions(
                cmbPublicationCode.GetSelectedString(),
                dtpStartDate.Date.Value, dtpEndDate.Date.Value,
                string.Empty, // txtExtract.Text,
                TAppSettingsManager.GetValue("NewDonorSubscriptions.AttributeGroupCode", "BEDANKUNG"),
                TAppSettingsManager.GetValue("NewDonorSubscriptions.AttributeDetailCode", "ERSTSPENDER"),
                true);

            if (FMainDS.LetterRecipient.Rows.Count == 0)
            {
                MessageBox.Show(Catalog.GetString("There are no letters with valid addresses for your current parameters"),
                    Catalog.GetString("Nothing to print"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // TODO: for some reason, the columns' initialisation in the constructor does not have any effect; need to do here again???
            grdDonorDetails.Columns.Clear();
            grdDonorDetails.AddTextColumn("Donor", FMainDS.AGift.ColumnDonorKey);
            grdDonorDetails.AddTextColumn("DonorShortName", FMainDS.AGift.ColumnDonorShortName);
            grdDonorDetails.AddTextColumn("Recipient", FMainDS.AGift.ColumnRecipientDescription);
            grdDonorDetails.AddDateColumn("Subscription Start", FMainDS.AGift.ColumnDateOfSubscriptionStart);
            grdDonorDetails.AddDateColumn("Date Gift", FMainDS.AGift.ColumnDateOfFirstGift);

            FMainDS.AGift.DefaultView.Sort = NewDonorTDSAGiftTable.GetDonorShortNameDBName();

            FMainDS.AGift.DefaultView.AllowNew = false;
            grdDonorDetails.DataSource = new DevAge.ComponentModel.BoundDataView(FMainDS.AGift.DefaultView);
            grdDonorDetails.AutoResizeGrid();

            PrintLetters();
        }

        private Int32 FNumberOfPages = 0;
        private TGfxPrinter FGfxPrinter = null;

        private void PrintLetters()
        {
            string letterTemplateFilename = txtPathHTMLTemplate.Text;

            if (letterTemplateFilename == String.Empty)
            {
                OpenFileDialog DialogOpen = new OpenFileDialog();
    
                DialogOpen.Filter = Catalog.GetString("HTML file (*.html)|*.html;*.htm");
                DialogOpen.RestoreDirectory = true;
                DialogOpen.Title = Catalog.GetString("Select the template for the letter to the new donors");
    
                if (DialogOpen.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                letterTemplateFilename = DialogOpen.FileName;
            }

            StreamReader sr = new StreamReader(letterTemplateFilename);

            string htmlTemplate = sr.ReadToEnd();

            sr.Close();

            TMNewDonorSubscriptionsNamespace TRemote = new TMNewDonorSubscriptionsNamespace();
            StringCollection Letters = TRemote.Server.WebConnectors.PrepareNewDonorLetters(ref FMainDS, htmlTemplate);

            System.Drawing.Printing.PrintDocument printDocument = new System.Drawing.Printing.PrintDocument();
            bool printerInstalled = printDocument.PrinterSettings.IsValid;

            if (printerInstalled)
            {
                string AllLetters = String.Empty;

                foreach (string letter in Letters)
                {
                    if (AllLetters.Length > 0)
                    {
                        // AllLetters += "<div style=\"page-break-before: always;\"/>";
                        string body = letter.Substring(letter.IndexOf("<body"));
                        body = body.Substring(0, body.IndexOf("</html"));
                        AllLetters += body;
                    }
                    else
                    {
                        // without closing html
                        AllLetters += letter.Substring(0, letter.IndexOf("</html"));
                    }
                }

                if (AllLetters.Length > 0)
                {
                    AllLetters += "</html>";

                    FGfxPrinter = new TGfxPrinter(printDocument, TGfxPrinter.ePrinterBehaviour.eFormLetter);
                    try
                    {
                        TPrinterHtml htmlPrinter = new TPrinterHtml(AllLetters,
                            System.IO.Path.GetDirectoryName(letterTemplateFilename),
                            FGfxPrinter);
                        FGfxPrinter.Init(eOrientation.ePortrait, htmlPrinter, eMarginType.ePrintableArea);
                        this.ppvLetters.InvalidatePreview();
                        this.ppvLetters.Document = FGfxPrinter.Document;
                        this.ppvLetters.Zoom = 1;
                        FGfxPrinter.Document.EndPrint += new PrintEventHandler(this.EndPrint);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
            }
        }

        private void EndPrint(object ASender, PrintEventArgs AEv)
        {
            FNumberOfPages = FGfxPrinter.NumberOfPages;
            RefreshPagePosition();
        }

        private void CreateExtract(object ASender, EventArgs AEv)
        {
#if TODO            
            TFrmCreateExtract newExtract = new TFrmCreateExtract();

            newExtract.BestAddress = FMainDS.BestAddress;
            newExtract.IncludeNonValidAddresses = false;
            newExtract.ShowDialog();
#endif
        }

        private void ExportAddresses(object ASender, EventArgs AEv)
        {
            if (FMainDS.LetterRecipient.Rows.Count == 0)
            {
                GenerateLetters(ASender, AEv);
            }

            XmlDocument doc = new XmlDocument();

            XmlNode docNode = doc.CreateElement("NewDonorLetter");
            doc.AppendChild(docNode);

            foreach (NewDonorTDSLetterRecipientRow row in FMainDS.LetterRecipient.Rows)
            {
                if (row.ValidAddress)
                {
                    XmlNode addressNode = doc.CreateElement("address");
                    XmlAttribute att = doc.CreateAttribute("PartnerKey");
                    att.Value = row.PartnerKey.ToString();
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Name");
                    att.Value = row.Name;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Title");
                    att.Value = row.Title;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("FirstName");
                    att.Value = row.Firstname;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Surname");
                    att.Value = row.Surname;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Email");
                    att.Value = row.Email;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Locality");
                    att.Value = row.Locality;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("StreetName");
                    att.Value = row.StreetName;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Building1");
                    att.Value = row.Building1;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Building2");
                    att.Value = row.Building2;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("Address3");
                    att.Value = row.Address3;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("PostalCode");
                    att.Value = row.PostalCode;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("City");
                    att.Value = row.City;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("CountryCode");
                    att.Value = row.CountryCode;
                    addressNode.Attributes.Append(att);
                    att = doc.CreateAttribute("CountryName");
                    att.Value = row.CountryName;
                    addressNode.Attributes.Append(att);
                    docNode.AppendChild(addressNode);
                }
            }

            if (TImportExportDialogs.ExportWithDialog(doc, Catalog.GetString("Export addresses for mail merge"), "xlsx"))
            {
                MessageBox.Show(Catalog.GetString("Address list has been stored"),
                    Catalog.GetString("Success"));

                if (MessageBox.Show(Catalog.GetString("Should we store the contact details?"), Catalog.GetString("Question"),
                                                      MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    AddContactHistory(ASender, AEv);
                }
            }
        }

        private void AddContactHistory(object ASender, EventArgs AEv)
        {
            List <Int64>partnerKeys = new List <long>();

            foreach (NewDonorTDSLetterRecipientRow row in FMainDS.LetterRecipient.Rows)
            {
                if (row.ValidAddress)
                {
                    partnerKeys.Add(row.PartnerKey);
                }
            }

            TMNewDonorSubscriptionsNamespace TRemote = new TMNewDonorSubscriptionsNamespace();
            if (TRemote.Server.WebConnectors.AddPartnerContact(partnerKeys,
                TAppSettingsManager.GetValue("NewDonorSubscriptions.AttributeGroupCode", "BEDANKUNG"),
                TAppSettingsManager.GetValue("NewDonorSubscriptions.AttributeDetailCode", "ERSTSPENDER"),
                DateTime.Now))
            {
                MessageBox.Show(Catalog.GetString("Partner Contact Details have been stored"),
                                Catalog.GetString("Success"));
            }

#if TODO
            // No Mailing code, because this is a form letter
            TRemote.MPartner.Partner.WebConnectors.AddContact(partnerKeys,
                DateTime.Today,
                MPartnerConstants.METHOD_CONTACT_FORMLETTER,
                Catalog.GetString("Letter for new donors announcing subscription to magazine"),
                SharedConstants.PETRAMODULE_FINANCE1,
                "");

            MessageBox.Show(Catalog.GetString("The partner contacts have been updated successfully!"),
                Catalog.GetString("Success"));
#endif
        }
    }
}