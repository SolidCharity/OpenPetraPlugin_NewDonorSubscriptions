//
// DO NOT REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
//
// @Authors:
//       timop
//
// Copyright 2004-2015 by OM International
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
using System.Data;
using System.Data.Odbc;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using GNU.Gettext;
using Ict.Common;
using Ict.Common.DB;
using Ict.Common.Data;
using Ict.Common.Printing;
using Ict.Petra.Shared.MPartner.Partner.Data;
using Ict.Petra.Shared.MCommon.Data;
using Ict.Petra.Shared.MFinance.Gift.Data;
using Ict.Petra.Shared.MPartner;
using Ict.Petra.Server.MPartner.Common;
using Ict.Petra.Server.MPartner.Partner.Data.Access;
using Ict.Petra.Server.MCommon.Data.Access;
using Ict.Petra.Server.App.Core.Security;
using Ict.Petra.Plugins.NewDonorSubscriptions.Data;

namespace Ict.Petra.Plugins.NewDonorSubscriptions.Server.WebConnectors
{
    /// <summary>
    /// a letter for a new donor telling him he can get a free subscription
    /// </summary>
    public class TNewDonorSubscriptionsWebConnector
    {
        /// <summary>
        /// return a table with the details of people that have a new subscriptions because they donated
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static NewDonorTDS GetNewDonorSubscriptions(
            string APublicationCode,
            DateTime ASubscriptionStartFrom,
            DateTime ASubscriptionStartUntil,
            string AExtractName,
            bool ADropForeignAddresses)
        {
            NewDonorTDS MainDS = new NewDonorTDS();

            bool NewTransaction;
            TDBTransaction transaction = DBAccess.GDBAccessObj.GetNewOrExistingTransaction(IsolationLevel.ReadUncommitted, out NewTransaction);

            try
            {
                string stmt = TDataBase.ReadSqlFile("Gift.NewDonorSubscription.sql");

                OdbcParameter parameter;

                List <OdbcParameter>parameters = new List <OdbcParameter>();
                parameter = new OdbcParameter("PublicationCode", OdbcType.VarChar);
                parameter.Value = APublicationCode;
                parameters.Add(parameter);
                parameter = new OdbcParameter("StartDate", OdbcType.Date);
                parameter.Value = ASubscriptionStartFrom;
                parameters.Add(parameter);
                parameter = new OdbcParameter("EndDate", OdbcType.Date);
                parameter.Value = ASubscriptionStartUntil;
                parameters.Add(parameter);
                parameter = new OdbcParameter("ExtractName", OdbcType.VarChar);
                parameter.Value = AExtractName;
                parameters.Add(parameter);

                DBAccess.GDBAccessObj.Select(MainDS, stmt, MainDS.AGift.TableName, transaction, parameters.ToArray());

                // drop all previous gifts, keep only the most recent one
                NewDonorTDSAGiftRow previousRow = null;

                Int32 rowCounter = 0;

                while (rowCounter < MainDS.AGift.Rows.Count)
                {
                    NewDonorTDSAGiftRow row = MainDS.AGift[rowCounter];

                    if ((previousRow != null) && (previousRow.DonorKey == row.DonorKey))
                    {
                        MainDS.AGift.Rows.Remove(previousRow);
                    }
                    else
                    {
                        rowCounter++;
                        NewDonorTDSLetterRecipientRow donor = MainDS.LetterRecipient.NewRowTyped();
                        donor.PartnerKey = row.DonorKey;
                        MainDS.LetterRecipient.Rows.Add(donor);
                    }

                    previousRow = row;
                }

                // get recipient description
                foreach (NewDonorTDSAGiftRow row in MainDS.AGift.Rows)
                {
                    if (row.RecipientDescription.Length == 0)
                    {
                        row.RecipientDescription = row.MotivationGroupCode + "/" + row.MotivationDetailCode;
                    }
                }

                foreach (NewDonorTDSLetterRecipientRow row in MainDS.LetterRecipient.Rows)
                {
                    if (!row.IsPartnerKeyNull())
                    {
                        PLocationTable Address;
                        string CountryNameLocal;
                        string emailAddress = TMailing.GetBestEmailAddressWithDetails(row.PartnerKey, out Address, out CountryNameLocal);
                        
                        if (emailAddress.Length > 0)
                        {
                            row.Email = emailAddress;
                        }
    
                        row.ValidAddress = (Address != null);
    
                        if (Address == null)
                        {
                            // no best address; only report if emailAddress is empty as well???
                            continue;
                        }
    
                        if (!Address[0].IsLocalityNull())
                        {
                            row.Locality = Address[0].Locality;
                        }
    
                        if (!Address[0].IsStreetNameNull())
                        {
                            row.StreetName = Address[0].StreetName;
                        }
    
                        if (!Address[0].IsBuilding1Null())
                        {
                            row.Building1 = Address[0].Building1;
                        }
    
                        if (!Address[0].IsBuilding2Null())
                        {
                            row.Building2 = Address[0].Building2;
                        }
    
                        if (!Address[0].IsAddress3Null())
                        {
                            row.Address3 = Address[0].Address3;
                        }
    
                        if (!Address[0].IsCountryCodeNull())
                        {
                            row.CountryCode = Address[0].CountryCode;
                        }
    
                        row.CountryName = CountryNameLocal;
    
                        if (!Address[0].IsPostalCodeNull())
                        {
                            row.PostalCode = Address[0].PostalCode;
                        }
    
                        if (!Address[0].IsCityNull())
                        {
                            row.City = Address[0].City;
                        }
                    }
                }

                // remove all invalid addresses
                rowCounter = 0;

                while (rowCounter < MainDS.LetterRecipient.Rows.Count)
                {
                    NewDonorTDSLetterRecipientRow row = MainDS.LetterRecipient[rowCounter];

                    if (!row.ValidAddress)
                    {
                        MainDS.LetterRecipient.Rows.Remove(row);
                    }
                    else
                    {
                        rowCounter++;
                    }
                }
            }
            finally
            {
                if (NewTransaction)
                {
                    DBAccess.GDBAccessObj.RollbackTransaction();
                }
            }

            return MainDS;
        }

        private static string GetStringOrEmpty(object obj)
        {
            if (obj == System.DBNull.Value)
            {
                return "";
            }

            return obj.ToString();
        }

        /// <summary>
        /// prepare HTML text for each new donor
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static StringCollection PrepareNewDonorLetters(ref NewDonorTDS AMainDS, string AHTMLTemplate)
        {
            TDBTransaction transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadUncommitted);

            // get the local country code
            string LocalCountryCode = TAddressTools.GetCountryCodeFromSiteLedger(transaction);

            DBAccess.GDBAccessObj.RollbackTransaction();

            StringCollection Letters = new StringCollection();

            foreach (NewDonorTDSLetterRecipientRow addressRow in AMainDS.LetterRecipient.Rows)
            {
                if (addressRow.ValidAddress == false)
                {
                    continue;
                }

                AMainDS.AGift.DefaultView.RowFilter = NewDonorTDSAGiftTable.GetDonorKeyDBName() + " = '" + addressRow.PartnerKey.ToString() + "'";
                NewDonorTDSAGiftRow row = (NewDonorTDSAGiftRow)AMainDS.AGift.DefaultView[0].Row;

                string donorName = row.DonorShortName;

                string msg = AHTMLTemplate;

                msg =
                    msg.Replace("#RECIPIENTNAME",
                        Calculations.FormatShortName(row.RecipientDescription, eShortNameFormat.eReverseWithoutTitle));
                msg =
                    msg.Replace("#RECIPIENTFIRSTNAME",
                        Calculations.FormatShortName(row.RecipientDescription, eShortNameFormat.eOnlyFirstname));
                msg = msg.Replace("#TITLE", Calculations.FormatShortName(donorName, eShortNameFormat.eOnlyTitle));
                msg = msg.Replace("#NAME", Calculations.FormatShortName(donorName, eShortNameFormat.eReverseWithoutTitle));
                msg = msg.Replace("#FORMALGREETING", Calculations.FormalGreeting(donorName));
                msg = msg.Replace("#STREETNAME", GetStringOrEmpty(addressRow.StreetName));
                msg = msg.Replace("#LOCATION", GetStringOrEmpty(addressRow.Locality));
                msg = msg.Replace("#ADDRESS3", GetStringOrEmpty(addressRow.Address3));
                msg = msg.Replace("#BUILDING1", GetStringOrEmpty(addressRow.Building1));
                msg = msg.Replace("#BUILDING2", GetStringOrEmpty(addressRow.Building2));
                msg = msg.Replace("#CITY", GetStringOrEmpty(addressRow.City));
                msg = msg.Replace("#POSTALCODE", GetStringOrEmpty(addressRow.PostalCode));
                msg = msg.Replace("#DATE", DateTime.Now.ToString("d. MMMM yyyy"));

                // according to German Post, there is no country code in front of the post code
                // if country code is same for the address of the recipient and this office, then COUNTRYNAME is cleared
                if (GetStringOrEmpty(addressRow.CountryCode) != LocalCountryCode)
                {
                    msg = msg.Replace("#COUNTRYNAME", GetStringOrEmpty(addressRow.CountryCode));
                }
                else
                {
                    msg = msg.Replace("#COUNTRYNAME", "");
                }

                // TODO: projects have names as well. different way to determine project gifts? motivation detail?
                if ((row.MotivationGroupCode.ToUpper() == "GIFT") && (row.MotivationDetailCode.ToUpper() == "SUPPORT"))
                {
                    msg = TPrinterHtml.RemoveDivWithClass(msg, "donationForProject");
                }
                else
                {
                    msg = TPrinterHtml.RemoveDivWithClass(msg, "donationForWorker");
                }

                Letters.Add(msg);
            }

            return Letters;
        }
    }
}