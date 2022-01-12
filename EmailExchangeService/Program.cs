using System;
using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Reflection;
using excel = Microsoft.Office.Interop.Excel;

namespace EmailExchangeService
{
    class Program
    {
        //documentation: https://docs.microsoft.com/pt-br/exchange/client-developer/exchange-web-services/how-to-access-email-as-a-delegate-by-using-ews-in-exchange

        static void Main(string[] args)
        {
        }

        public object [,] LerEmail(string email, string user, string password)
        {
            object[,] dados = null;
            var pathTemp = $"C:\\Users\\{Environment.UserName}\\AppData\\Local\\Temp\\";

            try
            {
                //Instalar no NuGet => Microsoft.Exchange
                //System.Net 
                NetworkCredential networkCredential = GetNetworkCredential(user, password);
                ExchangeService service = GetExchangeService(networkCredential);

                var lstItemId = GetIdItems(service, email);

                foreach (ItemId itemId in lstItemId)
                {
                    var emailItem = EmailMessage.Bind(service, itemId, new PropertySet(ItemSchema.Attachments));

                    foreach (Attachment attachment in emailItem.Attachments)
                    {
                        if(attachment is FileAttachment)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            
                            if(fileAttachment.Name == "")
                            {
                                fileAttachment.Load(pathTemp + fileAttachment.Name);

                                dados = ExtrairExcel(pathTemp + fileAttachment.Name);

                                //emailItem.IsRead = true;
                                //emailItem.Move(new FolderId(""));
                            }
                        }
                    }

                }

                return dados;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private List<ItemId> GetIdItems(ExchangeService service, string mail)
        {
            try
            {
                var idFoderMove = String.Empty;
                FindItemsResults<Item> items = null;
                FolderView folderView = new FolderView(100);
                ItemView view = new ItemView(999);
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                var searchFilters = new List<SearchFilter> { new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                                                             new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true)};
                SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilters);
                ExtendedPropertyDefinition extendedProperty = new ExtendedPropertyDefinition(0x6698, MapiPropertyType.Binary);
                folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, extendedProperty);
                Mailbox mailbox = new Mailbox(mail);
                FolderId folderId = new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox);
                Folder folder = Folder.Bind(service, folderId);
                folder.Load();

                foreach (Folder folder1 in folder.FindFolders(folderView))
                {
                    if (folder1.DisplayName == "pasta01")
                    {
                        Folder pasta01 = Folder.Bind(service, new FolderId(folder1.Id.ToString()));
                        pasta01.Load();

                        foreach (Folder folder2 in pasta01.FindFolders(folderView))
                        {
                            if (folder2.DisplayName == "pasta02")
                            {
                                FolderId idFolder = new FolderId(folder2.Id.ToString());
                                items = service.FindItems(idFolder, searchFilter, view);
                            }

                            if (folder2.DisplayName == "pastaMove")
                                idFoderMove = folder2.Id.ToString();
                        }
                    }
                }

                return items.Select(i => i.Id).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private ExchangeService GetExchangeService(NetworkCredential networkCredential)
        {
            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1, TimeZoneInfo.Local);
                service.UseDefaultCredentials = true;
                service.TraceFlags = TraceFlags.All;
                service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                service.TraceEnabled = true;
                service.Credentials = networkCredential;

                return service;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private NetworkCredential GetNetworkCredential(string user, string password)
        {
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                return new NetworkCredential(user, password);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool CertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            if (sslPolicyErrors == SslPolicyErrors.None)
                return true;

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus !=null)
                {
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        if (certificate.Subject == certificate.Issuer && status.Status == X509ChainStatusFlags.UntrustedRoot)
                            continue;
                        else
                            if (status.Status != X509ChainStatusFlags.NoError)
                            return false;
                    }
                }

                return true;
            }

            return false;
        }

        private object[,] ExtrairExcel(string path)
        {
            object missValue = Missing.Value;
            object[,] result = null;
            excel.Application app = new excel.Application();
            excel.Workbook wb = null;
            excel.Worksheet ws = null;

            try
            {
                wb = app.Workbooks.Open(path);
                ws = (excel.Worksheet)wb.Worksheets.get_Item(1);

                ws.Unprotect("upr");

                long lastRow = ws.get_Range("A1:L15").SpecialCells(excel.XlCellType.xlCellTypeLastCell).Row;
                long lastColumn = ws.get_Range("A1:L15").SpecialCells(excel.XlCellType.xlCellTypeLastCell).Column;

                excel.Range range = ws.Range["A1", ws.Cells[lastRow, lastColumn]];
                result = range.Value;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                wb.Close();
                app.Quit();
            }

            return result;
        }

    }
}
