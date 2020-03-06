// for a dry run that won't extract to files
//#define SIMULATE

using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExportTool
{
    class Program
    {
        private const string FileNameFormat = "Brett's Archive {0}.pst";
        private const string OutputDirectory = @"C:\Users\Brett\Documents\Outlook Files";
        private static readonly string[] FolderExclusions = new[] { "Deleted Items" };
        private static readonly DateTime LowerCutoffDate = new DateTime(2020, 1, 1);
        private static readonly DateTime UpperCutoffDate = new DateTime(2020, 4, 1);

        static void Main(string[] args)
        {
            if (!Directory.Exists(OutputDirectory))
            {
                throw new DirectoryNotFoundException();
            }

#if !SIMULATE
            Cleanup(OutputDirectory);
#endif
            Extract(OutputDirectory);

            //Console.WriteLine("Ready.");
            //Console.ReadKey();
        }

        private static void Cleanup(string directory)
        {
            var application = new Application();
            var ns = application.GetNamespace("MAPI");
            var files = Directory.GetFiles(directory);

            foreach (var file in files)
            {
                StoreHelper.RemoveStore(ns, file);
            }

            StoreHelper.Quit(ns);

            foreach (var file in files)
            {
                File.Delete(file);
            }
        }

        private static void Extract(string directory)
        {
            var application = new Application();
            var accounts = application.Session.Accounts;

            foreach (Account account in accounts)
            {
                account.DisplayName.WriteLineIndented(0);

                var folder = (Folder)account.DeliveryStore.GetRootFolder();
                EnumerateFolders(folder, 1);
            }
        }

        private static void EnumerateFolders(Folder folder, int level)
        {
            var childFolders = folder.Folders;

            foreach (Folder childFolder in childFolders)
            {
                $"{childFolder.FolderPath} [{childFolder.DefaultItemType}] ({childFolder.Items.Count})".WriteLineIndented(level);

                if (childFolder.IsExcluded(FolderExclusions))
                {
                    continue;
                }

                EnumerateItems(childFolder, level);

                // call EnumerateFolders using childFolder, to see if there are any subfolders within this one
                EnumerateFolders(childFolder, level + 1);
            }
        }

        private static void EnumerateItems(Folder folder, int level)
        {
            foreach (dynamic item in folder.Items)
            {
                dynamic copy = null;

                try
                {
                    DateTime date;

                    switch ((OlObjectClass)item.Class)
                    {
                        case OlObjectClass.olMail:
                        case OlObjectClass.olMeetingCancellation:
                        case OlObjectClass.olMeetingForwardNotification:
                        case OlObjectClass.olMeetingRequest:
                        case OlObjectClass.olMeetingResponseNegative:
                        case OlObjectClass.olMeetingResponsePositive:
                        case OlObjectClass.olMeetingResponseTentative:
                        case OlObjectClass.olReport:
                            date = GetMailItemDate(folder, item);
                            break;
                        case OlObjectClass.olAppointment:
                            date = item.Start;
                            break;
                        default:
                            throw new ArgumentException("Unsupported Outlook item type.");
                    }

                    // ignore items outside the range
                    if (date < LowerCutoffDate || date > UpperCutoffDate)
                    {
                        $"SKIP  {date} > {item.Subject}".WriteLineIndented(level + 1);
                        continue;
                    }

                    $"ADD  {date} > {item.Subject}".WriteIndented(level + 1);

#if !SIMULATE
                    var targetFolder = GetTargetFolder(folder, date);

                    copy = item.Copy();
                    copy.Move(targetFolder);
#endif
                }
                catch (System.Exception ex)
                {
                    if (copy != null)
                    {
                        copy.Remove();
                    }

                    Console.WriteLine("An item could not be extracted.");
                }
            }
        }

        private static DateTime GetMailItemDate(Folder folder, dynamic item) =>
            folder.Name.IsSentItems()
                ? item.SentOn
                : item.ReceivedTime;

        private static Folder GetTargetFolder(Folder sourceFolder, DateTime date)
        {
            var year = date.Year;
            var fileName = string.Format(FileNameFormat, year);
            var path = Path.Combine(OutputDirectory, fileName);
            var application = sourceFolder.Application;
            var ns = application.GetNamespace("MAPI");

            var store = StoreHelper.CreateOrGetStore(ns, path);
            Folder targetFolder = StoreHelper.GetFolder(store, sourceFolder);

            return targetFolder;
        }
    }
}
