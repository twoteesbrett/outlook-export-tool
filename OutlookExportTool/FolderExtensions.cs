using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExportTool
{
    public static class FolderExtensions
    {
        public static bool IsExcluded(this Folder folder, string[] exclusions)
        {
            var type = GetOlDefaultFolder(folder);

            switch (type)
            {
                // skip contacts, journals, notes and tasks
                case OlDefaultFolders.olFolderConflicts:
                case OlDefaultFolders.olFolderContacts:
                case OlDefaultFolders.olFolderDeletedItems:
                case OlDefaultFolders.olFolderDrafts:
                case OlDefaultFolders.olFolderJournal:
                case OlDefaultFolders.olFolderJunk:
                case OlDefaultFolders.olFolderLocalFailures:
                case OlDefaultFolders.olFolderManagedEmail:
                case OlDefaultFolders.olFolderNotes:
                case OlDefaultFolders.olFolderTasks:
                case OlDefaultFolders.olFolderToDo:
                case OlDefaultFolders.olPublicFoldersAllPublicFolders:
                    return true;
                // otherwise, process
                case OlDefaultFolders.olFolderInbox:
                case OlDefaultFolders.olFolderCalendar:
                case OlDefaultFolders.olFolderSentMail:
                    break;
                default:
                    throw new ArgumentException("Unsupported Outlook folder type.");
            }

            if (folder.FolderPath.Contains(exclusions))
            {
                return true;
            }

            return false;
        }

        public static OlDefaultFolders GetOlDefaultFolder(this Folder folder)
        {
            switch (folder.DefaultMessageClass)
            {
                case "IPM.Activity": return OlDefaultFolders.olFolderJournal;
                case "IPM.Appointment": return OlDefaultFolders.olFolderCalendar;
                case "IPM.Contact": return OlDefaultFolders.olFolderContacts;
                case "IPM.Note": return folder.Name.IsSentItems() ? OlDefaultFolders.olFolderSentMail : OlDefaultFolders.olFolderInbox;
                case "IPM.StickyNote": return OlDefaultFolders.olFolderNotes;
                case "IPM.Task": return OlDefaultFolders.olFolderTasks;
                default:
                    throw new ArgumentException("Unsupported Outlook default folder type.");
            }
        }

        // this doesn't work; it is too unreliable and very bad performance
        //private static bool ItemExists(Folder targetFolder, dynamic item)
        //{
        //    foreach (dynamic targetItem in targetFolder.Items)
        //    {
        //        // do not compare size as they may differ even for the same item
        //        if (/*item.Size == targetItem.Size && */item.CreationTime == targetItem.CreationTime && item.ReceivedTime == targetItem.ReceivedTime && item.Subject == targetItem.Subject && targetItem.Body.Length == item.Body.Length)
        //        {
        //            return true;
        //        }
        //    }

        //    return false;
        //}
    }
}
