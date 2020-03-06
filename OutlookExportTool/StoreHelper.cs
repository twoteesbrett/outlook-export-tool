using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExportTool
{
    public static class StoreHelper
    {
        public static void Quit(NameSpace ns)
        {
            // Outlook cannot be open to delete files
            ns.Session.Application.Quit();

            var processes = System.Diagnostics.Process.GetProcesses();
            var filtered = processes.Where(p => p.ProcessName.ToUpper().Contains("OUTLOOK"));

            foreach (var process in filtered)
            {
                Console.WriteLine("Waiting for an Outlook process to exit...");
                process.WaitForExit();
                //process.Kill();
            }
        }

        public static Store CreateOrGetStore(NameSpace ns, string fileName)
        {
            var store = GetStore(ns, fileName);

            if (store == null)
            {
                store = AddStore(ns, fileName);
            }

            return store;
        }

        public static Store GetStore(NameSpace ns, string fileName)
        {
            foreach (Store store in ns.Stores)
            {
                if (store.FilePath == fileName)
                {
                    return store;
                }
            }

            return null;
        }

        public static Store AddStore(NameSpace ns, string fileName)
        {
            ns.AddStore(fileName);

            var store = GetStore(ns, fileName);

            return store;
        }

        public static void RemoveStore(NameSpace ns, string fileName)
        {
            var store = GetStore(ns, fileName);

            if (store != null)
            {
                ns.Application.Session.RemoveStore(store.GetRootFolder());
            }
        }

        public static Folder GetFolder(Store store, Folder sourceFolder)
        {
            var rootFolders = store.GetRootFolder().Folders;

            foreach (Folder folder in rootFolders)
            {
                if (folder.Name == sourceFolder.Name)
                {
                    return folder;
                }
            }

            var defaultFolder = sourceFolder.GetOlDefaultFolder();

            // the default folder needs to be changed and olFolderSentMail is not accepted
            if (defaultFolder == OlDefaultFolders.olFolderSentMail)
            {
                defaultFolder = OlDefaultFolders.olFolderInbox;
            }

            Folder targetFolder = (Folder)rootFolders.Add(sourceFolder.Name, defaultFolder);

            return targetFolder;
        }
    }
}
