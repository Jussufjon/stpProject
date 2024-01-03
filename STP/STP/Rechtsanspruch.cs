using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using static STP.Rechtsanspruch;
using System.Diagnostics;

namespace STP
{
    public class Rechtsanspruch
    {
        private const int filesPathIndex = 11;
        private static String tamplatePath = AppDomain.CurrentDomain.BaseDirectory  + "tamplate.docx";
        private static String directoryPath = AppDomain.CurrentDomain.BaseDirectory;
        private static Word.Application appl = new Word.Application();
        private static Word.Document claim = new Word.Document();
        public enum Status
        {
            UNREVIEWED,
            ACCEPTED,
            REJECTED
        }
        /*
         * Funktion, um neues Rechtsanspruch zu erstellen. Name, Beschreibungstext werden in Bookmarks 
         * eingefügt und in einem Ordner, der dem Status des Ansprucher entspricht.
         * */
        public void createClaim(String name, String description, Status newFileStatus)
        {
            claim = appl.Documents.Add(tamplatePath);
            Word.Bookmarks bks = claim.Bookmarks;
            bks["name"].Range.Text = name;
            bks["description"].Range.Text = description;
            String path = directoryPath + newFileStatus + @"\" + name + ".pdf";
            Console.WriteLine(path);
            claim.SaveAs2(path, WdSaveFormat.wdFormatPDF);
        }

        /*
        Diese Funktion zeigt alle Ansprüche nach ihrem Status.
         */
        public String[] docsOverview(Status openedFileStatus)
        {
            String[] directory = Directory.GetFiles(directoryPath + openedFileStatus);
            String[] directoryFiles = new String[directory.Length - 1];

            for (int i=0;i<directory.Length;i++)
            {
                if (directory[i].EndsWith(".pdf"))
                {
                    directoryFiles.Append(directory[i].Split(@"\")[filesPathIndex].Replace(".pdf", ""));
                }
            }
           return directoryFiles;
        }

        /*
         * Öffnet bestimmtes Anspruch
         */
        public void openDoc(Status status, String name)
        {
            String fileName = directoryPath + status + @"\" + name + ".pdf";
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(fileName)
            {
                UseShellExecute = true
            };
            p.Start();
        }

        /*
         *Ändert das Status eines Anspruches 
         * */
        public void changeStatus(Status oldStatus, Status newStatus, String fileName)
        {
            String oldPath = directoryPath + oldStatus + @"\" + fileName + ".pdf";
            String newPath = directoryPath + newStatus + @"\" + fileName + ".pdf";
            File.Move(oldPath, newPath);
        }

        /*
         * Gibt Status nach seinem Index, da das Typ des Statuses enum ist und kann nicht als Integer betrachten werden.
         */
        public Status getStatus(int statusId)
        {
            switch (statusId)
            {
                case 1:
                    return Status.UNREVIEWED;
                    break;
                case 2:
                    return Status.ACCEPTED;
                    break;
                case 3:
                    return Status.REJECTED;
                    break;
                default:
                    return Status.UNREVIEWED;
                    break;
            }
        }
    }
}
