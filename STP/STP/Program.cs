    using System.Reflection.Metadata;
using Word = Microsoft.Office.Interop.Word;
using System.Text;
using System.Security.Claims;
using Microsoft.Office.Interop.Word;
using Rechtsanspruch = STP.Rechtsanspruch;

namespace STP
{
    public class Programm
    {
        public static void Main(String[] args)
        {
            Rechtsanspruch rechtsanspruch = new Rechtsanspruch();
            while (true)
            {
                Console.WriteLine("Wählen Sie ein Option(Geben Sie bitte Nummer):\n1.Neuen Rechtsanspruch erstellen.\n2.Übersicht von Dokumenten nach dem Status.\n3.Ein Anspruch öffnen\n4.Das Status angesehenes Dokumentes ändren.\n5.Programm beenden.");
                int input = int.Parse(Console.ReadLine());
                switch (input)
                {
                    /*
                     * Case, wo wir einen neuen Anspruch erstellen
                     */
                    case 1:
                        Console.WriteLine("Geben Sie den Name des Mandaten, ein Beschreibungstext und das Status des Dokumentes(1.Ungeprueft  2.Genehmigt  3.Abgelehnt)\nGeben Sie jedes Information nach Komma");
                        String[] newClaimInfos = Console.ReadLine().Replace(" ","").Split(",");
                        rechtsanspruch.createClaim(newClaimInfos[0], newClaimInfos[1], rechtsanspruch.getStatus(int.Parse(newClaimInfos[2])));
                        break;

                        /*
                         Übersicht von Ansprüche nach ihrem Status
                         */
                    case 2:
                        Console.WriteLine("Waehlen Sie das Status der Dokumenten, die Sie ansehen wollen(1.Ungeprueft  2.Genehmigt  3.Abgelehnt).");
                        Rechtsanspruch.Status openedFileStatus = rechtsanspruch.getStatus(int.Parse(Console.ReadLine().Replace(" ","")));
                        String[] directoryFiles = rechtsanspruch.docsOverview(openedFileStatus);
                        for (int i = 0;i<directoryFiles.Length;i++)
                        {
                            Console.WriteLine((i+1).ToString() + "." + directoryFiles[i]);
                        }
                        break;

                        /*
                         Einen bestimmten Anspruch öffnen
                         */
                    case 3:
                        Console.WriteLine("Welches File wollen Sie oefnnen?\nGeben Sie zuerst das Status des Anspruches und den Name nach dem Komma(1.Ungeprueft  2.Genehmigt  3.Abgelehnt)");
                        String[] inputs = Console.ReadLine().Split(",");
                        rechtsanspruch.openDoc(rechtsanspruch.getStatus(int.Parse(inputs[0])), inputs[1]);
                        break;
                        
                        /*
                        Das Status von einem Anspruch ändern 
                         */
                    case 4:
                        Console.WriteLine("Geben Sie Daten in solcher Reihenfolge :\nAltes Status des Dokumentes, Neues Status, Der Name des Mandaten.\nStatus-Id: 1.Ungeprueft  2.Genehmigt  3.Abgelehnt");
                        String[] infos = Console.ReadLine().Split(',');
                        rechtsanspruch.changeStatus(rechtsanspruch.getStatus(int.Parse(infos[0])), rechtsanspruch.getStatus(int.Parse(infos[1])), infos[2]);
                        break;

                        /*
                         Das Programm beenden
                         */
                    case 5:
                        Environment.Exit(0);
                        break;
                } 
            } 
        }
    }
}