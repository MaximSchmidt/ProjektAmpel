using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace ProjektAmpel
{
    public partial class ThisAddIn
    {
        private Visio.Document visioDocument;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Pfad zur Visio-Datei angeben
            string docPath = @"C:\Users\maxim.schmidt\Documents\Zeichnung3.vsdm";

            // Visio-Anwendung abrufen
            Visio.Application visioApp = Globals.ThisAddIn.Application;

            // Überprüfen, ob die Anwendung gültig ist
            if (visioApp != null)
            {
                // Versuchen, das Diagramm zu öffnen
                try
                {
                    // Öffnen des Visio-Dokuments
                    visioDocument = visioApp.Documents.Open(docPath);

                    // Optional: MQTT-Verarbeitung hier starten (ersetze dies durch deine eigene Logik)
                    StartMqttProcessing();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Fehler beim Öffnen des Diagramms: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("Visio-Anwendung konnte nicht abgerufen werden.");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Hier kannst du Code hinzufügen, der beim Herunterfahren des Add-Ins ausgeführt wird
        }

        private void StartMqttProcessing()
        {
            // Hier implementiere deine Logik zur Verarbeitung von MQTT-Nachrichten
            // Du kannst beispielsweise eine separate Klasse oder Methode für die MQTT-Verarbeitung erstellen und aufrufen
            // Füge deine MQTT-Logik hier hinzu
        }

        #region Von VSTO generierter Code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}