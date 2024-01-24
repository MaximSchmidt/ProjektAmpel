using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using MQTTnet;
using MQTTnet.Client;
using MQTTnet.Server;
using MQTTnet.Extensions;
using MQTTnet.Extensions.ManagedClient;
using MQTTnet.Channel;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools;
using System.Diagnostics;

namespace ProjektAmpel
{
    public partial class ThisAddIn
    {
        private IMqttClient mqttClient;

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
            // Broker-Verbindung konfigurieren
            var factory = new MqttFactory();
            mqttClient = factory.CreateMqttClient();

            var options = new MQTTnet.Client.Options.MqttClientOptionsBuilder()
                .WithTcpServer("localhost", 1883)
                .Build();

            await mqttClient.ConnectAsync(options);

            // Themen abonnieren
            await mqttClient.SubscribeAsync(new MqttTopicFilterBuilder().WithTopic("ampel/farbe").Build());
            mqttClient.UseApplicationMessageReceivedHandler(HandleReceivedMessage);


            Visio.Application visioApp = Globals.ThisAddIn.Application;
            visioApp.Documents.Open("C:\\Users\\maxim.schmidt\\Documents\\Zeichnung3.vsdm");

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
          

        }

        private void HandleReceivedMessage(MqttApplicationMessageReceivedEventArgs e)
        {
            string messagePayload = Encoding.UTF8.GetString(e.ApplicationMessage.Payload);
            Visio.Application visioApp = Globals.ThisAddIn.Application;
            if (visioApp == null)
            {
                // Fehlerbehandlung: Visio-Application-Objekt ist null
                return;
            }

            Visio.Document activeDocument = visioApp.ActiveDocument;
            if (activeDocument == null)
            {
                // Fehlerbehandlung: Kein aktives Dokument
                return;
            }

            // Annahme: "Zeichnung3.vsdm" ist der Seitenname
            Visio.Page page;
            try
            {
                page = activeDocument.Pages["Zeichenblatt-1"];
            }
            catch (Exception ex)
            {
                // Fehlerbehandlung: Seite nicht gefunden
                return;
            }

            Visio.Shape shapeToChange = null;
            try
            {
                shapeToChange = page.Shapes.get_ItemU("23");
            }
            catch (Exception ex)
            {
                // Fehlerbehandlung: Shape nicht gefunden
                return;
            }

            if (shapeToChange != null)
            {
                int shapeID = shapeToChange.ID;

                switch (messagePayload)
                {
                    case "red":
                        shapeToChange.Cells["FillForegnd"].FormulaU = "RGB(255, 0, 0)"; // Red color
                        break;
                    case "green":
                        shapeToChange.Cells["FillForegnd"].FormulaU = "RGB(0, 255, 0)"; // Green color
                        break;
                    case "yellow":
                        shapeToChange.Cells["FillForegnd"].FormulaU = "RGB(255, 255, 0)"; // Yellow color
                        break;
                }
            }
        }



        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        public IMqttClient MqttClient
        {
            get { return mqttClient; }
        }


        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Hier kannst du Code hinzufügen, der beim Herunterfahren des Add-Ins ausgeführt wird
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