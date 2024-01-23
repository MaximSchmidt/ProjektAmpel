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


namespace ProjektAmpel
{
    public partial class ThisAddIn
    {
        private IMqttClient mqttClient;
        private YourRibbonClass ribbon;

        private async void ThisAddIn_Startup(object sender, EventArgs e)
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

            ribbon = new YourRibbonClass();
            Globals.Ribbons.YourRibbonClass = ribbon;


        }


        public partial class YourRibbonClass
        {
            private void Red_Click(object sender, RibbonControlEventArgs e)
            {
                // Handle the click event for the Red button
            }

            private void Green_Click(object sender, RibbonControlEventArgs e)
            {
                // Handle the click event for the Green button
            }

            private void Yellow_Click(object sender, RibbonControlEventArgs e)
            {
                // Handle the click event for the Yellow button
            }
        }


        private void HandleReceivedMessage(MqttApplicationMessageReceivedEventArgs eventArgs)
        {
            // Callback-Methode, die aufgerufen wird, wenn eine MQTT-Nachricht empfangen wird
            string payload = Encoding.UTF8.GetString(eventArgs.ApplicationMessage.Payload);

            // Logik, um die Ampelfarbe basierend auf der empfangenen Nachricht zu ändern
            ChangeTrafficLightColor(payload);
        }

        private void ChangeTrafficLightColor(string color)
        {
            try
            {
                Visio.Application visioApp = Globals.ThisAddIn.Application;
                visioApp.Documents.Open("C:\\Users\\maxim.schmidt\\Documents\\Zeichnung3.vsdm");


                // Get the shape you want to modify
                int shapeID = 23;
                Visio.Shape trafficLightShape = visioApp.ActivePage.Shapes.ItemFromID[shapeID];

                // Check if the shape exists
                if (trafficLightShape != null)
                {
                    // Set the fill color based on the received MQTT color
                    switch (color.ToLower())
                    {
                        case "red":
                            trafficLightShape.CellsU["FillForegnd"].FormulaU = "RGB(255,0,0)";
                            break;

                        case "green":
                            trafficLightShape.CellsU["FillForegnd"].FormulaU = "RGB(0,255,0)";
                            break;

                        case "yellow":
                            trafficLightShape.CellsU["FillForegnd"].FormulaU = "RGB(255,255,0)";
                            break;

                        default:
                            // Handle unknown color
                            break;
                    }
                }
                else
                {
                    // Handle the case where the shape is not found
                    MessageBox.Show("Traffic light shape not found!");
                }
            }
            catch (System.Exception ex)
            {
                // Handle exceptions
                MessageBox.Show($"An error occurred: {ex.Message}");
            }

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