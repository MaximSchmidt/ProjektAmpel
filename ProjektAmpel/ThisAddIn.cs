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
using MQTTnet.Client.Options;

namespace ProjektAmpel
{
    public partial class ThisAddIn
    {
        private IMqttClient mqttClient;

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Broker-Verbindung konfigurieren
            var factory = new MqttFactory();
            mqttClient = factory.CreateMqttClient();

            var options = new MqttClientOptionsBuilder()
                .WithTcpServer("docker-host-ip", 1883)
                .Build();

            mqttClient.UseApplicationMessageReceivedHandler(HandleReceivedMessage);

            await mqttClient.ConnectAsync(options);

            // Themen abonnieren
            await mqttClient.SubscribeAsync(new TopicFilterBuilder().WithTopic("ampel/status").Build());
        }

        private void HandleReceivedMessage(MqttApplicationMessageReceivedEventArgs eventArgs)
        {
            // Verarbeite die empfangene Nachricht und ändere die Ampelfarbe entsprechend
            string message = Encoding.UTF8.GetString(eventArgs.ApplicationMessage.Payload);
            ChangeTrafficLightColor(message);
        }

        private void ChangeTrafficLightColor(string color)
        {
            // Implementiere die Logik, um die Ampelfarbe im VSTO-Plugin zu ändern
            // Verwende z.B. Visio-Objekte, um die Änderungen vorzunehmen
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