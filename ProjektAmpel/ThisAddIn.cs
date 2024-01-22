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

            var options = new MQTTnet.Client.Options.MqttClientOptionsBuilder()
                .WithTcpServer("localhost", 1883)
                .Build();

            await mqttClient.ConnectAsync(options);

            // Themen abonnieren
            await mqttClient.SubscribeAsync(new MqttTopicFilterBuilder().WithTopic("ampel/farbe").Build());
            mqttClient.UseApplicationMessageReceivedHandler(HandleReceivedMessage);
        }

        private void HandleReceivedMessage(MqttApplicationMessageReceivedEventArgs eventArgs)
        {
            // Callback-Methode, die aufgerufen wird, wenn eine MQTT-Nachricht empfangen wird
            string topic = eventArgs.ApplicationMessage.Topic;
            string payload = Encoding.UTF8.GetString(eventArgs.ApplicationMessage.Payload);

            // Logik, um die Ampelfarbe basierend auf der empfangenen Nachricht zu ändern
            ChangeTrafficLightColor(payload);
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