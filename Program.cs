namespace AudioProfileSwitcher
{
    using System;
    using System.Threading;
    using System.Xml;
    using NDesk.Options;

    internal class Program
    {
        private static readonly AudioWrapper.EDataFlow[] DataFlowTypesToWork =
        {
            AudioWrapper.EDataFlow.Render,
            AudioWrapper.EDataFlow.Capture,
        };

        private static readonly AudioWrapper.ERole[] RoleTypesToWork =
        {
            AudioWrapper.ERole.Console,
            AudioWrapper.ERole.Multimedia,
            AudioWrapper.ERole.Communications,
        };

        private static void SetWithTimeout(AudioWrapper wrapper, AudioWrapper.EDataFlow flow, AudioWrapper.ERole role,
            string devId, int timeout)
        {
            var endTime = Environment.TickCount + timeout*1000;
            for (;;)
            {
                wrapper.SetDefaultDevice(flow, role, devId);
                if (wrapper.GetDefaultDevice(flow, role) == devId || (timeout != -1 && Environment.TickCount > endTime))
                    return;
                Thread.Sleep(1000);
            }
        }

        public static void LoadSettings(string path, int timeout)
        {
            var wrapper = new AudioWrapper();

            var xml = new XmlDocument();
            xml.Load(path);

            foreach (var flow in DataFlowTypesToWork)
            {
                foreach (var role in RoleTypesToWork)
                {
                    var selectSingleNode = xml.SelectSingleNode(string.Format("/config/{0}-{1}", flow, role));
                    if (selectSingleNode == null) continue;
                    SetWithTimeout(wrapper, flow, role, selectSingleNode.InnerText, timeout);
                }
            }
        }

        public static void SaveSettings(string path, int timeout)
        {
            var wrapper = new AudioWrapper();

            var xml = new XmlDocument();
            xml.LoadXml("<config></config>");
            
            foreach (var flow in DataFlowTypesToWork)
            {
                foreach (var role in RoleTypesToWork)
                {
                    var node = xml.CreateElement(string.Format("{0}-{1}", flow, role));
                    node.InnerText = wrapper.GetDefaultDevice(flow, role);
                    xml.DocumentElement.AppendChild(node);
                }
            }

            xml.Save(path);
        }

        static void Main(string[] args)
        {
            var timeout = 10;
            var filename = "default.xml";
            Action<string, int> action = SaveSettings;

            var p = new OptionSet
            {
                {"t|timeout=", (int value) => timeout = value},
                {"l|load=", path => { filename = path; action = LoadSettings; }},
                {"s|save=", path => { filename = path; action = SaveSettings; }},
            };
            p.Add("?|help", v => p.WriteOptionDescriptions(Console.Out));
            p.Parse(args);

            action(filename, timeout);
        }
    }
}
