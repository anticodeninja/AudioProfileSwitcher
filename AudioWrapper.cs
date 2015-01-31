namespace AudioProfileSwitcher
{
    using System;
    using System.Runtime.InteropServices;

    public class AudioWrapper
    {
        #region # Types #

        public enum ERole
        {
            Console = 0,
            Multimedia = 1,
            Communications = 2,
            Count = 3
        }

        public enum EDataFlow
        {
            Render = 0,
            Capture = 1,
            All = 2,
            Count = 3
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDevice
        {
            int NotImpl1(); int NotImpl2();

            [PreserveSig]
            int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
        }

        [Guid("f8679f50-850a-41cf-9c72-430f290290c8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IPolicyConfig
        {
            int NotImpl1(); int NotImpl2(); int NotImpl3(); int NotImpl4(); int NotImpl5();
            int NotImpl6(); int NotImpl7(); int NotImpl8(); int NotImpl9(); int NotImpl10();

            [PreserveSig]
            int SetDefaultEndpoint(string pszDeviceName, ERole role);
        }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDeviceEnumerator
        {
            int NotImpl1();

            [PreserveSig]
            int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
        }

        [ComImport, Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
        internal class PolicyConfigClient
        {
        }

        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        internal class MMDeviceEnumerator 
        {
        }

        #endregion # Types #

        #region # Fields #

        private readonly IMMDeviceEnumerator _realEnumerator;
        private readonly IPolicyConfig _policyConfig;

        #endregion # Fields #

        #region # Properties #

        public int LoadTimeout { get; set; }

        #endregion # Properties #

        #region # Constructor #

        public AudioWrapper()
        {
            LoadTimeout = 10;
            _realEnumerator = new MMDeviceEnumerator() as IMMDeviceEnumerator;
            _policyConfig = new PolicyConfigClient() as IPolicyConfig;
        }

        #endregion # Constructor #

        #region # Methods #

        public string GetDefaultDevice(EDataFlow dataFlow, ERole role)
        {
            IMMDevice device = null;
            try
            {
                _realEnumerator.GetDefaultAudioEndpoint(dataFlow, role, out device);
                string devId;
                device.GetId(out devId);
                return devId;
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                if (device != null) Marshal.ReleaseComObject(device);
            }
        }

        public void SetDefaultDevice(EDataFlow dataFlow, ERole role, string devId)
        {
            try
            {
                _policyConfig.SetDefaultEndpoint(devId, role);
            }
            catch (Exception)
            {
            }
        }

        #endregion # Methods #
    }
}
