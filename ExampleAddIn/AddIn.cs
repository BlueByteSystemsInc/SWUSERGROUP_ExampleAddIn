using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExampleAddIn
{

    [Guid("FB357129-1D3D-4754-B301-CF7BFB3D138A")]
    [ComVisible(true)]
    public class AddIn : ISwAddin
    {
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            // swApp = Application.SldWorks - VBA
            var swApp = ThisSW as SldWorks;

            swApp.SendMsgToUser("Hello from add-in!");

			swApp.SetAddinCallbackInfo(0, this, Cookie);



			return true; 
        }

        public bool DisconnectFromSW()
        {
            return true; 
        }

		[ComRegisterFunctionAttribute]
		public static void RegisterFunction(Type t)
		{


			try
			{
				Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
				Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

				string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
				Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
				addinkey.SetValue(null, 0);

				addinkey.SetValue("Description", "ExampleAddIn");
				addinkey.SetValue("Title", "ExampleAddIn");

				keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
				addinkey = hkcu.CreateSubKey(keyname);
				addinkey.SetValue(null, 1, Microsoft.Win32.RegistryValueKind.DWord);
			}
			catch (System.NullReferenceException nl)
			{
				Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
				System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
			}

			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);

				System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
			}
		}

		[ComUnregisterFunctionAttribute]
		public static void UnregisterFunction(Type t)
		{
			try
			{
				Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
				Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

				string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
				hklm.DeleteSubKey(keyname);

				keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
				hkcu.DeleteSubKey(keyname);
			}
			catch (System.NullReferenceException nl)
			{
				Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
				System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
			}
			catch (System.Exception e)
			{
				Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
				System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
			}
		}


	}
}
