/*
 * DBJ COM Magic
 * (c) 2001 -2013 by Dusan B. Jovanovic
 */
namespace dbj
{
    namespace com
    {

        using System;
        using System.Runtime.InteropServices;
        internal sealed class OLE32
        {
            public enum retval { S_OK = 0 };
            /*
            WINOLEAPI ProgIDFromCLSID(REFCLSID clsid,LPOLESTR * lplpszProgID);
            */
            [DllImport("OLE32.DLL", EntryPoint = "ProgIDFromCLSID", SetLastError = true,
            CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern int ProgIDFromCLSID(ref Guid class_id, System.Text.StringBuilder progid);
            /*
             * wrapper for the above 
             */
            public static string ProgIDFromCLSID(Guid class_id)
            {
                System.Text.StringBuilder progid_ = new System.Text.StringBuilder(256);
                int retvalue = OLE32.ProgIDFromCLSID(ref class_id, progid_);
                switch (retvalue)
                {
                    case (int)retval.S_OK: return progid_.ToString();
                    default:
                        // int err_code = Marshal.GetLastWin32Error() ;
                        System.ComponentModel.Win32Exception myEx =
                            new System.ComponentModel.Win32Exception(retvalue);
                        System.Diagnostics.Debug.WriteLine("ERROR Message:\t" + myEx.Message);
                        System.Diagnostics.Debug.WriteLine("ERROR Code:\t" + myEx.ErrorCode);
                        System.Diagnostics.Debug.WriteLine("ERROR Source:\t" + myEx.Source);
                        break;
                }
                return string.Empty;
            }
#if TESTING
			public static void test (  )
			{
            string [] args = {
					"ca761232ed4211cebacd00aa0057b223", /* valid but non existent */
					"{D7A7D7C3-D47F-11D0-89D3-00A0C90833E6}" /* directanimation.pathcontrol */
            };
				foreach( string clsid in args )
				{
						System.Diagnostics.Debug.WriteLine(	"For CLSID { " + clsid + "}" ) ;
					try 
					{
						System.Diagnostics.Debug.WriteLine("PROGID = " + 
							dll.progid_from_clsid( new Guid(clsid) )
							) ;
					} 
					catch (Exception x) 
					{
						System.Diagnostics.Debug.WriteLine("dll.test() Exception") ;
						System.Diagnostics.Debug.WriteLine(x) ;
					}
				}
			}
#endif
        }
    } // com
} // dbj