/*
 * DBJ COM Magic
 * (c) 2001 -2013 by Dusan B. Jovanovic
 */
namespace dbj
{
    namespace com
    {
        using System;
        using System.Reflection;
        using System.Runtime.InteropServices;

        internal class comobject : IDisposable
        {
            public readonly string PROGID = null;
            public readonly string SERVER = null;

            private object the_instance_ = null;
            protected Type the_type_ = null;

            //---------------------------------------------------------------------------------------
            private string CLSID_ = null;
            public string CLSID { get { return this.CLSID_; } }
            //---------------------------------------------------------------------------------------
            protected comobject() { /* not allowed */ }
            //---------------------------------------------------------------------------------------
#if DEBUG
            public
#else
			protected 
#endif
 object the_instance()
            {
                if (the_instance_ == null)
                {
                    this.the_type_ = Type.GetTypeFromProgID(PROGID, SERVER, true);
                    the_instance_ = System.Runtime.InteropServices.Marshal.BindToMoniker("new:" + PROGID);
                    this.CLSID_ = this.the_type_.GUID.ToString("D");
                }
#if TESTING
				dll.test( 
					"ca761232ed4211cebacd00aa0057b223", /* valid but non existent */
					"{D7A7D7C3-D47F-11D0-89D3-00A0C90833E6}" /* directanimation.pathcontrol */
				) ;
#endif
                return the_instance_;
            }
            public comobject(string progid_)
            {
                this.PROGID = progid_;
                this.SERVER = "localhost";
                object dummy = this.the_instance();
            }

            public comobject(string progid_, string server_)
            {
                this.PROGID = progid_;
                this.SERVER = server_;
                object dummy = this.the_instance();
            }

            public comobject(object the_com_object)
            {
                this.the_type_ = the_com_object.GetType();
                if (!the_type_.IsCOMObject)
                    throw new ApplicationException(the_type_.FullName + ", is NOT a COM object");
                this.the_instance_ = the_com_object;
                this.PROGID = comobj2progid(this.the_instance_);
                this.SERVER = "localhost";
                this.CLSID_ = this.the_type_.GUID.ToString("D");
                // System.Runtime.InteropServices.Marshal.ReleaseComObject( the_com_object ) ;
            }

            public static string comobj2progid(object comobj)
            {
                return comobj.GetType().Name;
            }

            #region IDisposable Members

            public void Dispose()
            {
                if (the_instance_ != null)
                {
                    if (System.Runtime.InteropServices.Marshal.IsComObject(the_instance_))
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(the_instance_);
                    the_instance_ = null;
                }
            }

            #endregion


            public override string ToString()
            {
                return this.PROGID + (this.SERVER != null ? "@" + this.SERVER : string.Empty);
            }
        }
        /// <summary>
        /// represents an COM object obtained through it's progid
        /// whose methods may be called 
        /// </summary>
        sealed class dispatch_user : comobject
        {

            private dispatch_user() : base() { /* can't make this without an progid */ }

            public dispatch_user(object another_) : base(another_) { }

            public dispatch_user(string progid_) : base(progid_) { }

            public dispatch_user(string progid_, string server_) : base(progid_, server_) { }

            public object call(string method_name_, params object[] args)
            {
                return this.the_type_.InvokeMember(method_name_,
                    BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.IgnoreCase |
                    BindingFlags.Instance | BindingFlags.InvokeMethod,
                    null,
                    this.the_instance(),
                    args);
            }

            public object prop_get(string prop_name_)
            {
                return this.the_type_.InvokeMember(prop_name_,
                    BindingFlags.GetProperty | BindingFlags.GetField | BindingFlags.IgnoreCase,
                    null,
                    this.the_instance(),
                    null);
            }

            public object prop_set(string prop_name_, params object[] args)
            {
                return this.the_type_.InvokeMember(prop_name_,
                    BindingFlags.SetProperty | BindingFlags.SetField | BindingFlags.IgnoreCase,
                    null,
                    this.the_instance(),
                    args);
            }

        }


        internal sealed class dll
        {
            public enum retval { S_OK = 0 };
            /*
            WINOLEAPI ProgIDFromCLSID(REFCLSID clsid,LPOLESTR * lplpszProgID);
            */
            [DllImport("OLE32.DLL", EntryPoint = "ProgIDFromCLSID", SetLastError = true,
            CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern int ProgIDFromCLSID(ref Guid class_id, System.Text.StringBuilder progid);
            // wrapper for the above 
            public static string progid_from_clsid(Guid class_id)
            {
                System.Text.StringBuilder progid_ = new System.Text.StringBuilder(256);
                int retvalue = dll.ProgIDFromCLSID(ref class_id, progid_);
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
			public static void test ( params string [] args )
			{
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

    }
}