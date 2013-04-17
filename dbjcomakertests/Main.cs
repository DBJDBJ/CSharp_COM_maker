/*
 * DBJ COM Magic
 * (c) 2001 -2013 by Dusan B. Jovanovic
 */
namespace com_instancer
{
	using System;
	using System.Reflection;
    using NUnit.Framework;
	/// <summary>
	/// Various ways to make COM instances in C#
	/// </summary>
	class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
#if ! NUNIT
		static void Main(string[] args)
		{
            f1();
            f2();
			f3();
		}
#endif
        //----------------------------------------------------------------------------------------
#if UNSAFE
        unsafe 
#endif
        public static void f3() 
		{
#if UNSAFE
			Console.WriteLine("The size of ushort is {0}.", sizeof(ushort));
			Console.WriteLine("The size of uint is {0}.", sizeof(uint));
			Console.WriteLine("The size of ulong is {0}.", sizeof(ulong));
#endif
		}
		//----------------------------------------------------------------------------------------
		static void f2 ()
		{
			using ( dbj.com.ICallable wsh_shell = dbj.com.Facade.ObjectUser( "wscript.shell" ) )
			{
				object retval = null ; // wsh_shell.call( "Popup" , "Hello World, for 5 seconds!", 5 ) ;

				// returns WshEnvironment instance
				using ( dbj.com.ICallable wsh_environment = 
							dbj.com.Facade.ObjectUser( wsh_shell.prop_get("Environment")) )
				{
					int len = (int)wsh_environment.prop_get( "length" ) ;
						retval = wsh_shell.call( "Popup" , "WSH Environment collection has " + len + " items\n", 5 ) ;
				}
			}
		}

		//----------------------------------------------------------------------------------------
		static void f1 ()
		{
				object the_object = null ;
			try
			{
				// Use server localhost.
				string theServer="localhost";
				// Use  ProgID HKEY_CLASSES_ROOT\DirControl.DirList.1.
				string the_progid ="WSCRIPT.SHELL"; 
				// Make a call to the method to get the type information for the given ProgID.
				Type the_type =Type.GetTypeFromProgID(the_progid,theServer,true);
				Console.WriteLine("GUID for ProgID {0}, is {1}.", the_progid, the_type.GUID);
				the_object = System.Runtime.InteropServices.Marshal.BindToMoniker("new:" + the_progid) ;
				// Call a method.
				object rezult = the_type.InvokeMember("Popup", 
					BindingFlags.DeclaredOnly | BindingFlags.Public |  
					BindingFlags.Instance | BindingFlags.InvokeMethod, 
					null, 
					the_object, 
					new object [] {"Hello world!"});
			}
			catch(Exception e)
			{
				Console.WriteLine("**************************** An exception occurred");
				Console.WriteLine("Source: {0}" , e.Source);
				Console.WriteLine("Message: {0}" , e.Message);
			}
			finally 
			{
				if ( the_object != null )
				{
					if ( System.Runtime.InteropServices.Marshal.IsComObject( the_object ))
						System.Runtime.InteropServices.Marshal.ReleaseComObject( the_object ) ;
					the_object = null ;
				}
			}
		}
	}
}


