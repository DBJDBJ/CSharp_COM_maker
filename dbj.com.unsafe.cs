/*
 * DBJ COM Magic
 * (c) 2001 -2013 by Dusan B. Jovanovic
 */

//#define TESTING

namespace dbj 
{
	namespace com 
	{
		using System ;
		using System.Reflection ;
		using System.Runtime.InteropServices ;

		/// <summary>
		/// The unsafe version of the dbj.com.dll
		/// </summary>
		internal sealed unsafe class dllUnsafe
		{
			public enum retval { S_OK	= 0  };

			/*
						typedef struct _GUID 
						{ 
						unsigned long Data1;  
						unsigned short Data2;  
						unsigned short Data3;  
						unsigned char Data4[8];
						} GUID, UUID;
					*/
			[StructLayout(LayoutKind.Explicit)]
			unsafe struct _GUID 
			{ 
				[FieldOffset(0)]public ulong	Data1;  
				[FieldOffset(0)]public ushort	Data2;  
				[FieldOffset(0)]public ushort	Data3;  
				[FieldOffset(0)]public char []	Data4 ; // len = 8
			} ;

			/*
					WINOLEAPI ProgIDFromCLSID(REFCLSID clsid,LPOLESTR * lplpszProgID);
					*/
			[DllImport("OLE32.DLL", EntryPoint="ProgIDFromCLSID",  SetLastError=true,
				 CharSet=CharSet.Unicode, ExactSpelling=true, CallingConvention=CallingConvention.StdCall)]
			static unsafe extern int ProgIDFromCLSID( ref dllUnsafe._GUID class_id, char * progid);
			// wrapper for the above 
			public static unsafe string progid_from_clsid ( Guid class_id )
			{
				char * progid_ = stackalloc char [256] ;
				_GUID _guid = guid2guid( class_id );
				int retvalue = dllUnsafe.ProgIDFromCLSID(  ref _guid , progid_ );

				switch( retvalue )
				{
					case (int)retval.S_OK :		return new string(progid_,0,39) ;
					default :
						// int err_code = Marshal.GetLastWin32Error() ;
						System.ComponentModel.Win32Exception myEx = 
							new System.ComponentModel.Win32Exception(retvalue);
						System.Diagnostics.Debug.WriteLine("ERROR Message:\t" + myEx.Message);
						System.Diagnostics.Debug.WriteLine("ERROR Code:\t" + myEx.ErrorCode);
						System.Diagnostics.Debug.WriteLine("ERROR Source:\t" + myEx.Source);
						break ;
				}
				return string.Empty ;
			}
			/*
             * TBD -- obviously 
             */
			static unsafe dllUnsafe._GUID guid2guid ( Guid managed_guid )
			{
					byte [] bits = managed_guid.ToByteArray() ;

					dllUnsafe._GUID retval = new _GUID()  ;
 
					return retval ;
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
							dllUnsafe.progid_from_clsid( new Guid(clsid) )
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