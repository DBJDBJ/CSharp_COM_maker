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

        internal class DisposableObject : IDisposable
        {
            public readonly string PROGID = null;
            public readonly string SERVER = null;

            private object the_instance_ = null;
            protected Type the_type_ = null;

            //---------------------------------------------------------------------------------------
            private string CLSID_ = null;
            public string CLSID { get { return this.CLSID_; } }
            //---------------------------------------------------------------------------------------
            protected DisposableObject() { /* not allowed */ }
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

                return the_instance_;
            }
            public DisposableObject(string progid_)
            {
                this.PROGID = progid_;
                this.SERVER = "localhost";
                object dummy = this.the_instance();
            }

            public DisposableObject(string progid_, string server_)
            {
                this.PROGID = progid_;
                this.SERVER = server_;
                object dummy = this.the_instance();
            }

            public DisposableObject(object the_com_object)
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
        /// therefore this works only for COM objects implementing 
        /// IDispatch COM interface
        /// </summary>
        sealed class ObjectUser : DisposableObject, dbj.com.ICallable
        {

            private ObjectUser() : base() { /* can't make this without an progid */ }

            public ObjectUser(object another_) : base(another_) { }

            public ObjectUser(string progid_) : base(progid_) { }

            public ObjectUser(string progid_, string server_) : base(progid_, server_) { }

            object ICallable.call(string method_name_, params object[] args)
            {
                return this.the_type_.InvokeMember(method_name_,
                    BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.IgnoreCase |
                    BindingFlags.Instance | BindingFlags.InvokeMethod,
                    null,
                    this.the_instance(),
                    args);
            }

            object ICallable.prop_get(string prop_name_)
            {
                return this.the_type_.InvokeMember(prop_name_,
                    BindingFlags.GetProperty | BindingFlags.GetField | BindingFlags.IgnoreCase,
                    null,
                    this.the_instance(),
                    null);
            }

            object ICallable.prop_set(string prop_name_, params object[] args)
            {
                return this.the_type_.InvokeMember(prop_name_,
                    BindingFlags.SetProperty | BindingFlags.SetField | BindingFlags.IgnoreCase,
                    null,
                    this.the_instance(),
                    args);
            }
        }
    }
}