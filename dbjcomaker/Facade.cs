
namespace dbj.com
{
    using System;

    public interface ICallable : IDisposable
    {
        object call(string method_name_, params object[] args);
        object prop_get(string prop_name_);
        object prop_set(string prop_name_, params object[] args);
    }

    public class Facade
    {
        public static ICallable ObjectUser(string prog_id, params string[] server )
        {
            if (server.Length > 0)
                return new ObjectUser(prog_id, server[0]);
            return new ObjectUser(prog_id);
        }

        public static ICallable ObjectUser(object another_)
        {
            return new ObjectUser(another_);
        }
    }
}
