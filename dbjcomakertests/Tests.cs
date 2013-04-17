
namespace dbj.tests
{
    using NUnit.Framework;
    using TestClass = NUnit.Framework.TestFixtureAttribute;
    using TestMethod = NUnit.Framework.TestAttribute;
    using TestInitialize = NUnit.Framework.SetUpAttribute;
    using TestCleanup = NUnit.Framework.TearDownAttribute;
    using TestContext = System.String;
    using DeploymentItem = NUnit.Framework.DescriptionAttribute;

    [TestFixture]
    public class dbjComInstancerTest
    {
        [SetUp]
        public void Init()
        {
           // disp_user = new com.ObjectUser("WSCRIPT.SHELL");
        }

        [Test]
        public void wsh_popup()
        {
            using (dbj.com.ICallable wsh_shell = dbj.com.Facade.ObjectUser("wscript.shell"))
            {
                object retval = null; 
                
                // wsh_shell.call( "Popup" , "Hello World, for 5 seconds!", 5 ) ;

                // returns WshEnvironment instance
                using (dbj.com.ICallable wsh_environment =
                            dbj.com.Facade.ObjectUser(wsh_shell.prop_get("Environment")))
                {
                    int len = (int)wsh_environment.prop_get("length");
                    retval = wsh_shell.call("Popup", "WSH Environment collection has " + len + " items\n", 5);
                }
            }
        }

    }
}