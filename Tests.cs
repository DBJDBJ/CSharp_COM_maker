#if !NUNIT
using Microsoft.VisualStudio.TestTools.UnitTesting;
#else
  using NUnit.Framework;
  using TestClass = NUnit.Framework.TestFixtureAttribute;
  using TestMethod = NUnit.Framework.TestAttribute;
  using TestInitialize = NUnit.Framework.SetUpAttribute;
  using TestCleanup = NUnit.Framework.TearDownAttribute;
  using TestContext = System.String;
  using DeploymentItem = NUnit.Framework.DescriptionAttribute;
#endif
namespace dbj.tests
{
    using NUnit.Framework;

    [TestFixture]
    public class dbjComInstancerTest
    {
        [SetUp]
        public void Init()
        {
           // disp_user = new com.dispatch_user("WSCRIPT.SHELL");
        }

        [Test]
        public void wsh_popup()
        {
            using (dbj.com.dispatch_user wsh_shell = new dbj.com.dispatch_user("wscript.shell"))
            {
                object retval = null; 
                
                // wsh_shell.call( "Popup" , "Hello World, for 5 seconds!", 5 ) ;

                // returns WshEnvironment instance
                using (dbj.com.dispatch_user wsh_environment =
                            new dbj.com.dispatch_user(wsh_shell.prop_get("Environment")))
                {
                    int len = (int)wsh_environment.prop_get("length");
                    retval = wsh_shell.call("Popup", "WSH Environment collection has " + len + " items\n", 5);
                }
            }
        }

    }
}