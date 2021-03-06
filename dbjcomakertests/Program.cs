// ***********************************************************************
// Copyright (c) 2009 Charlie Poole
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
// 
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// ***********************************************************************
#if NUNIT

namespace NUnitLite.Tests
{
using System;
using System.IO;
using NUnitLite.Runner;
    class Program
    {
        // The main program executes the tests. Output may be routed to
        // various locations, depending on the arguments passed.
        //
        // Arguments:
        //
        //  Arguments may be names of assemblies or options prefixed with '/'
        //  or '-'. Normally, no assemblies are passed and the calling
        //  assembly (the one containing this Main) is used. The following
        //  options are accepted:
        //
        //    -test:<testname>  Provides the name of a test to be exected.
        //                      May be repeated. If this option is not used,
        //                      all tests are run.
        //
        //    -out:PATH         Path to a file to which output is written.
        //                      If omitted, Console is used, which means the
        //                      output is lost on a platform with no Console.
        //
        //    -full             Print full report of all tests.
        //
        //    -result:PATH      Path to a file to which the XML test result is written.
        //
        //    -explore[:Path]   If specified, list tests rather than executing them. If a
        //                      path is given, an XML file representing the tests is written
        //                      to that location. If not, output is written to tests.xml.
        //
        //    -noheader,noh     Suppress display of the initial message.
        //
        //    -wait             Wait for a keypress before exiting.
        //
        static void Main(string[] args)
        {
            try
            {
                new TextUI().Execute(args);
            }
            catch (System.Reflection.ReflectionTypeLoadException ex)
            {
                foreach (var item in ex.LoaderExceptions)
                {
                    System.Console.WriteLine(item.Message.ToString());
                }
            }
            catch (Exception x)
            {
                System.Console.WriteLine(x.ToString());
            }
        }
    }
}
#endif