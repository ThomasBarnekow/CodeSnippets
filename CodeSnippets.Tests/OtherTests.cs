using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace CodeSnippets.Tests
{
    public class OtherTests
    {
        [Fact]
        public void Test1()
        {
            MyMethod(new MemoryStream(), 10);
        }

        private void MyMethod(IDisposable db, int myInt)
        {
            try
            {
                using (db)
                {
                    var bar = OtherMethod(myInt);
                }
            }
            catch
            {
                // Ignore
            }
        }

        private int OtherMethod(int myInt)
        {
            return myInt;
        }
    }
}
