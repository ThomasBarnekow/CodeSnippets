using System;
using System.IO;
using Xunit;

namespace CodeSnippets.Windows.Tests
{
    public class OtherTests
    {
        private void MyMethod(IDisposable db, int myInt)
        {
            try
            {
                using (db)
                {
                    int bar = OtherMethod(myInt);
                }
            }
            catch (Exception ex)
            {
                // Do something with the exception.
                throw;
            }
        }

        private int OtherMethod(int myInt)
        {
            return myInt;
        }

        [Fact]
        public void Test1()
        {
            MyMethod(new MemoryStream(), 10);
        }
    }
}
