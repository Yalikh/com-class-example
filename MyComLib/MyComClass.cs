using System;
using System.Runtime.InteropServices;

namespace MyComLib
{
    [ComVisible(true)]
    [Guid("07bf25e6-ad4c-4bf7-890c-213c7f5d01cf")]
    [ClassInterface(ClassInterfaceType.None)]
    public class MyComClass
    {
        public string DoSomeWork(string name)
        {
            return "It works, " + name + "! Version: 1";
        }
    }
}
