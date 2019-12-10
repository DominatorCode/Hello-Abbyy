using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Hello
{
    static class AssertMe
    {
        public static void assert(bool condition)
        {
            if (!condition)
            {
                StackTrace st = new StackTrace(1, true);
                StackFrame sf = st.GetFrame(0);
                throw new AssertionFailedException(sf);
            }
        }
    };


    class AssertionFailedException : System.Exception
    {
        public StackFrame Frame;

        public AssertionFailedException(StackFrame sf)
        {
            Frame = sf;
        }

        public override string Message
        {
            get
            {
                return "ERROR: Assertion Failed: " +
                    Frame.GetMethod().Name + ", line " + Frame.GetFileLineNumber();
            }
        }

       
    };
}
