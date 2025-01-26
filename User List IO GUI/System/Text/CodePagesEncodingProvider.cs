// This system needs to be trained to properly convert various vendors user access lists
// to a single standard format with the same header names for like data.
// The disctionary in this code determines if each header name is in the dictionary, and if it is,
// it will convert the header name to the standard header name. That is brought into a dataTable, and then
// carried through the rest of the process.


namespace System.Text
{
    internal class CodePagesEncodingProvider
    {
        public static EncodingProvider Instance { get; internal set; }
    }
}