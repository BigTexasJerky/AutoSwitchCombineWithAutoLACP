namespace System.Runtime.CompilerServices
{
    [System.AttributeUsage(System.AttributeTargets.All, AllowMultiple = false, Inherited = false)]
    internal sealed class NullableAttribute : System.Attribute
    {
        public NullableAttribute(byte value) { }
        public NullableAttribute(byte[] value) { }
    }

    [System.AttributeUsage(System.AttributeTargets.All, AllowMultiple = false, Inherited = false)]
    internal sealed class NullableContextAttribute : System.Attribute
    {
        public NullableContextAttribute(byte value) { }
    }
}