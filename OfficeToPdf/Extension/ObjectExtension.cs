namespace OfficeToPdf
{
    public static class ObjectExtension
    {
        public static bool IsNotNull(this object obj)
        {
            return (obj != null);
        }
        public static bool IsNull(this object obj)
        {
            return (obj == null);
        }
    }
}
