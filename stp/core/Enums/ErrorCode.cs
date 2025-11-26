namespace core.Enums
{
    public enum errorCode
    {
        None = 0,
        
    }
    public static class ErrorCodeExtensions
    {
        public static bool IsSystemError(this errorCode code) => (int)code >=200;
    }
}