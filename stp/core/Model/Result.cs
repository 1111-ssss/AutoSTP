using core.Enums;

namespace core.Model
{
    public abstract class ResultBase
    {
        public errorCode ErrorCode { get; internal set; }
        public bool IsCompleted { get; internal set; }
        public string? MessageForUser { get; internal set; }
    }
    public class TResult<T> : ResultBase
    {
        public T? Value { get; private set; }
        public static TResult<T> CompletedOperation(T value)
        {
            return new TResult<T>
            {
                Value = value,
                IsCompleted = true,
                ErrorCode = errorCode.None
            };
        }
        public static TResult<T> FailedOperation(errorCode errorCode, string? messageForUser = null)
        {
            var result = new TResult<T>
            {
                IsCompleted = false,
                ErrorCode = errorCode,
                MessageForUser = messageForUser
            };
            return result;
        }
        public static implicit operator TResult(TResult<T> value) => new TResult
        {
            ErrorCode = value.ErrorCode,
            IsCompleted = value.IsCompleted,
            MessageForUser = value.MessageForUser
        };

    }
    public class TResult : ResultBase
    {
        public static TResult CompletedOperation()
        {
            return new TResult
            {
                IsCompleted = true,
                ErrorCode = errorCode.None
            };
        }
        public static TResult FailedOperation(errorCode errorCode, string? messageForUser = null)
        {
            var result = new TResult
            {
                IsCompleted = false,
                ErrorCode = errorCode,
                MessageForUser = messageForUser
            };
            return result;
        }

    }
}