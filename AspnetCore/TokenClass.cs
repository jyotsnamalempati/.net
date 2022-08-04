using Microsoft.JSInterop;

namespace AspnetCore
{
    public class TokenClass
    {
        [JSInvokable]
        public string GetHelloMessage() => $"Hello!";
    }
}
