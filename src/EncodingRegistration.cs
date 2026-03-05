using System.Text;
using System.Threading;

namespace Nefdev.PptToPptx
{
    internal static class EncodingRegistration
    {
        private static int _registered;

        public static void EnsureCodePages()
        {
            if (Interlocked.Exchange(ref _registered, 1) == 0)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            }
        }
    }
}

