using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetCore
{
    public class ImageData
    {
        public IReadOnlyCollection<byte> Data { get; set; } = Array.Empty<byte>();
        public bool IsDataAvailable { get; set; } = false;
        public byte[] RiskMatrixPic { get; set; }
        public byte[] RiskMatrixPicSpanish { get; set; }    
        public byte[] RiskMatrixPicPortuguese { get; set; }
    }
}
