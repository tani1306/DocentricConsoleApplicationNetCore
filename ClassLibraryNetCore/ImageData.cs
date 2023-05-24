using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetCore
{
    public class ImageData
    {
        public byte[] Data { get; set; }
        public bool IsDataAvailable { get; set; } = false;
    }
}
