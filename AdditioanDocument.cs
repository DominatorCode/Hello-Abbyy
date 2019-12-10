using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hello
{
    class AdditionalDocument
    {               
        public string nameDocumentNumber { get; set; }
        public int numberDocumentPage { get; set; }

        public AdditionalDocument(string paramDocumentNumber, int paramPageNumber)
        {
            nameDocumentNumber = paramDocumentNumber;
            numberDocumentPage = paramPageNumber;
        }

    }
}
