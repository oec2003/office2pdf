
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeToPdf
{
    interface IOfficeConverter
    {
        bool OnWork(Messages message = null);
    }
}
