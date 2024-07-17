using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ConEd.HeavyBid.Importer.Functionality
{
    interface IHandler
    {
        Task Do();
    }
}
