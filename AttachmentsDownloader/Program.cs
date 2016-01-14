using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace AttachmentsDownloader
{
    static class Program
    {
        
        static void Main()
        {

#if  DEBUG
            AttachmentsDownloader ad = new AttachmentsDownloader();
            ad.OnDebug();
#else 
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
			{ 
				new AttachmentsDownloader() 
			};
            ServiceBase.Run(ServicesToRun); 
#endif
        }
    }
}
