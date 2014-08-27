using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuildNotifyService
{
    class Config
    {
        public string FromMail { get; set; }
        public string ToMail { get; set; }
        public string SuccessImgPath { get; set; }
        public string FailedImgPath { get; set; }
        public string PartiallyImgPath { get; set; }
        public string DefaultImgPath { get; set; }
        public string StoppedImgPath { get; set; }
        public string OutputPath { get; set; }
        public string TeamProject { get; set; }
        public string Url { get; set; }

    }
}
