using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace UofTranslatorActiveX
{
    class StatusIcons
    {
        private static Bitmap ok = IconResource.OK.ToBitmap();
        private static Bitmap error = IconResource.Error.ToBitmap();
        private static Bitmap stop = IconResource.Stop.ToBitmap();
        private static Bitmap ready = IconResource.Ready.ToBitmap();
        private static Bitmap warning = IconResource.Warning.ToBitmap();

        internal static Bitmap OK
        {
            get
            {
                return ok;
            }
        }

        internal static Bitmap Error
        {
            get
            {
                return error;
            }
        }

        internal static Bitmap Stop
        {
            get
            {
                return stop;
            }
        }

        internal static Bitmap Ready
        {
            get
            {
                return ready;
            }
        }

        internal static Bitmap Warning
        {
            get
            {
                return warning;
            }
        }
    }
}
