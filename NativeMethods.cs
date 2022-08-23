using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Academic_Pro
{
	class NativeMethods
	{
		[DllImport("shell32.dll")]
		static internal extern IntPtr ExtractAssociatedIcon(IntPtr hInst, StringBuilder lpIconPath,
			out ushort lpiIcon);
	}
}
