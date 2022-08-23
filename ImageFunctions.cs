using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Academic_Pro
{
	public static class ImageFunctions
	{
		/// <summary>
		/// Get the image for the specified file location.
		/// If the image does not exist, a default one will be provided.
		/// </summary>
		/// <param name="fileLocation">The location of the file</param>
		public static ImageSource GetAppImage(this string fileLocation)
		{
			ushort index;
			var icon = NativeMethods.ExtractAssociatedIcon(IntPtr.Zero, new StringBuilder(fileLocation), out index);
			if (icon != null)
			{
				var options = BitmapSizeOptions.FromWidthAndHeight(256, 256);
				var image = Imaging.CreateBitmapSourceFromHIcon(icon,
					new Int32Rect(), options);
				return image;
			}
			else
			{
				return new BitmapImage(new Uri("Application.png", UriKind.Relative));
			}
		}
	}
}
