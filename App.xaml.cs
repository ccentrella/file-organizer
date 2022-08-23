using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using Microsoft.VisualBasic.ApplicationServices;

namespace Academic_Pro
{
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application
	{

		/// <summary>
		/// Application Entry Point.
		/// </summary>
		[System.STAThreadAttribute()]
		[System.Diagnostics.DebuggerNonUserCodeAttribute()]
		[System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
		public static void Main(string[] args)
		{
			bool mutexIsNew;
			using (var mutex = new Mutex(true, "{CEFF9936-5CB4-43C9-BF63-677921E73147}", out mutexIsNew))
			{
				if (args.Length == 0 && mutexIsNew)
				{
					SplashScreen splashScreen = new SplashScreen("academic%20pro%20logo.png");
					splashScreen.Show(true);
					Academic_Pro.App app = new Academic_Pro.App();
					app.InitializeComponent();
					app.Run();
				}

				else if (args.Length > 0)
					foreach (var arg in args)
					{
						try
						{
							System.Diagnostics.Process.Start(arg);
						}
						catch (IOException)
						{
							MessageBox.Show("The file (\"" + Path.GetFileName(arg) + "\") could not be opened.", "File Error",
								MessageBoxButton.OK, MessageBoxImage.Error);
						}
						catch (System.ComponentModel.Win32Exception)
						{
							MessageBox.Show("The file (\"" + Path.GetFileName(arg) + "\") could not be found.", "File Error",
								MessageBoxButton.OK, MessageBoxImage.Error);
						}
					}
			}
		}


	}
}
