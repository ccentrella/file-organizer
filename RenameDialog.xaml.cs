using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Academic_Pro
{
	/// <summary>
	/// Interaction logic for RenameDialog.xaml
	/// </summary>
	public partial class RenameDialog : Window
	{
		public RenameDialog()
		{
			InitializeComponent();
		}

		public RenameDialog(string fileName)
		{
			InitializeComponent();

			FileNameLabel.Content = "Please enter a new file name for \"" + fileName + "\".";
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			this.DialogResult = true;
		}

		private void Button_Click_1(object sender, RoutedEventArgs e)
		{
			this.DialogResult = false;
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			FileName.Focus();
		}
	}
}
