namespace Academic_Pro
{
	using Microsoft.Win32;
	using System;
	using System.Collections.Generic;
	using System.ComponentModel;
	using System.Diagnostics;
	using System.IO;
	using System.Linq;
	using System.Security;
	using System.Text;
	using System.Threading.Tasks;
	using System.Windows;
	using System.Windows.Controls;
	using System.Windows.Input;
	using System.Windows.Interop;
	using System.Windows.Media.Animation;
	using System.Windows.Media.Imaging;
	using System.Windows.Shell;
	using System.Windows.Threading;

	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}
		object lockObject = new object();
		string defaultPath = Path.Combine(
Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Homework");
		string defaultAppPath = Path.Combine(
Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Autosoft", "Apps", "Academic Pro");
		const string registryLocation = @"Hkey_Current_User\Software\Autosoft\Academic Pro";
		string location;
		string appLocation;
		bool reloadEnabled = false;
		OpenFileDialog dialog1 = new OpenFileDialog()
		{
			AddExtension = false,
			CheckFileExists = true,
			CheckPathExists = true,
			Title = "Select Image - Academic Pro",
			Filter = "Image Files | *.png; *.ico; *.jpg; *.jpeg; *.tiff; *.gif; *.bmp; *.wmf | All Files | *.*"
		};
		System.Windows.Forms.FolderBrowserDialog folderDialog1 = new System.Windows.Forms.FolderBrowserDialog()
		{
			ShowNewFolderButton = true,
			Description = "Select the path to use as the file location"
		};
		private void Close_Click(object sender, RoutedEventArgs e)
		{
			// Close the application
			Application.Current.Shutdown();
		}

		private void browseImage_Click(object sender, RoutedEventArgs e)
		{
			if (dialog1.ShowDialog() == true)
			{
				imageTextBox.Text = dialog1.FileName;
				ImageUpdate(true); // Save the image.
			}
		}

		private void browseFile_Click(object sender, RoutedEventArgs e)
		{
			// Attempt to change the file location, but only if the user click OK
			if (folderDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
				return;

			sourceTextBox.Text = folderDialog1.SelectedPath;
			if (SetRegistryValue("defaultLocation", sourceTextBox.Text) == true)
				location = sourceTextBox.Text;
			else
				sourceTextBox.Text = GetRegistryValue("defaultLocation", false);
		}

		private void ImageUpdate()
		{
			if (!File.Exists(imageTextBox.Text))
			{
				var newBitmapImage = new BitmapImage(new Uri("Generic Avatar (Unisex).png", UriKind.Relative));
				avatar.Source = newBitmapImage;
			}
			else
			{
				try
				{
					Registry.SetValue(registryLocation, "imageLocation", imageTextBox.Text);
					var bitmapImage = new BitmapImage(new Uri(imageTextBox.Text));
					avatar.Source = bitmapImage;
				}
				catch (UnauthorizedAccessException)
				{
					MessageBox.Show("Advanced security settings are enabled on this computer. "
					+ "Please check security settings and try running this program as an administrator.",
					"Security Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (UriFormatException)
				{
					MessageBox.Show("The image may not have been loaded properly.",
						"Image Error", MessageBoxButton.OK, MessageBoxImage.Warning);
				}
			}
		}

		private void ImageUpdate(bool showWarnings)
		{
			if (!File.Exists(imageTextBox.Text))
			{
				var newBitmapImage = new BitmapImage(new Uri("Generic Avatar (Unisex).png", UriKind.Relative));
				avatar.Source = newBitmapImage;
				if (showWarnings && !string.IsNullOrEmpty(imageTextBox.Text))
					MessageBox.Show("An invalid image has been entered. Please enter a valid image.", "Invalid Image",
						MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			else
			{
				try
				{
					SetRegistryValue("imageLocation", imageTextBox.Text);
					var bitmapImage = new BitmapImage(new Uri(imageTextBox.Text, UriKind.Absolute));
					avatar.Source = bitmapImage;
				}
				catch (IOException)
				{
					MessageBox.Show("The image could not be loaded.",
					"Security Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (UriFormatException)
				{
					MessageBox.Show("The image could not be loaded.",
					"Security Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}
		}

		private void imageTextBox_LostFocus(object sender, RoutedEventArgs e)
		{
			ImageUpdate(true); // Save the image.
		}

		private void Load()
		{
			string imageLocation = (string)Registry.GetValue(registryLocation, "imageLocation", "");
			location = GetRegistryValue("defaultLocation", defaultPath, true);
			appLocation = GetRegistryValue("appLocation", defaultAppPath, true);

			// Read registry values
			birthday.Text = GetRegistryValue("birthDate", true);
			startDate.Text = GetRegistryValue("startDate", true);
			endDate.Text = GetRegistryValue("endDate", true);
			firstName.Text = GetRegistryValue("firstName", true);
			lastName.Text = GetRegistryValue("lastName", true);
			motto.Text = GetRegistryValue("motto", "Study Hard", true);

			// Show personal information
			var fullName = firstName.Text + " " + lastName.Text;
			avatar.ToolTip = fullName;
			mottoLabel.Content = motto.Text;
			sourceTextBox.Text = location;

			// Warn the user if the location is empty.
			if (string.IsNullOrWhiteSpace(location))
			{
				MessageBox.Show("The file location is empty or consists only of whitespace.", "Invalid File Location", MessageBoxButton.OK, MessageBoxImage.Warning);
				tabControl1.SelectedIndex = 1;
				sourceTextBox.Focus();
			}

			LoadApps();// Reload all applications
			DeleteEmptyDirectories(); // Delete all empty directories.
			UpdateProgress(); // Update the progress bar

			// Now load the image.
			imageTextBox.Text = imageLocation;
			ImageUpdate();
		}

		/// <summary>
		/// Load all applications
		/// </summary>
		private void LoadApps()
		{
			ParallelQuery<RPApp> apps = null;
			double incrementValue = 0;

			// Only continue if the app folder exists.
			if (!Directory.Exists(appLocation))
				return;

			reloadEnabled = false;
			AppsPane.Children.Clear(); // Clear all items in the App Pane

			#region LoadApps
			try
			{
				apps = from app in Directory.EnumerateFiles(
							  appLocation, "*", SearchOption.AllDirectories).AsParallel().AsOrdered()
					   let appName = Path.GetFileNameWithoutExtension(app)
					   let attributes = File.GetAttributes(app)
					   where !attributes.HasFlag(FileAttributes.Hidden)
					   select new RPApp(app, appName);
			}
			catch (IOException)
			{
				MessageBox.Show("File Error", "The apps could not be loaded.",
					MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (UnauthorizedAccessException)
			{
				MessageBox.Show("File Error", "The apps could not be loaded. Access was denied.",
				MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (SecurityException)
			{
				MessageBox.Show("File Error", "The apps could not be loaded. "
				+"The program does not have the required permission.",
					MessageBoxButton.OK, MessageBoxImage.Warning);
			}

			if (apps != null)
			{
				if (apps.Count() > 0)
					incrementValue = 1d / apps.Count();

				// Add each app
				foreach (var app in apps)
				{
				Dispatcher.BeginInvoke(DispatcherPriority.Background,
							new Action(() =>
							{
								AddApp(app.AppLocation, app.Title);
							}));
				}
			}

			reloadEnabled = true;
			#endregion
		}

		/// <summary>
		/// Add the specified app, allowing the user to specify a custom ToolTip
		/// </summary>
		/// <param name="location">The full location of the App</param>
		/// <param name="title">The App text that will be displayed to the user</param>
		private void AddApp(string location, string title)
		{
			// Create the image object
			Image appImage = new Image() { Source = location.GetAppImage() };


			// Now add the app button
			Button newItem = new Button()
			{
				Tag = location,
				ToolTip = title,
				Content = appImage,
			};
			newItem.Click += app_Click; // Attach the event handler
			AppsPane.Children.Add(newItem); // Add the button to the Apps Pane
		}

		private void app_Click(object sender, RoutedEventArgs e)
		{
			var button = (Button)sender;
			var location = button.Tag.ToString();
			Debug.Assert(location != null);
			try
			{
				Process.Start(location);
			}
			catch (Win32Exception)
			{
				MessageBox.Show("App could not be started", 
					"An error has occurred. The app could not be opened.", 
					MessageBoxButton.OK, MessageBoxImage.Warning);
			}
		}

		// Update the progress bar
		private void UpdateProgress()
		{
			if (endDate.SelectedDate == null | startDate.SelectedDate == null)
				return;

			DateTime minuendDate = (DateTime)endDate.SelectedDate;
			DateTime subtrahendDate = (DateTime)startDate.SelectedDate;
			TimeSpan dDiff = minuendDate - subtrahendDate;
			TimeSpan tdDiff = DateTime.Today - subtrahendDate;
			double totalDays = dDiff.Days;
			double daysCompleted = tdDiff.Days;
			double percentage = daysCompleted / totalDays;
			double roundedPercentage = Math.Round(percentage * 100);
			string progressMessage = "Estimated Progress: ";

			// Ensure there are no negative percentages
			if (roundedPercentage < 0)
				roundedPercentage = 0;

			progressMessage += roundedPercentage + "% Completed"; // Finish creating the progress message
			mainProgressBar.Value = Math.Round(percentage, 2); // Update the progress value

			// Display the progress message
			if (percentage < 1)
				mainProgressBar.ToolTip = progressMessage;
			else
				mainProgressBar.ToolTip = "Completed";
		}

		private void DeleteEmptyDirectories()
		{
			// Only continue if the location is not null or consists only of whitespace.
			if (string.IsNullOrWhiteSpace(location))
				return;

			// Search for every empty directory in parallel.
			try
			{
				var t = new Task(() =>
				{
					Parallel.ForEach(Directory.EnumerateDirectories(location, "*", SearchOption.AllDirectories), (folder) =>
				   {
					   try
					   {
						   var dInfo = new DirectoryInfo(folder);
						   var files = dInfo.GetFiles("*", SearchOption.AllDirectories);

						   // Since the code runs in parallel, ensure it exists first.
						   if (files.Length == 0 && Directory.Exists(folder))
							   Directory.Delete(folder, true);
					   }
					   catch (ArgumentException)
					   {
						   MessageBox.Show("The file location is invalid. Empty directories may not have been deleted.",
							"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
					   }
					   catch (IOException)
					   {
						   MessageBox.Show("An I/O error occurred while searching for empty directories. " +
						   "Empty directories may not have been deleted.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
					   }

					   catch (UnauthorizedAccessException)
					   {
						   MessageBox.Show("The search location is inaccessible. Empty directories may not have been deleted." +
						   "If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
					   }
				   });
				});
				t.Start(); // Start the task
			}
			catch (AggregateException ex)
			{
				foreach (var exception in ex.InnerExceptions)
				{
					if (exception.GetType() == typeof(ArgumentException))
					{
						MessageBox.Show("The file location is invalid. Empty directories may not have been deleted.",
											 "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
					}
					else if (exception.GetType() == typeof(IOException))
					{
						MessageBox.Show("An I/O error occurred while searching for empty directories. " +
							"Empty directories may not have been deleted.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
					}
					else if (exception.GetType() == typeof(UnauthorizedAccessException))
					{
						MessageBox.Show("The search location is inaccessible. Empty directories may not have been deleted." +
						"If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
					}
					else
						throw;
				}
			}
		}
		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			if (Registry.GetValue(registryLocation, "firstRun", null) == null)
			{
				Registry.SetValue(registryLocation, "firstRun", true);
			}

			// We're now loading values and initializing the program.
			Load();

			// Only continue if the string consists of characters other than whitespace
			if (string.IsNullOrWhiteSpace(location))
				return;
	
				try
				{
					Directory.CreateDirectory(location);
					Directory.CreateDirectory(appLocation);
				}
				catch (IOException)
				{
					MessageBox.Show("A required directory could not be created. The program will now terminate.",
						"Critical Error", MessageBoxButton.OK, MessageBoxImage.Error);
					this.Close();
				}
		}

		private void firstNameEnter(object sender, RoutedEventArgs e)
		{
			// Attempt to set the registry value. If it fails, attempt to update the textbox.
			if (!SetRegistryValue("firstName", firstName.Text))
			{
				firstName.Text = GetRegistryValue("firstName", false);
			}
			else
			{
				var fullName = firstName.Text + " " + lastName.Text;
				avatar.ToolTip = fullName;
			}
		}

		/// <summary>
		/// Attempts to set the appropriate value in the registry, returning true if the operation succeeds.
		/// </summary>
		/// <param name="FunctionName">The name of the key to be set</param>
		/// <param name="Value">The new value for the key</param>
		/// <returns>True if the operation succeeded, and false if it failed</returns>
		private bool SetRegistryValue(string FunctionName, string Value)
		{
			// Only continue if a valid function has been entered
			if (FunctionName == null)
				return false;

			try
			{
				Registry.SetValue(registryLocation, FunctionName, Value);
				return true;
			}
			catch (ArgumentException)
			{
				MessageBox.Show("The registry key contains invalid characters or is too long.",
					"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (SecurityException)
			{
				MessageBox.Show("The registry key is inaccessible. The value could not be set.",
									"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (UnauthorizedAccessException)
			{
				MessageBox.Show("The registry key is inaccessible. The value could not be set.",
									"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			return false;
		}

		/// <summary>
		/// Attempts to get the appropriate value from the registry
		/// </summary>
		/// <param name="FunctionName">The name of the key that contains the value to retrieve</param>
		/// <param name="showWarnings">Specifies whether the program should show warnings to the user</param>
		/// <returns>The value of the registry key or null if the operation failed</returns>
		private string GetRegistryValue(string FunctionName, bool showWarnings)
		{
			// Only continue if a valid function has been entered
			if (FunctionName == null)
				return null;

			try
			{
				return Registry.GetValue(registryLocation, FunctionName, "").ToString();
			}
			catch (ArgumentException)
			{
				if (showWarnings)
					MessageBox.Show("The registry key contains invalid characters or is too long.",
					"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (SecurityException)
			{
				if (showWarnings)
					MessageBox.Show("The registry key is inaccessible. The value could not be retrieved.",
										"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (IOException)
			{
				if (showWarnings)
					MessageBox.Show("The registry key has been marked for deletion. The value could not be retrieved.",
										"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}

			return null;
		}

		/// <summary>
		/// Attempts to get the appropriate value from the registry
		/// </summary>
		/// <param name="FunctionName">The name of the key that contains the value to retrieve</param>
		/// <param name="showWarnings">Specifies whether the program should show warnings to the user</param>
		///<param name="defaultValue">Specifies the default value to return if the key doesn't exist</param>
		/// <returns>The value of the registry key or null if the operation failed</returns>
		private string GetRegistryValue(string FunctionName, string defaultValue, bool showWarnings)
		{
			// Only continue if a valid function has been entered
			if (FunctionName == null)
				return null;

			try
			{
				return Registry.GetValue(registryLocation, FunctionName, defaultValue).ToString();
			}
			catch (ArgumentException)
			{
				if (showWarnings)
					MessageBox.Show("The registry key contains invalid characters or is too long.",
					"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (SecurityException)
			{
				if (showWarnings)
					MessageBox.Show("The registry key is inaccessible. The value could not be retrieved.",
										"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (IOException)
			{
				if (showWarnings)
					MessageBox.Show("The registry key has been marked for deletion. The value could not be retrieved.",
										"Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}

			return null;
		}

		private void lastNameEnter(object sender, RoutedEventArgs e)
		{
			// Attempt to set the registry value. If it fails, attempt to update the textbox.
			if (!SetRegistryValue("lastName", lastName.Text))
			{
				lastName.Text = GetRegistryValue("lastName", false);
			}
			else
			{
				var fullName = firstName.Text + " " + lastName.Text;
				avatar.ToolTip = fullName;
			}
		}

		private void birthDateEnter(object sender, RoutedEventArgs e)
		{
			// Attempt to set the registry value. If it fails, attempt to update the textbox.
			if (!SetRegistryValue("birthDate", birthday.Text))
				birthday.Text = GetRegistryValue("birthDate", false);
		}

		private void startDateEnter(object sender, RoutedEventArgs e)
		{
			// Attempt to set the registry value. If it fails, attempt to update the textbox.
			if (!SetRegistryValue("startDate", startDate.Text))
				startDate.Text = GetRegistryValue("startDate", false);
			else
				UpdateProgress();
		}

		private void endDateEnter(object sender, RoutedEventArgs e)
		{
			// Attempt to set the registry value. If it fails, attempt to update the textbox.
			if (!SetRegistryValue("endDate", endDate.Text))
				endDate.Text = GetRegistryValue("endDate", false);
			else
				UpdateProgress();
		}

		private void mottoEnter(object sender, RoutedEventArgs e)
		{
			// Attempt to set the registry value. If it fails, attempt to update the textbox.
			if (!SetRegistryValue("motto", motto.Text))
				motto.Text = GetRegistryValue("motto", false);
			else
				mottoLabel.Content = motto.Text;
		}

		/// <summary>
		/// Opens all items that are selected
		/// </summary>
		private void OpenFiles()
		{
			// Ensure an item is selected.
			if (items.SelectedItem == null)
				return;
			while (items.SelectedItems.Count > 0)
			{
				ListBoxItem item = items.SelectedItems[0] as ListBoxItem;

				// Ensure the item is not null
				if (item == null & items.SelectedItems.Count > 0)
				{
					items.SelectedItems.RemoveAt(0);
					continue;
				}

				var ToolTipEndIndex = item.ToolTip.ToString().IndexOf("\n");
				var ToolTipData = item.ToolTip.ToString().Substring(0, ToolTipEndIndex);
				string currentDocumentURL = ToolTipData;
				try
				{
					System.Diagnostics.Process.Start(currentDocumentURL);
					if (items.SelectedItems.Count > 0)
						items.SelectedItems.RemoveAt(0); // Remove the selection from the item
				}
				catch (IOException)
				{
					MessageBox.Show("The file (\"" + item.Content + "\") could not be opened.", "File Error",
						MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (Win32Exception)
				{
					MessageBox.Show("The file (\"" + item.Content + "\") could not be found.", "File Error",
						MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}
		}

		/// <summary>
		/// Searches asynchronously.
		/// </summary>
		private void Search()
		{
			#region Variables
			string currentCourse = course.Text.ToUpper(); // The selected course
			string currentWeek = week.Text.ToUpper(); // The selected week
			string currentDay = day.Text.ToUpper(); // The selected day
			string searchText = searchBox.Text.ToUpper(); // The text that the user has entered
			var quoteList = new List<string>(); // The list containing all quotes
			#endregion

			SearchPopup.IsOpen = false; // Close the search popup
			items.Items.Clear(); // Clear all items in the tree view.

			// Check if the folder exists, and terminate operation if it doesn't.
			if (!Directory.Exists(location))
				return;
			try
			{
				#region Check Regular Text
				int startIndex = 0;
				string actualText = searchText.Replace("\"", ""); // Get a copy of the search text without quotation marks
				var stringList = new List<string>();

				// Search for all words
				while (startIndex < actualText.Length)
				{
					int endIndex = actualText.IndexOf(" ", startIndex);
					if (endIndex == -1)
					{
						endIndex = actualText.Length;
					}
					int length = endIndex - startIndex;
					if (length > 0)
					{
						string match = actualText.Substring(startIndex, length);
						stringList.Add(match); // Add the item to the match list
					}
					startIndex = endIndex + 1;
				}
				#endregion

				#region Check Quotes
				int quoteLocation = 0;
				while (quoteLocation < searchText.Length & quoteLocation > -1)
				{
					quoteLocation = searchText.IndexOf("\"", quoteLocation) + 1;

					// Only continue if there is another quote
					if (quoteLocation == 0)
						break;

					int newLocation = searchText.IndexOf("\"", quoteLocation); // The location of the second quote

					// If the user forgets to put the last quotation mark, that's okay
					if (newLocation == -1)
						newLocation = searchText.Length;
					int length = newLocation - quoteLocation;


					// Ensure that there is a valid word
					if (length <= 0)
						continue;

					string text = searchText.Substring(quoteLocation, length); // Get the string
					quoteList.Add(text); // Add the quote to the quote list
					quoteLocation = newLocation + 1; // Update the quote location
				}
				#endregion
				foreach (var file in Directory.EnumerateFiles(location, "*", SearchOption.AllDirectories))
				{
					#region Variables
					bool matchesCourse = false; // True if the file name contains the selected course
					bool matchesWeek = false; // True if the file name contains the selected week
					bool matchesDay = false; // True if the file name contains the selected day
					bool matchesText = true; // True if at least 1/2 of the words are found in the file name
					int count = 0; // The total number of words that the user has entered
					int matches = 0; // The total number of words found in the file name
					double percentage = 0; // The percentage of words that are correct
					double roundedPercentage = 0; // The percentage of words that are correct rounded to an integral value.
					string fileName = file.ToUpper(); // Converts the file name to uppercase.
					#endregion

					#region Check Regular Text


					// Enumerate through each word
					stringList.ForEach((x) =>
					{
						if (fileName.Contains(x))
							matches++;
					});
					count = stringList.Count; // Update the word count
					percentage = (double)matches / count * 100; // Calculate the percentage
					roundedPercentage = Math.Round(percentage);
					#endregion

					#region Check Quotes
					quoteList.ForEach((quote) =>
					{
						if (!file.ToUpper().Contains(quote))
							matchesText = false; // Notify the program that this match does not work
					});
					#endregion

					// Check if the selected course has been found
					if (fileName.Contains("\\" + currentCourse + "\\") || currentCourse == "ANY")
						matchesCourse = true;

					// Check if the selected week has been found
					if (fileName.Contains("\\" + currentWeek + "\\") || currentWeek == "ANY")
						matchesWeek = true;

					// Check if the selected week has been found
					if (fileName.Contains("\\" + currentDay + "\\") || currentDay == "ANY")
						matchesDay = true;

					// Check if at least 1/2 of the words have been found
					if (roundedPercentage < 50 && !string.IsNullOrWhiteSpace(searchText))
						matchesText = false;

					// If all the required fields return true, add the file.
					if (matchesCourse & matchesWeek & matchesDay & matchesText)
					{
						var fileCreationDate = File.GetCreationTime(file);
						var fileCompletionDate = File.GetLastWriteTime(file);
						var lastAccessTime = File.GetLastAccessTime(file);
						var newTooltip = file + "\nThis assignment was started: " + fileCreationDate + "\nThis assignment was completed: "
							+ fileCompletionDate + "\nThis assignment was last accessed: " + lastAccessTime;
						var newListItem = new ListBoxItem() { ToolTip = newTooltip, Content = System.IO.Path.GetFileName(file) };
						items.Items.Add(newListItem);
					}
				}
			}
			catch (ArgumentException)
			{
				MessageBox.Show("An invalid search path has been entered. The courses could not be located. " +
								"To correct this problem, change the search path under Account.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (IOException)
			{
				MessageBox.Show("An invalid search path has been entered. The courses could not be located. " +
								"To correct this problem, change the search path under Account.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (SecurityException)
			{
				MessageBox.Show("The search location is inaccessible. The courses could not be located. " +
				"If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (UnauthorizedAccessException)
			{
				MessageBox.Show("The search location is inaccessible. The courses could not be located. " +
			"If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);

			}
		}

		private void CreateFile()
		{
			#region Variables
			string currentCourse = course.Text;
			string currentWeek = week.Text;
			string currentDay = day.Text;
			StringBuilder fileLocation = new StringBuilder(location + "\\");
			string fileLocationStr;
			string parentDirectory;
			#endregion

			// Only continue if the file location consists of characters other than whitespace
			if (string.IsNullOrWhiteSpace(location))
			{
				MessageBox.Show("The file location is empty or consists only of whitespace.", "Invalid File Location", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			SearchPopup.IsOpen = false; // Close the search popup

			// Update the file location
			if (currentCourse != "Any")
				fileLocation.Append(currentCourse + "\\");
			if (currentWeek != "Any")
				fileLocation.Append(currentWeek + "\\");
			if (currentDay != "Any")
				fileLocation.Append(currentDay + "\\");

			// If the search box is not blank, see if an extension must be added.
			if (!string.IsNullOrWhiteSpace(searchBox.Text))
			{
				bool foundExtension = ContainsExtension(searchBox.Text);
				fileLocation.Append(searchBox.Text);

				// Add the .docx extension if none has been found
				if (!foundExtension)
					fileLocation.Append(".docx");

				// Convert the string builder to a string
				fileLocationStr = fileLocation.ToString();
			}

				// If the search box is blank, find a suitable name.
			else
			{
				fileLocationStr = fileLocation.ToString();
				string tempLocation = fileLocationStr + "Assignment.docx";
				bool fileFound = false;

				// Find a match that doesn't already exist.
				if (!File.Exists(tempLocation))
					fileLocationStr = tempLocation;
				else
				{
					for (int i = 2; i <= 1000000; i++)
					{
						tempLocation = fileLocationStr + "Assignment " + i + ".docx";
						if (!File.Exists(tempLocation))
						{
							fileLocationStr = tempLocation;
							fileFound = true;
							break;
						}
					}

					// If all matches already exist, notify the user.
					if (!fileFound)
						MessageBox.Show("No name has been entered and all auto-generated matches already exist. Please enter a valid name.", "Enter Valid Name",
					MessageBoxButton.OK, MessageBoxImage.Warning);
				}
			}

			if (File.Exists(fileLocationStr) && MessageBox.Show("The file already exists. Do you want to overwrite the existing file?",
	"Overwrite File?", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes)
				return;

			try
			{
				// Attempt to create the file and all directories in its path
				parentDirectory = Path.GetDirectoryName(fileLocationStr);
				Directory.CreateDirectory(parentDirectory);
				File.WriteAllText(fileLocationStr, "");

				// Add the file, which is done by searching.
				Search();
			}
			catch (ArgumentException)
			{
				MessageBox.Show("Invalid characters have been entered. Please enter only valid characters. ",
								 "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (NotSupportedException)
			{
				MessageBox.Show("Invalid characters have been entered. Please enter only valid characters. ",
								 "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (IOException)
			{
				MessageBox.Show("An invalid search path has been entered. The courses could not be located. " +
								"To correct this problem, change the search path under Account.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (UnauthorizedAccessException)
			{
				MessageBox.Show("The search location is inaccessible. The courses could not be located. " +
				"If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (SecurityException)
			{
				MessageBox.Show("The search location is inaccessible. The courses could not be located. " +
				"If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
		}

		private void items_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			OpenFiles();
		}

		private void sourceTextBox_LostFocus(object sender, RoutedEventArgs e)
		{
			if (SetRegistryValue("defaultLocation", sourceTextBox.Text))
				location = sourceTextBox.Text;
			else
				sourceTextBox.Text = GetRegistryValue("defaultLocation", false);
		}

		private void searchButtonClick(object sender, RoutedEventArgs e)
		{
			// Open the search popup
			createButton.Visibility = Visibility.Collapsed;
			searchButton.Visibility = Visibility.Visible;
			headerLabel.Content = "Search";
			SearchPopup.IsOpen = true;
		}

		private void searchButton_Click(object sender, RoutedEventArgs e)
		{
			Search(); // Search for assignments
		}

		private void SearchPopup_Opened(object sender, EventArgs e)
		{
			try
			{
				// Only continue if the file location consists of characters other than whitespace
				if (string.IsNullOrWhiteSpace(location))
					return;

				var oldText = course.Text;
				course.Items.Clear(); // Clear all items in the combobox
				course.Items.Add("Any");

				// Reselect the appropriate item
				if (oldText != null)
					course.Text = oldText;
				else
					course.SelectedIndex = 0;

				// Add each course to the combobox asynchronously
					foreach (var directory in Directory.EnumerateDirectories(location))
					{
						var newInfo = new DirectoryInfo(directory);
						course.Items.Add(new ComboBoxItem() { Content = newInfo.Name });
					}
		}
			catch (ArgumentException)
			{
				SearchPopup.IsOpen = false;
				MessageBox.Show("An invalid search path has been entered. The courses could not be located. " +
								"To correct this problem, change the search path under Account.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (IOException)
			{
				SearchPopup.IsOpen = false;
				MessageBox.Show("An invalid search path has been entered. The courses could not be located. " +
								"To correct this problem, change the search path under Account.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (SecurityException)
			{
				MessageBox.Show("The search location is inaccessible. The courses could not be located. " +
				"If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
			catch (UnauthorizedAccessException)
			{
				MessageBox.Show("The search location is inaccessible. The courses could not be located. " +
			"If the problem persists, try running the program as an Administrator.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
			}
		}

		private void items_KeyDown(object sender, KeyEventArgs e)
		{
			// If the delete key was pressed, delete the selected files.
			if (e.Key == Key.Delete)
				DeleteFiles();

			// If the enter key was pressed, open the selected files.
			else if (e.Key == Key.Enter)
				OpenFiles();
		}

		/// <summary>
		/// Deletes all selected files
		/// </summary>
		private void DeleteFiles()
		{
			if (MessageBox.Show("Are you sure you want to delete these files?", "Delete Files?",
			MessageBoxButton.YesNo, MessageBoxImage.Exclamation) == MessageBoxResult.No)
			{
				return;
			}

			// Delete each selected item in the ListBox.
			while (items.SelectedItems.Count > 0)
			{
				ListBoxItem file = (ListBoxItem)items.SelectedItems[0];
				int ToolTipEndIndex = file.ToolTip.ToString().IndexOf("\n");
				string deletionLocation = file.ToolTip.ToString().Substring(0, ToolTipEndIndex);
				bool exceptionOccurred = true;

				try
				{
					File.Delete(deletionLocation);
					exceptionOccurred = false;
					items.Items.Remove(file); // Remove the item from the ListBox.
				}

				catch (ArgumentException)
				{
					MessageBox.Show("The file (\"" + file.Content + "\") could not be deleted.",
						"File Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (SecurityException)
				{
					MessageBox.Show("The file (\"" + file.Content + "\") could not be deleted. "
				+ "The program does not have the appropriate permissions.\n"
				+ "Try running the program as an administrator.",
				"Security Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (UnauthorizedAccessException)
				{
					MessageBox.Show("The file (\"" + file.Content + "\") could not be deleted. "
					+ "Ensure the file is not currently in use.\n"
					+ "Try running the program as an administrator.",
					"Security Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (IOException)
				{
					MessageBox.Show("The file (\"" + file.Content + "\") could not be deleted.", "File Error",
						MessageBoxButton.OK, MessageBoxImage.Error);
				}

				// If an exception occurred, remove the item from the selected items list.
				if (exceptionOccurred)
					items.SelectedItems.RemoveAt(0);
				DeleteEmptyDirectories(); // Delete all empty directories
			}
		}

		private void items_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			// Smoothly slide the file panel in and out
			if (items.SelectedItems.Count > 0)
			{
				DoubleAnimation dAnimation = new DoubleAnimation()
				{
					To = filePanel.ActualHeight,
					Duration = new TimeSpan(0, 0, 0, 0, 500)
				};
				filePanel.BeginAnimation(StackPanel.HeightProperty, dAnimation);
			}
			else
			{
				DoubleAnimation dAnimation = new DoubleAnimation()
				{
					To = 0,
					Duration = new TimeSpan(0, 0, 0, 0, 500)
				};
				filePanel.BeginAnimation(StackPanel.HeightProperty, dAnimation);
			}
		}

		private void deleteButtonClick(object sender, RoutedEventArgs e)
		{
			DeleteFiles(); // Delete all selected items
		}

		private void openButtonClick(object sender, RoutedEventArgs e)
		{
			OpenFiles(); // Open all items that are selected
		}

		private void openParentDirectoryButton_Click(object sender, RoutedEventArgs e)
		{
			ShowInFolder(); // Open the parent directory for each selected item
		}

		/// <summary>
		/// Open the parent directory for each selected item
		/// </summary>
		private void ShowInFolder()
		{
			foreach (ListBoxItem file in items.SelectedItems)
			{
				var ToolTipEndIndex = file.ToolTip.ToString().IndexOf("\n");
				var ToolTipData = file.ToolTip.ToString().Substring(0, ToolTipEndIndex);
				string folderLocation = ToolTipData; // The path of the file
				try
				{
					var parentDirectory = Path.GetDirectoryName(folderLocation);
					Process.Start(parentDirectory);
				}
				catch (ArgumentException)
				{
					MessageBox.Show("The file (\"" + file.Content + "\") path is tool long or is in an invalid format. "
					+ "If the problem continues, immediately seek assistance.", "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (Win32Exception)
				{
					MessageBox.Show("An error occurred (\"" + file.Content + "\") when trying to open the folder.",
				"Warning", MessageBoxButton.OK, MessageBoxImage.Error);
				}
				catch (IOException)
				{
					MessageBox.Show("The folder (\"" + file.Content + "\") could not be opened.", "File Error",
						MessageBoxButton.OK, MessageBoxImage.Error);
				}
			}
		}

		private void addButton_Click(object sender, RoutedEventArgs e)
		{
			// Open the search popup
			searchButton.Visibility = Visibility.Collapsed;
			createButton.Visibility = Visibility.Visible;
			headerLabel.Content = "Create File";
			SearchPopup.IsOpen = true;
		}

		private void createButton_Click(object sender, RoutedEventArgs e)
		{
			CreateFile(); // Create the new file
		}

		private void openFolder_Click(object sender, RoutedEventArgs e)
		{
			// Only continue if the file location consists of characters other than whitespace
			if (string.IsNullOrWhiteSpace(location))
				return;
			try
			{
				System.Diagnostics.Process.Start(location);
			}
			catch (IOException)
			{
				MessageBox.Show("The folder could not be opened.", "File Error",
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
			catch (Win32Exception)
			{
				MessageBox.Show("The folder could not be found.", "File Error",
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void Window_KeyDown(object sender, KeyEventArgs e)
		{
			// If the escape key is pressed, close the search popup.
			if (e.Key == Key.Escape)
				SearchPopup.IsOpen = false;
		}

		private void renameButton_Click(object sender, RoutedEventArgs e)
		{
			bool renamingOccurred = false; // True if one or more files was renamed
			foreach (var item in items.SelectedItems)
			{
				ListBoxItem file = item as ListBoxItem;
				int ToolTipEndIndex = file.ToolTip.ToString().IndexOf("\n");
				string renamingLocation = file.ToolTip.ToString().Substring(0, ToolTipEndIndex);

				// Only continue if the selected item is a ListBoxItem
				if (file == null)
				{
					items.SelectedItems.Remove(item);
					continue;
				}

				// Now rename the file
				if (RenameFile(renamingLocation, items.SelectedItems.IndexOf(file)))
					renamingOccurred = true;
			}

			if (renamingOccurred)
				Search(); // Search after renaming the files
		}

		/// <summary>
		/// Checks if the file name contains an extension
		/// </summary>
		/// <param name="fileName">The name of the file to be examined</param>
		/// <returns>True if the file name contains an extension. Otherwise, false.</returns>
		private bool ContainsExtension(string fileName)
		{
			int startIndex = 0;
			while (startIndex < fileName.Length)
			{
				int periodLoc = fileName.IndexOf(".", startIndex);
				if (periodLoc == -1)
					return false;

				/* If the period is not the last character and the character following the
				 * extension is not a space, then an extension has been found. */
				char periodFChar = fileName[periodLoc + 1];
				if (!char.IsWhiteSpace(periodFChar))
				{
					return true;
				}
				else
				{
					startIndex = periodLoc + 1;
				}
			}
			return false;
		}

		/// <summary>
		/// Prompts the user to rename the given file
		/// </summary>
		/// <param name="file">The file to rename</param>
		/// <param name="index">The index of the item to be renamed</param>
		/// <returns>Whether or not any renaming took place.</returns>
		private bool RenameFile(string file, int index)
		{
			#region Variables
			string parentDirectory = Path.GetDirectoryName(file);
			string oldFileName = Path.GetFileName(file);
			string newFilePath;
			#endregion

			// Ensure the path does not contain invalid names
			foreach (var c in Path.GetInvalidPathChars())
			{
				if (parentDirectory.Contains(c.ToString()))
				{
					MessageBox.Show("Invalid characters are present in the path. Please contact the administrator. "
					+ "The file (\"" + oldFileName + "\") could not be renamed.",
					"File Error", MessageBoxButton.OK, MessageBoxImage.Error);
					return false;
				}
			}

			var rDialog = new RenameDialog(oldFileName);

			// If the user clicks cancel, don't continue
			if (rDialog.ShowDialog() != true)
				return false;

			// Ensure the user does not enter an invalid name
			foreach (var c in Path.GetInvalidFileNameChars())
			{
				if (rDialog.FileName.Text.Contains(c.ToString()))
				{
					MessageBox.Show("Invalid characters are present. The file (\"" + oldFileName + "\") could not be renamed.",
					"File Error", MessageBoxButton.OK, MessageBoxImage.Error);
					return false;
				}
			}

			// Create the new file name
			newFilePath = Path.Combine(parentDirectory, rDialog.FileName.Text);

			// Add the .docx extension if none has been found
			// Check if the  if the filename contains an extension
			if (!ContainsExtension(rDialog.FileName.Text))
				newFilePath = newFilePath + ".docx";

			// Only continue if the new file name differs from the original
			if (newFilePath == file)
				return false;

			try
			{
				// Attempt to rename the file
				File.Copy(file, newFilePath);
				File.Delete(file);
			}
			catch (ArgumentException)
			{
				MessageBox.Show("Invalid characters are present. The file (\"" + oldFileName + "\") could not be renamed.",
					"File Error", MessageBoxButton.OK, MessageBoxImage.Error);
			}
			catch (UnauthorizedAccessException)
			{
				MessageBox.Show("The file could not be renamed. "
				+ "Ensure the file is not currently in use.\n"
				+ "Try running the program as an administrator.",
				"Security Error", MessageBoxButton.OK, MessageBoxImage.Error);
			}
			catch (PathTooLongException)
			{
				MessageBox.Show("The file name (\"" + oldFileName + "\") is too long and could not be renamed.", "File Error",
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
			catch (NotSupportedException)
			{
				MessageBox.Show("The file name (\"" + oldFileName + "\") is in an invalid format and could not be renamed.", "File Error",
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
			catch (IOException)
			{
				MessageBox.Show("The file (\"" + oldFileName + "\") could not be renamed. Ensure the file is not in use.", "File Error",
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
			return true;
		}

		private void appLocation_Click(object sender, RoutedEventArgs e)
		{
			// Only continue if the app location consists of characters other than whitespace
			if (string.IsNullOrWhiteSpace(appLocation))
				return;
			try
			{
				System.Diagnostics.Process.Start(appLocation);
			}
			catch (IOException)
			{
				MessageBox.Show("The folder could not be opened.", "File Error",
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
			catch (Win32Exception)
			{
				MessageBox.Show("The folder could not be found.", "File Error",
					MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
		{
			e.CanExecute = reloadEnabled; // Enabled the reload button only if no applications are currently loading
		}

		private void CommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
		{
			LoadApps(); // Reload all applications
		}

		private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
		{
			// Allow the user to start a new search
			course.SelectedIndex = 0;
			week.SelectedIndex = 0;
			day.SelectedIndex = 0;
			searchBox.Clear();
		}
	}
}
