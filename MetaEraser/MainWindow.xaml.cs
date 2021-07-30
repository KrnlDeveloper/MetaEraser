using GroupDocs.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MetaEraser
{
	public partial class MainWindow : Window
	{
		private Boolean overwriteExistingFile = false;

		public MainWindow()
		{
			InitializeComponent();

			Uri iconUri = new Uri("pack://application:,,,/MetaEraser.ico", UriKind.RelativeOrAbsolute);
			this.Icon = BitmapFrame.Create(iconUri);

			this.AllowDrop = true;
			
			this.saveRadioBtn.IsChecked = true;

			this.CreateElementHandlers();

			iconUri = new Uri("pack://application:,,,/MetaEraser.png", UriKind.RelativeOrAbsolute);
			this.formImg.Source = BitmapFrame.Create(iconUri);
		}

		private void CreateElementHandlers()
		{
			this.Drop += MainWindow_Drop;

			this.overwriteRadioBtn.Checked += OverwriteRadioBtn_Checked;
			this.saveRadioBtn.Checked += SaveRadioBtn_Checked;

			this.closeButton.Click += CloseButton_Click;
			this.minButton.Click += MinButton_Click;

			this.MouseLeftButtonDown += MainWindow_MouseLeftButtonDown;
		}

		private void MainWindow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			this.DragMove();
		}

		private void MinButton_Click(object sender, RoutedEventArgs e)
		{
			this.WindowState = WindowState.Minimized;
		}

		private void CloseButton_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		private void SaveRadioBtn_Checked(object sender, RoutedEventArgs e)
		{
			this.overwriteRadioBtn.IsChecked = false;
			this.overwriteExistingFile = false;
		}

		private void OverwriteRadioBtn_Checked(object sender, RoutedEventArgs e)
		{
			this.saveRadioBtn.IsChecked = false;
			this.overwriteExistingFile = true;
		}

		private void MainWindow_Drop(object sender, DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
			{
				String files = String.Empty;

				String[] droppedFilePaths = (String [])e.Data.GetData(DataFormats.FileDrop, true);
				foreach (var name in droppedFilePaths)
				{
					files += (name + Environment.NewLine);
				}

				Int32 numOfFiles = droppedFilePaths.Count();
				String insertStr = String.Empty;

				if(numOfFiles == 1)
				{
					insertStr = "file ";
				}

				else
				{
					insertStr = numOfFiles.ToString() + " files";
				}

				String message = "Metadata from the following " + insertStr + " will be erased " + Environment.NewLine + files;
				MessageBoxResult dlgRes = MessageBox.Show(message, "Delete metadata?", MessageBoxButton.OKCancel, MessageBoxImage.Question);

				if(dlgRes == MessageBoxResult.OK)
				{
					foreach(var file in droppedFilePaths)
					{
						try
						{
							Int32 pointPos = file.LastIndexOf('.');
							Int32 slashPos = file.LastIndexOf('\\');

							String newFileName = file.Substring(slashPos + 1, pointPos - slashPos - 1) + "_cln";
							MsOfficeMetaDataEraser eraser = new MsOfficeMetaDataEraser(file, newFileName, this.overwriteExistingFile);
							eraser.Sanitize();
						}
						
						catch(NotSupportedException exp)
						{
							MessageBox.Show(exp.Message, "File type not supported", MessageBoxButton.OK, MessageBoxImage.Error);
						}

						catch(Exception exp)
						{
							MessageBox.Show(exp.Message, "An error occured", MessageBoxButton.OK, MessageBoxImage.Error);
						}
					}
				}
			}		
		}
	}
}
