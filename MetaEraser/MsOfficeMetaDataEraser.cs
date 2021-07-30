using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using Microsoft.VisualBasic.FileIO;
using System.Xml;
using System.Windows;
using System.Text.RegularExpressions;
using System.IO;

namespace MetaEraser
{
	class MsOfficeMetaDataEraser
	{
		private String srcFileName;
		private String dstFileName;
		private String extName;
		private String dirName;
		private String[] pathToMetaData = { "\\docProps\\core.xml", "\\docProps\\app.xml" };
		private String zipExt = ".zip";
		private Boolean overwriteExistingFile;

		private static String[] SupportedFileTypes = { "docx", "xlsx", "pptx", "vsdx"};


		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		public MsOfficeMetaDataEraser(String SrcFileName, String DstFileName, Boolean OverwriteExistingFile)
		{
			this.srcFileName = SrcFileName;

			this.overwriteExistingFile = OverwriteExistingFile;

			if(this.overwriteExistingFile || DstFileName.Equals(String.Empty))
			{
				Int32 pointPos = this.srcFileName.LastIndexOf('.');
				Int32 slashPos = this.srcFileName.LastIndexOf('\\');
				this.dstFileName = this.srcFileName.Substring(slashPos + 1, pointPos - slashPos - 1);

				if(!this.overwriteExistingFile)
				{
					this.overwriteExistingFile = true;
				}
			}

			else
			{
				this.dstFileName = DstFileName;
			}

			Int32 pos = this.srcFileName.LastIndexOf('.');
			this.extName = this.srcFileName.Substring(pos + 1, this.srcFileName.Length - pos - 1);
			this.dirName = this.srcFileName.Substring(0, pos) + "_" + Guid.NewGuid().ToString();

			if(!MsOfficeMetaDataEraser.SupportedFileTypes.Contains(this.extName))
			{
				throw new NotSupportedException("File type not supported!" + Environment.NewLine + this.srcFileName);
			}
		}

		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		public void Sanitize()
		{
			this.UnzipFile();
			this.ZeroAllElementsInXml();
			this.ZipFile();
			this.RenameArchieve();
			this.DeleteTmpFiles();
		}

		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		private void UnzipFile()
		{
			System.IO.Compression.ZipFile.ExtractToDirectory(this.srcFileName, this.dirName);
		}

		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		private void ZeroAllElementsInXml()
		{
			for (Int32 i = 0; i < this.pathToMetaData.Count(); i++)
			{
				String fileContent = File.ReadAllText(this.dirName + this.pathToMetaData[i]);

				String[] elementsToDelete = {   "<dc:title>.*</dc:title>",
												"<dc:creator>.*</dc:creator>",
												"<cp:lastModifiedBy>.*</cp:lastModifiedBy>",
												"<cp:revision>.*</cp:revision>",
												"<Application>.*</Application>"
											};

				foreach(var item in elementsToDelete)
				{
					fileContent = Regex.Replace(fileContent, item, String.Empty, RegexOptions.Singleline);
				}

				File.WriteAllText(this.dirName + this.pathToMetaData[i], fileContent);
			}
		}

		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		private void ZipFile()
		{
			System.IO.Compression.ZipFile.CreateFromDirectory(this.dirName, this.dirName + this.zipExt);
		}

		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		private void RenameArchieve()
		{
			if(this.overwriteExistingFile)
			{
				File.Delete(this.srcFileName);
			}

			String newName = this.dstFileName + "." + this.extName;
			FileSystem.RenameFile(this.dirName + this.zipExt, newName);
		}

		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		private void DeleteTmpFiles()
		{
			Directory.Delete(this.dirName, true);
		}
	}
}
