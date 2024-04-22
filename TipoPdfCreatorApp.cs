using BitMiracle.LibTiff.Classic;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace App.TipoPdfCreator
{
    class TipoPdfCreatorApp
	{
		public List<string> ErrorList { get; private set; }
		public bool ResumeOnError { get; set; }
		public int Skip { get; set; }
		public bool SkipIfExists { get; set; }
		private string _creator = System.Configuration.ConfigurationManager.AppSettings["Creator"];
		public TipoPdfCreatorApp()
		{
			this.ErrorList = new List<string>();
			this.ResumeOnError = false;
			this.SkipIfExists = true;
			this.Skip = 0;
		}

		public void Main(string patImgDirNames, string outputDirName, string patImgDir = @"path/to/tiff", string patPdfDir = @"path/to/paf")
		{
			foreach (var srcFolder in patImgDirNames.Split(','))
			{
				bool issued = srcFolder.StartsWith("PP");
				int volOffset = issued ? 2 : 4;
				int volNo = Int32.Parse(srcFolder.Substring(srcFolder.Length - volOffset, 2));
				int month = Int32.Parse(outputDirName.Substring(outputDirName.Length - 2));

				bool valid = issued && (month * 3 - volNo < 3) && Regex.IsMatch(outputDirName, @"\d{6}") ||
							!issued && (month * 2 - volNo < 2) && Regex.IsMatch(outputDirName, @"app\d{6}");
				if (!valid) throw new Exception("CD 名稱和目的資料夾名稱不符!");

				srcFolder.Dump();
				BuildPdfFromTipoTiff(Path.Combine(patImgDir, srcFolder), Path.Combine(patPdfDir, outputDirName), issued);
			}
			if (ErrorList.Count > 0) ErrorList.Dump("Error List");
		}

		void BuildPdfFromTipoTiff(string srcFolder, string dstFolder, bool issued)
		{
			Directory.CreateDirectory(dstFolder);
			var targets = Directory.EnumerateDirectories(srcFolder);

			int idx = 0;
			foreach (var pnDir in targets)
			{
				if (Path.GetFileName(pnDir).ToLower().EndsWith("corrections")) continue;

				idx++;
				if (idx <= this.Skip)
				{
					Console.Write("\rSkipping " + idx.ToString());
					continue;
				}

				var fileNameList = getOrderedFileNameList(Directory.GetFiles(pnDir));

				var pn = Path.GetFileName(pnDir);
				var pdfFileName = Path.Combine(dstFolder, "TW" + pn + ".pdf");
				var tempPdfFileName = pdfFileName + ".tmp";
				var errorPdfFileName = pdfFileName + ".err";
				var watermarkPdfFileName = pdfFileName + ".wm";

				if (this.SkipIfExists && File.Exists(pdfFileName))
				{
					String.Format("[{0:0000}] {1} exists, skipping!", idx, pdfFileName).Dump();
					continue;
				}

				try
				{
					File.Delete(tempPdfFileName);
					File.Delete(errorPdfFileName);
					mergeImages(fileNameList, tempPdfFileName, pnDir, pn, issued);
					addWatermarkText(tempPdfFileName, _creator, watermarkPdfFileName);
                    File.Delete(pdfFileName);
                    File.Delete(tempPdfFileName);
                    File.Move(watermarkPdfFileName, pdfFileName);
					String.Format("[{0:0000}] {1} ok", idx, pdfFileName).Dump();
				}
				catch (Exception ex)
				{
					if (ex is InvalidDataException)
					{
						File.Delete(errorPdfFileName);
						File.Move(tempPdfFileName, errorPdfFileName);

						ErrorList.Add(errorPdfFileName + "\r\n\t" + ex.Message);
					}

					if (!ResumeOnError)
					{
						ErrorList.Dump();
						throw;
					}
				}
			}
		}

		void mergeImages(IList<string> inputImageFileNames, string outputPdfFileName, string pnDir,  string pn, bool issued)
		{
			using (var doc = new iTextSharp.text.Document())
			{
				var w = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, File.OpenWrite(outputPdfFileName));
				
				w.SetLinearPageMode();

				var Author = "InnoVue Corp.";
				var Creator = "Webpat Pdf Builder 2022";
				var Subject = Path.GetFileNameWithoutExtension(outputPdfFileName);
				var Title = Subject;
				var Keywords = Subject;

				doc.AddTitle(Title);
				doc.AddSubject(Subject);
				doc.AddKeywords(Keywords);
				doc.AddAuthor(Author);
				doc.AddCreator(Creator);


				doc.Open();
				iTextSharp.text.Rectangle a4 = iTextSharp.text.PageSize.A4;
				for (int i = 0; i < inputImageFileNames.Count; i++)
				{
					var fileName = inputImageFileNames[i];

					Console.Write("\r{0} : {1}/{2}", outputPdfFileName, i + 1, inputImageFileNames.Count);
					var imgFileName = makeSureImageFormatAndSize(fileName);
					try
					{
						iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imgFileName);
						img.SetAbsolutePosition(0, 0);
						img.ScaleAbsolute(a4.Width, a4.Height);

						doc.NewPage();
						doc.Add(img);
					}
					catch (Exception ex)
					{
						throw new InvalidDataException("Failed reading image: " + imgFileName, ex);
					}
				}

				doc.Close();
			}
		}

		private void addWatermarkText(string tempPdfFileName, string text, string watermarkPdfFileName)
		{
			using (Stream inputPdfStream = File.OpenRead(tempPdfFileName))
			using (FileStream outputPdfStream = new FileStream(watermarkPdfFileName, FileMode.Create))
			{
				using (PdfReader reader = new PdfReader(inputPdfStream))
				using (PdfStamper pdfStamper = new PdfStamper(reader, outputPdfStream))
				{

					for (var i = 1; i <= reader.NumberOfPages; i++)
					{
						PdfContentByte pdfPageContents = pdfStamper.GetOverContent(i);

						//文字調整成透明
						pdfPageContents.SetGState(new PdfGState() { FillOpacity = 0.0f });

						ColumnText.ShowTextAligned(pdfPageContents, Element.ALIGN_LEFT, new Phrase(text), 0, 0, 0);
					}
				}
			}

		}

		List<string> getOrderedFileNameList(IList<string> fileNameList)
		{
			var dupStore = new HashSet<int>();
			Func<int, int, int> getDuplicateMark = (type, page) =>
			{
				int val = (type * 10000) + page;
				if (dupStore.Contains(val)) return 1;

				dupStore.Add(val);
				return 0;
			};

			var q = from f in fileNameList.AsQueryable()
					let ext = Path.GetExtension(f).ToLower()
                    where ext != ".txt" && ext != ".db" && ext != ".nrg"

					let name = Path.GetFileNameWithoutExtension(f)

					let parts = Regex.Split(name, "-")
					let type = Int32.Parse(parts[1]) % 100
					let ver = Int32.Parse(parts[1]) / 100
					let page = Int32.Parse(parts[2])
					let date = parts.Length > 3 ? parts[3] : ""

					let showOrder = getShowOrder(type)
					orderby showOrder, page, ver, date

					let dupMark = getDuplicateMark(type, page)
					orderby dupMark, showOrder, page, ver, date

					select new { f, dupMark, type, showOrder, page, ver, date };

			return q.Select(x => x.f).ToList();
		}

		int getShowOrder(int type)
		{
			// type order: 1, 2, 4, 12, 8, 3, 5, 6, 7, 9, 10, 11
			if (type == 4) return 21;
			if (type == 12) return 22;
			if (type == 8) return 23;

			return type * 10;
		}

		string makeSureImageFormatAndSize(string fileName)
		{
			if (Path.GetExtension(fileName).ToLower() == ".tif")
			{
				int bitsPerSample;
				int samplesPerPixel;
				bool isBW;
				bool isCCITTFAX4;
				bool isLZW;

                var tiff = Tiff.Open(fileName, "r");
                if (tiff == null)
                {
                    // move original file to ./orig/{fileName}
                    var origFileName = backupOrigImage(fileName);
                    // save bw image in CCITT4 format
                    saveAsTiff(origFileName, fileName);
                }
                else
                {
                    tiff.Dispose();
                }


                using (var tiff2 = Tiff.Open(fileName, "r"))
                {
                    bitsPerSample = tiff2.GetFieldDefaulted(TiffTag.BITSPERSAMPLE)[0].ToInt();
                    samplesPerPixel = tiff2.GetFieldDefaulted(TiffTag.SAMPLESPERPIXEL)[0].ToInt();
                    isBW = bitsPerSample == 1 && samplesPerPixel == 1;
                    int compression = tiff2.GetFieldDefaulted(TiffTag.COMPRESSION)[0].ToInt();
                    isCCITTFAX4 = compression == (int)Compression.CCITTFAX4;
                    isLZW = compression == (int)Compression.LZW;
                }
                   
				if (isBW && isCCITTFAX4)
				{
					return fileName;
				}
				else if (isBW)
				{
					// move original file to ./orig/{fileName}
					var origFileName = backupOrigImage(fileName);

					// save bw image in CCITT4 format
					saveAsTiff(origFileName, fileName);
					return fileName;
				}
				else
				{
					// save image in png format
					var pngFileName = Path.ChangeExtension(fileName, ".png");
					if (saveAsPngWhenSmallerEnough(fileName, pngFileName))
					{
						// move original file to ./orig/{fileName}
						backupOrigImage(fileName);
						return pngFileName;
					}
					else
					{
						// ex: AP2013110201\201302165 (skip = 116)
						return fileName;
					}
				}
			}
			return fileName;
		}

		private static string backupOrigImage(string fileName)
		{
			var origFileName = Path.GetDirectoryName(fileName) + "\\orig\\" + Path.GetFileName(fileName);
			Directory.CreateDirectory(Path.GetDirectoryName(origFileName));
            if (!File.Exists(origFileName))
            {
                File.Move(fileName, origFileName);
            }
			
			return origFileName;
		}

		private bool saveAsPngWhenSmallerEnough(string srcFileName, string dstFileName, int sizeDiff = 20000)
		{
			using (var img = System.Drawing.Image.FromFile(srcFileName))
			using (var ms = new MemoryStream())
			{
				img.Save(ms, ImageFormat.Png);
				var imgBytes = ms.ToArray();

				// if new size if smaller the original
				long origSize = new FileInfo(srcFileName).Length;
				if (origSize - imgBytes.Length > sizeDiff)
				{
					File.WriteAllBytes(dstFileName, imgBytes);
					return true;
				}
				return false;
			}
		}

		private void saveAsTiff(string srcFileName, string dstFileName)
		{
			using (var img = System.Drawing.Image.FromFile(srcFileName))
			{
				var encoder = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/tiff");
				var encoderParams = new EncoderParameters(1);
				encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
				img.Save(dstFileName, encoder, encoderParams);
			}
		}
	}

	public static class DumpExtension
    {
		public static T Dump<T>(this T input, string name = "")
		{
			Console.WriteLine("\r{0} {1} {2}", DateTime.Now.ToString("HH:mm:ss.f"), name, input);
			return input;
		}
	}
}
