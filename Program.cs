using NMaier.GetOptNet;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace App.TipoPdfCreator
{
    class Program
	{
		[GetOptOptions(OnUnknownArgument = UnknownArgumentsAction.Throw,
			UsageEpilog = "ex:\r\n  App.TipoPdfCreator.exe -s PP1024020 -d 201307\r\n  App.TipoPdfCreator.exe -s AP2013111401 -d app201307")]
		class Opts : GetOpt
		{
			[Argument("srcDirPrefix", HelpText = @"PatImg dir prefix, ex: path/to/tiff")]
			[ShortArgument('i')]
			public string PatImgDirPrefix = @"path/to/tiff";

			[Argument("outputDirPrefix", HelpText = @"PatPdf dir prefix, ex: path/to/pdf")]
			[ShortArgument('o')]
			public string OutputDirPrefix = @"path/to/pdf";


			private string _patImgDirNames;
			[Argument("src", Required = true, HelpText = "PatImg dir names, ex: PP1024016,PP1024017,PP1024018")]
			[ShortArgument('s')]
			public string PatImgDirNames
			{
				get { return _patImgDirNames; }
				set
				{
					if (value.Split(',').Any(s => !Regex.IsMatch(s, @"(PP|AP)\d+")))
					{
						throw new InvalidValueException("Incorrect PatImg dir names.");
					}
					_patImgDirNames = value;
				}
			}

			private string _outputDirName;
			[Argument("dst", Required = true, HelpText = "PatPdf dir name, ex: 201306")]
			[ShortArgument('d')]
			public string OutputDirName
			{
				get { return _outputDirName; }
				set
				{
					if (!Regex.IsMatch(value, @"(app)?\d{6}"))
					{
						throw new InvalidValueException("Incorrect PatPdf dir name.");
					}
					_outputDirName = value;
				}
			}


			[Argument("resumeOnError", HelpText = "Resume next on error, default is false.")]
			[ShortArgument('r')]
			[FlagArgument(true)]
			public bool ResumeOnError = false;
		}

		static void Main(string[] args)
		{
			Opts opts = new Opts();
			try
			{
				opts.Parse(args);

				var app = new TipoPdfCreatorApp();
				app.ResumeOnError = opts.ResumeOnError;
                app.Main(opts.PatImgDirNames, opts.OutputDirName, opts.PatImgDirPrefix, opts.OutputDirPrefix);
			}
			catch (GetOptException)
			{
				Console.WriteLine("\r\nTipoPdfCreator v" + typeof(Program).Assembly.GetName().Version.ToString() + "\r\n");
				opts.PrintUsage();
			}
		}

	}
}
