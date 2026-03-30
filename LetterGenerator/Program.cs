using CsvHelper;
using ICSharpCode.SharpZipLib.Zip;
using iText.Kernel.Pdf;
using System.Collections.Concurrent;
using System.Globalization;

namespace Webvoto.LetterGenerator;

record Arguments(FileInfo LettersFile, FileInfo TemplateFile, FileInfo OutputFile);

record Letter(string Identifier, string Name, string Address, string CepNetCode, string Password) {

	public Dictionary<string, string> GetFields() => new() {
		[nameof(Identifier)] = Identifier,
		[nameof(Name)] = Name,
		[nameof(Address)] = Address.Replace("\\n", "\n").Replace("\\t", "    "),
		[nameof(CepNetCode)] = CepNetCode,
		[nameof(Password)] = Password,
	};
}

record LetterFile(string Name, List<Letter> Letters);

class Program {

	private const int GroupSize = 500;
	private const int ThreadCount = 16;

	static int Main(string[] args) {
		try {
			Task.Run(() => runAsync(args)).GetAwaiter().GetResult();
			Console.WriteLine("Done.");
			return 0;
		} catch (Exception ex) {
			Console.WriteLine($"FATAL: {ex}");
			return 1;
		}
	}

	static async Task runAsync(string[] args) {

		// Parse arguments

		var arguments = parse(args);

		// Read password from console

		var password = readPassword();

		// Read letter data

		var letters = read(arguments.LettersFile, password);

		// Divide letters into files

		var files = group(letters);

		// Generate letter files

		await generateAsync(files, arguments.TemplateFile, arguments.OutputFile.Directory);

		// Pack generated files into ZIP file using same password

		pack(files, arguments.OutputFile, password);
	}

	static Arguments parse(string[] args) {

		var lettersFilePath = args.ElementAtOrDefault(0);
		var templateFilePath = args.ElementAtOrDefault(1);
		var outputFilePath = args.ElementAtOrDefault(2);

		if (string.IsNullOrEmpty(lettersFilePath) || string.IsNullOrEmpty(templateFilePath) || string.IsNullOrEmpty(outputFilePath)) {
			throw new Exception("Syntax: LetterGenerator <letters file path> <PDF template file path> <output file path>");
		}

		var lettersFileInfo = new FileInfo(lettersFilePath);
		var templateFileInfo = new FileInfo(templateFilePath);
		var outputFileInfo = new FileInfo(outputFilePath);

		ensureExists(lettersFileInfo);
		ensureExists(templateFileInfo);
		ensureExists(outputFileInfo.Directory);

		return new Arguments(lettersFileInfo, templateFileInfo, outputFileInfo);
	}

	static List<Letter> read(FileInfo zipFileInfo, string password) {

		using var zipFileStream = zipFileInfo.OpenRead();
		using var zipFile = new ZipFile(zipFileStream) { Password = password };

		var csvFileEntry = zipFile.Cast<ZipEntry>().FirstOrDefault(e => e.Name.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) ?? throw new Exception("Letters ZIP file does not contain a CSV file");

		using var csvStream = zipFile.GetInputStream(csvFileEntry);
		using var csvStreamReader = new StreamReader(csvStream);
		using var csvReader = new CsvReader(csvStreamReader, CultureInfo.CurrentCulture);

		return [.. csvReader.GetRecords<Letter>()];
	}

	static List<LetterFile> group(List<Letter> letters) {

		List<LetterFile> files = [];

		for (var offset = 0; offset < letters.Count; offset += GroupSize) {
			var length = Math.Min(GroupSize, letters.Count - offset);
			var fileName = $"{(offset + 1):D6}-{(offset + length):D6}.pdf";
			files.Add(new LetterFile(fileName, [.. letters.Skip(offset).Take(length)]));
		}

		return files;
	}

	static async Task generateAsync(List<LetterFile> files, FileInfo templateFileInfo, DirectoryInfo outputDirectory) {
		var fileBag = new ConcurrentBag<LetterFile>(files);
		var tasks = Enumerable.Range(0, ThreadCount).Select(_ => Task.Run(() => consume(fileBag, templateFileInfo, outputDirectory)));
		await Task.WhenAll(tasks);
	}

	static void consume(ConcurrentBag<LetterFile> source, FileInfo templateFileInfo, DirectoryInfo outputDirectory) {

		var pdfTemplate = PdfTemplate.Create(File.ReadAllBytes(templateFileInfo.FullName));

		while (source.TryTake(out var file)) {

			var fileInfo = getLetterFileInfo(file, outputDirectory);
			if (fileInfo.Exists) {
				Console.WriteLine($"{file.Name} : already exists");
			} else {
				Console.WriteLine($"{file.Name} : generating ...");
				using var fileStream = fileInfo.Create();
				var pdfWriter = new PdfWriter(fileStream);
				pdfWriter.SetSmartMode(true);

				var pdfDocument = new PdfDocument(pdfWriter);
				pdfDocument.InitializeOutlines();

				foreach (var letter in file.Letters) {

					using var buffer = new MemoryStream();

					using var letterPdf = pdfTemplate.Generate(buffer, letter.GetFields(), leaveStreamOpen: true);

					buffer.Position = 0;

					var letterDoc = new PdfDocument(new PdfReader(buffer));
					letterDoc.CopyPagesTo(1, letterDoc.GetNumberOfPages(), pdfDocument);
					letterDoc.Close();
				}

				pdfDocument.Close();
				
				Console.WriteLine($"{file.Name} : generated");
			}
		}
	}

	static void pack(List<LetterFile> files, FileInfo outputFileInfo, string password) {

		Console.WriteLine("Generating ZIP file with PDFs ...");

		using var outputFileStream = outputFileInfo.Create();
		using var zipStream = new ZipOutputStream(outputFileStream) { Password = password };

		foreach (var file in files.OrderBy(p => p.Name)) {
			zipStream.PutNextEntry(new ZipEntry(file.Name));
			var fileInfo = getLetterFileInfo(file, outputFileInfo.Directory);
			using (var fileStream = fileInfo.OpenRead()) {
				fileStream.CopyTo(zipStream);
			}
			fileInfo.Delete();
		}

		zipStream.Close();
	}

	#region Helper methods

	static void ensureExists(FileSystemInfo item) {
		if (!item.Exists) {
			throw new Exception($"Not found: {item.FullName}");
		}
	}

	static string readPassword() {

		Console.WriteLine("Enter the ZIP file password:");

		var password = string.Empty;

		ConsoleKeyInfo key;

		do {
			key = Console.ReadKey(true); // Intercept the keypress (do not echo)

			// Process the key
			if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter && !char.IsControl(key.KeyChar)) {
				// Append the character to the password and optionally display a mask
				password += key.KeyChar;
				Console.Write("*"); // Display an asterisk
			} else if (key.Key == ConsoleKey.Backspace && password.Length > 0) {
				// Handle backspace: remove the last character and erase the asterisk
				password = password[..^1];
				Console.Write("\b \b"); // Move back, write space, move back again
			}
			// Ignore other control keys (like arrow keys, function keys, etc.)
		} while (key.Key != ConsoleKey.Enter); // Stop when Enter is pressed

		Console.WriteLine(); // Add a new line after the password entry is complete
		return password;
	}

	private static FileInfo getLetterFileInfo(LetterFile letterFile, DirectoryInfo outputDirectory)
		=> new(Path.Combine(outputDirectory.FullName, letterFile.Name));

	#endregion
}
