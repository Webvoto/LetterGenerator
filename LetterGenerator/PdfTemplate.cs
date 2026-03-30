using iText.Forms;
using iText.Kernel.Pdf;
using System.Text;
using System.Text.RegularExpressions;
using PdfWriter = iText.Kernel.Pdf.PdfWriter;

namespace Webvoto.LetterGenerator;

class PdfTemplate {

	private PdfTemplate(byte[] template) {
		Template = template;
	}

	public byte[] Template { get; }

	public List<string> TextFields { get; set; } = new List<string>();

	public void AddTextField(string key) {
		TextFields.Add(key);
	}

	public PdfDocument Generate(Stream outputStream, Dictionary<string, string> data, bool leaveStreamOpen = false) {
		using var inputStream = new MemoryStream(Template);

		var reader = new PdfReader(inputStream);
		var writer = new PdfWriter(outputStream);
		writer.SetCloseStream(!leaveStreamOpen);

		var pdf = new PdfDocument(reader, writer);

		var form = PdfAcroForm.GetAcroForm(pdf, false);

		if (form == null) {
			pdf.Close();
			return pdf;
		}

		var fields = form.GetAllFormFields();

		foreach (var field in fields) {
			var fieldContent = field.Value.GetValueAsString();

			var sb = new StringBuilder(fieldContent);
			foreach (var dataToUse in data) {
				sb.Replace($"{{{dataToUse.Key}}}", dataToUse.Value);
			}

			var finalContent = sb.ToString();
			field.Value.SetValue(finalContent);
		}

		form.FlattenFields();

		pdf.Close();

		return pdf;
	}

	#region Static methods
	private static List<string> getTemplateVariables(string template) {
		var templateRegex = new Regex(@"{([^}{]+?)}", RegexOptions.Compiled);
		return templateRegex.Matches(template).Select(m => m.Groups[1].Value).ToList();
	}

	public static PdfTemplate Create(byte[] bytes) {
		using var memoryStream = new MemoryStream(bytes);
		var result = new PdfTemplate(bytes);

		loadFieldsFromStream(memoryStream, result);

		return result;
	}

	private static void loadFieldsFromStream(Stream stream, PdfTemplate result) {
		var reader = new PdfReader(stream);
		var pdf = new PdfDocument(reader);
		var form = PdfAcroForm.GetAcroForm(pdf, false);

		if (form == null) {
			pdf.Close();
			return;
		}

		var fields = form.GetAllFormFields();
		foreach (var fieldName in fields.Keys) {
			var field = fields[fieldName];
			var value = field.GetValueAsString();

			// do something only if the type of the field is FIELD_TYPE_TEXT; if not, ignore the field.
			if (string.IsNullOrEmpty(value)) {
				result.AddTextField(fieldName);
			} else {
				getTemplateVariables(value).ForEach(f => result.TextFields.Add(f));
			}
		}

		pdf.Close();
	}
	#endregion
}
