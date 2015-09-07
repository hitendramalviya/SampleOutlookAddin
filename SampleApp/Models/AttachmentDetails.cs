using Newtonsoft.Json;

namespace SampleApp.Models
{
	public class AttachmentDetails
	{
		[JsonProperty(PropertyName = "attachmentType")]
		public string AttachmentType { get; set; }

		[JsonProperty(PropertyName = "contentType")]
		public string ContentType { get; set; }

		[JsonProperty(PropertyName = "id")]
		public string Id { get; set; }

		[JsonProperty(PropertyName = "isInline")]
		public bool IsInline { get; set; }

		[JsonProperty(PropertyName = "name")]
		public string Name { get; set; }

		[JsonProperty(PropertyName = "size")]
		public int Size { get; set; }
	}
}