using Newtonsoft.Json;
using System.Collections.Generic;

namespace SampleApp.Models
{
	public class AttachmentProcessRequest
	{
		[JsonProperty(PropertyName = "token")]
		public string Token { get; set; }

		[JsonProperty(PropertyName = "ewsUrl")]
		public string EwsUrl { get; set; }

		[JsonProperty(PropertyName = "attachment")]
		public AttachmentDetails Attachment { get; set; }

		[JsonProperty(PropertyName = "documentServiceUrl")]
		public string DocumentServiceUrl { get; set; }

		[JsonProperty(PropertyName = "documentServiceToken")]
		public string DocumentServiceToken { get; set; }
	}
}