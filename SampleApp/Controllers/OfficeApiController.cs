using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using SampleApp.Models;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Web;
using Newtonsoft;
using Newtonsoft.Json;

namespace SampleApp.Controllers
{
    public class OfficeApiController : ApiController
	{
		[Route("OfficeApi/ProcessMailAttachment")]
		[HttpPost]
		public async Task<HttpResponseMessage> ProcessMailAttachment(AttachmentProcessRequest request)
		{
			var response = await ProcessAttachmentRequest(request.Attachment.Id, request.EwsUrl, request.Token);
			return response;
		}

		private async Task<HttpResponseMessage> ProcessAttachmentRequest(string attachmentId, string ewsUrl, string token)
		{
			//var soapString = string.Format(AttachmentSoapRequest, attachmentId);
			var client = new HttpClient
			{
				DefaultRequestHeaders =
				{
					Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token)
				},
				Timeout = TimeSpan.FromDays(1)
			};
			var stream = Stream.Null;
			try
			{
				var uri = new Uri(ewsUrl);
				//var contentRequest = new StringContent(soapString, Encoding.UTF8, "text/xml");
				var response = await client.GetStreamAsync(uri);

				var sReader = new StreamReader(response);
				var jsonTextReader = new JsonTextReader(sReader);
				var serializer = new JsonSerializer();
				var attData = serializer.Deserialize<AttachmentData>(jsonTextReader);

				//new Bina
				//stream = await response.Content.ReadAsStreamAsync();
				//var attachmentStream = new AttachmentStream(stream);
				var res = Request.CreateResponse();
				//res.Content = new PushStreamContent(attachmentStream.WriteToStream());
				res.Content = new StreamContent(response);
				res.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
				return res;
			}
			catch (Exception)
			{
				return new HttpResponseMessage(HttpStatusCode.InternalServerError);
			}
			finally
			{
				stream.Close();
			}
		}
	}

	public class AttachmentData
	{
		public string ContentBytes { get; set; }
		public string Id { get; set; }
		public string Name { get; set; }
	}

	public class AttachmentStream
	{
		private readonly Stream _stream;

		public AttachmentStream(Stream stream)
		{
			_stream = stream;
		}

		public Action<Stream, HttpContent, TransportContext> WriteToStream()
		{
			return async (outputStream, content, context) =>
			{
				//This controls how many bytes to read at a time and send to the client
				var bytesToRead = 10000;

				// Buffer to read bytes in chunk size specified
				var buffer = new byte[bytesToRead];
				try
				{
					int length;
					do
					{
						length = _stream.Read(buffer, 0, bytesToRead);
						await outputStream.WriteAsync(buffer, 0, length);
						//Clear the buffer
						buffer = new byte[bytesToRead];
					} while (length > 0);
				}
				catch (HttpException)
				{
					return;
				}
				finally
				{
					outputStream.Close();
				}
			};
		}
	}
}
