using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using SampleApp.Models;
using System.Net;
using System.Text;
using System.Web;

namespace SampleApp.Controllers
{
    public class OfficeApiController : ApiController
	{
		[Route("OfficeApi/ProcessMailAttachment")]
		[HttpPost]
		public async Task<IHttpActionResult> ProcessMailAttachment(AttachmentProcessRequest request)
		{
			var response = await ProcessAttachmentRequest(request.Attachments.First().Id, request.EwsUrl, request.Token);
			return Ok(response);
		}

		private async Task<HttpResponseMessage> ProcessAttachmentRequest(string attachmentId, string ewsUrl, string token)
		{
			var soapString = string.Format(AttachmentSoapRequest, attachmentId);
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
				var contentRequest = new StringContent(soapString, Encoding.UTF8, "text/xml");
				var response = await client.PostAsync(uri, contentRequest);
				stream = await response.Content.ReadAsStreamAsync();
				var attachmentStream = new AttachmentStream(stream);
				var res = Request.CreateResponse();
				res.Content = new PushStreamContent(attachmentStream.WriteToStream());
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

		private const string AttachmentSoapRequest =
			@"<?xml version=""1.0"" encoding=""utf-8""?>
		<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
		 xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
		 xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
		 xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
		<soap:Header>
		<t:RequestServerVersion Version=""Exchange2013"" />
		</soap:Header>
			<soap:Body>
				<GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
				 xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
					<AttachmentShape/>
						<AttachmentIds>
						<t:AttachmentId Id=""{0}""/>
					</AttachmentIds>
				</GetAttachment>
			</soap:Body>
		</soap:Envelope>";
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
