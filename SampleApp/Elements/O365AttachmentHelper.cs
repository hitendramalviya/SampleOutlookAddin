using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using Newtonsoft.Json;
using Gecko.Elements.Models;

namespace Gecko.Elements.Helpers
{
	public static class O365AttachmentHelper
	{
		public static async Task<UploadResponse> GetAttachmentFromO365AndUpload(AttachmentRequest request)
		{
			//Debug & StopWatch introduced just to calculate & overview performance with response time. Need remove once have stable.
			Debug.WriteLine("AttachmentUrl={0}{1}", request.ApiUrl, string.Empty);
			Debug.WriteLine("ApiToken=Bearer {0}{1}", request.Token, string.Empty);
			var uri = new Uri(request.ApiUrl);
			using (var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false })
			{
				DefaultRequestHeaders = { Authorization = new AuthenticationHeaderValue("Bearer", request.Token) }
			})
			{
				client.DefaultRequestHeaders.Add("x-AnchorMailbox", request.UserEmail);
				var watch = Stopwatch.StartNew();
				using (var response = await client.GetStreamAsync(uri))
				{
					watch.Stop();
					Debug.WriteLine("Time taken to get attachment from o365 reset api {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
					watch = Stopwatch.StartNew();
					using (var streamReader = new StreamReader(response))
					{
						watch.Stop();
						Debug.WriteLine("Time taken to initialize StreamReader {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
						watch = Stopwatch.StartNew();
						using (var jsonTextReader = new JsonTextReader(streamReader))
						{
							watch.Stop();
							Debug.WriteLine("Time taken to initialize JsonTextReader {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var serializer = new JsonSerializer();
							var attachmentResponse = serializer.Deserialize<AttachmentResponse>(jsonTextReader);
							watch.Stop();
							Debug.WriteLine("Time taken to Deserialize {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var bytes = Convert.FromBase64String(attachmentResponse.ContentBytes);
							watch.Stop();
							Debug.WriteLine("Time taken to ConvertBase64 to bytes array {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var byteArrayContent = new ByteArrayContent(bytes);
							watch.Stop();
							Debug.WriteLine("Time taken to prepare to ByteArrayContent {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var result = await PostDocumentAsync(byteArrayContent, request.DocumentServiceUrl,
								request.DocumentServiceToken);
							watch.Stop();
							Debug.WriteLine("Time taken to post and upload on ncore server {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							return result;
						}
					}
				}
			}
		}

		private static async Task<UploadResponse> PostDocumentAsync(HttpContent content, string url, string token)
		{
			var uri = new Uri(url);
			using (var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false })
			{
				BaseAddress = new Uri($"{uri.Scheme}://{uri.Host}"),
				DefaultRequestHeaders = { Authorization = new AuthenticationHeaderValue("Bearer", token) }
			})
			{
				using (var response = await client.PostAsync(uri, content))
				{
					return await response.Content.ReadAsAsync<UploadResponse>();
				}
			}
		}

		#region Experimental
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

		public static async Task<UploadResponse> GetAttachmentFromEwsAndUpload(AttachmentRequest request)
		{
			var soapString = string.Format(AttachmentSoapRequest, request.Attachment.Id);
			using (var client = new HttpClient
			{
				DefaultRequestHeaders =
				{
					Authorization = new AuthenticationHeaderValue("Bearer", request.EwsToken)
				},
				Timeout = TimeSpan.FromDays(1)
			})
			{
				client.DefaultRequestHeaders.Add("x-AnchorMailbox", request.UserEmail);
				var uri = new Uri(request.EwsUrl);
				var watch = Stopwatch.StartNew();
				var contentRequest = new StringContent(soapString, Encoding.UTF8, "text/xml");
				using (var response = await client.PostAsync(uri, contentRequest))
				{
					Debug.WriteLine("Time taken to post to EWS {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
					watch = Stopwatch.StartNew();
					using (var stream = await response.Content.ReadAsStreamAsync())
					{
						Debug.WriteLine("Time taken to read Stream {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
						watch = Stopwatch.StartNew();
						ByteArrayContent contentToUpload = null;
						using (var xmlReader = XmlReader.Create(stream))
						{
							Debug.WriteLine("Time create xml reader {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds,
								request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var contentFound = false;
							while (xmlReader.Read())
							{
								if (contentFound && xmlReader.NodeType == XmlNodeType.Text)
								{
									watch.Stop();
									Debug.WriteLine("Time taken to Reach to cotent in xml stream {0}ms. AttachmentName={1} Size={2}",
										watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
									watch = Stopwatch.StartNew();
									var bytes = Convert.FromBase64String(xmlReader.Value);
									watch.Stop();
									Debug.WriteLine("Time taken to ConvertBase64 to bytes array {0}ms. AttachmentName={1} Size={2}",
										watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
									watch = Stopwatch.StartNew();
									contentToUpload = new ByteArrayContent(bytes);
									watch.Stop();
									Debug.WriteLine("Time taken to prepare to ByteArrayContent {0}ms. AttachmentName={1} Size={2}",
										watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
									break;
								}
								if (!xmlReader.Name.Equals("t:Content")) continue;
								contentFound = true;
							}

							watch = Stopwatch.StartNew();
							var result = await PostDocumentAsync(contentToUpload, request.DocumentServiceUrl,
								request.DocumentServiceToken);
							watch.Stop();
							Debug.WriteLine("Time taken to post and upload on ncore server {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							return result;
						}
					}
				}
			}
		}

		public static async Task<UploadResponse> GetAttachmentFromO365AndUpload2(AttachmentRequest request)
		{
			//Debug & StopWatch introduced just to calculate & overview performance with response time. Need remove once have stable.
			Debug.WriteLine("AttachmentUrl={0}{1}", request.ApiUrl, string.Empty);
			Debug.WriteLine("ApiToken=Bearer {0}{1}", request.Token, string.Empty);
			var uri = new Uri(request.ApiUrl);
			using (var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false })
			{
				DefaultRequestHeaders = { Authorization = new AuthenticationHeaderValue("Bearer", request.Token) }
			})
			{
				client.DefaultRequestHeaders.Add("x-AnchorMailbox", request.UserEmail);
				var watch = Stopwatch.StartNew();
				using (var response = await client.GetStreamAsync(uri))
				{
					watch.Stop();
					Debug.WriteLine("Time taken to get attachment from o365 reset api {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
					watch = Stopwatch.StartNew();
					using (var streamReader = new StreamReader(response))
					{
						watch.Stop();
						Debug.WriteLine("Time taken to initialize StreamReader {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
						watch = Stopwatch.StartNew();
						using (var jsonTextReader = new JsonTextReader(streamReader))
						{
							watch.Stop();
							Debug.WriteLine("Time taken to initialize JsonTextReader {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var serializer = new JsonSerializer();
							var attachmentResponse = serializer.Deserialize<AttachmentResponse>(jsonTextReader);
							watch.Stop();
							Debug.WriteLine("Time taken to Deserialize {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var bytes = Convert.FromBase64String(attachmentResponse.ContentBytes);
							watch.Stop();
							Debug.WriteLine("Time taken to ConvertBase64 to bytes array {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var byteArrayContent = new ByteArrayContent(bytes);
							watch.Stop();
							Debug.WriteLine("Time taken to prepare to ByteArrayContent {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							watch = Stopwatch.StartNew();
							var result = await PostDocumentAsync(byteArrayContent, request.DocumentServiceUrl,
								request.DocumentServiceToken);
							watch.Stop();
							Debug.WriteLine("Time taken to post and upload on ncore server {0}ms. AttachmentName={1} Size={2}", watch.ElapsedMilliseconds, request.Attachment.Name, request.Attachment.Size);
							return result;
						}
					}
				}
			}
		}
		#endregion
	}
}