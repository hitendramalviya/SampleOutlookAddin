using System.Threading.Tasks;
using System.Web.Http;
using Gecko.Elements.Models;
using Gecko.Elements.Helpers;

namespace Gecko.Elements.Controllers
{
	public class OfficeApiController : ApiController
	{
		[Route("ProcessAttachement")]
		[HttpPost]
		public async Task<IHttpActionResult> ProcessMailAttachment(AttachmentRequest request)
		{
			if (request.ServiceType == "soap")
			{
				return Ok(await O365AttachmentHelper.GetAttachmentFromEwsAndUpload(request));
			}
			else
			{
				return Ok(await O365AttachmentHelper.GetAttachmentFromO365AndUpload(request));
			}
		}
	}
}
