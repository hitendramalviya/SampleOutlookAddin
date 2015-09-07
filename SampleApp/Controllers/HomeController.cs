﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SampleApp.Controllers
{
	public class HomeController : Controller
	{
		[Route("")]
		public ActionResult Index()
		{
			ViewData["appUrl"] = ToAbsoluteUrlWithoutQuery(Request.Url);
			return View();
		}

		private static string ToAbsoluteUrlWithoutQuery(Uri uri)
		{
			var url = uri;
			var port = (url.Port != 80) && (url.Scheme.Equals("https") && url.Port != 443) ? (":" + url.Port) : string.Empty;
			return $"{url.Scheme}://{url.Host}{port}{uri.AbsolutePath}";
		}
	}
}