using System;
using System.Web;
using System.Web.Optimization;

namespace SampleApp
{
	public class BundleConfig
	{
		// For more information on bundling, visit http://go.microsoft.com/fwlink/?LinkId=301862
		public static void RegisterBundles(BundleCollection bundles)
		{
			bundles.IgnoreList.Clear();
			AddDefaultIgnorePatterns(bundles.IgnoreList);

			//bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
			//			"~/Scripts/jquery-{version}.js"));

			//bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
			//			"~/Scripts/jquery.validate*"));

			//// Use the development version of Modernizr to develop with and learn from. Then, when you're
			//// ready for production, use the build tool at http://modernizr.com to pick only the tests you need.
			//bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
			//			"~/Scripts/modernizr-*"));

			//bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
			//		  "~/Scripts/bootstrap.js"));

			//bundles.Add(new StyleBundle("~/Content/css").Include(
			//		  "~/Content/bootstrap.css",
			//		  "~/Content/site.css"));

			bundles.Add(
				new ScriptBundle("~/Scripts/vendor")
					.Include("~/Scripts/jquery-{version}.js")
					.Include("~/Scripts/bootstrap.js")
					.Include("~/Scripts/knockout-{version}.js")
					.Include("~/Scripts/q.js")
					.Include("~/Scripts/MicrosoftAjax.js")
				);

			bundles.Add(
				new StyleBundle("~/Content/css")
					.Include("~/Content/ie10mobile.css")
					.Include("~/Content/bootstrap.min.css")
					.Include("~/Content/font-awesome.min.css")
					.Include("~/Content/durandal.css")
					.Include("~/Content/starterkit.css")
				);

		}

		public static void AddDefaultIgnorePatterns(IgnoreList ignoreList)
		{
			if (ignoreList == null)
			{
				throw new ArgumentNullException("ignoreList");
			}

			ignoreList.Ignore("*.intellisense.js");
			ignoreList.Ignore("*-vsdoc.js");
			ignoreList.Ignore("*.debug.js", OptimizationMode.WhenEnabled);
			//ignoreList.Ignore("*.min.js", OptimizationMode.WhenDisabled);
			//ignoreList.Ignore("*.min.css", OptimizationMode.WhenDisabled);
		}
	}
}
