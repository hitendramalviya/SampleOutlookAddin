import app = require("durandal/app");
import viewLocator = require("durandal/viewLocator");
import system = require("durandal/system");
declare var appUrl: any;
export function start(): void {
	if (window.location.href.indexOf("https://outlook") > -1) {
		window.location.href = appUrl;
		return;
	}

	//>>excludeStart("build", true);
	system.debug(true);
	//>>excludeEnd("build");

	app.title = 'Durandal Starter Kit';

	app.configurePlugins({
		router: true,
		dialog: true
	});

	Q(app.start()).then(() => {
		//Replace 'viewmodels' in the moduleId with 'views' to locate the view.
		//Look for partial views in a 'views' folder in the root.
		viewLocator.useConvention();

		//Show the app by setting the root view model for our application with a transition.
		app.setRoot('viewmodels/shell', 'entrance');
	}).done();
}