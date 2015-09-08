define("jquery", () => jQuery);
define("knockout", () => ko);

requirejs.config({
	paths: {
		'text': '../Scripts/text',
		'durandal': '../Scripts/durandal',
		'plugins': '../Scripts/durandal/plugins',
		'transitions': '../Scripts/durandal/transitions'
	}
});

require(["bootstrapper"], (bootstrapper) => {
	bootstrapper.start();
});