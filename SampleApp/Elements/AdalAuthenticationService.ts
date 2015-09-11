import Config = require("config/Config");
import AdalWrapper = require("wrapper/AdalWrapper");

class AdalAuthenticationService {
	private oAuthData: any = {};
	private adal: AdalWrapper;
	constructor() {
		//extend o365 auth to window pop up setting
		Config.current.o365.displayCall = this.displayCall;
		this.oAuthData.isAuthenticated = false;
		this.oAuthData.loginError = "";
		this.oAuthData.profile = "";
		this.oAuthData.userName = "";
	}
	private displayCall(url) {
		var newwindow = window.open(url, "Authentication", "height=500, width=700");
		if (window.focus) {
			newwindow.focus();
		}
	}

	private updateDataFromCache(resource: any) {
		// only cache lookup here to not interrupt with events
		var token = this.adal.getCachedToken(resource);
		this.oAuthData.isAuthenticated = token !== null && token.length > 0;
		var user = this.adal.getCachedUser() || { userName: "" };
		this.oAuthData.userName = user.userName;
		this.oAuthData.profile = user.profile;
		this.oAuthData.loginError = this.adal.getLoginError();
	}

	init() {
		var configOptions = Config.current.o365;

		// redirect and logout_redirect are set to current location by default
		var existingHash = window.location.hash;
		var pathDefault = window.location.href;
		if (existingHash) {
			pathDefault = pathDefault.replace(existingHash, "");
		}
		configOptions.redirectUri = configOptions.redirectUri || pathDefault;
		configOptions.postLogoutRedirectUri = configOptions.postLogoutRedirectUri || pathDefault;

		// create instance with given config
		this.adal = new AdalWrapper(configOptions);
	}

	login() {
		//if (!this.oAuthData.isAuthenticated && !this.adal.isCallback(window.location.hash)) {
			
		//}
		var callbackFunctionName = "oAuthCallback";

		window[callbackFunctionName] = (hash: string, win: Window) => {
			// Check For & Handle Redirect From AAD After Login
			if (hash) {
				this.adal.handleWindowCallbackNew(win, !!(win.opener));

				// loginresource is used to set authenticated status
				this.updateDataFromCache(this.adal.config.loginResource);
			}
		};

		this.adal.login();
	}

	acquireToken(resource): Q.Promise<string> {
		var deferred = Q.defer<string>();
		this.adal.acquireToken(resource, function (error, tokenOut) {
			if (error) {
				deferred.reject(error);
			} else {
				deferred.resolve(tokenOut);
			}
		});
		return deferred.promise;
	}
}
export = AdalAuthenticationService;