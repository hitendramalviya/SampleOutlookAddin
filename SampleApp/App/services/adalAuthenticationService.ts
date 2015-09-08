
class adalAuthenticationService {
	static AuthConfig: any;
	static AuthContect: adalAuthenticationService;
	adal: any = null;
	private oauthData = { isAuthenticated: false, userName: '', loginError: '', profile: '' };

	private updateDataFromCache(resource) {
		// only cache lookup here to not interrupt with events
		var token = this.adal.getCachedToken(resource);
		this.oauthData.isAuthenticated = token !== null && token.length > 0;
		var user = this.adal.getCachedUser() || { userName: '' };
		this.oauthData.userName = user.userName;
		this.oauthData.profile = user.profile;
		this.oauthData.loginError = this.adal.getLoginError();
	}

	init(configOptions) {
		if (configOptions) {
			// redirect and logout_redirect are set to current location by default
			var existingHash = window.location.hash;
			var pathDefault = window.location.href;
			if (existingHash) {
				pathDefault = pathDefault.replace(existingHash, '');
			}
			configOptions.redirectUri = configOptions.redirectUri || pathDefault;
			configOptions.postLogoutRedirectUri = configOptions.postLogoutRedirectUri || pathDefault;

			//if (httpProvider && httpProvider.interceptors) {
			//	httpProvider.interceptors.push('ProtectedResourceInterceptor');
			//}

			// create instance with given config
			this.adal = new AuthenticationContext(configOptions);
		} else {
			throw new Error('You must set configOptions, when calling init');
		}

		// Check For & Handle Redirect From AAD After Login
		var isCallback = this.adal.isCallback(window.location.hash);
		this.adal.handleWindowCallback();

		if (isCallback && !this.adal.getLoginError()) {
			window.location = this.adal._getItem(this.adal.CONSTANTS.STORAGE.LOGIN_REQUEST);
		}

		// loginresource is used to set authenticated status
		this.updateDataFromCache(this.adal.config.loginResource);

		adalAuthenticationService.AuthContect = this;
		adalAuthenticationService.AuthConfig = configOptions;
	}

	login() {
		if (!this.oauthData.isAuthenticated && !this.adal.isCallback(window.location.hash)) {
			this.adal.login();
		}
	}
}
export = adalAuthenticationService;