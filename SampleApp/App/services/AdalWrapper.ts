class AdalWrapper extends AuthenticationContext {
    constructor(options: any) {
        super(options);
    }

    getRequestInfoNew(hash, win: Window) {
        hash = super._getHash(hash);
        var parameters = super._deserialize(hash);
        var requestInfo = {
            valid: false,
            parameters: {},
            stateMatch: false,
            stateResponse: '',
            requestType: this.REQUEST_TYPE.UNKNOWN
        };
        if (parameters) {
            requestInfo.parameters = parameters;
            if (parameters.hasOwnProperty(this.CONSTANTS.ERROR_DESCRIPTION) ||
                parameters.hasOwnProperty(this.CONSTANTS.ACCESS_TOKEN) ||
                parameters.hasOwnProperty(this.CONSTANTS.ID_TOKEN)) {

                requestInfo.valid = true;

                // which call
                var stateResponse = '';
                if (parameters.hasOwnProperty('state')) {
                    super._logstatus('State: ' + parameters.state);
                    stateResponse = parameters.state;
                } else {
                    super._logstatus('No state returned');
                }

                requestInfo.stateResponse = stateResponse;

                // async calls can fire iframe and login request at the same time if developer does not use the API as expected
                // incoming callback needs to be looked up to find the request type
                switch (stateResponse) {
                    case super._getItem(this.CONSTANTS.STORAGE.STATE_LOGIN):
                        requestInfo.requestType = this.REQUEST_TYPE.LOGIN;
                        requestInfo.stateMatch = true;
                        break;

                    case super._getItem(this.CONSTANTS.STORAGE.STATE_IDTOKEN):
                        requestInfo.requestType = this.REQUEST_TYPE.ID_TOKEN;
                        super._saveItem(this.CONSTANTS.STORAGE.STATE_IDTOKEN, '');
                        requestInfo.stateMatch = true;
                        break;
                }

                // external api requests may have many renewtoken requests for different resource          
                if (!requestInfo.stateMatch && win.parent && win.parent.AuthenticationContext()) {
                    var statesInParentContext = win.parent.AuthenticationContext()._renewStates;
                    for (var i = 0; i < statesInParentContext.length; i++) {
                        if (statesInParentContext[i] === requestInfo.stateResponse) {
                            requestInfo.requestType = this.REQUEST_TYPE.RENEW_TOKEN;
                            requestInfo.stateMatch = true;
                            break;
                        }
                    }
                }
            }
        }

        return requestInfo;
    }

    handleWindowCallbackNew(win: Window, isPopup: boolean) {
        // This is for regular javascript usage for redirect handling
        // need to make sure this is for callback
        var hash = win.location.hash;
        if (this.isCallback(hash)) {
            var requestInfo = this.getRequestInfoNew(hash, win);
            super._log(super._getResourceFromState(requestInfo.stateResponse), 'returned from redirect url');
            this.saveTokenFromHash(requestInfo);
            var callback = null;
            if ((requestInfo.requestType === this.REQUEST_TYPE.RENEW_TOKEN ||
                requestInfo.requestType === this.REQUEST_TYPE.ID_TOKEN) &&
                !isPopup) {
                // iframe call but same single page
                super._logstatus('Window is in iframe');
                callback = win.parent.callBackMappedToRenewStates[requestInfo.stateResponse];
                win.src = '';
            } else if (win && win.oauth2Callback) {
                super._logstatus('Window is redirecting');
                callback = this.callback;
            }

            //window.location.hash = '';
            //window.location = super._getItem(this.CONSTANTS.STORAGE.LOGIN_REQUEST);
            if (requestInfo.requestType === this.REQUEST_TYPE.RENEW_TOKEN) {
                callback(super._getItem(this.CONSTANTS.STORAGE.ERROR_DESCRIPTION), requestInfo.parameters[this.CONSTANTS.ACCESS_TOKEN] || requestInfo.parameters[this.CONSTANTS.ID_TOKEN]);
                return;
            } else if (requestInfo.requestType === this.REQUEST_TYPE.ID_TOKEN) {
                // JS context may not have the user if callback page was different, so parse idtoken again to callback
                callback(super._getItem(this.CONSTANTS.STORAGE.ERROR_DESCRIPTION), super._createUser(super._getItem(this.CONSTANTS.STORAGE.IDTOKEN)));
                return;
            }
        }
    }    
}
export = AdalWrapper;