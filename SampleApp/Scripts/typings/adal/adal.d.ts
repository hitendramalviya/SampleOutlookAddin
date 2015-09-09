declare class AuthenticationContext {
    constructor(config: any);
    isCallback(hash: string): boolean;
    acquireToken(resourceUrl: string, callback: (error: string, token: string) => void);
    _getHash(hash: string): string;
    _deserialize(hash: string): any;
    _logstatus(status: string): void;
    _getItem(key: any): any;
    _saveItem(key: any, obj: any): void;
    _log(resource: any, message: string): void;
    _getResourceFromState(state: string);
    _createUser(idToken: any);
    saveTokenFromHash(requestedInfo: any);

    callback: any;
    //Constants
    REQUEST_TYPE: any;
    CONSTANTS: any;
}

interface Window {
    AuthenticationContext(): any; 
    src: string; //for iframe
    callBackMappedToRenewStates: any;
    oauth2Callback: any;
}