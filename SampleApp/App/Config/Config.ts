class Config {
	adalConfig: any;
	constructor(config: any) {
		//to do to read adal config from config variable
		this.adalConfig = {
			tenant: "gecko.no",
			'clientId': '3791c89b-4c16-4c18-b996-2fdb7588451d',
			'cacheLocation': 'localStorage',
			"endpoints": {
				'https://outlook.office.com/api/v1.0/me': 'https://outlook.office.com'
			},
			"displayCall": function (url) {
				var ref = window.open(url, '_blank', 'location=no,hidden=no');
				ref.addEventListener('loadstart', function (event) {
					var redirectUri = "https://localhost:44300/";
					if ((event.url).indexOf(redirectUri) === 0) {
						ref.close();
						window.location.href = event.url.replace(redirectUri);
						window.location.reload();
					}
				});
			}
		};
	}
}
export = Config;