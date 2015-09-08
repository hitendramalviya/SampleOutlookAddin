import http = require("plugins/http");
import app = require("durandal/app");
import ko = require("knockout");
import OfficeAPI = require("services/OfficeAPI");
import adalAuthenticationService = require("services/adalAuthenticationService");
import Config = require("Config/Config");

class Attachment {
	displayName: string = "Attachemtns";
	attachments: Office.AttachmentDetails[];
	private officeApi: OfficeAPI;

	constructor() {
		this.officeApi = new OfficeAPI();
	}

	activate() {
		this.attachments = this.officeApi.attachments;
	}

	initOutlook(): Q.Promise<OfficeAPI> {
		var deferred = Q.defer<OfficeAPI>();

		Office.initialize = () => {
			deferred.resolve(new OfficeAPI());
		};

		return deferred.promise;
	}

	getFile = (attachment: Office.AttachmentDetails) => { 
		return this.officeApi.getAttachmentDetails(attachment)
			.then((response) => {
				console.log("Attachment received", response);
			});
	}
}

export = Attachment;