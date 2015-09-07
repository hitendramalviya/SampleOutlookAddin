class OfficeAPI {
	isReceivedItem: boolean;
	userProfile: Office.UserProfile;
	subject: string;
	sender: Office.EmailAddressDetails;
	toRecipients: Office.EmailAddressDetails[] = [];
	ccRecipients: Office.EmailAddressDetails[] = [];
	attachments: Office.AttachmentDetails[] = [];

	private messageRead: Office.Types.MessageRead;
	private ewsServiceToken: string;
	constructor() {
		this.userProfile = Office.context.mailbox.userProfile;
		this.messageRead = Office.cast.item.toMessageRead(Office.context.mailbox.item);
		this.isReceivedItem = this.userProfile.emailAddress === this.messageRead.from.emailAddress;

		//Read mail item properties
		this.subject = this.messageRead.subject;
		this.sender = this.messageRead.sender;
		this.toRecipients = this.messageRead.to;
		this.ccRecipients = this.messageRead.cc;
		this.attachments = this.messageRead.attachments;
	}

	getBodyContent(): Q.Promise<string> {
		var deferred = Q.defer<string>();
		var mailBox = Office.context.mailbox;
		mailBox.makeEwsRequestAsync(OfficeAPI.getBodyRequest(mailBox.item.itemId), (result: Office.AsyncResult) => {
			if (result.status === Office.AsyncResultStatus.Succeeded) {
				var response = jQuery.parseXML(result.value);
				//HKM: used jquery selector below to make it cross browser supported (instead of response.getElementsByTagName("t:Body"))
				var body = jQuery(response).find("Body:gt(0)");
				if (body.length > 0) {
					deferred.resolve(body[0].textContent);
				}
			}
		});
		return deferred.promise;
	}

	private static getBodyRequest(id: string): string {
		// Return a GetItem operation request for the body of the specified item. 
		var result = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
			"<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" " +
			"xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " +
			"xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" " +
			"xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"> " +
			"<soap:Header>" +
			"<t:RequestServerVersion Version=\"Exchange2013\" />" +
			"</soap:Header>" +
			"<soap:Body>" +
			"<m:GetItem>" +
			"<m:ItemShape>" +
			"<t:BaseShape>IdOnly</t:BaseShape>" +
			"<t:AdditionalProperties>" +
			"<t:FieldURI FieldURI=\"item:Body\" />" +
			"</t:AdditionalProperties>" +
			"</m:ItemShape>" +
			"<m:ItemIds>" +
			"<t:ItemId Id=\"" + id + "\" />" +
			"</m:ItemIds>" +
			"</m:GetItem>" +
			"</soap:Body>" +
			"</soap:Envelope>";
		return result;
	}

	getAttacmentContent(): Q.Promise<any> {
		var deferred = Q.defer<any>();
		var mailBox = Office.context.mailbox;
		this.setServiceToken().then(token => {
			var request = {
				token: token,
				ewsUrl: mailBox.ewsUrl,
				attachments: [],
				documentServiceUrl: "",
				documentServiceToken: ""
			};
			_.each(this.attachments, attachment => {
				request.attachments.push({
					attachmentType: attachment.attachmentType,
					contentType: attachment.contentType,
					id: attachment.id,
					isInline: attachment.isInline,
					name: attachment.name,
					size: attachment.size
				});
			});
			var ajaxOptions = {
                url: appConfig.absoluteUrl + "OfficeApi/ProcessMailAttachment",
				type: "post",
				data: request,
				contentType: "text/xml"
			};
			$.ajax(ajaxOptions).done(function (data) {
				deferred.resolve(data);
			}).fail(function (jqXHR, textStatus, errorThrown) {
				deferred.reject(errorThrown);
			});
		});
		return deferred.promise;
	}

	private setServiceToken(): Q.Promise<string> {
		var deferred = Q.defer<string>();
		if (this.ewsServiceToken) {
			deferred.resolve(this.ewsServiceToken);
		}
		else {
			Office.context.mailbox.getCallbackTokenAsync((result: Office.AsyncResult) => {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					this.ewsServiceToken = result.value;
					deferred.resolve(this.ewsServiceToken);
				}
			});
		}
		return deferred.promise;
	}
}

export = OfficeAPI;