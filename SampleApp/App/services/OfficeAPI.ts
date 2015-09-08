import http = require("plugins/http");
import adalAuthenticationService = require("services/adalAuthenticationService");

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

	//getAttachmentDetails(attachment: Office.AttachmentDetails): Q.Promise<any> {
	//	var deferred = Q.defer<any>();
	//	var mailBox = Office.context.mailbox;
	//	this.setServiceToken().then(token => {
	//		var request = {
	//			token: token,
	//			ewsUrl: mailBox.ewsUrl,
	//			attachment: this.getAttachemntObject(attachment),
	//			documentServiceUrl: "",
	//			documentServiceToken: ""
	//		};
	//		http.post(appConfig.absoluteUrl + "OfficeApi/ProcessMailAttachment", request)
	//			.then((response) => {
	//				deferred.resolve(response);
	//			}).fail((jqXHR, textStatus, errorThrown) => {
	//				deferred.reject(errorThrown);
	//			});
	//	});
	//	return deferred.promise;
	//}

	getAttachmentDetails(attachment: Office.AttachmentDetails): Q.Promise<any> {
		adalAuthenticationService.AuthContect.adal.acquireToken("https://outlook.office.com", (error, token) => {

		});
		var deferred = Q.defer<any>();
		var mailBox = Office.context.mailbox;
		var token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSIsImtpZCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzhlOGU3YmYwLWUxOTgtNDFhNi04OTc2LTJiZDNhOWNhMGJhZi8iLCJpYXQiOjE0NDE3MTkwNTQsIm5iZiI6MTQ0MTcxOTA1NCwiZXhwIjoxNDQxNzIyOTU0LCJ2ZXIiOiIxLjAiLCJ0aWQiOiI4ZThlN2JmMC1lMTk4LTQxYTYtODk3Ni0yYmQzYTljYTBiYWYiLCJvaWQiOiI1MWYyYzA4NC03N2QyLTRhY2YtYWVlNS0xZjM1MzY2MGI3MmUiLCJ1cG4iOiJoa21AZ2Vja28ubm8iLCJwdWlkIjoiMTAwMzAwMDA4QjA3NjQ1MyIsInN1YiI6IldJU0tKZ3VsN2wyMHluUGQ2UUhUa3FFbDc0THdma3dtcHRjUV96eVc3M3MiLCJnaXZlbl9uYW1lIjoiSGl0ZW5kcmEgS3VtYXIiLCJmYW1pbHlfbmFtZSI6Ik1hbHZpeWEiLCJuYW1lIjoiSGl0ZW5kcmEgS3VtYXIgTWFsdml5YSIsImFtciI6WyJwd2QiXSwidW5pcXVlX25hbWUiOiJoa21AZ2Vja28ubm8iLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMTk5Mzk2Mjc2My04NTQyNDUzOTgtMTcwODUzNzc2OC04MjM2IiwiYXBwaWQiOiIzNzkxYzg5Yi00YzE2LTRjMTgtYjk5Ni0yZmRiNzU4ODQ1MWQiLCJhcHBpZGFjciI6IjAiLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBNYWlsLlJlYWRXcml0ZSIsImFjciI6IjEiLCJpcGFkZHIiOiIxMTEuOTMuMTI4LjI2In0.Kkyu5FjyCNNSkxLVQMEf1FrUz5dSP06ggoUJh-6hsXBzXgy3DRisYI3VH3LjqWISmqbdkVjVwpvOZlDFTYofAhAaDmO_FUCuFV6h29msGsbHws69Pr7R8F7nKuOGoXH-IjJmARGJ177RYvKIS8SeHgkGUrypEB6zzYVOjTHVVvz-k5Qvv3pdEdy53tFX1qKbLTqVUjj8qA0TcrYWR1htEcVKqZH3kYfxUcDe1ORm8TlReFfEhlX2IHygE0x5P6DgrexumOzSD5R_QTbMpCoqixBiuzGaTOl6_Im4VVl_nrwEeWirM-pxlvVTDN_V21g3jVEIIlMi9NvEciAI-SOBJg";
		var url = "https://outlook.office.com/api/v1.0/me/messages/" + this.cleanId(this.messageRead.itemId) + "/attachments/" + this.cleanId(attachment.id);
				var request = {
				token: token,
				ewsUrl: url,
				attachment: this.getAttachemntObject(attachment),
				documentServiceUrl: "",
				documentServiceToken: ""
		};
			http.post(appConfig.absoluteUrl + "OfficeApi/ProcessMailAttachment", request)
				.then((response) => {
					deferred.resolve(response);
				}).fail((jqXHR, textStatus, errorThrown) => {
					deferred.reject(errorThrown);
				});
		//http.get(url, null, {
		//	"Authorization": "Bearer " + token
		//})
		//	.then((response) => {
		//		deferred.resolve(response);
		//	}).fail((jqXHR, textStatus, errorThrown) => {
		//		deferred.reject(errorThrown);
		//	});
		//adalAuthenticationService.AuthContect.adal.acquireToken("https://outlook.office.com", (error, token) => {
		//	if (error || !token) {
		//		deferred.reject(error);
		//	}
		//	var url = "https://outlook.office.com/api/v1.0/me/messages/" + this.cleanId(this.messageRead.itemId) + "/attachments/" + this.cleanId(attachmentId);
		//	http.get(url, null, {
		//		"Authorization": "Bearer " + token
		//	})
		//		.then((response) => {
		//			deferred.resolve(response);
		//		}).fail((jqXHR, textStatus, errorThrown) => {
		//			deferred.reject(errorThrown);
		//		});
		//});
		return deferred.promise;
	}

	private cleanId(id:any) : any {
        return id.replace(/\+/g, "_").replace(/\//g, "-");
    }

	private getAttachemntObject(attachment: Office.AttachmentDetails): any {
		return {
			attachmentType: attachment.attachmentType,
			contentType: attachment.contentType,
			id: attachment.id,
			isInline: attachment.isInline,
			name: attachment.name,
			size: attachment.size
		}
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