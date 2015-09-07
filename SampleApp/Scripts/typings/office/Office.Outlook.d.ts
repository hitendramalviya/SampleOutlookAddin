// Based on outlook-15.js build 15.0.4128.1005

declare module Office.MailboxEnums {

	export enum BodyType {
		/**
		 * The body is in HTML format
		 */
		HTML,
		/**
		 * The body is in text format
		 */
		text
	}

	export enum EntityType { 
		/**
		 * Specifies that the entity is a meeting suggestion
		 */
		MeetingSuggestion,
		/**
		 * Specifies that the entity is a task suggestion
		 */
		TaskSuggestion,
		/**
		 * Specifies that the entity is a postal address
		 */
		Address,
		/**
		 * Specifies that the entity is SMTP email address
		 */
		EmailAddress,
		/**
		 * Specifies that the entity is an Internet URL
		 */
		Url,
		/**
		 * Specifies that the entity is US phone number
		 */
		PhoneNumber,
		/**
		 * Specifies that the entity is a contact
		 */
		Contact
	}

	export enum ItemType {
		/**
		 * A meeting request, response, or cancellation
		 */
		Message,
		/**
		 * Specifies an appointment item
		 */
		Appointment
	}

	export enum ResponseType { 
		/**
		 * There has been no response from the attendee
		 */
		None,
		/**
		 * The attendee is the meeting organizer
		 */
		Organizer,
		/**
		 * The meeting request was tentatively accepted by the attendee
		 */
		Tentative,
		/**
		 * The meeting request was accepted by the attendee
		 */
		Accepted,
		/**
		 * The meeting request was declined by the attendee
		 */
		Declined
	}

	export enum RecipientType {
		/**
		 * Specifies that the recipient is not one of the other recipient types
		 */
		Other,
		/**
		 * Specifies that the recipient is a distribution list containing a list of email addresses
		 */
		DistributionList,
		/**
		 * Specifies that the recipient is an SMTP email address that is on the Exchange server
		 */
		User,
		/**
		 * Specifies that the recipient is an SMTP email address that is not on the Exchange server
		 */
		ExternalUser
	}

	export enum AttachmentType {
		/**
		 * The attachment is a file
		 */
		File,
		/**
		 * The attachment is an Exchange item
		 */
		Item
	}
}

declare module Office {
	export module Types {
		export interface ItemRead extends Office.Item {
			subject: any;

			/**
			 * Displays a reply form that includes the sender and all the recipients of the selected message
			 * @param htmlBody A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
			 */
			displayReplyAllForm(htmlBody: string): void;
			/**
			 * Displays a reply form that includes only the sender of the selected message
			 * @param htmlBody A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
			 */
			displayReplyForm(htmlBody: string): void;
			/**
			 * Gets an array of entities found in an message
			 */
			getEntities(): Office.Entities;
			/**
			 * Gets an array of entities of the specified entity type found in an message
			 * @param entityType One of the EntityType enumeration values
			 */
			getEntitiesByType(entityType: Office.MailboxEnums.EntityType): Office.Entities;
			/**
			 * Returns well-known entities that pass the named filter defined in the manifest XML file
			 * @param name  A TableData object with the headers and rows 
			 */
			getFilteredEntitiesByName(name: string): Office.Entities;
			/**
			 * Returns string values in the currently selected message object that match the regular expressions defined in the manifest XML file
			 */
			getRegExMatches(): string[];
			/**
			 * Returns string values that match the named regular expression defined in the manifest XML file
			 */
			getRegExMatchesByName(name: string): string[];
		}

		export interface ItemCompose extends Office.Item {
			body: Office.Body;
			subject: any;

			/**
			 * Adds a file to a message as an attachment
			 * @param uri The URI that provides the location of the file to attach to the message. The maximum length is 2048 characters
			 * @param attachmentName The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			addFileAttachmentAsync(uri: string, attachmentName: string, options?: any, callback?: (result: AsyncResult) => void): void;
			/**
			 * Adds an Exchange item, such as a message, as an attachment to the message
			 * @param itemId The Exchange identifier of the item to attach. The maximum length is 100 characters
			 * @param attachmentName The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			addItemAttachmentAsync(itemId: any, attachmentName: string, options?: any, callback?: (result: AsyncResult) => void): void;
			/**
			 * Removes an attachment from a message
			 * @param attachmentIndex The index of the attachment to remove. The maximum length of the string is 100 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			removeAttachmentAsync(attachmentIndex: string, option?: any, callback?: (result: AsyncResult) => void): void;
		}

		export interface MessageCompose extends Office.Message {
			attachments: Office.AttachmentDetails[];
			body: Office.Body;
			bcc: Office.Recipients;
			cc: Office.Recipients;
			subject: Office.Subject;
			to: Office.Recipients;

			/**
			 * Adds a file to a message as an attachment
			 * @param uri The URI that provides the location of the file to attach to the message. The maximum length is 2048 characters
			 * @param attachmentName The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			addFileAttachmentAsync(uri: string, attachmentName: string, options?: any, callback?: (result: AsyncResult) => void): void;
			/**
			 * Adds an Exchange item, such as a message, as an attachment to the message
			 * @param itemId The Exchange identifier of the item to attach. The maximum length is 100 characters
			 * @param attachmentName The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			addItemAttachmentAsync(itemId: any, attachmentName: string, options?: any, callback?: (result: AsyncResult) => void): void;
			/**
			 * Removes an attachment from a message
			 * @param attachmentIndex The index of the attachment to remove. The maximum length of the string is 100 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			removeAttachmentAsync(attachmentIndex: string, option?: any, callback?: (result: AsyncResult) => void): void;
		}

		export interface MessageRead extends Office.Message {
			attachments: Office.AttachmentDetails[];
			cc: Office.EmailAddressDetails[];
			from: Office.EmailAddressDetails;
			internetMessageId: string;
			normalizedSubject: string;
			sender: Office.EmailAddressDetails;
			subject: string;
			to: Office.EmailAddressDetails[];

			/**
			 * Displays a reply form that includes the sender and all the recipients of the selected message
			 * @param htmlBody A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
			 */
			displayReplyAllForm(htmlBody: string): void;
			/**
			 * Displays a reply form that includes only the sender of the selected message
			 * @param htmlBody A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
			 */
			displayReplyForm(htmlBody: string): void;
			/**
			 * Gets an array of entities found in an message
			 */
			getEntities(): Office.Entities;
			/**
			 * Gets an array of entities of the specified entity type found in an message
			 * @param entityType One of the EntityType enumeration values
			 */
			getEntitiesByType(entityType: Office.MailboxEnums.EntityType): Office.Entities;
			/**
			 * Returns well-known entities that pass the named filter defined in the manifest XML file
			 * @param name  A TableData object with the headers and rows 
			 */
			getFilteredEntitiesByName(name: string): Office.Entities;
			/**
			 * Returns string values in the currently selected message object that match the regular expressions defined in the manifest XML file
			 */
			getRegExMatches(): string[];
			/**
			 * Returns string values that match the named regular expression defined in the manifest XML file
			 */
			getRegExMatchesByName(name: string): string[];
		}

		export interface AppointmentCompose extends Office.Appointment {
			body: Office.Body;
			end: Office.Time;
			location: Office.Location;
			optionalAttendees: Office.Recipients;
			requiredAttendees: Office.Recipients;
			start: Office.Time;
			subject: Office.Subject;

			/**
			 * Adds a file to an appointment as an attachment
			 * @param uri The URI that provides the location of the file to attach to the appointment. The maximum length is 2048 characters
			 * @param attachmentName The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			addFileAttachmentAsync(uri: string, attachmentName: string, options?: any, callback?: (result: AsyncResult) => void): void;
			/**
			 * Adds an Exchange item, such as a message, as an attachment to the appointment
			 * @param itemId The Exchange identifier of the item to attach. The maximum length is 100 characters
			 * @param attachmentName The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			addItemAttachmentAsync(itemId: any, attachmentName: string, options?: any, callback?: (result: AsyncResult) => void): void;
			/**
			 * Removes an attachment from a appointment
			 * @param attachmentIndex The index of the attachment to remove. The maximum length of the string is 100 characters
			 * @param options Any optional parameters or state data passed to the method
			 * @param callback The optional callback method
			 */
			removeAttachmentAsync(attachmentIndex: string, option?: any, callback?: (result: AsyncResult) => void): void;
		}

		export interface AppointmentRead extends Office.Appointment {
			attachments: Office.AttachmentDetails[];
			end: Date;
			location: string;
			normalizedSubject: string;
			optionalAttendees: Office.EmailAddressDetails;
			organizer: Office.EmailAddressDetails;
			requiredAttendees: Office.EmailAddressDetails;
			resources: string[];
			start: Date;
			subject: string;

			/**
			 * Displays a reply form that includes the organizer and all the attendees of the selected appointment item
			 * @param htmlBody A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
			 */
			displayReplyAllForm(htmlBody: string): void;
			/**
			 * Displays a reply form that includes only the organizer of the selected appointment item
			 * @param htmlBody A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
			 */
			displayReplyForm(htmlBody: string): void;
			/**
			 * Gets an array of entities found in an appointment
			 */
			getEntities(): Office.Entities;
			/**
			 * Gets an array of entities of the specified entity type found in an appointment
			 * @param entityType One of the EntityType enumeration values
			 */
			getEntitiesByType(entityType: Office.MailboxEnums.EntityType): Office.Entities;
			/**
			 * Returns well-known entities that pass the named filter defined in the manifest XML file
			 * @param name  A TableData object with the headers and rows 
			 */
			getFilteredEntitiesByName(name: string): Office.Entities;
			/**
			 * Returns string values in the currently selected appointment object that match the regular expressions defined in the manifest XML file
			 */
			getRegExMatches(): string[];
			/**
			 * Returns string values that match the named regular expression defined in the manifest XML file
			 */
			getRegExMatchesByName(name: string): string[];
		}
	}

	export var cast: CastHelper;

	export interface CastHelper {
		item: ItemCastHelper;
	}

	export interface ItemCastHelper {
		toAppointmentCompose(item: Office.Item): Office.Types.AppointmentCompose;
		toAppointmentRead(item: Office.Item): Office.Types.AppointmentRead;
		toAppointment(item: Office.Item): Office.Appointment;
		toMessageCompose(item: Office.Item): Office.Types.MessageCompose;
		toMessageRead(item: Office.Item): Office.Types.MessageRead;
		toMessage(item: Office.Item): Office.Message;
		toItemCompose(item: Office.Item): Office.Types.ItemCompose;
		toItemRead(item: Office.Item): Office.Types.ItemRead;
	}

	export interface AttachmentDetails {
		attachmentType: Office.MailboxEnums.AttachmentType;
		contentType: string;
		id: string;
		isInline: boolean;
		name: string;
		size: number;
	}

	export interface Contact {
		personName: string;
		businessName: string;
		phoneNumbers: PhoneNumber[];
		emailAddresses: string[];
		urls: string[];
		addresses: string[];
		contactString: string;
	}
	
	export interface Context {
		mailbox: Mailbox;
		roamingSettings: RoamingSettings;
	}

	export interface CustomProperties {
		/**
		 * Returns the value of the specified custom property
		 * @param name The name of the property to be returned
		 */
		get(name: string): any;
		/**
		 * Sets the specified property to the specified value
		 * @param name The name of the property to be set
		 * @param value The value of the property to be set
		 */
		set(name: string, value: string): void;
		/**
		 * Removes the specified property from the custom property collection.
		 * @param name The name of the property to be removed
		 */
		remove(name: string): void;
		/**
		 * Saves the custom property collection to the server
		 * @param callback The optional callback method
		 * @param userContext Optional variable for any state data that is passed to the saveAsync method
		 */
		saveAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
	}

	export interface EmailAddressDetails {
		emailAddress: string;
		displayName: string;
		appointmentResponse: Office.MailboxEnums.ResponseType;
		recipientType: Office.MailboxEnums.RecipientType;
	}

	export interface EmailUser {
		name: string;
		userId: string;
	}

	export interface Entities {
		addresses: string[];
		taskSuggestions: string[];
		meetingSuggestions: MeetingSuggestion[];
		emailAddresses: string[];
		urls: string[];
		phoneNumbers: PhoneNumber[];
		contacts: Contact[];
	}

	export interface Item {
		dateTimeCreated: Date;
		dateTimeModified: Date;
		itemClass: string;
		itemId: string;
		itemType: Office.MailboxEnums.ItemType;

		/**
		 * Asynchronously loads custom properties that are specific to the item and a app for Office
		 * @param callback The optional callback method
		 * @param userContext Optional variable for any state data that is passed to the asynchronous method
		 */
		loadCustomPropertiesAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
	}

	export interface Appointment extends Item {
	}

	export interface Body {
		/**
		 * Gets a value that indicates whether the content is in HTML or text format
		 * @param tableData  A TableData object with the headers and rows 
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the getTypeAsync method returns
		 */
		getTypeAsync(options?: any, callback?: (result: AsyncResult) => void): void;
		/**
		 * Adds the specified content to the beginning of the item body
		 * @param data The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		prependAsync(data: string, options?: any, callback?: (result: AsyncResult) => void): void;
		/**
		 * Replaces the selection in the body with the specified text
		 * @param data The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		setSelectedDataAsync(data: string, options?: any, callback?: (result: AsyncResult) => void): void;
	}

	export interface Location {
		/**
		 * Begins an asynchronous request for the location of an appointment
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		getAsync(options?: any, callback?: (result: AsyncResult) => void): void;
		/**
		 * Begins an asynchronous request to set the location of an appointment
		 * @param data The location of the appointment. The string is limited to 255 characters
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the location is set
		 */
		setAsync(location: string, options?: any, callback?: (result: AsyncResult) => void): void;
	}

	export interface Mailbox {
		item: Item;
		userProfile: UserProfile;

		/**
		 * Gets a Date object from a dictionary containing time information
		 * @param timeValue A Date object
		 */
		convertToLocalClientTime(timeValue: Date): any;
		/**
		 * Gets a dictionary containing time information in local client time
		 * @param input A dictionary containing a date. The dictionary should contain the following fields: year, month, date, hours, minutes, seconds, time zone, time zone offset
		 */
		convertToUtcClientTime(input: any): Date;
		/**
		 * Displays an existing calendar appointment
		 * @param itemId The Exchange Web Services (EWS) identifier for an existing calendar appointment
		 */
		displayAppointmentForm(itemId: any): void;
		/**
		 * Displays an existing message
		 * @param itemId The Exchange Web Services (EWS) identifier for an existing message
		 */
		displayMessageForm(itemId: any): void;
		/**
		 * Displays a form for creating a new calendar appointment
		 * @param requiredAttendees An array of strings containing the email addresses or an array containing an EmailAddressDetails object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries
		 * @param optionalAttendees An array of strings containing the email addresses or an array containing an EmailAddressDetails object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries
		 * @param start A Date object specifying the start date and time of the appointment
		 * @param end A Date object specifying the end date and time of the appointment
		 * @param location A string containing the location of the appointment. The string is limited to a maximum of 255 characters
		 * @param resources An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries
		 * @param subject A string containing the subject of the appointment. The string is limited to a maximum of 255 characters
		 * @param body The body of the appointment message. The body content is limited to a maximum size of 32 KB
		 */
		displayNewAppointmentForm(requiredAttendees: any, optionalAttendees: any, start: Date, end: Date, location: string, resources: string[], subject: string, body: string): void;
		/**
		 * Gets a string that contains a token used to get an attachment or item from an Exchange Server
		 * @param callback The optional method to call when the string is inserted
		 * @param userContext Optional variable for any state data that is passed to the asynchronous method
		 */
		getCallbackTokenAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
		/**
		 * Gets a token identifying the user and the app for Office
		 * @param callback The optional method to call when the string is inserted
		 * @param userContext Optional variable for any state data that is passed to the asynchronous method
		 */
		getUserIdentityTokenAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
		/**
		 * Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox
		 * @param data The EWS request
		 * @param callback The optional method to call when the string is inserted
		 * @param userContext Optional variable for any state data that is passed to the asynchronous method
		 */
		makeEwsRequestAsync(data: any, callback?: (result: AsyncResult) => void, userContext?: any): void;

		GetIsRead(): boolean;

		ewsUrl: string;
	}

	export interface Message extends Item {
		conversationId: string;
	}

	export interface MeetingRequest extends Message {
		start: Date;
		end: Date;
		location: string;
		optionalAttendees: EmailAddressDetails[];
		requiredAttendees: EmailAddressDetails[];
	}

	export interface MeetingSuggestion {
		meetingString: string;
		attendees: EmailAddressDetails[];
		location: string;
		subject: string;
		start: Date;
		end: Date;
	}

	export interface PhoneNumber {
		phoneString: string;
		originalPhoneString: string;
		type: string;        
	}

	export interface Recipients {
		/**
		 * Begins an asynchronous request to add a recipient list to an appointment or message
		 * @param recipients The recipients to add to the recipients list
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		addAsync(recipients: any, options?: any, callback?: (result: AsyncResult) => void): void;
		/**
		 * Begins an asynchronous request to get the recipient list for an appointment or message
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		getAsync(options?: any, callback?: (result: AsyncResult) => void): void;
		/**
		 * Begins an asynchronous request to set the recipient list for an appointment or message
		 * @param recipients The recipients to add to the recipients list
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		setAsync(recipients: any, options?: any, callback?: (result: AsyncResult) => void): void;
	}

	export interface RoamingSettings {
		/**
		 * Retrieves the specified setting
		 * @param name The case-sensitive name of the setting to retrieve
		 */
		get(name: string): any;
		/**
		 * Removes the specified setting
		 * @param name The case-sensitive name of the setting to remove
		 */
		remove(name: string): void;
		/**
		 * Saves the settings
		 * @param callback A function that is invoked when the callback returns, whose only parameter is of type AsyncResult
		 */
		saveAsync(callback?: (result: AsyncResult) => void): void;
		/**
		 * Sets or creates the specified setting
		 * @param name The case-sensitive name of the setting to set or create
		 * @param value Specifies the value to be stored
		 */
		set(name: string, value: any): void;
	}

	export interface Subject {
		/**
		 * Begins an asynchronous request to get the subject of an appointment or message
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		getAsync(options?: any, callback?: (result: AsyncResult) => void): void;
		/**
		 * Begins an asynchronous call to set the subject of an appointment or message
		 * @param data The subject of the appointment. The string is limited to 255 characters
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		setAsync(data: string, options?: any, callback?: (result: AsyncResult) => void): void;
	}

	export interface TaskSuggestion {
		assignees: EmailUser[];
		taskString: string;
	}

	export interface Time {
		/**
		 * Begins an asynchronous request to get the start or end time
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		getAsync(options?: any, callback?: (result: AsyncResult) => void): void;
		/**
		 * Begins an asynchronous request to set the start or end time
		 * @param dateTime A date-time object in Coordinated Universal Time (UTC)
		 * @param options Any optional parameters or state data passed to the method
		 * @param callback The optional method to call when the string is inserted
		 */
		setAsync(dateTime: Date, options?: any, callback?: (result: AsyncResult) => void): void;
	}

	export interface UserProfile {
		displayName: string;
		emailAddress: string;
		timeZone: string;
	}
}