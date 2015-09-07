// Based on Office.js build 16.0.2420.1000

declare module Office {
	export var context: Context;
	/**
	 * This method is called after the Office API was loaded.
	 * @param reason Indicates how the app was initialized
	 */
	export function initialize(reason: InitializationReason): void;
	/**
	 * Indicates if the large namespace for objects will be used or not.
	 * @param useShortNamespace  Indicates if 'true' that the short namespace will be used
	 */
	export function useShortNamespace(useShortNamespace: boolean): void;

	// Enumerations

	export enum AsyncResultStatus {
		/**
		 * Operation succeeded
		 */
		Succeeded,
		/**
		 * Operation failed, check error object
		 */
		Failed
	}

	export enum InitializationReason {
		/**
		 * Indicates the app was just inserted in the document
		 */
		Inserted,
		/**
		 * Indicates if the extension already existed in the document
		 */
		DocumentOpened
	}

	// Objects

	export interface AsyncResult {
		asyncContext: any;
		status: AsyncResultStatus;
		error: Error;
		value: any;
	}

	export interface Context {
		contentLanguage: string;
		displayLanguage: string;
		license: string;
	}

	export interface Error {
		message: string;
		name: string;
	}
}