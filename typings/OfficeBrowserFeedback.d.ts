// Type definitions for OfficeBrowserFeedback
// Project: OfficeBrowserFeedback
// Definitions by: ocefeedback@microsoft.com

declare namespace OfficeBrowserFeedback {
	// The init options.
	let initOptions: IInitOptions;

	// Holder for floodgate-related APIs
	let floodgate: IFloodgate;

	// Create a single feedback dialog, where the feedback type is predetermined.
	// launchOptions properties are optional.
	// This actually returns a promise, typecast explicitly to use as promise.
	function singleFeedback(feedbackType: string, launchOptions?: ILaunchOptionsInAppFeedback): void;

	// Create a multi feedback dialog, where the feedback type is selected by the user.
	// launchOptions properties are optional.
	// This actually returns a promise, typecast explicitly to use as promise.
	function multiFeedback(launchOptions?: ILaunchOptionsInAppFeedback): void;

	//Determines if feedback should be enabled or not. This is a method to help the host application 
	//integrating the SDK know if they should hide or show their feedback entry point. 
	function getFeedbackStatus(): FeedbackStatus;

	// With AADC rule, function return true if Feedback and Survey is enabled.
	// Otherwise, return false.
	function isFeedbackSurveyEnabledByAADC(): boolean;

	/**
	 * Enum to indicate possible feedback status values.
	 */
	export type FeedbackStatus = 0 | 1 | 2;

	interface IFloodgate {
		// Init options
		initOptions: IInitOptionsFloodgate;
		// Initialize Floodgate engine
		initialize: () => Promise<any>;
		// Show custom survey
		showCustomSurvey: (survey: ICustomSurvey) => Promise<any>;
		// Start Floodgate engine
		start: () => Promise<any>;
		// Stop Floodgate engine
		stop: () => void;
		// Get Floodgate engine
		getEngine: () => FloodgateEngine;
		// Get a list of setting names and ids
		getSettingIdMap: () => IFloodgateSettingIdMap;
	}

	interface ICustomSurvey {
		// Comment question
		commentQuestion: string;
		// Campaign Id
		campaignId: string;
		// Is the rating component zero based? i.e. is the lowest rating value 0, if not it's assumed to be 1
		isZeroBased: boolean;
		// Prompt no button text
		promptNoButtonText: string;
		// Prompt question
		promptQuestion: string;
		// Prompt yes button text
		promptYesButtonText: string;
		// Rating question
		ratingQuestion: string;
		// Rating values (ascending)
		ratingValuesAscending: string[];
		// Should the email component be shown
		showEmailRequest: boolean;
		// Should the prompt be shown? If not UI jumps straight to the survey screen.
		showPrompt: boolean;
		// Survey type. Feedback = 0, Nps = 1, Psat = 2, Bps = 3, Fps = 4
		surveyType: number;
		// Title
		title: string;
	}

	interface FloodgateEngine {
		// Gets the IActivityListener logging interface for callers that want to log directly rather than through telemetry
		getActivityListener(): IActivityListener;
	}

	interface IContextData {
		// Activity name that was used for passing context 
		activityName: string;
		// Context information that was passed along with activity
		context: unknown;
	}

	interface ISurveyParams {
		// Context information of all the activities that activated a campaign
		contextInfo: IContextData[];
	}

	interface IOnSurveyActivatedCallback {
		/**
		 * Gets called when the survey has been activated. To launch the UI call the launch() method of the launcher.
		 * @param launcher The launcher object.
		 * @param survey Survey object.
		 * @param surveyParams Survey context and other params.
		 */
		onSurveyActivated(launcher: { launch: () => void }, survey?: ISurveyForm, surveyParams?: ISurveyParams): void;
		/**
		 * Allows host to render custom UX based on dynamic web campaign UX
		 * @param uxOptions UX options.
		 */
		onTeachingCampaignRender?(uxOptions: any): Promise<any>;
	}

	interface IFloodgateSettingStorageCallback {
		/**
		 * Gets a collection of items for a given setting
		 * @param setting The setting
		 */
		readSettingList(setting: string): IFloodgateSetting;
		/**
		 * Update or insert an item of a given setting
		 * @param setting The setting
		 * @param itemKey The setting identifier
		 * @param itemValue The setting value
		 */
		upsertSettingItem(setting: string, itemKey: string, itemValue: string): void;
		/**
		 * Delete an item of a given setting
		 * @param setting The setting
		 * @param itemKey The setting item key
		 */
		deleteSettingItem(setting: string, itemKey: string): void;
	}

	interface IFloodgateAuthTokenCallback {
		/**
		 * Gets auth token for a given azure app Id
		 * @param azure app Id
		 */
		getAuthToken?(appId: string): Promise<string>;
		/**
		 * Gets user Puid as getAuthToken fallback
		 */
		getUserPuid?(): Promise<string>;
	}

	interface IFloodgateSetting {
		// A collection settings
		[key: string]: string;
	}

	interface IFloodgateSettingIdMap {
		// A collection of setting ids keyed by names
		[key: string]: number
	}

	/**
	 * Interface for a Survey to allow users access to Survey Components
	 */
	interface ISurveyForm {
		// gets the Survey Component given a Type
		getComponent: (componentType: ISurveyComponent.Type) => ISurveyComponent;
		// gets the back end campaign Id
		getCampaignId: () => string;
		// gets the ClientFeedbackId
		getClientFeedbackId(): string;
		// Gets the launcher type string from the campaign definition
		getLauncherType: () => string;
		// Gets the survey type
		getType: () => ISurvey.Type;
		// submits responses set by the user
		submit: () => void;
	}

	/**
	 * Interface for a SurveyComponent (i.e. a question/widget to show the user
	 * in a survey form, and that typically requires a response value of some kind)
	 */
	interface ISurveyComponent {
		getType: () => ISurveyComponent.Type;
	}

	module ISurvey {
		export type Type =
			0 | // Feedback = 0
			// An NPS (net promoter score) survey. Asks user to rate "whether or not they would recommend this product to family/friends".
			// Contains a prompt, question, and rating
			1 | // Nps = 1,
			// A PSAT (product satisfaction) survey. Asks user to rate "overall, based on their experience, how satisifed are they with this app"
			// Contains a prompt, question, and rating
			2 | // Psat = 2,
			// A BPS (build promotion) survey. Asks user to choose between Yes and No options of promoting the current build to the next audience ring
			// Contains a prompt, question, and rating (Yes/No)
			3 | // Bps = 3,
			// A FPS (feature promotion) survey. Asks user to rate a given app feature.
			// Contains a prompt, question, and rating
			4 | // Fps = 4,
			// A NLQS (net language quality score) survey. Asks user to rate the language quality.
			// Contains a prompt, question, and rating
			5 | // Nlqs = 5,
			// An intercept survey. Asks user if they want to talk to a Microsoft engineer to give feedback.
			// User can dismiss it or click on it to go to the intercept website, where the experience continues.
			6 | // Intercept = 6,
			// A Generic surface survey that uses content metadata to render a surface.
			// As of 4th Feb 2019 there are 11 types defined in Mso hence giving a value of 12.
			12;  // GenericMessagingSurface = 12,
	}

	module ISurveyComponent {
		export type Type =
			"Prompt" |
			"Comment" |
			"Rating" |
			"CVSurvey" |
			"MultipleChoice" |
			"Intercept";
	}

	/**
	 * A comment is a simple question posed to the user, expecting a
	 * free-form text response
	 */
	interface IComment extends ISurveyComponent {
		getQuestion: () => string;
		getSubmittedText: () => string;
		setSubmittedText: (submittedText: string) => void;
	}

	/**
	 * A rating is a question posed to the user, expecting a
	 * selection index from an ordered list of values.  The response
	 * is typically normalized to a decimal value (e.g. 1 out of 5 is 0.2).
	 * If using the setSelectedRatingIndex method to set index value, will
	 * set the array index as the selectedIndex i.e. 0th element maps to the first element.
	 * Example: setSelectedRatingIndex(2) maps to the 3rd element in the array
	 */
	interface IRating extends ISurveyComponent {
		getQuestion: () => string;
		getRatingValuesAscending: () => string[];
		getSelectedRating: () => string;
		getSelectedRatingIndex: () => number;
		setSelectedRatingIndex: (selected: number) => void;
	}

	/**
	 * A prompt is question posed to the user about whether or not they would like to
	 * participate in an activated survey.  It can be intrusive or passive, and includes
	 * a title, question, and yes/no labels
	 */
	interface IPrompt extends ISurveyComponent {
		getTitle: () => string;
		getQuestion: () => string;
		getYesButtonText: () => string;
		getNoButtonText: () => string;
		getButtonSelected: () => IPrompt.PromptButton;
		setButtonSelected: (buttonSelected: IPrompt.PromptButton) => void;
	}

	module IPrompt {
		// Values for the prompt button
		export type PromptButton =
			0 | // Unselected
			1 | // Yes
			2;  // No
	}

	/**
	* A multiple choice question posed to the user, expecting one or more
	* selection index from an ordered list of values.
	*/
	interface IMultipleChoice extends ISurveyComponent {
		// The question for the Multiple Choice component
		getQuestion: () => string;
		// Gets the array of options specified for the component
		getAvailableOptions: () => string[];
		// Gets the state (true/false) of the checkboxes for each option in the option array
		getOptionSelectedStates: () => boolean[];
		// Sets the state (true/false) of the checkboxes for each option in the option array
		setOptionSelectedStates: (selectedState: boolean[]) => void;
		// Gets the minimum number of options that must be selected by the user
		getMinNumberofSelectedOptions: () => number;
		// Gets the maximum number of options that a user can select
		getMaxNumberofSelectedOptions: () => number;
		// Returns true if user selected at least minimum number of options
		ValidateMinNumberofSelectedOptions: () => boolean;
		// Returns true if user selected less than or equal to maximum number of options
		ValidateMaxNumberofSelectedOptions: () => boolean;
	}

	/**
	 * An intercept notification posed to the user. The goal is to direct the user
	 * to the intercept URL.
	 */
	interface IIntercept extends ISurveyComponent {
		// Title for the intercept dialog
		getTitle(): string;

		// Text to show the user
		getQuestion(): string;

		// URL to redirect the user to
		getUrl(): string;
	}

	interface IActivityListener {
		/**
		 * Log an activity to Floodgate, incrementing its occurrence count by the given number if specified,
		 * otherwise incrementing its occurrence count by one as default
		 */
		logActivity(activityName: string, increment?: number): void;
		/**
		 * Start an activity timer (overwriting any previously unclosed start).
		 * NOTE: Does not increment the activity count.
		 */
		logActivityStartTime(activityName: string, startTime?: Date): void;
		/**
		 * Stop an activity timer and clears the previous start time.
		 * Adds the elapsed seconds between this stop and the previous start into the count for this activity
		 * \note If no previous start was logged, or start is somehow in the future, results in 0 count increment
		 */
		logActivityStopTime(activityName: string, stopTime?: Date): void;
	}

	interface IInitOptions extends IInitOptionsCommon, IInitOptionsInAppFeedback { }

	/**
	 * Interface for launch options for inAppFeedback
	 */
	interface ILaunchOptionsInAppFeedback {
		applicationGroup?: IManifestDataApplication;
		telemetryGroup?: IManifestDataTelemetry;
		webGroup?: IManifestDataWeb;
		categories?: ICategoryOptions;
	}

	/**
	 * Interface for common initialization options
	 */
	interface IInitOptionsCommon {
		// An id uniquely identifying the host app in the OCV world
		appId: number;
		// The url at which the styles reside
		stylesUrl?: string;
		// The url at which the internationalization files reside
		intlUrl?: string;
		// The filename for the internationalization strings, if different than the default filename
		intlFilename?: string;
		// Environment
		environment?: number;
		// A string uniquely identifying the current session
		sessionId?: string;
		// cid string which is in-scope for GDPR compliance for export and deletion request
		cid?: string;
		// A build version for your app
		build?: string;
		// The primary colour in hex form #[0-9a-f]{6}
		primaryColour?: string;
		// The secondary colour in hex form #[0-9a-f]{6}
		secondaryColour?: string;
		// The UI locale name of the calling page
		locale?: string;
		// A callback which will be called when the SDK encounters errors
		onError?: (err: string) => void;
		// Application related metadata
		applicationGroup?: IManifestDataApplication;
		// Telemetry related metadata
		telemetryGroup?: IManifestDataTelemetry;
		// Web related metadata
		webGroup?: IManifestDataWeb;
		// User Email
		userEmail?: string;
		// The json string to define event sampling
		eventSampling?: ISamplingEvent[];
		// The timeout for sending response in petrol
		petrolTimeout?: number;
		// check to see if user is consumer/commercial to dynamically display UI
		isCommercialHost?: boolean;
		// Option to see if the host will be providing their own resources such as strings or CSS
		customResourcesSetExternally?: CustomResources;
		// the email policy value from OCPS to determines if the user can see an option to include their email when they submit feedback
		emailPolicyValue?: OCPSValues;
		// the screenshot policy value from OCPS to determines if the user can see an option to include their screenshot when they submit feedback
		screenshotPolicyValue?: OCPSValues;
		// AgeGroup enum which user is classfied as
		ageGroup?: AgeGroup;
		// Authentication Type of user
		authenticationType?: AuthenticationType;
		// Host application settings
		applicationSettings?: IApplicationSettings;
	}

	/**
	 * Interface for host application settings
	 */
	interface IApplicationSettings {
		readonly [key: string]: unknown;
	}

	/**
	 * Interface for inAppFeedback initialization options
	 */
	interface IInitOptionsInAppFeedback {
		// Method called on feedback submission
		onDismiss?: (submitted: boolean) => void;
		// Bug form toggle
		bugForm?: boolean;
		// email enabled
		showEmailAddress?: boolean;
		// User email
		userEmail?: string;
		// Screenshot toggle
		screenshot?: boolean;
		// Flag to enable/disable transistion on the UI.
		transitionEnabled?: boolean;
		// The switch to show a thank you panel
		isShowThanks?: boolean;
		// the feedback policy value from OCPS to determines if the user can submit feedback
		// NotConfigured = 0 | Enabled = 1 | Disabled = 2
		sendFeedbackPolicyValue?: OCPSValues;
		// FeedbackForumUrl
		feedbackForumUrl?: string;
		// Show normal Thanks Page or with Feedback portal link
		myFeedbackForumUrl?: string;
	}

	/**
	 * Interface for floodgate event sampling options
	 */
	export interface ISamplingEvent {
		type: string;
		name: string;
		sampleRate: number;
	}

	/**
	 * Enum mask flag to indicate the resources that are going to be set by the host app.
	 * Example: Set to 3 in order to specify both Css and Strings (b11)
	 */
	export type CustomResources =
		0 | // None
		1 | // Css - b01
		2;  // Strings - b10

	/**
	 * Enum to indicate OCPS policies possible values that are going to be set by the host app.
	 */
	export type OCPSValues =
		0 | // NotConfigured
		1 | // Enabled
		2;  // Disabled

	/**
	 * Interface for floodgate initialization options
	 */
	interface IInitOptionsFloodgate {
		// Duration after which floodgate prompt automatically dismissed if not clicked? Default is no auto-dismiss.
		// 0: No auto-dismiss (default), 1: 7s, 2: 14s, 3: 21s, 4: 28s, 5: 60s, 6: 90s, 7: 120s, 8: 150s
		autoDismiss?: number;

		// The campaign definitions, see the documentation for the schema
		campaignDefinitions?: any;

		// Flights for TMS service
		campaignFlights?: string;

		// Additional query parameters for TMS service
		campaignQueryParameters?: string;

		// Delegate to be called when the floodgate UI is dismissed, whether a submission happened or not.
		onDismiss?: (campaignId: string, submitted: boolean) => void;

		// The callback to be executed when a survey is activated
		onSurveyActivatedCallback?: IOnSurveyActivatedCallback;

		// The callback to be executed when host-based data operations are needed
		settingStorageCallback?: IFloodgateSettingStorageCallback;

		// Optionally provide a method which will accept the custom strings in Campaign Definitions and return actual
		// strings that must be displayed. Used for localization.
		uIStringGetter?: (str: string) => string;

		// The callback to be executed when host-based auth token is needed
		authTokenCallback?: IFloodgateAuthTokenCallback;

		// The boolean option to enforce if surveys are allowed by a policy, such as an OCPS policy for a tenant.
		// True: Surveys are enabled, floodgate works normally; False: Surveys are disabled
		surveyEnabled?: boolean;

		// The boolean option to enforce if email address collection in a survey is allowed by a policy, such as an OCPS policy for a tenant.
		// True: email address collection is enabled, floodgate works normally; False: email address collection is disabled
		showEmailAddress?: boolean;

		// Interface to pass the params for the Customer Voice Survey
		customerVoiceSurveyParams?: ICVSurveyParams;

		// The boolean option to enable the usage of Governance Service. It's default as false if not defined.
		// True: Call to Governance Service to get the survey trigger permission before showing the survey ; False: Do not use Governance Service
		governanceServiceEnabled?: boolean;

		// The Governance Service configuration options. It's only taken into account if governanceServiceEnable is set to true.
		governanceServiceConfig?: IGovernanceServiceConfig;

		// is personalizer enabled on client side.
		personalizerEnabled?: boolean;
	}

	interface IManifestDataApplication {
		appData?: string;
		feedbackTenant?: string;
	}

	export interface ICVSurveyParams {
		cvFlights: string,
		isCVSurveyEnabled: boolean,
		productName: string,
		uiHost: string,
	}

	interface IManifestDataTelemetry {
		audience?: string;
		audienceGroup?: string;
		channel?: string;
		deviceId?: string;
		deviceType?: string;
		errorClassification?: string;
		errorCode?: string;
		errorName?: string;
		featureArea?: string;
		flights?: string;
		flightSource?: string;
		fundamentalArea?: string;
		installationType?: string;
		isLogIncluded?: boolean;
		isUserSubscriber?: boolean;
		loggableUserId?: string;
		officeArchitecture?: string;
		officeBuild?: string;
		officeEditingLang?: number;
		officeUILang?: number;
		osBitness?: number;
		osBuild?: string;
		osUserLang?: number;
		platform?: string;
		processSessionId?: string;
		ringId?: number;
		sku?: string;
		sourceContext?: string;
		systemProductName?: string;
		systemManufacturer?: string;
		tenantId?: string;
	}

	interface IManifestDataWeb {
		browser?: string;
		browserVersion?: string;
		dataCenter?: string;
		sourcePageName?: string;
		sourcePageURI?: string;
	}

	interface ICategoryOptions {
		// Wether to show categories drop down or not
		show: boolean;

		// Custom categories
		customCategories?: string[];
	}
	
	interface IGovernanceServiceConfig {
		// limit number of retry
		retry?: number;

		// timeout in milliseconds
		timeout?: number;

		// In case of failure call, option to whether continue display survey or fall off
		forceServicelessSurveyDisplay?: boolean;
	}

	export type AgeGroup =
		0 | // Undefined​ - Unknown age
		1 | // MinorWithoutParentalConsent - Minor under the age of consent
		2 | // MinorWithParentalConsent - Minor under age of consent
		3 | // Adult - Adult
		4 | // NotAdult - Minor above age of consent
		5;  // MinorNoParentalConsentRequired - Minor above age of consent

	export type AuthenticationType =
		0 | // MSA,
		1 | // AAD,
		2;  // Unauthenticated
}