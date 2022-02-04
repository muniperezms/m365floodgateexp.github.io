# Change Log
All notable changes to this project will be documented in this file.
The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [Unreleased]

## [1.4.4] - 2019-07-16
### Added
- Added option to pass custom IRIS flights to TMS service
- Added method to return PUID in auth callback for getting targeted campaigns

## [1.4.3] - 2019-07-01
### Changed
- Handle multiple start 'async' calls in performant manner with better event tracking

## [1.4.2] - 2019-05-23
- Fixed Floodgate urgent governed channel remaining closed after one trigger

## [1.4.1] - 2019-05-16
### Changed
- Improved resiliency to rapid Floodgate engine start() calls
- Improved survey UX accessibility

## [1.4.0] - 2019-04-29
### Added
- Added Targeted Messaging campaign definition provider that fetches campaigns from TMS service
- Added new campaignQueryParameters option to pass custom query parameters to TMS service
- Added TmsClient which is responsible for service retrieval of campaign and governance information and caching it
- Added Generic message surface implementation to dynamically fetch UX package and render surfaces

## [1.3.2] - 2019-04-02
### Added
- Client side sampling of the user selected event

## [1.3.1] - 2019-02-11
### Added
- Added a method in Floodgate to return a list of setting ids keyed by names
- Added a new MultipleChoice survey componet (without UX) which allows answers to be submitted in a multiple choice format to a question
- Added functionality to pass a LauncherType string from the campaign definition. This string can be accessed from getLauncherType method on the Survey
- Added functionality to enable access to survey components and ability to submit responses via OnSurveyActivatedCallback

### Changed
- Changed functionality for FPS to have a mandatory Prompt component and at least one or more of a single instance of any of the other components (Rating, Comment, MultipleChoice)

## [1.3.0] - 2018-12-06
### Changed
- Improved Floodgate campaign state tracking with host-based storage provider option

## [1.2.9] - 2018-11-27
### Changed
- Updated dependent Aria compact SDK to latest v1.2.2
- Updated urgent channel delay to 0 second from 4 hours

### Added
- A callback in Floodgate initOptions to enable host-based storage provider option

## [1.2.8] - 2018-11-21
### Fixed
- Enabled anonymous cross origin when loading ui strings js files

## [1.2.7] - 2018-10-18
### Fixed
- Fix checkbox placement to be aligned with text and padded the space between text and email editbox

## [1.2.6] - 2018-10-04
### Added
- Added support for enabling or disabling User Email component to show up in the survey
### Fixed
- Fix high contrast mode invisible icon bug on Edge and Internet Explorer

## [1.2.5] - 2018-08-23
- Version bump

## [1.2.4] - 2018-08-21
- Version bump

## [1.2.3] - 2018-08-20
### Added
- Added sdkVersion property to 'window.OfficeFeedback' object

### Fixed
- Fixed bug which caused stop time to be not set with default

## [1.2.2] - 2018-08-16
### Added
- Added support locale support for 'language + region' in addition to 'language' alone

## [1.2.1] - 2018-06-29
### Added
- Added cid to initOptions

## [1.2.0] - 2018-06-25
### Added
- Added audience group to telemetry
- Added new SurveyTemplate types: NLQS and NPS (for custom NPS)

### Changed
- (Breaking change) Auto-dismiss for the floodgate prompt can now be customized with a set of duration values instead of the fixed 7s earlier. The auto-dismiss init option now takes an int-based enum instead of boolean.

## [1.1.6] - 2018-06-08
### Added
- The ability to launch the floodgate UI on demand with custom strings/properties

### Changed
- Update UserVoice UI and remove links to terms/privacy

## [1.1.5] - 2018-05-09
### Added
- An option to auto-dismiss the floodgate prompt after 7s of inactivity
- Ability to 'esc' out of the floodgate toast

### Changed
- The floodgate toast title is now themed

## [1.1.4] - 2018-05-03
### Added
- An onDismiss parameter for floodgate

## [1.1.3] - 2018-04-19
### Changed
- The floodgate.initialize() promise rejects if any campaign definition is not valid

### Added
- An onError callback parameter to initOptions which gets called when the sdk hits errors

### Fixed
- Fixed the initialize() promise not rejecting when there is an error loading the intl files
- Fixed custom category value length validation to allow 20 characters

## [1.1.2] - 2018-04-09
### Fixed
- (Accessibility) Use aria-labelledby instead of aria-label whenever possible
- Bug which caused floodgate responses in v1.1.1 to not be ingested by OCV

## [1.1.1] - 2018-04-04
### Fixed
- Bug which caused screenshots in v1.1.0 to not be ingested by OCV

## [1.1.0] - 2018-03-07
### Changed
- Removed dependency on zip, drastically reducing package size

### Fixed
- No more UI blocking while screenshot is rendered
- InitOptions are no longer passed 'by reference'

## [1.0.5] - 2018-02-27
### Fixed
- UI spacing refinements (sharepoint review)
- Added missing surveyId in telemetry logs

## [1.0.4] - 2018-02-12
### Fixed
- Added onfocus styling for checkboxes and links
- Improved the line wrapping behaviour of the "include screenshot" control (applicable for locales with long strings)

## [1.0.3] - 2018-01-29
### Changed
- Updated telemetry for GDPR compliance

## [1.0.2] - 2017-12-22
### Fixed
- Screenshot bug where host page security policy causing screenshot rendering to hang

## [1.0.1] - 2017-12-07
### Changed
- The locale "bn" has been replaced by "bn-BD"

### Fixed
- Accessibility bug where the main dialog did not have role specified

## [1.0.0] - 2017-09-15
### Added
- Floodgate (System Initiated Survey Feedback)

## [0.2.5] - 2017-08-11
### Changed
- All assets (intl, scripts, styles) were converted to lowercase

## [0.2.4] - 2017-07-26
### Fixed
- fixed bad npm package publish where vsts-npm-auth was included as a dependancy

## [0.2.3] - 2017-06-30
### Fixed
- fixed Chinese language support

### Added
- enabled category support

### Changed
- removed 13 language locales that are no longer supported by Office

## [0.2.2] - 2017-05-10
### Added
- enabled Npm scoped package support

## [0.2.1] - 2017-05-04
### Added
- new initOptions property intlFilename

## [0.2.0] - 2017-04-27
### Fixed
- fixed tab/shift tab issue so focus will stay to elements on the feedback dialog
- updated build number to allow a single octet and up to 9 digits in each octet

### Added
- enabled support for V2 feedback schema with additional properties support
- enabled locale "ll-cc" fallback to matching "ll".

### Changed
- the list of supported locales has changed (see README)

## [0.1.14] - 2016-12-28
### Fixed
- fixed a IE/Edge perf issue due to creation of canvas elements

## [0.1.13] - 2016-12-06
### Added
- language support for pt-BR.

### Fixed
- do not fallback on the global instance of require() if it exists

## [0.1.12] - 2016-11-29
### Changed
- UI tweaks.

## [0.1.11] - 2016-11-22
### Fixed
- Added return value types to methods in OfficeBrowserFeedback.d.ts

## [0.1.10] - 2016-11-22
### Added
- A typescript definitions for the package
- CHANGELOG.md
- LICENSE.md

### Changed
- Passing unknown locale will default to "en" as opposed to throw an exception
- Uservoice will not be shown for locales other than "en"

### Fixed
- Go to uservoice button was opening a blank page
