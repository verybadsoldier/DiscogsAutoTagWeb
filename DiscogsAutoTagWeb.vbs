Option Explicit
'
' Discogs Tagger Script for MediaMonkey ( Let & eepman & crap_inhuman )
'
Const VersionStr = "v5.84"

'Changes from 5.83 to 5.84 by crap_inhuman in 01.2024
'	Added option to select all images


'Changes from 5.82 to 5.83 by crap_inhuman in 09.2023
'	Changed the detection of URL in search string


'Changes from 5.81 to 5.82 by crap_inhuman in 06.2022
'	Added option to change search-mode (Using album name for initial search Yes/no)


'Changes from 5.80 to 5.81 by crap_inhuman in 08.2021
'	Request of the user collection changed (making them much faster)


'Changes from 5.79 to 5.80 by crap_inhuman in 03.2021
'	Fixed bug with styles option
'	Added a Please-Wait button


'Changes from 5.78 to 5.79 by crap_inhuman in 01.2021
'	Bug removed from last update


'Changes from 5.77 to 5.78 by crap_inhuman in 01.2021
'	Added "Digital Media" to Mediatype List (Thanks to rvonder)
'	Changes in the new matching engine: If only one song was choosen for tagging, you have to select the track from the album in the upper list


'Changes from 5.76 to 5.77 by crap_inhuman in 10.2020
'	Added option to show the qty of release format in format tag (i.e. 2 x CD)
'	Added "Digital Media" to Mediatype List (Thanks to rvonder)


'Changes from 5.75 to 5.76 by crap_inhuman in 09.2020
'	Bug with additional images while tagging one track removed
'	Option to choose matching engine added


'Changes from 5.74 to 5.75 by crap_inhuman in 05.2020
'	Some changes made in matching method


'Changes from 5.73 to 5.74 by crap_inhuman in 04.2020
'	The following chars ´ , ` , ’ , ‘ , ” , “ are replaced by '
'	Fixed tracknumber bug
'	Changed the matching method (Thanks to DJ Samhein)


'Changes from 5.72 to 5.73 by crap_inhuman in 04.2020
'	Fixed trackname-bug
'	A ’ will now replaced by ' , too


'Changes from 5.71 to 5.72 by crap_inhuman in 04.2020
'	Added option to replace a ´ and ` with a '
'	Added option to search at Metal-Archives.com
'	Removed issues with multi-disc albums, need MediaMonkey Version 4.1.28.1905 or above


'Changes from 5.70 to 5.71 by crap_inhuman in 09.2019
'	Small bugfix from new release yesterday


'Changes from 5.69 to 5.70 by crap_inhuman in 09.2019
'	Bug removed: Track-numbering didn't work with more than one char in rare cases


'Changes from 5.68 to 5.69 by crap_inhuman in 07.2019
'	Added 'Video' to the CD tag list


'Changes from 5.67 to 5.68 by crap_inhuman in 05.2019
'	Now the track-numbering start with the Disc-numbering


'Changes from 5.66 to 5.67 by crap_inhuman in 05.2019
'	Changed the Side to Disc function (2 Vinyl sides are one disc)
'	Added a "Deselect all" button


'Changes from 5.65 to 5.66 by crap_inhuman in 04.2019
'	Changed the Authorize links to work with Windows 10


'Changes from 5.64 to 5.65 by crap_inhuman in 04.2019
'	Added option to change media format separator


'Changes from 5.63 to 5.64 by crap_inhuman in 11.2018
'	Bug removed with mixed media format
'	Select Grouping (not ready yet)
'	Add Start/Stop button for track refresh


'Changes from 5.62 to 5.63 by crap_inhuman in 09.2018
'	Musicbrainz: New Bug in Label subroutine removed


'Changes from 5.61 to 5.62 by crap_inhuman in 09.2018
'	Musicbrainz: Bug removed: Label with no name stop the script


'Changes from 5.60 to 5.61 by crap_inhuman in 07.2018
'	Bug removed: Format checkbox didn't work


'Changes from 5.59 to 5.60 by crap_inhuman in 06.2018
'	Added new option to store the name of grouped tracks into the grouping tag
'	Added changeable separator for sub-tracks


'Changes from 5.58 to 5.59 by crap_inhuman in 04.2018
'	The script now use secure channel (https) while fetching data
'	Workaround for Discogs date of master release bug: date is now empty instead number 0


'Changes from 5.57 to 5.58 by crap_inhuman in 02.2018
'	Repaired the artist tagging


'Changes from 5.56 to 5.57 by crap_inhuman in 02.2018
'	Choose artist for tagging back to his old place
'	Choose Track# and Disc# for tagging back on the main window


'Changes from 5.55 to 5.56 by crap_inhuman in 02.2018
'	Changed some visual things to be more user-friendly
'	Authorize the Discogs Tagger should now be more user-friendly
'	Every individual genre and style can now selected
'	Bug removed: if last track is a sub-track, it didn't recognized as sub-track
'	Updating the track titles can now be turned off


'Changes from 5.54 to 5.55 by crap_inhuman in 01.2018
'	Now choose your favourite tagger (Discogs or MusicBrainz) from options in the right upper corner
'	Bug removed: a slash (/) in folder-names will not stop the script anymore


'Changes from 5.53 to 5.54 by crap_inhuman in 12.2017
'	Discogs Tagger use the title-name of the first selected track for search at Discogs, if no album name was found


'Changes from 5.52 to 5.53 by crap_inhuman in 12.2017
'	Improved identification of Release/Master/Label/Artist Discogs Numbers e.g. [r12345], [m12345], [l12345], [a12345]
'	Fixed identification of master url 


'Changes from 5.51 to 5.52 by crap_inhuman in 09.2017
'	Now check status-code while image download
'	Return to 'MSXML2.ServerXMLHTTP.6.0' for MusicBrainz search
'	Improved some code


'Changes from 5.50 to 5.51 by crap_inhuman in 09.2017
'	Secure channel error: Replaced all 'MSXML2.ServerXMLHTTP.6.0' with 'MSXML2.XMLHTTP.6.0'


'Changes from 5.49 to 5.50 by crap_inhuman in 09.2017
'	Secure channel error: Replaced 'MSXML2.ServerXMLHTTP.6.0' with 'MSXML2.XMLHTTP.6.0' for image download
'	Added third option for date field


'Changes from 5.48 to 5.49 by crap_inhuman in 08.2017
'	Changed the 'where to store date' option


'Changes from 5.47 to 5.48 by crap_inhuman in 06.2017
'	Added forgotten blank
'	Added "Ignore featuring artist" option


'Changes from 5.46 to 5.47 by crap_inhuman in 06.2017
'	Added the feature "Add leading zero (Disc#)"


'Changes from 5.45 to 5.46 by crap_inhuman in 05.2017
'	Skip Extra-Artists without artistname in musicbrainz-tagger
'	Changed Musicbrainz logo
'	"Feat."-Function repaired
'	Searching at Musicbrainz should now be better


'Changes from 5.44 to 5.45 by crap_inhuman in 11.2016
'	Bug with leading zero removed
'	Added option for limiting releases


'Changes from 5.43 to 5.44 by crap_inhuman in 11.2016
'	Bug with Release Date / Original Date removed
'	In Advanced Search hitting enter will trigger search
'	Changed error message if no release found


'Changes from 5.42 to 5.43 by crap_inhuman in 11.2016
'	Leading/Trailing spaces will now detected and deleted in album-name, too.
'	Changing the release will reset the selected tracks
'	Added option for choosing where to store Release Date/Original Date


'Changes from 5.41 to 5.42 by crap_inhuman in 10.2016
'	Leading/Trailing spaces will now detected and deleted in track-names and the release number
'	will be send to mysql db, a busy bee will then fix it at Discogs (Hopefully) ;)


'Changes from 5.40 to 5.41 by crap_inhuman in 10.2016
'	Small bugfixes


'Changes from 5.39 to 5.40 by crap_inhuman in 06.2016
'	Discogs: Removed bug with featuring artist in the albumartist
'	Musicbrainz: Removed one of the two blank character in featuring artist
'	Trackname removed from first search


'Changes from 5.38 to 5.39 by crap_inhuman in 04.2016
'	The check if it's already in Discogs Collection can be turned off
'	Added option to delete duplicated entries in tags


'Changes from 5.37 to 5.38 by crap_inhuman in 03.2016
'	Redirected back to my webspace
'	Added new feature: Add the selected album to your Discogs Collection


'Changes from 5.36a to 5.37 by crap_inhuman in 03.2016
'	Changed the authorize script


'Changes from 5.36 to 5.36a by crap_inhuman in 03.2016
'	Changed from HTTP to HTTPS - using free webspace discogstagger.hol.es


'Changes from 5.35 to 5.36 by crap_inhuman in 01.2016
'	Bug with empty Label removed


'Changes from 5.34 to 5.35 by crap_inhuman in 01.2016
'	Changed feat. Artist function: Now you can use ";" for separator
'	Fixed some Artist-Separator bugs


'Changes from 5.33 to 5.34 by crap_inhuman in 01.2016
'	Artist search now use discogs artist-id


'Changes from 5.32 to 5.33 by crap_inhuman in 12.2015
'	Added advanced search button
'	Added option: Move The in artist name to the end


'Changes from 5.31 to 5.32 by crap_inhuman in 12.2015
'	Search improved for more accurate results


'Changes from 5.30 to 5.31 by crap_inhuman in 11.2015
'	Choose what kind of search after entering search string


'Changes from 5.29 to 5.30 by crap_inhuman in 11.2015
'	Fixed new bug with ampersand


'Changes from 5.28 to 5.29 by crap_inhuman in 10.2015
'	Added Relationship-Attributes for Musicbrainz Credits
'	Fixed bug with additional musicbrainz images
'	Added support for foreign characters
'	Fixed bug with ampersand in artistname


'Changes from 5.27 to 5.28 by crap_inhuman in 07.2015
'	Removed bug with unselecting tracks


'Changes from 5.26 to 5.27 by crap_inhuman in 07.2015
'	Removed bug with joint artists


'Changes from 5.25 to 5.26 by crap_inhuman in 07.2015
'	Comma removed after artist


'Changes from 5.24 to 5.25 by crap_inhuman in 06.2015
'	Comment Tag added to the Release info
'	Back to original Cover-Image saving-routine
'	Easier Discogs authorization


'Changes from 5.23 to 5.24 by crap_inhuman in 04.2015
'	Added check for image before try download it.
'	Image-Proxy removed


'Changes from 5.22 to 5.23 by crap_inhuman in 03.2015
'	Again a saving cover-image bug removed


'Changes from 5.21 to 5.22 by crap_inhuman in 03.2015
'	Bug removed: mm hangs while downloading covers
'	Bug removed: the tags will be written, while cover is saving
'	The common filename masks are implemented (Title, Artist, AlbumArtist,…)


'Changes from 5.20 to 5.21 by crap_inhuman in 03.2015
'	Saving cover-image bug removed


'Changes from 5.19 to 5.20 by crap_inhuman in 03.2015
'	Removed CheckImmedSaveImage, the image(s) will now saved immediately
'	Changed Cover-Image saving-routine to store the images in the cache


'Changes from 5.18 to 5.19 by crap_inhuman in 03.2015
'	Changed Image download due to recent changes on accessing images at discogs
'	Added check for new version once a day


'Changes from 5.17 to 5.18 by crap_inhuman in 02.2015
'	Removed bug with Catalog/Release tag
'	Removed bug with featuring artist


'Changes from 5.16 to 5.17 by crap_inhuman in 01.2015
'	Choosing "Master-Release" shows the Master-Release. Yes, really. ;)
'	Choosing "Versions of Master" shows all Versions (Releases) of the selected Master-Release.
'	If you selected a Release, the corresponding Master of this Release will choosen.
'	Added ISRC to CatalogTag


'Changes from 5.15 to 5.16 by crap_inhuman in 01.2015
'	Removed bug with Extra-Artists


'Changes from 5.14 to 5.15 by crap_inhuman in 01.2015
'	MusicBrainz: Tags with no value will no longer crash the script
'	MusicBrainz: The manual search now works
'	MusicBrainz: Some options which are not necessary are now hidden
'	Some small bugfixes


'Changes from 5.13 to 5.14 by crap_inhuman in 11.2014
'	The Tagger now detect OAuth authentication error


'Changes from 5.12 to 5.13 by crap_inhuman in 11.2014
'	More than one space between track positions doesn't stop the script anymore ;)
'	Changed subtrack error detection


'Changes from 5.11 to 5.12 by crap_inhuman in 10.2014
'	Removed a bug with empty keyword fields in the options menu
'	Now only the search requests use oauth


'Changes from 5.10 to 5.11 by crap_inhuman in 10.2014
'	New option "Save selected 'more images' after closing popup" fixed
'	New option "Don't copy empty values to non-empty fields" now works for genres/styles too
'	More Debug Output to Logfile


'Changes from 5.01 to 5.10 by crap_inhuman in 10.2014
'	Changed unclear text
'	Change order in dropdown list of the search result, put label at the end
'	Added option to enter unwanted tags in involved people
'	Added option to save selected "More images" after closing the popup
'	Added option "Don't copy empty values to non-empty fields"
'	Now show the TrackCount of every release in the search result (only with musicbrainz)
'	Some small changes to the layout


'Changes from 5.0 to 5.01 by crap_inhuman in 09.2014
'	Removed bug with search result
'	Removed bug if no release found


'Changes from 4.52 to 5.0 by crap_inhuman in 09.2014
'	Changed OAuth Authorization procedure (now wait 30 seconds for authorize)
'	Added MusicBrainz for searching
'	Removed some small bugs


'Changes from 4.51 to 4.52 by crap_inhuman in 08.2014
'	Changed OAuth Authorization procedure


'Changes from 4.50 to 4.51 by crap_inhuman in 07.2014
'	Removed bug with & character in searchstring
'	Small bugfixes
'	Removed bug with empty results


'Changes from 4.48 to 4.50 by crap_inhuman in 07.2014
'	Bug removed with utf-8 characters in searchstring (with big help from tillmanj !!)


'Changes from 4.47 to 4.48 by crap_inhuman in 07.2014
'	In the options menu you can now enter the access token manually
'	Bug removed in Keywords routine


'Changes from 4.46 to 4.47 by crap_inhuman in 07.2014
'	Changed the Delay - Function, WScript.sleep didn't work on all windows plattforms.


'Changes from 4.45 to 4.46 by crap_inhuman in 07.2014
'	The default settings for saving the Cover Images can now be changed in the options menu
'	Bug removed: Empty format-tag produced an error
'	Bug removed: Parsing wrong Artist Roles if a comma is between box bracket
'	Added OAuth authentication
'	Added option: Using Metal-Archives for release search (BETA)
'	Now it's possible to use * as wildcard in the Keywords
'	Added option: Print every involved people in a single line


'Changes from 4.44 to 4.45 by crap_inhuman in 05.2014
'	Bug removed: Didn't display the additional Image
'	Adjust the script for fetching the small album art
'	Adjust the script for removing leading and trailing spaces in Extra Artists
'	Add option to turn off subtrack detection


'Changes from 4.43 to 4.44 by crap_inhuman in 04.2014
'	Added simple routine to check and remove point in track positions (1. , 2. , 3. )
'	Bug removed: track position part
'	Max count for releases is set to 250


'Changes from 4.42 to 4.43 by crap_inhuman in 04.2014
'	Bug removed: Filter now work correctly
'	There's no max count for release results
'	Bug removed: Artist releases and Label releases work again


'Changes from 4.41 to 4.42 by crap_inhuman in 03.2014
'	Bug removed: Sub-Track do not select(set) the song
'	Added the option for switching the last artist separator ("&" or "chosen separator")


'Changes from 4.40 to 4.41 by crap_inhuman in 03.2014
'	Removed bug with more than one artist for a title
'	Added Artist separator to options menu
'	Added simple routine to check for false position separators


'Changes from 4.39 to 4.40 by crap_inhuman in 03.2014
'	featuring Keywords are now not case sensitive


'Changes from 4.38 to 4.39 by crap_inhuman in 03.2014
'	Keywords are now not case sensitive
'	Added Set Locale for supporting more countries


'Changes from 4.37 to 4.38 by crap_inhuman in 02.2014
'	Added the Featuring Keywords
'	Fixed a bug with the new submission form of discogs


'Changes from 4.36 to 4.37 by crap_inhuman in 02.2014
'	Changed the image access method


'Changes from 4.35 to 4.36 by crap_inhuman in 11.2013
'	The script now shows the filtered total and the matched total


'Changes from 4.34 to 4.35 by crap_inhuman in 11.2013
'	Raise the max count of release results to 100
'	Display the number of matched releases and which one you are viewing in the search bar


'Changes from 4.33 to 4.34 by crap_inhuman in 11.2013
'	Now it's possible to change the search string in the top bar


'Changes from 4.32 to 4.33 by crap_inhuman in 10.2013
'	Fixed a bug with the Separator


'Changes from 4.31 to 4.32 by crap_inhuman in 10.2013
'	Removed bug in extra artist assignment
'	Added 'Don't save' and 4 more fields for saving release-number


'Changes from 4.30 to 4.31 by crap_inhuman in 09.2013
'	Removed bug: Sub track name will not recognized if it is the last track
'	Removed bug: Script-Error occurred after closing the script-window, when no release found
'	Background of filter dropdown menu change to red if filter is selected (For better recognition)


'Changes from 4.00 to 4.30 by crap_inhuman in 07-09.2013
'	Added Sub tracks option.
'	Added option 'Unselect tracks without track-number'
'		Some albums at discogs have 'Index-Tracks'.
'		These tracks aren't song-tracks (e.g. Track-Name: 'Bonus track' or 'Live side')
'		This option unselect these tracks automatically
'		-------------------------------------------------------
'	Show a warning if the number of songs are different
'	For the catalog-number, release-country and media-format you can choose "Don't save" in the option menu, if you don't need it.
'	You can edit the keywords for linking the composer, producer, conductor,... tags with discogs
'	included DiscogsImages: you can choose more than one image for an album
'	New Option: Check 'Save Image' Checkbox only if release have no image
'	New Option: Choose another field for saving Style


'Changes from 3.65 to 4.00 by crap_inhuman in 07.2013
'	Bug removed with releases having leading zero in track-position
'	Added option for "Force NO Disc Usage". Helpful if a release have tracks with varying track-numbers (e.g. http://www.discogs.com/release/2942314 )
'		Without the option the script translate the varying track position to disc sides
'	Added option to show the original discogs track position
'	Moved the options to the left side for more place for the tracklisting
'	Moving the mouse-pointer over a checkbox now show more information about the usage
'	Now the chosen filters will be saved with the options
'		Choose one MediaType, MediaFormat, Country or Year from the drop-down list and save the options
'		or press one of the "Set ... Filter" button to select more than one Mediatype, MediaFormat, Country or Year
'		Choosing "Use ... Filter" in the drop-down list uses the custom filter-settings
'		Choosing "No ... Filter" from the drop-down list stop filtering the result
'		The Filter settings will only be saved if you press the "Save Options" button
'	The Custom Tags for saving the release, catalog, country and format will now be chosen in the options -> Discogs Tagger or during script installation
'	Showing the Data Quality of the Discogs release

'Changes from 3.64 to 3.65 by crap_inhuman in 07.2013
'	Bug removed: bug with additional artists removed, which only occur in rare cases

'Changes from 3.63 to 3.64 by crap_inhuman in 06.2013
'	Bug removed: selecting "Sides To Disc" and "Add Leading Zero", zero is dropped from track number and is displayed as a single digit

'Changes from 3.62 to 3.63 by crap_inhuman in 02.2013 (not released)
'	Bug removed: no search result -> no output

'Changes from 3.61 to 3.62 by crap_inhuman in 02.2013
'	Insert code for supporting french language machines
'	Comments will now be saved
'	Delete some unused but declared variables
'	Name for "feat." can be edit
'	Some small bugfixes

'Changes from 3.6 to 3.61 by crap_inhuman in 02.2013
'	Removed a bug in the option 'Featuring Artist behind title'
'	Better implementation of the option 'Featuring Artist behind title'
'	Inserting Master and Release URLs now work in the Search-Panel

'Changes from 3.5 to 3.6 by crap_inhuman in 02.2013
'	Implementation of eepman's JSON-Parser
'	Now read the user-specific Separator characters and use it for separating
'	Label / Artist / Master Search now using the JSON Parser too
'	Some bugfixes

'Changes from 3.3 to 3.5 by crap_inhuman in 01.2013
'	Now you can choose which Custom Tag will be used for the Tags: ReleaseID, Catalog, Country and Format
'	The "Credits for ExtraArtists in tracks" will now saved in MediaMonkey !
'	Added option for "Add Leading zero to Tracknumbers"
'	Added option for "Include Producer"
'	Added JSON Parser for the new Discogs-API

'	Some bugfixes (Filter now working correct)
'	Added the option to choose the place for Featuring Artist (Artist or Title)
'	e.g. Aaliyah - We Need a Resolution (ft. Timbaland) -or- Aaliyah (ft. Timbaland) - We Need a Resolution
'	Changeable Name for "Various" Artists (Various Artists)
'	Added option for "Adding comment"
'	Get OriginalDate from Master-Release if available
'	The Script now reads the saved Discogs Release-ID from the chosen Release-Tag

' ToDo: Add more tooltips to the html
'		Mediamonkey Bug: Erster und letzter Buchstabe in SearchArtist fehlt (wenn nächster Buchstabe blank ist e.g. "3 doors down", "Miss may i")
'		Wrong Publisher, Producer, etc. in Subtracks. The script only take the info from the first subtrack..
'		Adding Artist-Alias in Musicbrainz search


' WebBrowser is visible browser object with display of discogs album info
Dim WebBrowser

' decoded json object representing currently selected release
Dim CurrentRelease, QueryPage

Dim UI

Dim Results, ResultsReleaseID, NewResult
Dim CurrentReleaseId, CurrentResultId
Dim ini

Dim CheckAlbum, CheckArtist, CheckAlbumArtist, CheckAlbumArtistFirst, CheckLabel, CheckDate, CheckOrigDate, CheckGenre
Dim CheckCountry, CheckCover, CheckSmallCover, SmallCover, CheckCatalog, CheckRelease, CheckInvolved, CheckLyricist, CheckGrouping
Dim CheckComposer, CheckConductor, CheckProducer, CheckDiscNum, CheckTrackNum, CheckFormat, CheckUseAnv, CheckYearOnlyDate, CheckReplaceAccentChar
Dim CheckForceNumeric, CheckSidesToDisc, CheckForceDisc, CheckNoDisc, CheckLeadingZero, CheckLeadingZeroDisc, CheckVarious, TxtVarious
Dim CheckTitleFeaturing, CheckComment, CheckFeaturingName, TxtFeaturingName, CheckOriginalDiscogsTrack, CheckSaveImage, CheckLimitReleases
Dim CheckTurnOffSubTrack, CheckInvolvedPeopleSingleLine, CheckDontFillEmptyFields, CheckTheBehindArtist, CheckNewMatchingEngine, CheckSearchWithAlbum
Dim CheckDeleteDuplicatedEntry, StoreDate, OriginalDateRead, ReleaseDateRead
Dim CheckIgnoreFeatArtist, CheckShowQtyFormat, SubTrackSeparator, trackRefresh, DiffNumberTracks
REM Dim CheckUserCollection
Dim DiscogsUsername, GUID
Dim SubTrackNameSelection
Dim CountryFilterList, MediaTypeFilterList, MediaFormatFilterList, YearFilterList
Dim LyricistKeywords, ConductorKeywords, ProducerKeywords, ComposerKeywords, FeaturingKeywords, UnwantedKeywords

Dim SavedReleaseId
Dim NewSearchTerm, NewSearchArtist, NewSearchAlbum, NewSearchTrack
Dim SavedSearchArtist, SavedSearchAlbum
Dim SavedMasterID, SavedArtistID, SavedLabelID

Dim FilterMediaType, FilterCountry, FilterYear, FilterMediaFormat, CurrentLoadType
Dim NewGenre, GenresList, GenresSelect
Dim MediaTypeList, MediaFormatList, CountryList, CountryCode, YearList, AlternativeList, LoadList, RelationAttrList
Dim FormatSeparator, ArtistSeparator, ArtistLastSeparator, LimitReleases, Grouping

Dim FirstTrack, Errormessage
Dim AlbumArtURL, AlbumArtThumbNail
Dim iMaxTracks
Dim iAutoTrackNumber, iAutoDiscNumber, iAutoDiscFormat
Dim LastDisc
Dim SelectAll
Dim UnselectedTracks(1000)
Dim UnselectedTrackNames(1000)
Dim CheckNewVersion, LastCheck
Dim DiscTrackOrder1()
Dim DiscTrackOrder2()
Dim DiscTrackOrder3()
Dim NewTrackOrder()
Dim NewTrackOrder2()
Dim NewTrackOrder3()

Dim ReleaseTag, CountryTag, CatalogTag, FormatTag
Dim OriginalDate, Separator, FirstRun, Button
Dim OptionsChanged, UserAgent
Dim AccessToken, AccessTokenSecret

Dim fso, loc, logf

REM QueryPage = "Discogs"
REM QueryPage = "MusicBrainz"
REM QueryPage = "MetalArchives"

REM SearchFor = 1 = Artist
REM SearchFor = 2 = Album
REM SearchFor = 3 = Release

'----------------------------------DiscogsImages----------------------------------------
Dim SaveImageType, SaveImage, CoverStorage, CoverStorageName
Dim FileNameList, ImageTypeList, ImageList, ImageLocal
Dim list
Dim ImagesCount
Dim SaveMoreImages
Dim WebBrowser3
Dim SelectedSongsGlobal

Set ImageTypeList = SDB.NewStringList

ImageTypeList.Add SDB.Localize("Not specified")		'0
ImageTypeList.Add SDB.Localize("Cover (front)")		'3
ImageTypeList.Add SDB.Localize("Cover (back)")		'4
ImageTypeList.Add SDB.Localize("Leaflet Page")
ImageTypeList.Add SDB.Localize("Media Label")
ImageTypeList.Add SDB.Localize("Lead Artist")
ImageTypeList.Add SDB.Localize("Artist")
ImageTypeList.Add SDB.Localize("Conductor")
ImageTypeList.Add SDB.Localize("Band")
ImageTypeList.Add SDB.Localize("Composer")
ImageTypeList.Add SDB.Localize("Lyricist")
ImageTypeList.Add SDB.Localize("Recording Location")
ImageTypeList.Add SDB.Localize("During Recording")
ImageTypeList.Add SDB.Localize("During Performance")
ImageTypeList.Add SDB.Localize("Video Screen Capture")
ImageTypeList.Add SDB.Localize("Illustration")
ImageTypeList.Add SDB.Localize("Band Logotype")
ImageTypeList.Add SDB.Localize("Publisher Logotype")		'20
'----------------------------------DiscogsImages----------------------------------------

' Easier access of SDB.UI
Set UI = SDB.UI
Dim sTemp
sTemp = SDB.TemporaryFolder

' MediaMonkey calls this method whenever a search is started using this script
Sub StartSearchType(Panel, SearchTerm, SearchArtist, SearchAlbum, SearchType)

	Dim tmpCountry, tmpCountry2, tmpMediaType, tmpMediaType2, tmpMediaFormat, tmpMediaFormat2, tmpYear, tmpYear2
	Dim i, a, tmp, ret
	Set CountryFilterList = SDB.NewStringList
	Set MediaTypeFilterList = SDB.NewStringList
	Set MediaFormatFilterList = SDB.NewStringList
	Set YearFilterList = SDB.NewStringList

	OptionsChanged = False
	NewResult = True
	If SearchType = 1 Then QueryPage = "Discogs"
	If SearchType = 2 Then QueryPage = "MusicBrainz"
	If SearchType = 3 Then QueryPage = "MetalArchives"

	'*FilterList.Item(0) = "0" -> No Filter
	'*FilterList.Item(0) = "1" -> Custom Filter
	'*FilterList.Item(0) = "2" -> Selected Country/MediaType/MediaFormat/Year
	
	FirstRun = True
	
	If (SDB.VersionLo = 1 And SDB.VersionRelease < 28) Or (SDB.VersionLo = 1 And SDB.VersionRelease = 28 And SDB.VersionBuild < 1905) Then
		SDB.MessageBox "You need at least MediaMonkey Version 4.1.28.1905. Your Version is " & SDB.VersionHi & "." & SDB.VersionLo & "." & SDB.VersionRelease & "." & SDB.VersionBuild & " Please update!", mtInformation, Array(mbOk)
	End If
	Set ini = SDB.IniFile
	If Not (ini Is Nothing) Then
		'We init default settings only if they do not exist in ini file yet
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckArtist") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbumArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLabel") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLabel") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDate") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckOrigDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckGenre") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckGenre") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckCountry") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckCountry") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 1
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSmallCover") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckCatalog") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckRelease") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckRelease") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckInvolved") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckGrouping") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckGrouping") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLyricist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckComposer") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckComposer") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckConductor") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckConductor") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckProducer") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckProducer") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDiscNum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTrackNum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckFormat") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckFormat") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckUseAnv") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckForceNumeric") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSidesToDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckForceDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckNoDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = False
		End If
		REM If ini.StringValue("DiscogsAutoTagWeb","CheckUserCollection") = "" Then
			REM ini.BoolValue("DiscogsAutoTagWeb","CheckUserCollection") = False
		REM End If
		If ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = "Custom2"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = "Custom3"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CountryTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CountryTag") = "Custom4"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","FormatTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FormatTag") = "Custom5"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLeadingZero") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLeadingZeroDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZeroDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckVarious") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckVarious") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = "Various Artists"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckFeaturingName") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckFeaturingName") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = "feat."
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckComment") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckComment") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","SubTrackNameSelection") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = "Lyrics By,Words By"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = "Conductor"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = "Producer,Arranged By,Recorded By"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = "Composed By,Score,Written-By,Written By,Music By,Programmed By,Songwriter"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "featuring,feat.,ft.,ft ,feat ,Rap,Rap [Featuring],Vocals [Featuring]"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","UnwantedKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","UnwantedKeywords") = ""
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = ", "
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ArtistLastSeparator") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","FormatSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FormatSeparator") = ", "
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack") = False
		End If

		If ini.ValueExists("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") Then
			ini.DeleteKey "DiscogsAutoTagWeb","CheckNotAlwaysSaveimage"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","AccessToken") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","AccessToken") = ""
		End If
		If ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = ""
		End If
		REM If ini.StringValue("DiscogsAutoTagWeb","QueryPage") = "" Then
			REM ini.StringValue("DiscogsAutoTagWeb","QueryPage") = "Discogs"
		REM End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDontFillEmptyFields") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDontFillEmptyFields") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckNewVersion") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckNewVersion") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","LastCheck") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","LastCheck") = "1"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTheBehindArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTheBehindArtist") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","DiscogsUsername") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","DiscogsUsername") = ""
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDeleteDuplicatedEntry") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDeleteDuplicatedEntry") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","StoreDate") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","StoreDate") = 0
		End If
		If ini.ValueExists("DiscogsAutoTagWeb","StoreOrgDate") Then
			ini.DeleteKey "DiscogsAutoTagWeb","StoreOrgDate"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","LimitReleases") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","LimitReleases") = 0
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckIgnoreFeatArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckIgnoreFeatArtist") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","SubTrackSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","SubTrackSeparator") = ", "
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckReplaceAccentChar") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckReplaceAccentChar") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckNewMatchingEngine") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckNewMatchingEngine") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckShowQtyFormat") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckShowQtyFormat") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSearchWithAlbum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSearchWithAlbum") = True
		End If



		'----------------------------------DiscogsImages----------------------------------------
		CoverStorage = ini.StringValue("PreviewSettings","DefaultCoverStorage")
		'Coverstorage = 0 -> Save image to tag (if possible) otherwise save to file folder
		'Coverstorage = 1 -> Save image to file folder
		'Coverstorage = 2 -> Save image to cover folder (is deprecated and will not be supported !!)
		'Coverstorage = 3 -> Save image to tag (if possible) and to file folder
		If CoverStorage = 2 Then
			Call SDB.MessageBox("Discogs Images: Your Cover Storage is not supported by DiscogsImages !",mtError,Array(mbOk))
			Exit Sub
		End If
		CoverStorageName = ini.StringValue("AAMasks","Mask1")
		
		'----------------------------------DiscogsImages----------------------------------------

	End If

	CheckAlbum = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum")
	CheckArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckArtist")
	CheckAlbumArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist")
	CheckAlbumArtistFirst = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst")
	CheckLabel = ini.BoolValue("DiscogsAutoTagWeb","CheckLabel")
	CheckDate = ini.BoolValue("DiscogsAutoTagWeb","CheckDate")
	CheckOrigDate = ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate")
	CheckGenre = ini.BoolValue("DiscogsAutoTagWeb","CheckGenre")
	CheckCountry = ini.BoolValue("DiscogsAutoTagWeb","CheckCountry")
	CheckSaveImage = ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage")
	CheckSmallCover = ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover")
	CheckCatalog = ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog")
	CheckRelease = ini.BoolValue("DiscogsAutoTagWeb","CheckRelease")
	CheckInvolved = ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved")
	CheckGrouping = ini.BoolValue("DiscogsAutoTagWeb","CheckGrouping")
	CheckLyricist = ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist")
	CheckComposer = ini.BoolValue("DiscogsAutoTagWeb","CheckComposer")
	CheckConductor = ini.BoolValue("DiscogsAutoTagWeb","CheckConductor")
	CheckProducer = ini.BoolValue("DiscogsAutoTagWeb","CheckProducer")
	CheckDiscNum = ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum")
	CheckTrackNum = ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum")
	CheckFormat = ini.BoolValue("DiscogsAutoTagWeb","CheckFormat")
	CheckUseAnv = ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv")
	CheckYearOnlyDate = ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate")
	CheckForceNumeric = ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric")
	CheckSidesToDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc")
	CheckForceDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc")
	CheckOriginalDiscogsTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack")
	CheckNoDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc")
	REM CheckUserCollection = ini.BoolValue("DiscogsAutoTagWeb","CheckUserCollection")
	DiscogsUsername = ini.StringValue("DiscogsAutoTagWeb","DiscogsUsername")
	ReleaseTag = ini.StringValue("DiscogsAutoTagWeb","ReleaseTag")
	CatalogTag = ini.StringValue("DiscogsAutoTagWeb","CatalogTag")
	CountryTag = ini.StringValue("DiscogsAutoTagWeb","CountryTag")
	FormatTag = ini.StringValue("DiscogsAutoTagWeb","FormatTag")
	CheckLeadingZero = ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero")
	CheckLeadingZeroDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZeroDisc")
	CheckVarious = ini.BoolValue("DiscogsAutoTagWeb","CheckVarious")
	TxtVarious = ini.StringValue("DiscogsAutoTagWeb","TxtVarious")
	CheckTitleFeaturing = ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing")
	CheckFeaturingName = ini.boolValue("DiscogsAutoTagWeb","CheckFeaturingName")
	TxtFeaturingName = ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName")
	CheckComment = ini.BoolValue("DiscogsAutoTagWeb","CheckComment")
	SubTrackNameSelection = ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection")
	Separator = ini.StringValue("Appearance","MultiStringSeparator")
	LyricistKeywords = ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords")
	ConductorKeywords = ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords")
	ProducerKeywords = ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords")
	ComposerKeywords = ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords")
	FeaturingKeywords = ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords")
	UnwantedKeywords = ini.StringValue("DiscogsAutoTagWeb","UnwantedKeywords")
	ArtistSeparator = ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator")
	FormatSeparator = ini.StringValue("DiscogsAutoTagWeb","FormatSeparator")
	ArtistLastSeparator = ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator")
	CheckTurnOffSubTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack")
	AccessToken = ini.StringValue("DiscogsAutoTagWeb","AccessToken")
	AccessTokenSecret = ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret")
	REM QueryPage = ini.StringValue("DiscogsAutoTagWeb","QueryPage")
	CheckInvolvedPeopleSingleLine = ini.BoolValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine")
	CheckDontFillEmptyFields = ini.BoolValue("DiscogsAutoTagWeb","CheckDontFillEmptyFields")
	CheckNewVersion = ini.BoolValue("DiscogsAutoTagWeb","CheckNewVersion")
	LastCheck = ini.StringValue("DiscogsAutoTagWeb","LastCheck")
	CheckTheBehindArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckTheBehindArtist")
	CheckDeleteDuplicatedEntry = ini.BoolValue("DiscogsAutoTagWeb","CheckDeleteDuplicatedEntry")
	StoreDate = ini.StringValue("DiscogsAutoTagWeb","StoreDate")
	CheckLimitReleases = ini.StringValue("DiscogsAutoTagWeb","LimitReleases")
	CheckIgnoreFeatArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckIgnoreFeatArtist")
	SubTrackSeparator = ini.StringValue("DiscogsAutoTagWeb","SubTrackSeparator")
	CheckReplaceAccentChar = ini.BoolValue("DiscogsAutoTagWeb","CheckReplaceAccentChar")
	CheckNewMatchingEngine = ini.BoolValue("DiscogsAutoTagWeb","CheckNewMatchingEngine")
	CheckShowQtyFormat = ini.BoolValue("DiscogsAutoTagWeb","CheckShowQtyFormat")
	CheckSearchWithAlbum = ini.BoolValue("DiscogsAutoTagWeb","CheckSearchWithAlbum")

	Separator = Left(Separator, Len(Separator)-1)
	Separator = Right(Separator, Len(Separator)-1)

	CheckAlbumArtistFirst = False

	SelectAll = True

	Set MediaTypeList = SDB.NewStringList
	Set MediaFormatList = SDB.NewStringList
	Set CountryList = SDB.NewStringList
	Set CountryCode = SDB.NewStringList
	Set YearList = SDB.NewStringList
	Set AlternativeList = SDB.NewStringList
	Set LoadList = SDB.NewStringList
	Set RelationAttrList = SDB.NewStringList

	If CheckLimitReleases = 0 Then
		LimitReleases = 50
	ElseIf CheckLimitReleases = 1 Then
		LimitReleases = 100
	ElseIf CheckLimitReleases = 2 Then
		LimitReleases = 200
	ElseIf CheckLimitReleases = 3 Then
		LimitReleases = 400
	End If

	RelationAttrList.Add "Additional"
	RelationAttrList.Add "Assistant"
	RelationAttrList.Add "Associate"
	RelationAttrList.Add "Bonus"
	RelationAttrList.Add "Co"
	RelationAttrList.Add "Cover"
	RelationAttrList.Add "Executive"
	RelationAttrList.Add "Founder"
	RelationAttrList.Add "Guest"
	RelationAttrList.Add "Instrumental"
	RelationAttrList.Add "Live"
	RelationAttrList.Add "Medley"
	RelationAttrList.Add "Solo"

	LoadList.Add "Search Results"
	LoadList.Add "Master Release"
	LoadList.Add "Versions of Master"
	LoadList.Add "Releases of Artist"
	LoadList.Add "Releases of Label"

	MediaTypeList.Add "None"
	MediaTypeList.Add "Vinyl"
	MediaTypeList.Add "CD"
	MediaTypeList.Add "DVD"
	MediaTypeList.Add "Blu-Ray"
	MediaTypeList.Add "Digital Media"
	MediaTypeList.Add "Cassette"
	MediaTypeList.Add "DAT"
	MediaTypeList.Add "Minidisc"
	MediaTypeList.Add "File"
	MediaTypeList.Add "Acetate"
	MediaTypeList.Add "Flexi-disc"
	MediaTypeList.Add "Lathe Cut"
	MediaTypeList.Add "Shellac"
	MediaTypeList.Add "Pathé Disc"
	MediaTypeList.Add "Edison Disc"
	MediaTypeList.Add "Cylinder"
	MediaTypeList.Add "CDr"
	MediaTypeList.Add "CDV"
	MediaTypeList.Add "DVDr"
	MediaTypeList.Add "HD DVD"
	MediaTypeList.Add "HD DVD-R"
	MediaTypeList.Add "Blue-ray-R"
	MediaTypeList.Add "4-Track Cartridge"
	MediaTypeList.Add "8-Track Cartridge"
	MediaTypeList.Add "DCC"
	MediaTypeList.Add "Microcassette"
	MediaTypeList.Add "Reel-To-Reel"
	MediaTypeList.Add "Betamax"
	MediaTypeList.Add "VHS"
	MediaTypeList.Add "Video 2000"
	MediaTypeList.Add "Laserdisc"
	MediaTypeList.Add "SelectaVision"
	MediaTypeList.Add "VHD"
	MediaTypeList.Add "MVD"
	MediaTypeList.Add "UMD"
	MediaTypeList.Add "Floppy Disk"
	MediaTypeList.Add "Memory Stick"
	MediaTypeList.Add "Hybrid"
	MediaTypeList.Add "Box Set"

	MediaFormatList.Add "None"
	MediaFormatList.Add "Album"
	MediaFormatList.Add "Mini-Album"
	MediaFormatList.Add "Compilation"
	MediaFormatList.Add "Single"
	MediaFormatList.Add "Maxi-Single"
	MediaFormatList.Add "7"""
	MediaFormatList.Add "12"""
	MediaFormatList.Add "LP"
	MediaFormatList.Add "EP"
	MediaFormatList.Add "Single Sided"
	MediaFormatList.Add "Enhanced"
	MediaFormatList.Add "Limited Edition"
	MediaFormatList.Add "Reissue"
	MediaFormatList.Add "Remastered"
	MediaFormatList.Add "Repress"
	MediaFormatList.Add "Test Pressing"
	MediaFormatList.Add "Unofficial"
	MediaFormatList.Add "Promo"
	MediaFormatList.Add "White Label"
	MediaFormatList.Add "Mixed"
	MediaFormatList.Add "Sampler"
	MediaFormatList.Add "MP3"
	MediaFormatList.Add "FLAC"
	MediaFormatList.Add "16"""
	MediaFormatList.Add "11"""
	MediaFormatList.Add "10"""
	MediaFormatList.Add "9"""
	MediaFormatList.Add "8"""
	MediaFormatList.Add "6"""
	MediaFormatList.Add "5"""
	MediaFormatList.Add "4"""
	MediaFormatList.Add "3"""
	MediaFormatList.Add "45 RPM"
	MediaFormatList.Add "78 RPM"
	MediaFormatList.Add "Shape"
	MediaFormatList.Add "Card Backed"
	MediaFormatList.Add "Etched"
	MediaFormatList.Add "Picture Disc"
	MediaFormatList.Add "Stereo"
	MediaFormatList.Add "Mono"
	MediaFormatList.Add "Quadraphonic"
	MediaFormatList.Add "Ambisonic"
	MediaFormatList.Add "Mispress"
	MediaFormatList.Add "Misprint"
	MediaFormatList.Add "Partially Mixed"
	MediaFormatList.Add "Unofficial Release"
	MediaFormatList.Add "Partially Unofficial"
	MediaFormatList.Add "Copy Protected"

	CountryList.Add "None"
	CountryList.Add "Australia"
	CountryList.Add "Belgium"
	CountryList.Add "Brazil"
	CountryList.Add "Canada"
	CountryList.Add "China"
	CountryList.Add "Cuba"
	CountryList.Add "France"
	CountryList.Add "Germany"
	CountryList.Add "Italy"
	CountryList.Add "Ireland"
	CountryList.Add "India"
	CountryList.Add "Jamaica"
	CountryList.Add "Japan"
	CountryList.Add "Mexico"
	CountryList.Add "Netherlands"
	CountryList.Add "New Zealand"
	CountryList.Add "Spain"
	CountryList.Add "Sweden"
	CountryList.Add "Switzerland"
	CountryList.Add "UK"
	CountryList.Add "US"
	CountryList.Add "=========="
	CountryList.Add "Worldwide"
	CountryList.Add "Africa"
	CountryList.Add "Asia"
	CountryList.Add "Australasia"
	CountryList.Add "Benelux"
	CountryList.Add "Central America"
	CountryList.Add "Europe"
	CountryList.Add "Gulf Cooperation Council"
	CountryList.Add "North America"
	CountryList.Add "Scandinavia"
	CountryList.Add "South America"
	CountryList.Add "==========="
	CountryList.Add "Afghanistan"
	CountryList.Add "Albania"
	CountryList.Add "Algeria"
	CountryList.Add "American Samoa"
	CountryList.Add "Andorra"
	CountryList.Add "Angola"
	CountryList.Add "Anguilla"
	CountryList.Add "Antarctica"
	CountryList.Add "Antigua and Barbuda"
	CountryList.Add "Argentina"
	CountryList.Add "Armenia"
	CountryList.Add "Aruba"
	CountryList.Add "Austria"
	CountryList.Add "Azerbaijan"
	CountryList.Add "Bahamas"
	CountryList.Add "Bahrain"
	CountryList.Add "Bangladesh"
	CountryList.Add "Barbados"
	CountryList.Add "Belarus"
	CountryList.Add "Belize"
	CountryList.Add "Benin"
	CountryList.Add "Bermuda"
	CountryList.Add "Bhutan"
	CountryList.Add "Bolivia, Plurinational State of"
	CountryList.Add "Bonaire, Sint Eustatius and Saba"
	CountryList.Add "Bosnia and Herzegovina"
	CountryList.Add "Botswana"
	CountryList.Add "Bouvet Island"
	CountryList.Add "British Indian Ocean Territory"
	CountryList.Add "Brunei Darussalam"
	CountryList.Add "Bulgaria"
	CountryList.Add "Burkina Faso"
	CountryList.Add "Burundi"
	CountryList.Add "Cabo Verde"
	CountryList.Add "Cambodia"
	CountryList.Add "Cameroon"
	CountryList.Add "Cayman Islands"
	CountryList.Add "Central African Republic"
	CountryList.Add "Chad"
	CountryList.Add "Chile"
	CountryList.Add "Christmas Island"
	CountryList.Add "Cocos (Keeling) Islands"
	CountryList.Add "Colombia"
	CountryList.Add "Comoros"
	CountryList.Add "Congo"
	CountryList.Add "Congo (the Democratic Republic of the)"
	CountryList.Add "Cook Islands"
	CountryList.Add "Costa Rica"
	CountryList.Add "Côte d'Ivoire"
	CountryList.Add "Croatia"
	CountryList.Add "Curaçao"
	CountryList.Add "Cyprus"
	CountryList.Add "Czech Republic"
	CountryList.Add "Czechoslovakia"
	CountryList.Add "Denmark"
	CountryList.Add "Djibouti"
	CountryList.Add "Dominica"
	CountryList.Add "Dominican Republic"
	CountryList.Add "Ecuador"
	CountryList.Add "Egypt"
	CountryList.Add "El Salvador"
	CountryList.Add "Equatorial Guinea"
	CountryList.Add "Eritrea"
	CountryList.Add "Estonia"
	CountryList.Add "Ethiopia"
	CountryList.Add "Falkland Islands [Malvinas]"
	CountryList.Add "Faroe Islands"
	CountryList.Add "Fiji"
	CountryList.Add "Finland"
	CountryList.Add "French Guiana"
	CountryList.Add "French Polynesia"
	CountryList.Add "French Southern Territories"
	CountryList.Add "Gabon"
	CountryList.Add "Gambia"
	CountryList.Add "Georgia"
	CountryList.Add "Ghana"
	CountryList.Add "Gibraltar"
	CountryList.Add "Greece"
	CountryList.Add "Greenland"
	CountryList.Add "Grenada"
	CountryList.Add "Guadeloupe"
	CountryList.Add "Guam"
	CountryList.Add "Guatemala"
	CountryList.Add "Guernsey"
	CountryList.Add "Guinea"
	CountryList.Add "Guinea-Bissau"
	CountryList.Add "Guyana"
	CountryList.Add "Haiti"
	CountryList.Add "Heard Island and McDonald Islands"
	CountryList.Add "Holy See [Vatican City State]"
	CountryList.Add "Honduras"
	CountryList.Add "Hong Kong"
	CountryList.Add "Hungary"
	CountryList.Add "Iceland"
	CountryList.Add "Indonesia"
	CountryList.Add "Iran (the Islamic Republic of)"
	CountryList.Add "Iraq"
	CountryList.Add "Isle of Man"
	CountryList.Add "Israel"
	CountryList.Add "Jersey"
	CountryList.Add "Jordan"
	CountryList.Add "Kazakhstan"
	CountryList.Add "Kenya"
	CountryList.Add "Kiribati"
	CountryList.Add "Korea (the Democratic People's Republic of)"
	CountryList.Add "Korea (the Republic of)"
	CountryList.Add "Kuwait"
	CountryList.Add "Kyrgyzstan"
	CountryList.Add "Lao People's Democratic Republic"
	CountryList.Add "Latvia"
	CountryList.Add "Lebanon"
	CountryList.Add "Lesotho"
	CountryList.Add "Liberia"
	CountryList.Add "Libya"
	CountryList.Add "Liechtenstein"
	CountryList.Add "Lithuania"
	CountryList.Add "Luxembourg"
	CountryList.Add "Macao"
	CountryList.Add "Macedonia (the former Yugoslav Republic of)"
	CountryList.Add "Madagascar"
	CountryList.Add "Malawi"
	CountryList.Add "Malaysia"
	CountryList.Add "Maldives"
	CountryList.Add "Mali"
	CountryList.Add "Malta"
	CountryList.Add "Marshall Islands"
	CountryList.Add "Martinique"
	CountryList.Add "Mauritania"
	CountryList.Add "Mauritius"
	CountryList.Add "Mayotte"
	CountryList.Add "Micronesia (the Federated States of)"
	CountryList.Add "Moldova (the Republic of)"
	CountryList.Add "Monaco"
	CountryList.Add "Mongolia"
	CountryList.Add "Montenegro"
	CountryList.Add "Montserrat"
	CountryList.Add "Morocco"
	CountryList.Add "Mozambique"
	CountryList.Add "Myanmar"
	CountryList.Add "Namibia"
	CountryList.Add "Nauru"
	CountryList.Add "Nepal"
	CountryList.Add "New Caledonia"
	CountryList.Add "Nicaragua"
	CountryList.Add "Niger"
	CountryList.Add "Nigeria"
	CountryList.Add "Niue"
	CountryList.Add "Norfolk Island"
	CountryList.Add "Northern Mariana Islands"
	CountryList.Add "Norway"
	CountryList.Add "Oman"
	CountryList.Add "Pakistan"
	CountryList.Add "Palau"
	CountryList.Add "Palestine, State of"
	CountryList.Add "Panama"
	CountryList.Add "Papua New Guinea"
	CountryList.Add "Paraguay"
	CountryList.Add "Peru"
	CountryList.Add "Philippines"
	CountryList.Add "Pitcairn"
	CountryList.Add "Poland"
	CountryList.Add "Portugal"
	CountryList.Add "Puerto Rico"
	CountryList.Add "Qatar"
	CountryList.Add "Réunion"
	CountryList.Add "Romania"
	CountryList.Add "Russian Federation"
	CountryList.Add "Rwanda"
	CountryList.Add "Saint Barthélemy"
	CountryList.Add "Saint Helena, Ascension and Tristan da Cunha"
	CountryList.Add "Saint Kitts and Nevis"
	CountryList.Add "Saint Lucia"
	CountryList.Add "Saint Martin (French part)"
	CountryList.Add "Saint Pierre and Miquelon"
	CountryList.Add "Saint Vincent and the Grenadines"
	CountryList.Add "Samoa"
	CountryList.Add "San Marino"
	CountryList.Add "Sao Tome and Principe"
	CountryList.Add "Saudi Arabia"
	CountryList.Add "Senegal"
	CountryList.Add "Serbia"
	CountryList.Add "Seychelles"
	CountryList.Add "Sierra Leone"
	CountryList.Add "Singapore"
	CountryList.Add "Sint Maarten (Dutch part)"
	CountryList.Add "Slovakia"
	CountryList.Add "Slovenia"
	CountryList.Add "Solomon Islands"
	CountryList.Add "Somalia"
	CountryList.Add "South Africa"
	CountryList.Add "South Georgia and the South Sandwich Islands"
	CountryList.Add "South Sudan "
	CountryList.Add "Sri Lanka"
	CountryList.Add "Sudan"
	CountryList.Add "Suriname"
	CountryList.Add "Svalbard and Jan Mayen"
	CountryList.Add "Swaziland"
	CountryList.Add "Syrian Arab Republic"
	CountryList.Add "Taiwan (Province of China)"
	CountryList.Add "Tajikistan"
	CountryList.Add "Tanzania, United Republic of"
	CountryList.Add "Thailand"
	CountryList.Add "Timor-Leste"
	CountryList.Add "Togo"
	CountryList.Add "Tokelau"
	CountryList.Add "Tonga"
	CountryList.Add "Trinidad and Tobago"
	CountryList.Add "Tunisia"
	CountryList.Add "Turkey"
	CountryList.Add "Turkmenistan"
	CountryList.Add "Turks and Caicos Islands"
	CountryList.Add "Tuvalu"
	CountryList.Add "Uganda"
	CountryList.Add "Ukraine"
	CountryList.Add "United Arab Emirates"
	CountryList.Add "United States Minor Outlying Islands"
	CountryList.Add "Uruguay"
	CountryList.Add "Uzbekistan"
	CountryList.Add "Vanuatu"
	CountryList.Add "Venezuela, Bolivarian Republic of "
	CountryList.Add "Viet Nam"
	CountryList.Add "Virgin Islands (British)"
	CountryList.Add "Virgin Islands (U.S.)"
	CountryList.Add "Wallis and Futuna"
	CountryList.Add "Western Sahara"
	CountryList.Add "Yemen"
	CountryList.Add "Zambia"
	CountryList.Add "Zimbabwe"

	CountryCode.Add ""
	CountryCode.Add "AU"
	CountryCode.Add "BE"
	CountryCode.Add "BR"
	CountryCode.Add "CA"
	CountryCode.Add "CN"
	CountryCode.Add "CU"
	CountryCode.Add "FR"
	CountryCode.Add "DE"
	CountryCode.Add "IT"
	CountryCode.Add "IE"
	CountryCode.Add "IN"
	CountryCode.Add "JM"
	CountryCode.Add "JP"
	CountryCode.Add "MX"
	CountryCode.Add "NL"
	CountryCode.Add "NZ"
	CountryCode.Add "ES"
	CountryCode.Add "SE"
	CountryCode.Add "CH"
	CountryCode.Add "GB"
	CountryCode.Add "US"
	CountryCode.Add ""
	CountryCode.Add "XW"
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add "XE"
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add ""
	CountryCode.Add "AF"
	CountryCode.Add "AL"
	CountryCode.Add "DZ"
	CountryCode.Add "AS"
	CountryCode.Add "AD"
	CountryCode.Add "AO"
	CountryCode.Add "AI"
	CountryCode.Add "AQ"
	CountryCode.Add "AG"
	CountryCode.Add "AR"
	CountryCode.Add "AM"
	CountryCode.Add "AW"
	CountryCode.Add "AT"
	CountryCode.Add "AZ"
	CountryCode.Add "BS"
	CountryCode.Add "BH"
	CountryCode.Add "BD"
	CountryCode.Add "BB"
	CountryCode.Add "BY"
	CountryCode.Add "BZ"
	CountryCode.Add "BJ"
	CountryCode.Add "BM"
	CountryCode.Add "BT"
	CountryCode.Add "BO"
	CountryCode.Add "BQ"
	CountryCode.Add "BA"
	CountryCode.Add "BW"
	CountryCode.Add "BV"
	CountryCode.Add "IO"
	CountryCode.Add "BN"
	CountryCode.Add "BG"
	CountryCode.Add "BF"
	CountryCode.Add "BI"
	CountryCode.Add "CV"
	CountryCode.Add "KH"
	CountryCode.Add "CM"
	CountryCode.Add "KY"
	CountryCode.Add "CF"
	CountryCode.Add "TD"
	CountryCode.Add "CL"
	CountryCode.Add "CX"
	CountryCode.Add "CC"
	CountryCode.Add "CO"
	CountryCode.Add "KM"
	CountryCode.Add "CG"
	CountryCode.Add "CD"
	CountryCode.Add "CK"
	CountryCode.Add "CR"
	CountryCode.Add "CI"
	CountryCode.Add "HR"
	CountryCode.Add "CW"
	CountryCode.Add "CY"
	CountryCode.Add "CZ"
	CountryCode.Add "XC"
	CountryCode.Add "DK"
	CountryCode.Add "DJ"
	CountryCode.Add "DM"
	CountryCode.Add "DO"
	CountryCode.Add "EC"
	CountryCode.Add "EG"
	CountryCode.Add "SV"
	CountryCode.Add "GQ"
	CountryCode.Add "ER"
	CountryCode.Add "EE"
	CountryCode.Add "ET"
	CountryCode.Add "FK"
	CountryCode.Add "FO"
	CountryCode.Add "FJ"
	CountryCode.Add "FI"
	CountryCode.Add "GF"
	CountryCode.Add "PF"
	CountryCode.Add "TF"
	CountryCode.Add "GA"
	CountryCode.Add "GM"
	CountryCode.Add "GE"
	CountryCode.Add "GH"
	CountryCode.Add "GI"
	CountryCode.Add "GR"
	CountryCode.Add "GL"
	CountryCode.Add "GD"
	CountryCode.Add "GP"
	CountryCode.Add "GU"
	CountryCode.Add "GT"
	CountryCode.Add "GG"
	CountryCode.Add "GN"
	CountryCode.Add "GW"
	CountryCode.Add "GY"
	CountryCode.Add "HT"
	CountryCode.Add "HM"
	CountryCode.Add "VA"
	CountryCode.Add "HN"
	CountryCode.Add "HK"
	CountryCode.Add "HU"
	CountryCode.Add "IS"
	CountryCode.Add "ID"
	CountryCode.Add "IR"
	CountryCode.Add "IQ"
	CountryCode.Add "IM"
	CountryCode.Add "IL"
	CountryCode.Add "JE"
	CountryCode.Add "JO"
	CountryCode.Add "KZ"
	CountryCode.Add "KE"
	CountryCode.Add "KI"
	CountryCode.Add "KP"
	CountryCode.Add "KR"
	CountryCode.Add "KW"
	CountryCode.Add "KG"
	CountryCode.Add "LA"
	CountryCode.Add "LV"
	CountryCode.Add "LB"
	CountryCode.Add "LS"
	CountryCode.Add "LR"
	CountryCode.Add "LY"
	CountryCode.Add "LI"
	CountryCode.Add "LT"
	CountryCode.Add "LU"
	CountryCode.Add "MO"
	CountryCode.Add "MK"
	CountryCode.Add "MG"
	CountryCode.Add "MW"
	CountryCode.Add "MY"
	CountryCode.Add "MV"
	CountryCode.Add "ML"
	CountryCode.Add "MT"
	CountryCode.Add "MH"
	CountryCode.Add "MQ"
	CountryCode.Add "MR"
	CountryCode.Add "MU"
	CountryCode.Add "YT"
	CountryCode.Add "FM"
	CountryCode.Add "MD"
	CountryCode.Add "MC"
	CountryCode.Add "MN"
	CountryCode.Add "ME"
	CountryCode.Add "MS"
	CountryCode.Add "MA"
	CountryCode.Add "MZ"
	CountryCode.Add "MM"
	CountryCode.Add "NA"
	CountryCode.Add "NR"
	CountryCode.Add "NP"
	CountryCode.Add "NC"
	CountryCode.Add "NI"
	CountryCode.Add "NE"
	CountryCode.Add "NG"
	CountryCode.Add "NU"
	CountryCode.Add "NF"
	CountryCode.Add "MP"
	CountryCode.Add "NO"
	CountryCode.Add "OM"
	CountryCode.Add "PK"
	CountryCode.Add "PW"
	CountryCode.Add "PS"
	CountryCode.Add "PA"
	CountryCode.Add "PG"
	CountryCode.Add "PY"
	CountryCode.Add "PE"
	CountryCode.Add "PH"
	CountryCode.Add "PN"
	CountryCode.Add "PL"
	CountryCode.Add "PT"
	CountryCode.Add "PR"
	CountryCode.Add "QA"
	CountryCode.Add "RE"
	CountryCode.Add "RO"
	CountryCode.Add "RU"
	CountryCode.Add "RW"
	CountryCode.Add "BL"
	CountryCode.Add "SH"
	CountryCode.Add "KN"
	CountryCode.Add "LC"
	CountryCode.Add "MF"
	CountryCode.Add "PM"
	CountryCode.Add "VC"
	CountryCode.Add "WS"
	CountryCode.Add "SM"
	CountryCode.Add "ST"
	CountryCode.Add "SA"
	CountryCode.Add "SN"
	CountryCode.Add "RS"
	CountryCode.Add "SC"
	CountryCode.Add "SL"
	CountryCode.Add "SG"
	CountryCode.Add "SX"
	CountryCode.Add "SK"
	CountryCode.Add "SI"
	CountryCode.Add "SB"
	CountryCode.Add "SO"
	CountryCode.Add "ZA"
	CountryCode.Add "GS"
	CountryCode.Add "SS"
	CountryCode.Add "LK"
	CountryCode.Add "SD"
	CountryCode.Add "SR"
	CountryCode.Add "SJ"
	CountryCode.Add "SZ"
	CountryCode.Add "SY"
	CountryCode.Add "TW"
	CountryCode.Add "TJ"
	CountryCode.Add "TZ"
	CountryCode.Add "TH"
	CountryCode.Add "TL"
	CountryCode.Add "TG"
	CountryCode.Add "TK"
	CountryCode.Add "TO"
	CountryCode.Add "TT"
	CountryCode.Add "TN"
	CountryCode.Add "TR"
	CountryCode.Add "TM"
	CountryCode.Add "TC"
	CountryCode.Add "TV"
	CountryCode.Add "UG"
	CountryCode.Add "UA"
	CountryCode.Add "AE"
	CountryCode.Add "UM"
	CountryCode.Add "UY"
	CountryCode.Add "UZ"
	CountryCode.Add "VU"
	CountryCode.Add "VE"
	CountryCode.Add "VN"
	CountryCode.Add "VG"
	CountryCode.Add "VI"
	CountryCode.Add "WF"
	CountryCode.Add "EH"
	CountryCode.Add "YE"
	CountryCode.Add "ZM"
	CountryCode.Add "ZW"

	YearList.Add "None"
	For i=Year(Date) To 1900 Step -1
		YearList.Add i
	Next

	If ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = "" Then
		tmp = "0"
		For a = 1 To CountryList.Count - 1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = tmp
	End If

	If ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = "" Then
		tmp = "0"
		For a = 1 To MediaFormatList.Count - 1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = tmp
	End If

	If ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = "" Then
		tmp = "0"
		For a = 1 To MediaFormatList.Count - 1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = tmp
	End If

	If ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = "" Then
		tmp = "0"
		For a = Year(Date) To 1900 Step -1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmp
	End If
	
	tmpCountry = ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter")
	tmpCountry2 = Split(tmpCountry, ",")
	If UBound(tmpCountry2)+1 <> CountryList.Count Then
		tmp = "0"
		For a = 1 To CountryList.Count - 1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = tmp
		tmpCountry2 = Split(tmp, ",")
	End If

	tmpMediaType = ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter")
	tmpMediaType2 = Split(tmpMediaType, ",")
	If UBound(tmpMediaType2)+1 <> MediaTypeList.Count Then
		tmp = "0"
		For a = 1 To MediaTypeList.Count - 1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = tmp
		tmpMediaType2 = Split(tmp, ",")
	End If
	
	tmpMediaFormat = ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter")
	tmpMediaFormat2 = Split(tmpMediaFormat, ",")
	If UBound(tmpMediaFormat2)+1 <> MediaFormatList.Count Then
		tmp = "0"
		For a = 1 To MediaFormatList.Count - 1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = tmp
		tmpMediaFormat2 = Split(tmp, ",")
	End If
	
	tmpYear = ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter")
	tmpYear2 = Split(tmpYear, ",")
	If UBound(tmpYear2)+1 <> YearList.Count Then
		tmp = "0"
		For a = 1 To YearList.Count - 1
			tmp = tmp & ",0"
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmp
		tmpYear2 = Split(tmp, ",")
	End If

	For a = 0 To CountryList.Count - 1
		CountryFilterList.Add tmpCountry2(a)
	Next

	For a = 0 To MediaTypeList.Count - 1
		MediaTypeFilterList.Add tmpMediaType2(a)
	Next

	For a = 0 To MediaFormatList.Count - 1
		MediaFormatFilterList.Add tmpMediaFormat2(a)
	Next

	For a = 0 To YearList.Count - 1
		YearFilterList.Add tmpYear2(a)
	Next

	For i = 0 to 999
		UnselectedTracks(i) = ""
		UnselectedTrackNames(i) = ""
	Next

	trackrefresh = True


	SearchTerm = Trim(SearchTerm)

	WriteLogInit 'Initialize Log-File
	WriteLog " "
	WriteLog "SearchTerm=" & SearchTerm
	WriteLog "SearchArtist=" & SearchArtist
	WriteLog "SearchAlbum=" & SearchAlbum
	WriteLog " "

	If SearchArtist <> "" And SearchAlbum <> "" Then
		tmp = SDB.Tools.WebSearch.NewTracks.Item(0).AlbumArtistName
		If tmp = "Various" or tmp = "Various Artists" or tmp = TxtVarious Then
			AddAlternative TxtVarious & " - " & SearchAlbum
			SearchTerm = "Various " & SearchAlbum
			SearchArtist = "Various"
			WriteLog "SearchTerm changed to: " & SearchTerm
			WriteLog "SearchArtist changed to: Various"
		Else
			AddAlternative SearchTerm
			AddAlternative SearchArtist
			AddAlternative SearchAlbum
		End If
	End If

	NewSearchTerm = SearchTerm

	For i = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1
		AddAlternatives SDB.Tools.WebSearch.NewTracks.item(i)
	Next


	If MediaTypeFilterList.Item(0) = "0" Then
		FilterMediaType = "None"
	ElseIf MediaTypeFilterList.Item(0) = "1" Then
		FilterMediaType = "Use MediaType Filter"
	Else
		FilterMediaType = MediaTypeFilterList.Item(0)
	End If

	If MediaFormatFilterList.Item(0) = "0" Then
		FilterMediaFormat = "None"
	ElseIf MediaFormatFilterList.Item(0) = "1" Then
		FilterMediaFormat = "Use MediaFormat Filter"
	Else
		FilterMediaFormat = MediaFormatFilterList.Item(0)
	End If

	If CountryFilterList.Item(0) = "0" Then
		FilterCountry = "None"
	ElseIf CountryFilterList.Item(0) = "1" Then
		FilterCountry = "Use Country Filter"
	Else
		FilterCountry = CountryFilterList.Item(0)
	End If

	If YearFilterList.Item(0) = "0" Then
		FilterYear = "None"
	ElseIf YearFilterList.Item(0) = "1" Then
		FilterYear = "Use Year Filter"
	Else
		FilterYear = YearFilterList.Item(0)
	End If

	CurrentLoadType = "Search Results"

	If SDB.Tools.WebSearch.NewTracks.Count > 0 Then
		Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
		SavedReleaseId = get_release_ID(FirstTrack) 'get saved Release_ID from User-Defined Custom-Tag
		SavedSearchArtist = SearchArtist
		SavedSearchAlbum = SearchAlbum
		NewSearchArtist = SearchArtist
		NewSearchAlbum = SearchAlbum
	End If
	
	UserAgent = "MediaMonkeyDiscogsAutoTagWeb/" & Mid(VersionStr, 2) & " +http://www.mediamonkey.com"

	If CheckNewVersion = True Then
		Dim iDate, Version, colNodes, objNode
		iDate = CDate(Now())
		If Int(LastCheck) <> Int(DatePart("d",iDate)) Then
			Dim xmlDoc
			Set xmlDoc = CreateObject("Microsoft.XMLDOM")

			xmlDoc.Async = "False"
			xmlDoc.Load("https://www.germanc64.de/mm/DiscogsAutoTagWeb.xml")

			Set colNodes = xmlDoc.selectNodes("//SoftwareVersion/VersionMajor")
			For Each objNode in colNodes
				Version = "v" & objNode.Text
			Next
			Set colNodes = xmlDoc.selectNodes("//SoftwareVersion/VersionMinor")
			For Each objNode in colNodes
				Version = Version & "." & objNode.Text 
			Next
			Set colNodes = xmlDoc.selectNodes("//SoftwareVersion/VersionRelease")
			For Each objNode in colNodes
				Version = Version & objNode.Text
			Next
			ini.StringValue("DiscogsAutoTagWeb","LastCheck") = DatePart("d",iDate)

			If Version <> VersionStr Then
				SDB.MessageBox "A new version " & Version & " is released. Please download it via Tools->Extensions", mtInformation, Array(mbOk)
			End If
		End If
	End If

	Dim AlbumArt
	CheckCover = False
	SmallCover = False
	Set AlbumArt = FirstTrack.AlbumArt
	If CheckSaveImage = 0 Or CheckSaveImage = 1 Then
		If (AlbumArt.Count = 0 And CheckSaveImage = 1) Or CheckSaveImage = 0 Then
			If CheckSmallCover = True Then
				SmallCover = True
			End  If
			CheckCover = True
		End If
	End If

	Dim itmAlbum
	For a = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1
		Set tmp = SDB.Tools.WebSearch.NewTracks.item(a)
		Set itmAlbum = tmp.Album
		WriteLog "Disc=" & tmp.DiscNumberStr & " / Track=" & tmp.TrackOrderStr & " / AlbumID=" & itmAlbum.ID & " / Artist=" & tmp.ArtistName & " / Album=" & tmp.AlbumName & " / Title=" & tmp.Title
		If a = 0 Then NewSearchTrack = tmp.Title
	Next
	WriteLog " "
	
	' This is a web browser that we use to present results to the user
	Set WebBrowser = UI.NewActiveX(Panel, "Shell.Explorer")
	WebBrowser.Common.Align = 5
	WebBrowser.Common.ControlName = "WebBrowser"
	WebBrowser.Common.Top = 0
	WebBrowser.Common.Left = 0
	SDB.Objects("WebBrowser") = WebBrowser
	WebBrowser.Interf.Visible = true
	WebBrowser.Common.BringToFront

	If QueryPage = "Discogs" Then
		ret = authorize_script(False)
	End If

	If ret = True Or QueryPage <> "Discogs" Then
		FindResults NewSearchTerm, QueryPage
	End If

End Sub


Sub FindResults(SearchTerm, QueryPage)

	SearchTerm = LTrim(SearchTerm)
	WriteLog "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-"
	WriteLog "Start FindResults"
	WriteLog "SearchTerm=" & SearchTerm
	WriteLog "NewSearchArtist=" & NewSearchArtist
	WriteLog "NewSearchAlbum=" & NewSearchAlbum
	WriteLog "NewSearchTrack=" & NewSearchTrack

	Dim FilterFound, a, searchURL, searchURL_F, searchURL_L, SearchFor
	Dim SendArtist, SendAlbum, SendTrack, SendType, SendDBSearch, SendPerPage
	Dim tmp

	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	SDB.Tools.WebSearch.ClearTracksData
	Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)

	If (InStr(SearchTerm," - [search by release id]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search by release id]")-1)
	End If

	If (InStr(SearchTerm," - [search by release url]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search by release url]")-1)
	End If

	If (InStr(SearchTerm," - [currently tagged with this release]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [currently tagged with this release]")-1)
	End If

	If (InStr(SearchTerm," - [search returned no results]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search returned no results]")-1)
	End If

	If (InStr(SearchTerm," - [search that yielded error]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search that yielded error]")-1)
	End If

	' Handle direct urls

	If (InStr(SearchTerm,"/master/") > 0) Then
		CurrentLoadType = "Master Release"
		LoadMasterResults Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		Exit Sub
	End If

	If (InStr(SearchTerm,"/artist/") > 0) Then
		CurrentLoadType = "Releases of Artist"
		tmp = Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		tmp = Left(tmp, InStr(tmp,"-")-1)
		LoadArtistResults tmp
		Exit Sub
	End If

	If (InStr(SearchTerm,"/label/") > 0) Then
		CurrentLoadType = "Releases of Label"
		tmp = Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		tmp = Left(tmp, InStr(tmp,"-")-1)
		LoadLabelResults tmp
		Exit Sub
	End If
	
	If Left(SearchTerm, 2) = "[m" Then
		CurrentLoadType = "Master Release"
		LoadMasterResults Mid(SearchTerm, 3, Len(SearchTerm)-3)
		Exit Sub
	End If
	
	If Left(SearchTerm, 2) = "[a" Then
		CurrentLoadType = "Releases of Artist"
		LoadArtistResults Mid(SearchTerm, 3, Len(SearchTerm)-3)
		Exit Sub
	End If
	
	If Left(SearchTerm, 2) = "[l" Then
		CurrentLoadType = "Releases of Label"
		LoadLabelResults Mid(SearchTerm, 3, Len(SearchTerm)-3)
		Exit Sub
	End If
	
	If Left(SearchTerm, 2) = "[r" Then
		CurrentLoadType = "Search Results"
		SearchTerm = Mid(SearchTerm, 3, Len(SearchTerm)-3)
	End If

	SDB.ProcessMessages

	If SearchTerm = "" and NewSearchArtist = "" And NewSearchAlbum = "" And NewSearchTrack = "" Then
		ErrorMessage = "No search term"
		WriteLog "No search term"
	ElseIf IsNumeric(SearchTerm) Then
		Results.Add SearchTerm & " - [search by release id]"
		ResultsReleaseID.Add SearchTerm
	ElseIf Len(SearchTerm) = 36 And Mid(SearchTerm, 9, 1) = "-" And Mid(SearchTerm, 14, 1) = "-" Then
		Results.Add SearchTerm & " - [search by release id]"
		ResultsReleaseID.Add SearchTerm
	ElseIf (InStr(SearchTerm,"/release/") > 0) Then
		Results.Add SearchTerm & " - [search by release url]"
		tmp = Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		If InStr(tmp,"-") > 0 Then
			tmp = Left(tmp, InStr(tmp,"-")-1)
		End If
		ResultsReleaseID.Add tmp
	Else

		If QueryPage = "MetalArchives" Then
			'searchURL = "http://www.metal-archives.com/search/ajax-advanced/searching/albums?bandName=" & URLEncodeUTF8(CleanSearchString(SavedSearchArtist)) & "&exactBandMatch=1&releaseTitle=" & URLEncodeUTF8(CleanSearchString(SavedSearchAlbum)) & "&exactReleaseMatch=1&releaseYearFrom=&releaseMonthFrom=&releaseYearTo=&releaseMonthTo=&country=&location=&releaseLabelName=&genre=#albums"
			searchURL = "https://www.metal-archives.com/search/ajax-advanced/searching/albums/?bandName=" & URLEncodeUTF8(CleanSearchString(NewSearchArtist)) & "&exactBandMatch=1&releaseTitle=" & URLEncodeUTF8(CleanSearchString(NewSearchAlbum)) & "&exactReleaseMatch=1&releaseYearFrom=&releaseMonthFrom=&releaseYearTo=&releaseMonthTo=&country=&location=&releaseLabelName=&genre="
			JSONParser_find_result searchURL, "", "", "", "", "", "", "", "MetalArchives", False
		End If

		If QueryPage = "Discogs" Then
			If IsNumeric(SavedReleaseId) Then
				Results.Add FirstTrack.AlbumArtistName & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
				ResultsReleaseID.Add SavedReleaseId
			End If

			SendType = "release"
			SendPerPage = "100"

			If SearchTerm <> "" And NewSearchArtist = "" And NewSearchAlbum = "" Then
				REM Search from user-typed input
				WriteLog "Search from user-typed input"
				SendDBSearch = URLEncodeUTF8(CleanSearchString(SearchTerm))
			ElseIf NewSearchArtist <> "" Or NewSearchAlbum <> "" Or NewSearchTrack <> "" Then
				SendArtist = URLEncodeUTF8(CleanSearchString(NewSearchArtist))
				If CheckSearchWithAlbum = True Then
					SendAlbum = URLEncodeUTF8(CleanSearchString(NewSearchAlbum))
				Else
					SendAlbum = ""
				End If
				If SendAlbum = "" Then
					SendTrack = URLEncodeUTF8(CleanSearchString(NewSearchTrack))
				End If
			Else
				ErrorMessage = "No SearchTerm found"
				WriteLog "No SearchTerm found"
			End If

			If ErrorMessage = "" Then
				JSONParser_find_result "", "results", SendArtist, SendAlbum, SendTrack, SendType, SendDBSearch, SendPerPage, QueryPage, True
			End If
		End If

		If QueryPage = "MusicBrainz" Then
			If Len(SavedReleaseId) = 36 Then
				Results.Add FirstTrack.AlbumArtistName & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
				ResultsReleaseID.Add SavedReleaseId
			End If
			WriteLog "searchTerm=" & SearchTerm
			WriteLog "newsearchTerm=" & NewSearchTerm
			If SearchTerm <> "" And NewSearchArtist = "" And NewSearchAlbum = "" Then
				REM Search from user-typed input
				WriteLog "Search from user-typed input"
				SearchFor = ShowSearchFor()
				If SearchFor = 1 Then
					searchURL = "http://musicbrainz.org/ws/2/release?query=artist:" & Chr(34) & CleanSearchString(URLEncodeUTF8(SearchTerm)) & Chr(34) & "&limit=50&offset=0&fmt=json"
				ElseIf SearchFor = 2 Then
					searchURL = "http://musicbrainz.org/ws/2/release?query=release:" & Chr(34) & CleanSearchString(URLEncodeUTF8(SearchTerm)) & Chr(34) & "&limit=50&offset=0&fmt=json"
				ElseIf SearchFor = 3 Then
					a = Split(SearchTerm, " - ")
					If UBound(a) = 1 Then
						searchURL = "http://musicbrainz.org/ws/2/release?query=artist:" & Chr(34) & CleanSearchString(URLEncodeUTF8(a(0))) & Chr(34) & " AND release:" & Chr(34) & URLEncodeUTF8(CleanSearchString(a(1))) & Chr(34) & "&limit=50&offset=0&fmt=json"
					Else
						SDB.MessageBox "Please use this format: Artist - Album", mtInformation, Array(mbOk)
					End If
				End If
				WriteLog "SearchFor= " & SearchFor
			ElseIf (NewSearchArtist <> "" And NewSearchAlbum = "") Or (SearchTerm <> "" And SearchTerm = NewSearchArtist) Then
				searchURL = "http://musicbrainz.org/ws/2/release?query=artist:" & Chr(34) & CleanSearchString(URLEncodeUTF8(NewSearchArtist)) & Chr(34) & "&limit=50&offset=0&fmt=json"
			ElseIf NewSearchArtist <> "" And NewSearchAlbum <> "" Then
				searchURL = "http://musicbrainz.org/ws/2/release?query=artist:" & Chr(34) & CleanSearchString(URLEncodeUTF8(NewSearchArtist)) & Chr(34) & " AND release:" & Chr(34) & URLEncodeUTF8(CleanSearchString(NewSearchAlbum)) & Chr(34) & "&limit=50&offset=0&fmt=json"
			ElseIf NewSearchArtist = "" And NewSearchAlbum <> "" Then
				searchURL = "http://musicbrainz.org/ws/2/release?query=release:" & Chr(34) & CleanSearchString(URLEncodeUTF8(NewSearchAlbum)) & Chr(34) & "&limit=50&offset=0&fmt=json"
			End If

			If searchURL <> "" Then
				WriteLog "searchURL=" & searchURL
				JSONParser_find_result searchURL, "releases", "", "", "", "", "", "", "MusicBrainz", False
			End If
		End If

		If ResultsReleaseID.Count = 0 And ErrorMessage = "" Then
			ErrorMessage = "Search returned no results" & "<br/>"
			WriteLog "ResultsReleaseID=0"
			If FilterCountry <> "None" Then
				ErrorMessage = ErrorMessage & "Maybe wrong Country Filter set ?" & "<br/>"
			End If

			If FilterMediaType <> "None" Then
				ErrorMessage = ErrorMessage & "Maybe wrong MediaType Filter set ?" & "<br/>"
			End If

			If FilterMediaFormat <> "None" Then
				ErrorMessage = ErrorMessage & "Maybe wrong MediaFormat Filter set ?" & "<br/>"
			End If

			If FilterYear <> "None" Then
				ErrorMessage = ErrorMessage & "Maybe wrong Year Filter set ?" & "<br/>"
			End If

			Results.Add SearchTerm & " - [search returned no results]"
			ResultsReleaseID.Add ""
		End If

	End If

	SDB.ProcessMessages
	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub

Sub LoadMasterResults(MasterID)

	Dim searchURL
	Dim oXMLHTTP
	Dim json
	Set json = New VbsJson

	Dim title, v_year, artist, artistName, main_release, ReleaseDesc, currentArtist, AlbumArtistTitle, tmp

	WriteLog " "
	WriteLog "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+"

	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")

	WriteLog "Load MasterResult"
	ErrorMessage = ""

	If MasterID = "" Then
		ErrorMessage = "Cannot load empty master release"
	Else
		If QueryPage = "MusicBrainz" Then
			ErrorMessage = "Cannot load master release from MusicBrainz"

		ElseIf QueryPage = "Discogs" Then
			searchURL = "https://api.discogs.com/masters/" & MasterID
			WriteLog "searchURL=" & searchURL

			Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")   
			oXMLHTTP.open "GET", searchURL, false
			oXMLHTTP.setRequestHeader "Content-Type","application/json"
			oXMLHTTP.setRequestHeader "User-Agent",UserAgent
			oXMLHTTP.send()

			If oXMLHTTP.Status = 200 Then
				WriteLog "responseText=" & oXMLHTTP.responseText
				
				If InStr(oXMLHTTP.responseText, "OAuth client error") <> 0 Then
					ErrorMessage = oXMLHTTP.responseText
				Else
					Set currentRelease = json.Decode(oXMLHTTP.responseText)

					Set Results = SDB.NewStringList
					Set ResultsReleaseID = SDB.NewStringList

					title = ""
					v_year = ""
					artist = ""
					main_release = ""

					SDB.ProcessMessages

					title = CurrentRelease("title")
					tmp = getArtistsName(CurrentRelease, "artists", QueryPage)
					AlbumArtistTitle = tmp(0)

					If CurrentRelease.Exists("main_release") Then
						main_release = CurrentRelease("main_release")
					End If
					If CurrentRelease.Exists("year") Then
						v_year = CurrentRelease("year")
					End If

					If AlbumArtistTitle <> "" Then ReleaseDesc = AlbumArtistTitle End If
					If AlbumArtistTitle <> "" and title <> "" Then ReleaseDesc = ReleaseDesc & " -" End If
					If title <> "" Then ReleaseDesc = ReleaseDesc & " " & title End If
					If v_year <> "" Then ReleaseDesc = ReleaseDesc & " (" & v_year & ")" End If
					ReleaseDesc = ReleaseDesc & " (Master)"

					Results.Add ReleaseDesc
					ResultsReleaseID.Add CurrentRelease("id")

					SDB.Tools.WebSearch.SetSearchResults Results
					SDB.Tools.WebSearch.ResultIndex = 0

					ReloadResults
				End If
			Else
				ErrorMessage = "Try loading Master returns ErrorCode: " & oXMLHTTP.Status
				WriteLog "Try loading Master returns ErrorCode: " & oXMLHTTP.Status
			End If
		End If
	End If

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub


Sub LoadVersionResults(MasterID)

	WriteLog "VersionResult"

	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	If MasterID = "" Then
		ErrorMessage = "Cannot load empty master release"
	Else
		If IsNumeric(SavedReleaseId) Then
			Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
			Results.Add FirstTrack.AlbumArtistName & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add SavedReleaseId
		End If
		If QueryPage = "Discogs" Then
			JSONParser_find_result "https://api.discogs.com/masters/" & MasterID & "/versions?per_page=100", "versions", "", "", "", "", "", "", "Discogs", False
		Else
			ErrorMessage = "Cannot load master release from MusicBrainz"
		End If
	End If

	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub


Sub LoadArtistResults(ArtistId)

	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	If ArtistId = "" Then
		ErrorMessage = "Cannot load empty artist"
	Else
		If IsNumeric(SavedReleaseId) Then
			Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
			Results.Add FirstTrack.AlbumArtistName & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add SavedReleaseId
		End If

		If QueryPage = "Discogs" Then
			JSONParser_find_result "https://api.discogs.com/artists/" & ArtistId & "/releases?per_page=100", "releases", "", "", "", "", "", "", "Discogs", False
		ElseIf QueryPage = "MusicBrainz" Then
			JSONParser_find_result "http://musicbrainz.org/ws/2/release?artist=" & ArtistId & "&inc=artist-credits+release-groups+media&fmt=json&limit=100", "Artist", "", "", "", "", "", "", "MusicBrainz", False
		End If
	End If

	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub


Sub LoadLabelResults(LabelId)

	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	If LabelId = "" Then
		ErrorMessage = "Cannot load empty label"
	Else
		If IsNumeric(SavedReleaseId) Then
			Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
			Results.Add FirstTrack.AlbumArtistName & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add SavedReleaseId
		End If

		If QueryPage = "Discogs" Then
			JSONParser_find_result "https://api.discogs.com/labels/" & LabelId &  "/releases?per_page=100", "releases", "", "", "", "", "", "", "Discogs", False
		ElseIf QueryPage = "MusicBrainz" Then
			JSONParser_find_result "http://musicbrainz.org/ws/2/release?label=" & LabelId & "&inc=artist-credits+media&fmt=json&limit=100", "Label", "", "", "", "", "", "", "MusicBrainz", False
		End If
	End If

	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub


'For reloading results
Sub ReloadResults

	Dim Tracks, TracksNum, DiscogsTracksNum, TracksCD, ArtistTitles, InvolvedArtists, Lyricists, Composers, Conductors, Producers, Durations
	Dim AlbumArtist, AlbumArtistTitle, AlbumLyricist, AlbumComposer, AlbumConductor, AlbumProducer, AlbumInvolved, AlbumFeaturing, AlbumTitle
	Dim track, currentTrack, position, artist, currentArtist, artistName, extraArtist, extra
	Dim currentImage, currentLabel, currentFormat, theMaster, i, g, l, s, f, d, currentMedia, m, t, x
	Dim ReleaseDate, ReleaseSplit, theLabels, theCatalogs, theCountry, theFormat
	Dim rTrackPosition, rSubPosition
	Dim oXMLHTTP, searchURL
	Dim tmpArt, image
	Dim Genres, Comment, DataQuality
	Dim NoSubTrackUsing, oldSubTrackNumber
	Dim templateHTMLDoc
	Set Tracks = SDB.NewStringList
	Set TracksNum = SDB.NewStringList
	Set DiscogsTracksNum = SDB.NewStringList
	Set TracksCD = SDB.NewStringList
	Set ArtistTitles = SDB.NewStringList
	Set InvolvedArtists = SDB.NewStringList
	Set Lyricists = SDB.NewStringList
	Set Composers = SDB.NewStringList
	Set Conductors = SDB.NewStringList
	Set Producers = SDB.NewStringList
	Set Durations = SDB.NewStringList
	Set rTrackPosition = SDB.NewStringList
	Set rSubPosition = SDB.NewStringList
	Set GenresList = SDB.NewStringList
	REM Set GenresSelect = SDB.NewStringList

	'----------------------------------DiscogsImages----------------------------------------
	Set SaveImage = SDB.NewStringList
	Set SaveImageType = SDB.NewStringList
	Set FileNameList = SDB.NewStringList
	ImagesCount = 0
	'----------------------------------DiscogsImages----------------------------------------

	If OptionsChanged = True Then
		OptionsChanged = False
		WriteOptions()
	End If

	If not isnull(CurrentRelease) Then

		AlbumArtist = ""
		AlbumArtistTitle = ""
		AlbumLyricist = ""
		AlbumComposer = ""
		AlbumConductor = ""
		AlbumProducer = ""
		AlbumInvolved = ""
		AlbumArtURL = ""
		AlbumArtThumbNail = ""
		AlbumFeaturing = ""
		LastDisc = ""

		Dim iTrackNum, cSubTrack, subTrackTitle, aSubtrack
		Dim trackName, pos
		Dim role, role2, rolea, currentRole, NoSplit, zahl, zahltemp, zahl2, zahltemp2
		Dim CharSeparatorSubTrack
		ReDim Involved_R(0)
		Dim tmp, tmp2, tmp3, tmp4, tmp5, tmpResponse
		Dim rTrack
		Dim ret, ResponseHTML, TXTBegin, TXTEnd, Title
		Dim LeadingZeroTrackPosition
		Dim FormatCnt
		ReDim TrackRoles(0)
		ReDim TrackArtist2(0)
		ReDim TrackPos(0)
		ReDim Title_Position(0)
		ReDim TitleList(0)
		ReDim ArtistsList(0)
		ReDim Grouping(0)
		SavedArtistID = ""
		SavedLabelID = ""
		LeadingZeroTrackPosition = False
		Dim involvedArtist, involvedTemp, involvedRole, subTrack
		Dim TrackInvolvedPeople, TrackComposers, TrackConductors, TrackProducers, TrackLyricists, TrackFeaturing, currentHeading, currentCD
		Dim TrackArtist, artistList, min, sec, length

		WriteLog "Start ReloadResults"
		WriteLog "QueryPage=" & QueryPage
		SDB.Tools.WebSearch.ClearTracksData




		' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
		If QueryPage = "MetalArchives" Then
			ResponseHTML = CurrentRelease
			' Get Album art URL
			TXTBegin = InStr(ResponseHTML, "album_img")
			ResponseHTML = Mid(ResponseHTML, TXTBegin)
			TXTBegin = InStr(ResponseHTML, "href=")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 6)
			TXTEnd = InStr(ResponseHTML, Chr(34) & "><")
			AlbumArtURL = Left(ResponseHTML, TXTEnd - 1)
			AlbumArtThumbNail = Left(ResponseHTML, TXTEnd - 1)
			If AlbumArtThumbNail <> "" Then
				ret = getimages(AlbumArtThumbNail, sTemp & "cover.jpg")
				AlbumArtThumbNail = sTemp & "cover.jpg"
			End If

			' Get album title
			TXTBegin = InStr(ResponseHTML, "album_name")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 14)
			TXTBegin = InStr(ResponseHTML, Chr(34) & ">")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 2)
			TXTEnd = InStr(ResponseHTML, "</a>")
			AlbumTitle = PackSpaces(Left(ResponseHTML, TXTEnd - 1), False)
			WriteLog "AlbumTitle=" & AlbumTitle
			
			' Get artist title
			TXTBegin = InStr(ResponseHTML, "<a href")
			ResponseHTML = Mid(ResponseHTML, TXTBegin)
			TXTBegin = InStr(ResponseHTML, Chr(34) & ">")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 2)
			TXTEnd = InStr(ResponseHTML, "</a>")
			AlbumArtist = Left(ResponseHTML, TXTEnd - 1)
			AlbumArtistTitle = AlbumArtist

			If (Not CheckAlbumArtistFirst) Then
				AlbumArtist = AlbumArtistTitle
			End If

			If AlbumArtist = "Various" And CheckVarious Then
				AlbumArtist = TxtVarious
			End If
			If AlbumArtistTitle = "Various" And CheckVarious Then
				AlbumArtistTitle = TxtVarious
			End If
			Writelog "AlbumArtist=" & AlbumArtist
			Writelog "AlbumArtistTitle=" & AlbumArtistTitle
			
			' Get release year/date
			TXTBegin = InStr(ResponseHTML, "Release date:")
			ResponseHTML = Mid(ResponseHTML, TXTBegin)
			TXTBegin = InStr(ResponseHTML, "<dd>")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
			TXTEnd = InStr(ResponseHTML, "</dd>")
			ReleaseDate = Left(ResponseHTML, TXTEnd -1)
			If Len(ReleaseDate) > 4 Then
				If CheckYearOnlyDate Then
					ReleaseDate = Right(ReleaseDate, 4)
				End If
			End If
			WriteLog "ReleaseDate=" & ReleaseDate

			'Get catalog nr.
			TXTBegin = InStr(ResponseHTML, "Catalog ID:")
			ResponseHTML = Mid(ResponseHTML, TXTBegin)
			TXTBegin = InStr(ResponseHTML, "<dd>")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
			TXTEnd = InStr(ResponseHTML, "</dd>")
			theCatalogs = Left(ResponseHTML, TXTEnd -1)
			WriteLog "theCatalogs=" & theCatalogs

			'Get label
			TXTBegin = InStr(ResponseHTML, "Label:")
			ResponseHTML = Mid(ResponseHTML, TXTBegin)
			TXTBegin = InStr(ResponseHTML, Chr(34) & ">")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 2)
			TXTEnd = InStr(ResponseHTML, "</a>")
			theLabels = Left(ResponseHTML, TXTEnd -1)
			WriteLog "theLabels=" & theLabels
			
			'Get format
			TXTBegin = InStr(ResponseHTML, "Format:")
			ResponseHTML = Mid(ResponseHTML, TXTBegin)
			TXTBegin = InStr(ResponseHTML, "<dd>")
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
			TXTEnd = InStr(ResponseHTML, "</dd>")
			theFormat = Left(ResponseHTML, TXTEnd -1)
			WriteLog "theFormat=" & theFormat
			
			TXTBegin = InStr(ResponseHTML, "tbody")
			TXTEnd = InStr(ResponseHTML, "/tbody")
			ResponseHTML = Mid(ResponseHTML, TXTBegin, TXTEnd - TXTBegin - 1)

			currentHeading = ""
			currentCD = ""
			iTrackNum = 0
			iAutoTrackNumber = 1
			iAutoDiscNumber = 1
			If CheckForceDisc = True Then
				If currentCD = "" Then currentCD = iAutoDiscNumber
			End If
			If CheckNoDisc = True Then
				If currentCD <> "" Then currentCD = ""
			End If
			WriteLog "currentCD=" & currentCD

			Do While InStr(ResponseHTML, "wrapWords") > 0
				If InStr(ResponseHTML, "<tr class=" & Chr(34) & "discRow" & Chr(34) & ">") < InStr(ResponseHTML, "class=" & Chr(34) & "anchor" & Chr(34) & ">") And InStr(ResponseHTML, "<tr class=" & Chr(34) & "discRow" & Chr(34) & ">") <> 0 Then
					WriteLog "Neue CD erkannt"
					TXTBegin = InStr(ResponseHTML, "<tr class=" & Chr(34) & "discRow" & Chr(34) & ">")
					tmpResponse = Mid(ResponseHTML, TXTBegin + 22)
					TXTBegin = InStr(tmpResponse, ">")
					TXTEnd = InStr(tmpResponse, "<")
					tmpResponse = Mid(tmpResponse, TXTBegin + 2, TXTEnd - TXTBegin - 2)
					tmp = Split(tmpResponse, Chr(10))
					If Ubound(tmp) = 2 Then
						currentHeading = PackSpaces(Mid(tmp(2), 3), False)
						currentCD = tmp(1)
					End If
					If Ubound(tmp) = 1 Then
						currentCD = tmp(1)
					End If
					iAutoTrackNumber = 1
					iAutoDiscNumber = iAutoDiscNumber + 1
					If CheckForceDisc = True Then
						If currentCD = "" Then currentCD = iAutoDiscNumber
					End If
					If CheckNoDisc = True Then
						If currentCD <> "" Then currentCD = ""
					End If
				End If
				
				' TXTBegin = InStr(ResponseHTML, "<td class=" & Chr(34) & "wrapWords")
				TXTBegin = InStr(ResponseHTML, "class=" & Chr(34) & "anchor" & Chr(34) & ">")
				If TXTBegin = 0 Then Exit Do
				ResponseHTML = Mid(ResponseHTML, TXTBegin + 14)
				TXTBegin = InStr(ResponseHTML, "</a>")
				If TXTBegin = 0 Then Exit Do
				ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)

				TXTEnd = InStr(ResponseHTML, "</td>")
				position = Left(ResponseHTML, TXTEnd -2)
				Title_Position(UBound(Title_Position)) = position
				ReDim Preserve Title_Position(UBound(Title_Position)+1)
				DiscogsTracksNum.Add position
				WriteLog "position=" & position
				
				TXTBegin = InStr(ResponseHTML, Chr(34) & ">")
				ResponseHTML = Mid(ResponseHTML, TXTBegin + 2)
				TXTEnd = InStr(ResponseHTML, "</td>")
				Title = Left(ResponseHTML, TXTEnd - 1)
				Title = Replace(Title, Chr(10), "")
				Title = Trim(Replace(Title, Chr(9), " "))
				WriteLog "Title=" & Title & chr(34)
				TitleList(UBound(TitleList)) = Title
				Tracks.Add Title
				ReDim Preserve TitleList(UBound(TitleList)+1)
				
				TXTBegin = InStr(ResponseHTML, "<td align=" & Chr(34) & "right" & Chr(34) & ">")
				ResponseHTML = Mid(ResponseHTML, TXTBegin + 18)
				TXTEnd = InStr(ResponseHTML, "</td>")
				Durations.Add Left(ResponseHTML, TXTEnd - 1)
				WriteLog "Duration=" & Left(ResponseHTML, TXTEnd - 1)
				
				
				ArtistsList(UBound(ArtistsList)) = AlbumArtist
				ReDim Preserve ArtistsList(UBound(ArtistsList)+1)
				ArtistTitles.Add AlbumArtist

				If UnselectedTracks(iTrackNum) <> "x" Then
					If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
						tracksNum.Add "0" & iAutoTrackNumber
						tracksCD.Add LeadingZeroDisc(currentCD)
						WriteLog "Tracknumber=0" & iAutoTrackNumber
						WriteLog "currentCD=" & currentCD
					Else
						tracksNum.Add iAutoTrackNumber
						tracksCD.Add LeadingZeroDisc(currentCD)
						WriteLog "Tracknumber=" & iAutoTrackNumber
						WriteLog "currentCD=" & currentCD
					End If
					iAutoTrackNumber = iAutoTrackNumber + 1
					Lyricists.Add ""
					Composers.Add ""
					Conductors.Add ""
					Producers.Add ""
					InvolvedArtists.Add ""
					Grouping(UBound(Grouping)) = currentHeading
					ReDim Preserve Grouping(UBound(Grouping)+1)
					WriteLog "Grouping=" & currentHeading
				Else
					tracksNum.Add ""
					tracksCD.Add ""
					WriteLog "Tracknumber=Nothing"
				End If
				iTrackNum = iTrackNum + 1
			Loop

			' Get Comment
			' album_notes
			' If CurrentRelease.Exists("notes") Then
				' Comment = CurrentRelease("notes")
			' Else
				' Comment = ""
			' End If
			' WriteLog "Comment=" & Comment

		End If
		
		
		
		
		If QueryPage = "Discogs" Then
			'Get Track-List
			For Each track In CurrentRelease("tracklist")
				'track start at 0
				Set currentTrack = CurrentRelease("tracklist")(track)
				position = currentTrack("position")
				DiscogsTracksNum.Add position
				position = exchange_roman_numbers(position)
				Title_Position(UBound(Title_Position)) = position
				ReDim Preserve Title_Position(UBound(Title_Position)+1)
				TitleList(UBound(TitleList)) = currentTrack("title")
				ReDim Preserve TitleList(UBound(TitleList)+1)
				tmp = getArtistsName(currentTrack, "artists", QueryPage)
				ArtistsList(UBound(ArtistsList)) = tmp(0)
				ReDim Preserve ArtistsList(UBound(ArtistsList)+1)
				WriteLog "Track=" & track
				rTrackPosition.Add track
				If currentTrack("type_") = "heading" Then
					WriteLog "Heading Track found"
					currentHeading = PackSpaces(DecodeHtmlChars(currentTrack("title")), False)
					Grouping(UBound(Grouping)) = "|Heading|"
					ReDim Preserve Grouping(UBound(Grouping)+1)
				Else
					If currentHeading <> "" Then
						Grouping(UBound(Grouping)) = currentHeading
						ReDim Preserve Grouping(UBound(Grouping)+1)
					Else
						Grouping(UBound(Grouping)) = ""
						ReDim Preserve Grouping(UBound(Grouping)+1)
					End If
				End If
					
				If currentTrack.Exists("sub_tracks") Then
					WriteLog "SubTrack(s) found"
					rSubPosition.Add "NewSubTrack"
					For Each subtrack in currentTrack("sub_tracks")
						WriteLog "subTrack=" & subTrack
						Set aSubtrack = currentTrack("sub_tracks")(subtrack)
						position = aSubtrack("position")
						DiscogsTracksNum.Add position
						position = exchange_roman_numbers(position)
						Title_Position(UBound(Title_Position)) = position
						ReDim Preserve Title_Position(UBound(Title_Position)+1)
						rTrackPosition.Add track
						rSubPosition.Add subtrack
						TitleList(UBound(TitleList)) = aSubtrack("title")
						ReDim Preserve TitleList(UBound(TitleList)+1)
						tmp = getArtistsName(aSubtrack, "artists", QueryPage)
						ArtistsList(UBound(ArtistsList)) = tmp(0)
						ReDim Preserve ArtistsList(UBound(ArtistsList)+1)
						If aSubtrack("type_") = "heading" Then
							WriteLog "Heading SubTrack found"
							currentHeading = PackSpaces(DecodeHtmlChars(aSubtrack("title")), False)
							Grouping(UBound(Grouping)) = currentHeading
							ReDim Preserve Grouping(UBound(Grouping)+1)
						Else
							If currentHeading <> "" Then
								Grouping(UBound(Grouping)) = currentHeading
								ReDim Preserve Grouping(UBound(Grouping)+1)
							Else
								Grouping(UBound(Grouping)) = ""
								ReDim Preserve Grouping(UBound(Grouping)+1)
							End If
						End If
					Next
				Else
					rSubPosition.Add ""
				End If
				SDB.ProcessMessages
			Next
			WriteLog "rTrackPosition.Count=" & rTrackPosition.Count

			For i = 0 To UBound(TitleList)-1
				WriteLog i & " " & ArtistsList(i) &  " - " & TitleList(i)
			Next

			REM For i = 0 To rSubPosition.count-1
				REM WriteLog i & " " & rSubPosition.item(i)
			REM Next

			'Check for leading zero in track-position
			For i = 0 to UBound(Title_Position)-1
				If Title_Position(i) <> "" Then
					LeadingZeroTrackPosition = CheckLeadingZeroTrackPosition(Title_Position(i))
					Exit For
				End If
			Next
			WriteLog "LeadingZeroTrackPosition = " & LeadingZeroTrackPosition

			' Get artist title
			tmp = getArtistsName(CurrentRelease, "artists", QueryPage)

			AlbumArtist = tmp(2)
			If tmp(1) <> "" Then
				AlbumArtistTitle = tmp(0) & " " & tmp(1)
			Else
				AlbumArtistTitle = tmp(0)
			End If
			Writelog "AlbumArtistTitle=" & AlbumArtistTitle

			If (Not CheckAlbumArtistFirst) Then
				AlbumArtist = AlbumArtistTitle
			End If

			If AlbumArtist = "Various" And CheckVarious Then
				AlbumArtist = TxtVarious
			End If
			If AlbumArtistTitle = "Various" And CheckVarious Then
				AlbumArtistTitle = TxtVarious
			End If
			
			If CurrentRelease.Exists("format_quantity") Then
				FormatCnt = CurrentRelease("format_quantity")
			Else
				FormatCnt = 0
			End If

			WriteLog " "
			WriteLog "ExtraArtists"
			If currentRelease.Exists("extraartists") Then
				For Each extraArtist In CurrentRelease("extraartists")
					WriteLog " "
					Set currentArtist = CurrentRelease("extraartists")(extraArtist)
					If currentArtist("tracks") = "" Then
						If (currentArtist("anv") <> "") And Not CheckUseAnv Then
							artistName = CleanArtistName(currentArtist("anv"))
						Else
							artistName = CleanArtistName(currentArtist("name"))
						End If
						WriteLog "ArtistName=" & artistName
						WriteLog "Without Track Info"
						role = currentArtist("role")
						NoSplit = False
						If InStr(role, ",") = 0 Then
							currentRole = Trim(role)
							zahl = 1
							NoSplit = True
						Else
							rolea = CheckSpecialRole(role)
							zahl = UBound(rolea)
						End If

						WriteLog "Role count=" & zahl
						For zahltemp = 1 To zahl
							If NoSplit = False Then
								currentRole = Trim(rolea(zahltemp))
							End If
							WriteLog "currentRole=" & currentRole
							If LookForFeaturing(currentRole) And CheckIgnoreFeatArtist = False Then
								WriteLog "Featuring found"
								If InStr(AlbumFeaturing, artistName) = 0 Then
									If AlbumFeaturing = "" Then
										If CheckFeaturingName Then
											AlbumFeaturing = TxtFeaturingName & " " & artistName
										Else
											AlbumFeaturing = currentRole & " " & artistName
										End If
									Else
										AlbumFeaturing = AlbumFeaturing & ArtistSeparator & artistName
									End If
								End If
							Else
								Do
									tmp = searchKeyword(LyricistKeywords, currentRole, AlbumLyricist, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumLyricist = tmp
										WriteLog "AlbumLyricist=" & AlbumLyricist
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp = searchKeyword(ConductorKeywords, currentRole, AlbumConductor, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumConductor = tmp
										WriteLog "AlbumConductor=" & AlbumConductor
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp = searchKeyword(ProducerKeywords, currentRole, AlbumProducer, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumProducer = tmp
										WriteLog "AlbumProducer=" & AlbumProducer
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp = searchKeyword(ComposerKeywords, currentRole, AlbumComposer, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumComposer = tmp
										WriteLog "AlbumComposer=" & AlbumComposer
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp2 = search_involved(Involved_R, currentRole)
									If tmp2 = -2 Then
										WriteLog "Ignore Role: '" & currentRole & "' (Unwanted tag)"
									ElseIf tmp2 = -1 Then
										ReDim Preserve Involved_R(UBound(Involved_R)+1)
										Involved_R(UBound(Involved_R)) = currentRole & ": " & artistName
										WriteLog "New Role: " & currentRole & ": " & artistName
									Else
										If InStr(Involved_R(tmp2), artistName) = 0 Then
											Involved_R(tmp2) = Involved_R(tmp2) & ArtistSeparator & artistName
											WriteLog "Role updated: " & Involved_R(tmp2)
										Else
											WriteLog "artist already inside role"
										End If
									End If
									Exit Do
								Loop While True
							End If
						Next
					Else
						If Not CheckUseAnv And currentArtist("anv") <> "" Then
							artistName = CleanArtistName(currentArtist("anv"))
						Else
							artistName = CleanArtistName(currentArtist("name"))
						End If
						WriteLog "ArtistName=" & artistName
						role = currentArtist("role")
						rTrack = currentArtist("tracks")
						If Left(rTrack, 7) = "tracks:" Then
							rTrack = Trim(Mid(rTrack, 8))
						End If
						WriteLog "Track(s)=" & rTrack
						WriteLog "Role(s)=" & role
						NoSplit = False
						If InStr(role, ",") <> 0 Then
							rolea = CheckSpecialRole(role)
							zahl = UBound(rolea)
						ElseIf InStr(role, " & ") <> 0 Then
							rolea = Split(role, "&")
							zahl = UBound(rolea)
						Else
							involvedRole = Trim(role)
							zahl = 1
							NoSplit = True
						End If
						For zahltemp = 1 To zahl
							If NoSplit = False Then
								involvedRole = Trim(rolea(zahltemp))
							End If
							WriteLog "involvedRole=" & involvedRole
							If InStr(rTrack, ",") = 0 And InStr(rTrack, " to ") = 0 And InStr(rTrack, " & ") = 0 Then
								currentTrack = rTrack
								Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
							End If
							If InStr(rTrack, ",") <> 0 Then
								tmp = Split(rTrack, ",")
								zahl2 = UBound(tmp)
								For zahltemp2 = 0 To zahl2
									currentTrack = Trim(tmp(zahltemp2))
									If InStr(currentTrack, " to ") <> 0 Then
										Track_from_to currentTrack, artistName, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
									Else
										Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
									End If
								Next
							ElseIf InStr(rTrack, " to ") <> 0 Then
								currentTrack = Trim(rTrack)
								Track_from_to currentTrack, artistName, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
							ElseIf InStr(rTrack, " & ") <> 0 Then
								tmp = Split(rTrack, " & ")
								zahl2 = UBound(tmp)
								For zahltemp2 = 0 To zahl2
									currentTrack = Trim(tmp(zahltemp2))
									Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
								Next
							End If
						Next
					End If
				Next
			End If
			' Get track titles and track artists

			WriteLog " "
			WriteLog "--------------------------------------------------------------"
			iAutoTrackNumber = 1
			iAutoDiscNumber = 1
			iAutoDiscFormat = ""
			iTrackNum = 0
			cSubTrack = -1
			subTrackTitle = ""
			CharSeparatorSubTrack = 0
			Rem CharSeparatorSubTrack: 0 = nothing    1 = "."     2 = a-z
			Rem subTrackStart = 1 '0 = Song -1    1 = First Song

			'Workaround for using "." as separator at discogs -----------------------------------------------------------------------------------------------------------
			tmp = 0 : tmp2 = 0
			NoSubTrackUsing = False
			For t = 0 To UBound(Title_Position)-1
				If Title_Position(t) <> "" Then
					tmp2 = tmp2 + 1
				End If
				If InStr(Title_Position(t), ".") <> 0 Then tmp = tmp + 1
			Next
			If tmp = tmp2 And tmp <> 0 Then NoSubTrackUsing = True	'all tracks have "." in position tag, this can't be a subtrack
			WriteLog "NoSubTrackUsing = " & NoSubTrackUsing
			'Workaround for using "." as separator at discogs -----------------------------------------------------------------------------------------------------------

			WriteLog "Track count from Title_position=" & UBound(Title_Position)

			Dim NewSubTrackFound, cNewSubTrack
			NewSubTrackFound = False

			For t = 0 To UBound(Title_Position)-1
				WriteLog " "
				WriteLog "Process next track"
				WriteLog "Track=" & t
				WriteLog "Title_position=" & Title_Position(t)
				SDB.ProcessMessages
				If rSubPosition.Item(t) <> "" And rSubPosition.Item(t) <> "NewSubTrack" Then
					WriteLog "Processing Subtrack"
					Set currentTrack = CurrentRelease("tracklist")(CInt(rTrackPosition.Item(t)))("sub_tracks")(cInt(rSubPosition.Item(t)))
				Else
					Set currentTrack = CurrentRelease("tracklist")(CInt(rTrackPosition.Item(t)))
				End If

				position = currentTrack("position")
				If Right(position, 1) = "." Then position = Left(position, Len(position)-1)
				If NoSubTrackUsing = True Then position = Replace(position, ".", "-")
				trackname = PackSpaces(DecodeHtmlChars(TitleList(t)), True)
				WriteLog "Trackname=" & trackname
				WriteLog "Position=" & position
				' WriteLog "t=" & t
				WriteLog "rSubPosition(t)=" & rSubPosition.item(t)
				Durations.Add currentTrack("duration")
				position = exchange_roman_numbers(position)

				If rSubPosition.Item(t) <> "" And rSubPosition.Item(t) <> "NewSubTrack" Then
					If NewSubTrackFound = False Then
						cNewSubTrack = t - 1
						If ArtistsList(t) <> "" then
							subTrackTitle = ArtistsList(t) & " - " & trackName
						Else
							subTrackTitle = trackName
						End If
						If NewResult = True Then
							UnselectedTracks(iTrackNum) = "x"
							UnselectedTrackNames(iTrackNum) = "x"
						End If
						NewSubTrackFound = True
						WriteLog "New Subtrack"
					Else
						If ArtistsList(t) <> "" then
							subTrackTitle = subTrackTitle & SubTrackSeparator & ArtistsList(t) & " - " & trackName
						Else
							subTrackTitle = subTrackTitle & SubTrackSeparator & trackName
						End If
						If NewResult = True Then
							UnselectedTracks(iTrackNum) = "x"
							UnselectedTrackNames(iTrackNum) = "x"
						End If
						WriteLog "More Subtrack"
					End If
					If  t = (UBound(Title_Position) - 1) Then	'if last track and subtrack then add to subtrack songlist
						If SubTrackNameSelection = False Then
							Tracks.Item(cNewSubTrack) = Tracks.Item(cNewSubTrack) & " (" & subTrackTitle & ")"
						Else
							Tracks.Item(cNewSubTrack) = subTrackTitle
						End If
					End If
				Else
					If NewSubTrackFound = True Then
						If SubTrackNameSelection = False Then
							Tracks.Item(cNewSubTrack) = Tracks.Item(cNewSubTrack) & " (" & subTrackTitle & ")"
						Else
							Tracks.Item(cNewSubTrack) = subTrackTitle
						End If
						If NewResult = True Then
							UnselectedTracks(cNewSubTrack) = ""
							UnselectedTrackNames(cNewSubTrack) = ""
						End If
						NewSubTrackFound = False
						cNewSubTrack = -1
						WriteLog "Subtrack end"
					End If
				End If

				pos = 0
				If InStr(LCase(position), "-") > 0 And position <> "-" Then
					pos = InStr(LCase(position), "-")
				End If
				' Here comes the new track/disc numbering methods
				If position <> "" And rSubPosition.Item(t) = "" Then
					If CheckTurnOffSubTrack = False Then
						If (cSubTrack <> -1 And InStr(LCase(position), ".") = 0 And CharSeparatorSubTrack = 1) Or (cSubTrack <> -1 And IsNumeric(Right(position, 1)) And CharSeparatorSubTrack = 2) Or position = "-" Then
							WriteLog "End of Subtrack found"
							If SubTrackNameSelection = False Then
								Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
							Else
								Tracks.Item(cSubTrack) = subTrackTitle
							End If
							cSubTrack = -1
							subTrackTitle = ""
							CharSeparatorSubTrack = 0
						End If

						If NoSubTrackUsing = False Then
							WriteLog "Calling Subtrack Function"
							CharSeparatorSubTrack = 0
							'SubTrack Function ---------------------------------------------------------
							If InStr(LCase(position), ".") > 0 Then
								CharSeparatorSubTrack = 1
							ElseIf Not IsNumeric(Right(position, 1)) And Len(position) > 1 And position <> "Video" Then
								CharSeparatorSubTrack = 2
							End If
							WriteLog "CharSeparatorSubTrack = " & CharSeparatorSubTrack
							If CharSeparatorSubTrack <> 0 Then
								If cSubTrack <> -1 Then 'more subtrack
									WriteLog "More Subtrack"
									If CharSeparatorSubTrack = 1 Then
										tmp = Split(position, ".")
										If oldSubTrackNumber <> tmp(0) Then
											If SubTrackNameSelection = False Then
												Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
											Else
												Tracks.Item(cSubTrack) = subTrackTitle
											End If
											cSubTrack = -1
											subTrackTitle = ""
										End If
									ElseIf CharSeparatorSubTrack = 2 Then
										tmp2 = FindSubTrackSplit(position)
										If oldSubTrackNumber <> tmp2 Then
											If SubTrackNameSelection = False Then
												Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
											Else
												Tracks.Item(cSubTrack) = subTrackTitle
											End If
											cSubTrack = -1
											subTrackTitle = ""
										End If
									End If
								Else   'new subtrack
									WriteLog "New SubTrack found"
									If SubTrackNameSelection = False And iTrackNum > 0 Then
										cSubTrack = iTrackNum - 1
									Else
										cSubTrack = iTrackNum
									End If
									If CharSeparatorSubTrack = 1 Then
										tmp = Split(position, ".")
										oldSubTrackNumber = tmp(0)
									ElseIf CharSeparatorSubTrack = 2 Then
										oldSubTrackNumber = FindSubTrackSplit(position)
										If oldSubTrackNumber = "" Then oldSubTrackNumber = position
									End If
									WriteLog "oldSubTrackNumber=" & oldSubTrackNumber
								End If
								
								If subTrackTitle = "" Then
									If ArtistsList(t) <> "" Then
										subTrackTitle = ArtistsList(t) & " - " & trackName
									Else
										subTrackTitle = trackName
									End If
									If SubTrackNameSelection = False Then
										If iTrackNum > 0 Then
											If NewResult = True Then
												UnselectedTracks(iTrackNum-1) = ""
												UnselectedTracks(iTrackNum) = "x"
												UnselectedTrackNames(iTrackNum-1) = ""
												UnselectedTrackNames(iTrackNum) = "x"
											End If
										Else
											If NewResult = True Then
												UnselectedTracks(iTrackNum) = ""
												UnselectedTrackNames(iTrackNum) = ""
											End If
										End If
									Else
										If NewResult = True Then
											UnselectedTracks(iTrackNum) = ""
											UnselectedTrackNames(iTrackNum) = ""
										End If
									End If
								Else
									If ArtistsList(t) <> "" Then
										subTrackTitle = subTrackTitle & SubTrackSeparator & ArtistsList(t) & " - " & trackName
									Else
										subTrackTitle = subTrackTitle & SubTrackSeparator & trackName
									End If
									If NewResult = True Then
										UnselectedTracks(iTrackNum) = "x"
										UnselectedTrackNames(iTrackNum) = "x"
									End If
								End If

								'SubTrack Function ---------------------------------------------------------
							End If
						End If
					End If

					trackNumbering pos, position, TracksNum, TracksCD, iTrackNum, FormatCnt

				ElseIf (trackName = "-" And rSubPosition.Item(t) <> "NewSubTrack") Then
					tracksNum.Add ""
					tracksCD.Add ""
					If NewResult = True Then
						UnselectedTracks(iTrackNum) = "x"
						UnselectedTrackNames(iTrackNum) = "x"
					End If
				ElseIf currentTrack("type_") = "heading" Then
					WriteLog "Heading-Track erkannt"
					If NewResult = True Or UnselectedTracks(iTrackNum) = "x" Then
						UnselectedTracks(iTrackNum) = "x"
						UnselectedTrackNames(iTrackNum) = "x"
						tracksNum.Add ""
						tracksCD.Add ""
					End If
					If UnselectedTracks(iTrackNum) = "" Then
						trackNumbering pos, position, TracksNum, TracksCD, iTrackNum, FormatCnt
					End If
				Else ' Nothing specified
					trackNumbering pos, position, TracksNum, TracksCD, iTrackNum, FormatCnt
				End If

				ReDim Involved_R_T(0)

				TrackInvolvedPeople = ""
				TrackComposers = ""
				TrackConductors = ""
				TrackProducers = ""
				TrackLyricists = ""
				TrackFeaturing = AlbumFeaturing

				If UBound(Involved_R) > 0 Then
					For tmp = 1 To UBound(Involved_R)
						ReDim Preserve Involved_R_T(tmp)
						Involved_R_T(tmp) = Involved_R(tmp)
					Next
				End If

				For tmp = 1 To UBound(TrackPos)
					If TrackPos(tmp) = position And position <> "" Then
						WriteLog "trackpos(" & tmp & ")=" & trackpos(tmp)
						involvedRole = TrackRoles(tmp)
						involvedArtist = TrackArtist2(tmp)

						If LookForFeaturing(involvedRole) And CheckIgnoreFeatArtist = False Then
							If InStr(TrackFeaturing, involvedArtist) = 0 Then
								If TrackFeaturing = "" Then
									If CheckFeaturingName Then
										TrackFeaturing = TxtFeaturingName & " " & involvedArtist
									Else
										TrackFeaturing = involvedRole & " " & involvedArtist
									End If
								Else
									TrackFeaturing = TrackFeaturing & ArtistSeparator & involvedArtist
								End If
							End If
							WriteLog "TrackFeaturing=" & TrackFeaturing
						Else
							Do
								ret = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
								If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
									TrackLyricists = ret
									WriteLog "TrackLyricists=" & TrackLyricists
									Exit Do
								ElseIf ret = "ALREADY_INSIDE_ROLE" Then
									WriteLog "ALREADY_INSIDE_ROLE"
									Exit Do
								End If
								ret = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
								If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
									TrackConductors = ret
									WriteLog "TrackConductors=" & TrackConductors
									Exit Do
								ElseIf ret = "ALREADY_INSIDE_ROLE" Then
									WriteLog "ALREADY_INSIDE_ROLE"
									Exit Do
								End If
								ret = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
								If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
									TrackProducers = ret
									WriteLog "TrackProducers=" & TrackProducers
									Exit Do
								ElseIf ret = "ALREADY_INSIDE_ROLE" Then
									WriteLog "ALREADY_INSIDE_ROLE"
									Exit Do
								End If
								ret = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
								If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
									TrackComposers = ret
									WriteLog "TrackComposers=" & TrackComposers
									Exit Do
								ElseIf ret = "ALREADY_INSIDE_ROLE" Then
									WriteLog "ALREADY_INSIDE_ROLE"
									Exit Do
								End If
								tmp2 = search_involved(Involved_R_T, involvedRole)
								If tmp2 = -2 Then
									WriteLog "Ignore Role: '" & currentRole & "' (Unwanted tag)"
								ElseIf tmp2 = -1 Then
									ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
									Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & TrackArtist2(tmp)
									WriteLog "New Role: " & involvedRole & ": " & TrackArtist2(tmp)
								Else
									If InStr(Involved_R_T(tmp2), TrackArtist2(tmp)) = 0 Then
										Involved_R_T(tmp2) = Involved_R_T(tmp2) & ArtistSeparator & TrackArtist2(tmp)
										WriteLog "Role updated: " & Involved_R_T(tmp2)
									Else
										WriteLog "artist already inside role"
									End If
								End If
								Exit Do
							Loop While True
						End If
					End If
					SDB.ProcessMessages
				Next

				artistList = ""

				WriteLog " "
				WriteLog "Search for TrackArtist"
				If currentTrack.Exists("artists") Then
					tmp = getArtistsName(CurrentTrack, "artists", QueryPage)
					artistList = tmp(0)
					TrackFeaturing = tmp(1)
				End If
				If artistList = "" Then artistList = AlbumArtistTitle

				WriteLog "artistlist=" & artistlist
				WriteLog "rTrack=" & rTrackPosition.Item(t)
				REM WriteLog "ubound=" & UBound($currentTrack)

				If currentTrack.Exists("extraartists") Then
					WriteLog " "
					WriteLog "ExtraArtist found"
					For Each extra In currentTrack("extraartists")
						Set currentArtist = CurrentTrack("extraartists")(extra)
						If (currentArtist("anv") <> "") And Not CheckUseAnv Then
							involvedArtist = CleanArtistName(currentArtist("anv"))
						Else
							involvedArtist = CleanArtistName(currentArtist("name"))
						End If
						WriteLog "involvedArtist=" & involvedArtist
						If involvedArtist <> "" Then
							role = currentArtist("role")
							NoSplit = False
							If InStr(role, ",") = 0 Then
								involvedRole = Trim(role)
								zahl = 1
								NoSplit = True
							Else
								rolea = CheckSpecialRole(role)
								zahl = UBound(rolea)
							End If
							For zahltemp = 1 To zahl
								If NoSplit = False Then
									involvedRole = Trim(rolea(zahltemp))
								End If

								If LookForFeaturing(involvedRole) And CheckIgnoreFeatArtist = False Then
									If InStr(artistList, involvedArtist) = 0 Then
										If TrackFeaturing = "" Then
											If CheckFeaturingName Then
												TrackFeaturing = TxtFeaturingName & " " & involvedArtist
											Else
												TrackFeaturing = involvedRole & " " & involvedArtist
											End If
										Else
											If InStr(TrackFeaturing, involvedArtist) = 0 Then
												TrackFeaturing = TrackFeaturing & ArtistSeparator & involvedArtist
											End If
										End If
									End If
								Else
									Do
										tmp = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
										If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
											TrackLyricists = tmp
											WriteLog "TrackLyricists=" & TrackLyricists
											Exit Do
										ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
											WriteLog "ALREADY_INSIDE_ROLE"
											Exit Do
										End If
										tmp = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
										If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
											TrackConductors = tmp
											WriteLog "TrackConductors=" & TrackConductors
											Exit Do
										ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
											WriteLog "ALREADY_INSIDE_ROLE"
											Exit Do
										End If
										tmp = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
										If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
											TrackProducers = tmp
											WriteLog "TrackProducers=" & TrackProducers
											Exit Do
										ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
											WriteLog "ALREADY_INSIDE_ROLE"
											Exit Do
										End If
										tmp = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
										If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
											TrackComposers = tmp
											WriteLog "TrackComposers=" & TrackComposers
											Exit Do
										ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
											WriteLog "ALREADY_INSIDE_ROLE"
											Exit Do
										End If
										tmp2 = search_involved(Involved_R_T, involvedRole)
										If tmp2 = -2 Then
											WriteLog "Ignore Role: '" & currentRole & "' (Unwanted tag)"
										ElseIf tmp2 = -1 Then
											ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
											Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & involvedArtist
											WriteLog "New Role: " & involvedRole & ": " & involvedArtist
										Else
											If InStr(Involved_R_T(tmp2), involvedArtist) = 0 Then
												Involved_R_T(tmp2) = Involved_R_T(tmp2) & ArtistSeparator & involvedArtist
												WriteLog "Role updated: " & Involved_R_T(tmp2)
											Else
												WriteLog "artist already inside role"
											End If
										End If
										Exit Do
									Loop While True
								End If
							Next
						End If
						SDB.ProcessMessages
					Next
					WriteLog "ExtraArtist end"
				End If

				If TrackFeaturing <> "" And CheckIgnoreFeatArtist = False Then
					If CheckTitleFeaturing = True Then
						tmp = InStrRev(TrackFeaturing, ArtistSeparator)
						If tmp = 0 Or ArtistLastSeparator = False Then
							trackName = trackName & " (" & TrackFeaturing & ")"
						Else
							trackName = trackName & " (" &  Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+Len(ArtistSeparator)) & ")"
						End If
					Else
						tmp = InStrRev(TrackFeaturing, ArtistSeparator)
						If tmp = 0 Or ArtistLastSeparator = False Then
							If Left(TrackFeaturing, 1) = "," Or Left(TrackFeaturing, 1) = ";" Then
								artistList = artistList & TrackFeaturing
							Else
								artistList = artistList & " " & TrackFeaturing
							End If
						Else
							artistList = artistList & " " & Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+Len(ArtistSeparator))
						End If
					End If
				End If

				ArtistTitles.Add artistList

				If position <> "" Then
					TrackLyricists = FindArtist(TrackLyricists, AlbumLyricist)
					If AlbumLyricist <> "" and TrackLyricists <> "" Then
						Lyricists.Add AlbumLyricist & Separator & TrackLyricists
					Else
						Lyricists.Add AlbumLyricist & TrackLyricists
					End If
					TrackComposers = FindArtist(TrackComposers, AlbumComposer)
					If AlbumComposer <> "" and TrackComposers <> "" Then
						Composers.Add AlbumComposer & Separator & TrackComposers
					Else
						Composers.Add AlbumComposer & TrackComposers
					End If
					TrackConductors = FindArtist(TrackConductors, AlbumConductor)
					If AlbumConductor <> "" and TrackConductors <> "" Then
						Conductors.Add AlbumConductor & Separator & TrackConductors
					Else
						Conductors.Add AlbumConductor & TrackConductors
					End If

					TrackProducers = FindArtist(TrackProducers, AlbumProducer)
					If AlbumProducer <> "" and TrackProducers <> "" Then
						Producers.Add AlbumProducer & Separator & TrackProducers
					Else
						Producers.Add AlbumProducer & TrackProducers
					End If

					If UBound(Involved_R_T) > 0 Then
						For tmp = 1 To UBound(involved_R_T)
							TrackInvolvedPeople = TrackInvolvedPeople & Involved_R_T(tmp) & Separator
						Next
						TrackInvolvedPeople = Left(TrackInvolvedPeople, Len(TrackInvolvedPeople)-Len(Separator))
					Else
						TrackInvolvedPeople = ""
					End If

					InvolvedArtists.Add TrackInvolvedPeople
					If CheckReplaceAccentChar = True Then
						trackName = Replace(trackName, ChrW(8245), "'")
						trackName = Replace(trackName, ChrW(8246), "'")
						trackName = Replace(trackName, ChrW(8242), "'")
						trackName = Replace(trackName, ChrW(8243), "'")
						trackName = Replace(trackName, Chr(145), "'")
						trackName = Replace(trackName, Chr(146), "'")
						trackName = Replace(trackName, Chr(147), "'")
						trackName = Replace(trackName, Chr(148), "'")
					End If
					Tracks.Add trackName
				Else
					Lyricists.Add ""
					Composers.Add ""
					Conductors.Add ""
					Producers.Add ""
					InvolvedArtists.Add ""
					If CheckReplaceAccentChar = True Then
						trackName = Replace(trackName, ChrW(8245), "'")
						trackName = Replace(trackName, ChrW(8246), "'")
						trackName = Replace(trackName, ChrW(8242), "'")
						trackName = Replace(trackName, ChrW(8243), "'")
						trackName = Replace(trackName, Chr(145), "'")
						trackName = Replace(trackName, Chr(146), "'")
						trackName = Replace(trackName, Chr(147), "'")
						trackName = Replace(trackName, Chr(148), "'")
					End If
					Tracks.Add trackName
				End If
				iTrackNum = iTrackNum + 1
			Next


			If cSubTrack <> -1 Then
				If CheckReplaceAccentChar = True Then
					subTrackTitle = Replace(subTrackTitle, ChrW(8245), "'")
					subTrackTitle = Replace(subTrackTitle, ChrW(8246), "'")
					subTrackTitle = Replace(subTrackTitle, ChrW(8242), "'")
					subTrackTitle = Replace(subTrackTitle, ChrW(8243), "'")
					subTrackTitle = Replace(subTrackTitle, Chr(145), "'")
					subTrackTitle = Replace(subTrackTitle, Chr(146), "'")
					subTrackTitle = Replace(subTrackTitle, Chr(147), "'")
					subTrackTitle = Replace(subTrackTitle, Chr(148), "'")
				End If
				If SubTrackNameSelection = False Then
					Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
				Else
					Tracks.Item(cSubTrack) = subTrackTitle
				End If
				cSubTrack = -1
				subTrackTitle = ""
				CharSeparatorSubTrack = 0
			End If

			' Get album title
			AlbumTitle = PackSpaces(currentRelease("title"), True)

			' Get Album art URL
			If CurrentRelease.Exists("images") Then
				For Each i In CurrentRelease("images")
					Set currentImage = CurrentRelease("images")(i)

					If currentImage("type") = "primary" Or AlbumArtURL = "" Then
						AlbumArtURL = currentImage("resource_url")
						WriteLog "AlbumArtURL2=" & AlbumArtURL
						AlbumArtThumbNail = currentImage("uri150")
						WriteLog "AlbumArtThumbNail2=" & AlbumArtThumbNail
					End If
				Next
			End If
			If AlbumArtThumbNail <> "" Then
				ret = getimages(AlbumArtThumbNail, sTemp & "cover.jpg")
				AlbumArtThumbNail = sTemp & "cover.jpg"
			End If

			'----------------------------------DiscogsImages----------------------------------------
			Set ImageList = SDB.NewStringList
			Set SaveImageType = SDB.NewStringList
			Set SaveImage = SDB.NewStringList
			ImagesCount = 0

			If CurrentRelease.Exists("images") Then
				ImagesCount = CurrentRelease("images").Count
				WriteLog "ImagesCount=" & ImagesCount
				If ImagesCount > 1 Then
					For Each i In CurrentRelease("images")
						Set currentImage = CurrentRelease("images")(i)
						tmpArt = currentImage("resource_url")
						WriteLog tmpArt
						If AlbumArtURL <> tmpArt Then
							ImageList.add tmpArt
							SaveImageType.add "other"
							SaveImage.add "0"
						End If
					Next
				End If
			End If
			'----------------------------------DiscogsImages----------------------------------------

			' Get Master ID
			If CurrentRelease.Exists("master_id") Then
				theMaster = currentRelease("master_id")
				If SavedMasterID <> theMaster Then
					OriginalDateRead = ReloadMaster(theMaster)
					SavedMasterID = theMaster
				End If
			ElseIf CurrentRelease.Exists("main_release") Then	'Master
				If CurrentRelease.Exists("year") Then
					OriginalDateRead = CurrentRelease("year")
				End If
				SavedMasterID = currentRelease("id")
			Else
				theMaster = ""
				SavedMasterID = theMaster
				OriginalDateRead = ""
			End If

			' Get release year/date
			If CurrentRelease.Exists("released") Then
				ReleaseDateRead = CurrentRelease("released")
				If Len(ReleaseDateRead) > 4 Then
					ReleaseSplit = Split(ReleaseDateRead,"-")
					If ReleaseSplit(2) = "00" Then
						ReleaseDateRead = Left(ReleaseDateRead, 4)
					Else
						ReleaseDateRead = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
					End If
					If CheckYearOnlyDate Then
						ReleaseDateRead = Right(ReleaseDateRead, 4)
					End If
				End If
			Else
				ReleaseDateRead = ""
			End If
			WriteLog "ReleaseDateRead=" & ReleaseDateRead

			'Set OriginalDate
			If OriginalDateRead <> "" Then
				If Len(OriginalDateRead) > 4 Then
					ReleaseSplit = Split(OriginalDateRead,"-")
					If ReleaseSplit(2) = "00" Then
						OriginalDateRead = Left(OriginalDateRead, 4)
					Else
						OriginalDateRead = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
					End If
					If CheckYearOnlyDate Then
						OriginalDateRead = Right(OriginalDateRead, 4)
					End If
				End If
			End If
			WriteLog "OriginalDateRead=" & OriginalDateRead

			'Choose Date field saving
			If StoreDate = 0 Then
				ReleaseDate = ReleaseDateRead
				OriginalDate = OriginalDateRead
				WriteLog "Use StoreDate=Default"
			ElseIf StoreDate = 1 Then
				ReleaseDate = ReleaseDateRead
				OriginalDate = ReleaseDateRead
				WriteLog "Use release date for both date fields"
			ElseIf StoreDate = 2 Then
				ReleaseDate = OriginalDateRead
				OriginalDate = OriginalDateRead
				WriteLog "Use original date for both date fields"
			End If

			WriteLog "ReleaseDate=" & ReleaseDate
			WriteLog "OriginalDate=" & OriginalDate

			' Get genres
			If GenresSelect.Count > 0 Then
				tmp2 = true
			Else
				tmp2 = false
			End If
			For Each g In CurrentRelease("genres")
				tmp = false
				If GenresList.Count > 0 Then
					For x = 0 to GenresList.Count -1
						If LCase(GenresList.Item(x)) = LCase(CurrentRelease("genres")(g)) Then tmp = true
					Next
				End If
				If tmp = false Then 
					GenresList.Add CurrentRelease("genres")(g)
					If tmp2 = false Then
						GenresSelect.Add True
					End If
				End If
			Next

			' Get styles/moods/themes
			If CurrentRelease.Exists("styles") Then
				For Each s In CurrentRelease("styles")
					tmp = false
					If GenresList.Count > 0 Then
						For x = 0 to GenresList.Count -1
							If LCase(GenresList.Item(x)) = LCase(CurrentRelease("styles")(s)) Then tmp = true
						Next
					End If
					If tmp = false Then 
						GenresList.Add CurrentRelease("styles")(s)
						If tmp2 = false Then
							GenresSelect.Add True
						End If
					End If
				Next
			End If


			' Get Label
			If CurrentRelease.Exists("labels") Then
				For Each l in CurrentRelease("labels")
					Set currentLabel = CurrentRelease("labels")(l)
					If SavedLabelID = "" Then
						If currentLabel.Exists("id") Then
							SavedLabelID = currentLabel("id")
						End If
					End If
					If CheckDeleteDuplicatedEntry = True Then
						AddToFieldWD theLabels, CleanArtistName(currentLabel("name"))
						AddToFieldWD theCatalogs, currentLabel("catno")
					Else
						AddToField theLabels, CleanArtistName(currentLabel("name"))
						AddToField theCatalogs, currentLabel("catno")
					End If
				Next
			Else
				theLabels = ""
				theCatalogs = ""
			End If
			WriteLog "theLabels=" & theLabels
			WriteLog "theCatalogs=" & theCatalogs

			' Get Country
			If CurrentRelease.Exists("country") Then
				theCountry = CurrentRelease("country")
			Else
				theCountry = ""
			End If
			WriteLog "country=" & theCountry

			' Get Format
			If CurrentRelease.Exists("formats") Then
				For Each f in CurrentRelease("formats")
					Set currentFormat = CurrentRelease("formats")(f)
					If currentFormat("qty") > 1 And CheckShowQtyFormat = True Then
						AddToField theFormat, currentFormat("qty") & " x " & currentFormat("name")
					Else
						AddToField theFormat, currentFormat("name")
					End If
					If currentFormat.Exists("descriptions") Then
						For Each d in currentFormat("descriptions")
							theFormat = theFormat & FormatSeparator & currentFormat("descriptions")(d)
						Next
					End If
				Next
			Else
				theFormat = ""
			End If
			WriteLog "theformat=" & theFormat

			Dim theFormatSplit, albumType, numFormatSplits, albumSuffix, lastFormat
			theFormatSplit = Split(theFormat, ",")
			numFormatSplits = UBound(theFormatSplit)

			WriteLog "numFormatSplits=" & numFormatSplits

			' Promo, EP, Compilation, Single, CDM, CDS
			lastFormat = Trim(theFormatSplit(numFormatSplits))
			if (InStr(theFormat, "12""") OR InStr(theFormat, "7""") OR InStr(theFormat, "10""")) AND InStr(theFormat, "Album") = 0 AND InStr(theFormat, "EP") = 0 Then
				albumSuffix = theCatalogs
			Else
				Dim subTypes
				subTypes = Array("EP", "CDM", "Single", "CDS", "Compilation", "Mini-Album", "Promo")
				For Each t in subTypes
					if InStr(theFormat, t) > 0 Then
						albumSuffix = t
					End If
				Next

				'if numFormatSplits >= 2 Then
				'	if lastFormat <> "Album" Then
				'		albumSuffix = lastFormat
				'	End If
				'End If
			End If
			
			
			if albumSuffix <> "" Then
				Dim albumSplit
				albumSplit = Split(AlbumTitle, " ")
				if albumSplit(UBound(albumSplit)) <> albumSuffix Then
					AlbumTitle = AlbumTitle & " [" & albumSuffix  & "]"
				End If
			End If
      
			' Get Comment
			If CurrentRelease.Exists("notes") Then
				Comment = CurrentRelease("notes")
			Else
				Comment = ""
			End If
			WriteLog "Comment=" & Comment

			' Get data_quality
			If CurrentRelease.Exists("data_quality") Then
				DataQuality = CurrentRelease("data_quality")
			Else
				DataQuality = ""
			End If
			WriteLog "DataQuality=" & DataQuality
		End If

'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
		If QueryPage = "MusicBrainz" Then

			For Each m In CurrentRelease("media")
				Set currentMedia = CurrentRelease("media")(m)
				For Each track In CurrentMedia("tracks")
					Set currentTrack = CurrentMedia("tracks")(track)
					position = currentTrack("number")
					DiscogsTracksNum.Add position
					position = exchange_roman_numbers(position)
					ReDim Preserve Title_Position(UBound(Title_Position)+1)
					Title_Position(UBound(Title_Position)) = position
					ReDim Preserve Grouping(UBound(Grouping)+1)
					Grouping(UBound(Grouping)) = ""
				Next
			Next


			'Check for leading zero in track-position
			'LeadingZeroTrackPosition = CheckLeadingZeroTrackPosition(Title_Position(1))


			' Get release artist
			tmp = getArtistsName(CurrentRelease, "artist-credit", QueryPage)
			If tmp(1) <> "" Then
				AlbumArtistTitle = tmp(0) & " " & tmp(1)
			Else
				AlbumArtistTitle = tmp(0)
			End If
			AlbumArtist = tmp(2)

			Writelog "AlbumArtistTitle=" & AlbumArtistTitle

			If (Not CheckAlbumArtistFirst) Then
				AlbumArtist = AlbumArtistTitle
			End If

			If AlbumArtist = "Various Artists" And CheckVarious Then
				AlbumArtist = TxtVarious
			End If
			If AlbumArtistTitle = "Various Artists" And CheckVarious Then
				AlbumArtistTitle = TxtVarious
			End If


			WriteLog " "
			WriteLog "ExtraArtists for Release"
			If currentRelease.Exists("relations") Then
				If currentRelease("relations").Count > 0 Then
					For Each extraArtist In CurrentRelease("relations")
						WriteLog " "
						Set currentArtist = CurrentRelease("relations")(extraArtist)
						artistName = CleanArtistName(currentArtist("artist")("name"))
						WriteLog "ArtistName=" & artistName
						ret = RelationshipTypes(currentArtist)
						If ret(0) > 0 Then
							For i = 1 to ret(0)
								currentRole = ret(i)
								Do
									tmp = searchKeyword(LyricistKeywords, currentRole, AlbumLyricist, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumLyricist = tmp
										WriteLog "AlbumLyricist=" & AlbumLyricist
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp = searchKeyword(ConductorKeywords, currentRole, AlbumConductor, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumConductor = tmp
										WriteLog "AlbumConductor=" & AlbumConductor
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp = searchKeyword(ProducerKeywords, currentRole, AlbumProducer, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumProducer = tmp
										WriteLog "AlbumProducer=" & AlbumProducer
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp = searchKeyword(ComposerKeywords, currentRole, AlbumComposer, artistName)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										AlbumComposer = tmp
										WriteLog "AlbumComposer=" & AlbumComposer
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp2 = search_involved(Involved_R, currentRole)
									If tmp2 = -2 Then
										WriteLog "Ignore Role: '" & currentRole & "' (Unwanted tag)"
									ElseIf tmp2 = -1 Then
										ReDim Preserve Involved_R(UBound(Involved_R)+1)
										Involved_R(UBound(Involved_R)) = currentRole & ": " & artistName
										WriteLog "New Role: " & currentRole & ": " & artistName
									Else
										If InStr(Involved_R(tmp2), artistName) = 0 Then
											Involved_R(tmp2) = Involved_R(tmp2) & ArtistSeparator & artistName
											WriteLog "Role updated: " & Involved_R(tmp2)
										Else
											WriteLog "artist already inside role"
										End If
									End If
									Exit Do
								Loop While True
							Next
						End If
					Next
				End If
			End If
			' Get track titles and track artists

			WriteLog "End"
			iAutoTrackNumber = 1
			iAutoDiscNumber = 1
			iTrackNum = 0

			For Each m In CurrentRelease("media")
				Set currentMedia = CurrentRelease("media")(m)
				If CheckNoDisc = False Then
					iAutoTrackNumber = 1
				End If
				For Each t In CurrentMedia("tracks")
					Set currentTrack = currentMedia("tracks")(t)
					position = currentTrack("number")
					trackName = PackSpaces(DecodeHtmlChars(currentTrack("title")), False)
					length = currentTrack("length")
					If length <> "" And IsNumeric(length) Then
						min = Int(length / 60000)
						tmp = (length / 60000) - Int(length / 60000)
						If tmp <> 0 Then
							sec = Round(tmp * 60, 0)
						End If
						If sec < 10 Then
							length = CStr(min & ":" & "0" & sec)
						Else
							length = CStr(min & ":" & sec)
						End If
						Durations.Add length
					Else
						Durations.Add ""
					End If
					
					WriteLog " "
					WriteLog " "
					WriteLog "Position=" & position
					WriteLog "TrackName=" & trackname
					WriteLog "Duration=" & length

					' Here comes the new track/disc numbering methods
					If position <> "" Then
						If CurrentRelease("media").Count > 1 Then  ' More than 1 CD
							If CheckNoDisc = True Then 
								TracksCD.Add ""
							Else
								TracksCD.Add LeadingZeroDisc(CurrentMedia("position"))
							End If
						Else
							If CheckForceDisc = True Then
								TracksCD.Add LeadingZeroDisc("1")
							Else
								TracksCD.Add ""
							End If
						End If
						If UnselectedTracks(iTrackNum) <> "x" Then
							If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
								tracksNum.Add "0" & iAutoTrackNumber
								WriteLog "Tracknumber=" & iAutoTrackNumber
							Else
								tracksNum.Add iAutoTrackNumber
								WriteLog "Tracknumber=" & iAutoTrackNumber
							End If
							iAutoTrackNumber = iAutoTrackNumber + 1
						Else
							tracksNum.Add ""
						End If

					ElseIf currentTrack("length") = "" And currentTrack("title") = "-" Then
						tracksNum.Add ""
						tracksCD.Add ""
						UnselectedTracks(iTrackNum) = "x"
						UnselectedTrackNames(iTrackNum) = "x"
					Else ' Nothing specified
						If CheckForceNumeric And UnselectedTracks(iTrackNum) <> "x" Then
							If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
								tracksNum.Add "0" & iAutoTrackNumber
								WriteLog "Tracknumber=" & iAutoTrackNumber
							Else
								tracksNum.Add iAutoTrackNumber
								WriteLog "Tracknumber=" & iAutoTrackNumber
							End If
							iAutoTrackNumber = iAutoTrackNumber + 1
						Else
							tracksNum.Add ""
						End If
						If CheckForceDisc Then
							tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
						Else
							tracksCD.Add ""
						End If
					End If

					ReDim Involved_R_T(0)

					TrackInvolvedPeople = ""
					TrackComposers = ""
					TrackConductors = ""
					TrackProducers = ""
					TrackLyricists = ""
					TrackFeaturing = AlbumFeaturing
		
					If UBound(Involved_R) > 0 Then
						For tmp = 1 To UBound(Involved_R)
							ReDim Preserve Involved_R_T(tmp)
							Involved_R_T(tmp) = Involved_R(tmp)
						Next
					End If

					For tmp = 1 To UBound(TrackPos)
						If TrackPos(tmp) = position Then
							WriteLog "trackpos(" & tmp & ")=" & trackpos(tmp)
							involvedRole = TrackRoles(tmp)
							involvedArtist = TrackArtist2(tmp)
		
							If LookForFeaturing(involvedRole) And CheckIgnoreFeatArtist = False Then
								If InStr(TrackFeaturing, involvedArtist) = 0 Then
									If TrackFeaturing = "" Then
										If CheckFeaturingName Then
											TrackFeaturing = TxtFeaturingName & " " & involvedArtist
										Else
											TrackFeaturing = involvedRole & " " & involvedArtist
										End If
									Else
										TrackFeaturing = TrackFeaturing & ArtistSeparator & involvedArtist
									End If
								End If
								WriteLog "TrackFeaturing=" & TrackFeaturing
							Else
								Do
									ret = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
									If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
										TrackLyricists = ret
										WriteLog "TrackLyricists=" & TrackLyricists
										Exit Do
									ElseIf ret = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									ret = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
									If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
										TrackConductors = ret
										WriteLog "TrackConductors=" & TrackConductors
										Exit Do
									ElseIf ret = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									ret = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
									If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
										TrackProducers = ret
										WriteLog "TrackProducers=" & TrackProducers
										Exit Do
									ElseIf ret = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									ret = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
									If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
										TrackComposers = ret
										WriteLog "TrackComposers=" & TrackComposers
										Exit Do
									ElseIf ret = "ALREADY_INSIDE_ROLE" Then
										WriteLog "ALREADY_INSIDE_ROLE"
										Exit Do
									End If
									tmp2 = search_involved(Involved_R_T, involvedRole)
									If tmp2 = -2 Then
										WriteLog "Ignore Role: '" & currentRole & "' (Unwanted tag)"
									ElseIf tmp2 = -1 Then
										ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
										Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & TrackArtist2(tmp)
										WriteLog "New Role: " & involvedRole & ": " & TrackArtist2(tmp)
									Else
										If InStr(Involved_R_T(tmp2), TrackArtist2(tmp)) = 0 Then
											Involved_R_T(tmp2) = Involved_R_T(tmp2) & ArtistSeparator & TrackArtist2(tmp)
											WriteLog "Role updated: " & Involved_R_T(tmp2)
										Else
											WriteLog "artist already inside role"
										End If
									End If
									Exit Do
								Loop While True
							End If
						End If
					Next

					artistList = ""

					WriteLog " "
					WriteLog "Search for TrackArtist <> Release Artist"

					If currentTrack.Exists("artist-credit") Then
						tmp = getArtistsName(CurrentTrack, "artist-credit", QueryPage)
						artistList = tmp(0)
						TrackFeaturing = Trim(tmp(1))
					End If
					If artistList = "" Then artistList = AlbumArtistTitle

					WriteLog "artistlist=" & artistlist

					Set tmp = currentTrack("recording")
					If tmp.Exists("relations") Then
						If tmp("relations").Count > 0 Then
							For Each tmp2 in tmp("relations")
								WriteLog " "
								involvedArtist = ""
								Set tmp3 = tmp("relations")(tmp2)
								If tmp3.Exists("work") And tmp3("type") <> "other version" Then
									Set tmp3 = tmp("relations")(tmp2)("work")
									For Each tmp4 In tmp3("relations")
										If tmp3("relations")(tmp4).Exists("artist") Then
											involvedArtist = CleanArtistName(tmp3("relations")(tmp4)("artist")("name"))
										Else
											involvedArtist = ""
										End If
										involvedRole = ""
										Set tmp5 = tmp3("relations")(tmp4)
										If tmp5.Exists("attributes") Then
											ret = RelationshipTypes(tmp5)
											If ret(0) > 0 Then
												For i = 1 to ret(0)
													If involvedArtist <> "" Then getinvolvedRole involvedArtist, ret(i), artistList, TrackFeaturing, Involved_R_T, TrackComposers, TrackConductors, TrackProducers, TrackLyricists
												Next
											End If
										End If
									Next
								Else
									If tmp3.Exists("artist") Then
										involvedArtist = CleanArtistName(tmp3("artist")("name"))
									Else
										involvedArtist = ""
									End If
									ret = RelationshipTypes(tmp3)
									If ret(0) > 0 Then
										For i = 1 to ret(0)
											If involvedArtist <> "" Then getinvolvedRole involvedArtist, ret(i), artistList, TrackFeaturing, Involved_R_T, TrackComposers, TrackConductors, TrackProducers, TrackLyricists
										Next
									End If
								End If
							Next
						End If
					End If
					WriteLog "TrackArtist end"

					If TrackFeaturing <> "" And CheckIgnoreFeatArtist = False Then
						If CheckTitleFeaturing = True Then
							tmp = InStrRev(TrackFeaturing, ArtistSeparator)
							If tmp = 0 Or ArtistLastSeparator = False Then
								trackName = trackName & " (" & TrackFeaturing & ")"
							Else
								trackName = trackName & " (" &  Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+Len(ArtistSeparator)) & ")"
							End If
						Else
							tmp = InStrRev(TrackFeaturing, ArtistSeparator)
							If tmp = 0 Or ArtistLastSeparator = False Then
								If Left(TrackFeaturing, 1) = "," Or Left(TrackFeaturing, 1) = ";" Then
									artistList = artistList & TrackFeaturing
								Else
									artistList = artistList & " " & TrackFeaturing
								End If
							Else
								artistList = artistList & " " & Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+Len(ArtistSeparator))
							End If
						End If
					End If

					ArtistTitles.Add artistList
		
					TrackLyricists = FindArtist(TrackLyricists, AlbumLyricist)
					If AlbumLyricist <> "" and TrackLyricists <> "" Then
						Lyricists.Add AlbumLyricist & Separator & TrackLyricists
					Else
						Lyricists.Add AlbumLyricist & TrackLyricists
					End If
					TrackComposers = FindArtist(TrackComposers, AlbumComposer)
					If AlbumComposer <> "" and TrackComposers <> "" Then
						Composers.Add AlbumComposer & Separator & TrackComposers
					Else
						Composers.Add AlbumComposer & TrackComposers
					End If
					TrackConductors = FindArtist(TrackConductors, AlbumConductor)
					If AlbumConductor <> "" and TrackConductors <> "" Then
						Conductors.Add AlbumConductor & Separator & TrackConductors
					Else
						Conductors.Add AlbumConductor & TrackConductors
					End If
					TrackProducers = FindArtist(TrackProducers, AlbumProducer)
					If AlbumProducer <> "" and TrackProducers <> "" Then
						Producers.Add AlbumProducer & Separator & TrackProducers
					Else
						Producers.Add AlbumProducer & TrackProducers
					End If
		
					If UBound(Involved_R_T) > 0 Then
						For tmp = 1 To UBound(involved_R_T)
							TrackInvolvedPeople = TrackInvolvedPeople & Involved_R_T(tmp) & Separator
						Next
						TrackInvolvedPeople = Left(TrackInvolvedPeople, Len(TrackInvolvedPeople)- Len(Separator))
					Else
						TrackInvolvedPeople = ""
					End If
		
					InvolvedArtists.Add TrackInvolvedPeople
					Tracks.Add trackName
					iTrackNum = iTrackNum + 1
				Next
			Next


			' Get album title
			AlbumTitle = PackSpaces(currentRelease("title"), False)
			Dim json
			Set json = New VbsJson
			' Get Album art URL
			WriteLog " "
			WriteLog "Search for cover"
			If CurrentRelease.Exists("cover-art-archive") Then
				WriteLog CurrentRelease("cover-art-archive")("count") & " Images on CoverArtArchive found"
				If CurrentRelease("cover-art-archive")("count") > 0 And CurrentRelease("cover-art-archive")("front") = True Then
					searchURL = "http://coverartarchive.org/release/" & CurrentReleaseId
					WriteLog searchURL
					Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
					oXMLHTTP.Open "GET", searchURL, False
					oXMLHTTP.setRequestHeader "Content-Type", "application/json"
					oXMLHTTP.setRequestHeader "User-Agent",UserAgent
					oXMLHTTP.send ()

					If oXMLHTTP.Status = 200 Then
						Set currentImage = json.Decode(oXMLHTTP.responseText)
						For Each image In currentImage("images")
							Set tmp = currentImage("images")(image)
							If tmp("front") = True Then
								AlbumArtURL = tmp("image")
								AlbumArtThumbNail = tmp("thumbnails")("small")
								Exit For
							End If
						Next
					End If
				End If
			End If

			'----------------------------------CoverArtArchive Images----------------------------------------
			Set ImageList = SDB.NewStringList
			Set SaveImageType = SDB.NewStringList
			Set SaveImage = SDB.NewStringList
			ImagesCount = 0

			If CurrentRelease.Exists("cover-art-archive") And CurrentRelease("cover-art-archive")("count") <> "" Then
				If CurrentRelease("cover-art-archive")("count") > 1 Then
					searchURL = "http://coverartarchive.org/release/" & CurrentRelease("id")
					WriteLog searchURL
					ImagesCount = CurrentRelease("cover-art-archive")("count")
					Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
					oXMLHTTP.Open "GET", searchURL, False
					oXMLHTTP.setRequestHeader "Content-Type", "application/json"
					oXMLHTTP.setRequestHeader "User-Agent",UserAgent
					oXMLHTTP.send ()

					If oXMLHTTP.Status = 200 Then
						Set currentImage = json.Decode(oXMLHTTP.responseText)
						For Each image In currentImage("images")
							Set tmp = currentImage("images")(image)
							If tmp("image") <> AlbumArtURL Then
								ImageList.add tmp("thumbnails")("small")
								SaveImageType.add "other"
								SaveImage.add "0"
							End If
						Next
					End If
				End If
			End If
			'----------------------------------CoverArtArchive Images----------------------------------------

			' Get release year/date
			If CurrentRelease.Exists("date") And CurrentRelease("date") <> "" Then
				ReleaseDate = CurrentRelease("date")
				If Len(ReleaseDate) = 10 Then
					ReleaseSplit = Split(ReleaseDate,"-")
					ReleaseDate = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				ElseIf Len(ReleaseDate) = 7 Then
					ReleaseSplit = Split(ReleaseDate,"-")
					ReleaseDate = "00-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				End If
				If CheckYearOnlyDate Then
					ReleaseDate = Right(ReleaseDate, 4)
				End If
			Else
				ReleaseDate = ""
			End If
			WriteLog "ReleaseDate=" & ReleaseDate

			'Set OriginalDate
			set tmp = CurrentRelease("release-group")
			If tmp.Exists("first-release-date") And CurrentRelease("release-group")("first-release-date") <> "" Then
				OriginalDate = CurrentRelease("release-group")("first-release-date")
				If Len(OriginalDate) = 10 Then
					ReleaseSplit = Split(OriginalDate,"-")
					OriginalDate = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				ElseIf Len(OriginalDate) = 7 Then
					ReleaseSplit = Split(OriginalDate,"-")
					OriginalDate = "00-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				End If
				If CheckYearOnlyDate Then
					OriginalDate = Right(OriginalDate, 4)
				End If
			Else
				OriginalDate = ""
			End If
			WriteLog "OriginalDate=" & OriginalDate

			' Get Label
			If CurrentRelease.Exists("label-info") Then
				For Each l In CurrentRelease("label-info")
					Set currentLabel = CurrentRelease("label-info")(l)
					If SavedLabelID = "" And Not IsNull(currentLabel("label")) Then
						set tmp = currentLabel("label")
						If tmp.Exists("id") Then
							SavedLabelID = tmp("id")
						End If
						If Not IsNull(currentLabel("label")("name")) Then
							AddToField theLabels, CleanArtistName(currentLabel("label")("name"))
						End If
					End If
					If IsNull(currentLabel("catalog-number")) Then
						theCatalogs = ""
					Else
						AddToField theCatalogs, currentLabel("catalog-number")
					End If
				Next
			Else
				theLabels = ""
				theCatalogs = ""
			End If
			WriteLog "theLabels=" & theLabels
			WriteLog "theCatalogs=" & theCatalogs

			' Get Country
			If CurrentRelease.Exists("country") And CurrentRelease("country") <> "" Then
				theCountry = CurrentRelease("country")
				For f = 1 To CountryCode.Count
					If theCountry = CountryCode.Item(f) Then
						theCountry = CountryList.Item(f)
						Exit For
					End If
				Next
			Else
				theCountry = ""
			End If
			WriteLog "country=" & theCountry

			' Get Format
			If CurrentRelease.Exists("media") Then
				For Each f In CurrentRelease("media")
					Set currentFormat = CurrentRelease("media")(f)
					If currentFormat("format") <> "" And theFormat <> currentFormat("format") Then
						AddToField theFormat, currentFormat("format")
					End If
				Next
				WriteLog "theformat=" & theformat
			End If
			tmp = ""
			If CurrentRelease.Exists("release-group") Then
				tmp = CurrentRelease("release-group")("primary-type")
				If theFormat <> "" Then
					theFormat = theFormat & FormatSeparator & tmp
				Else
					theFormat = tmp
				End If
				Set tmp = CurrentRelease("release-group")
				If tmp.Exists("secondary-types") Then
					For Each f In tmp("secondary-types")
						If CurrentRelease("release-group")("secondary-types")(f) <> "" Then
							theFormat = theFormat & FormatSeparator & CurrentRelease("release-group")("secondary-types")(f)
						End If
					Next
				End If
			End If
			WriteLog "theformat=" & theFormat
		End If
	End If

	FormatSearchResultsViewer Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, GenresList, theLabels, theCountry, AlbumArtThumbNail, CurrentReleaseId, theCatalogs, Lyricists, Composers, Conductors, Producers, InvolvedArtists, theFormat, theMaster, comment, DiscogsTracksNum, DataQuality, Grouping

	Dim SelectedTracks, j, k, SelTracksCount, StartCount
	Set SelectedTracks = SDB.NewStringList
	Set SelectedSongsGlobal = SDB.NewSongList
	'Tracks = Stringliste mit den Songnamen
	WriteLog "-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+"


	WriteLog "DiffNumberTracks=" & DiffNumberTracks
	
	
	If CheckNewMatchingEngine = False Then
		For i = 0 To Tracks.Count - 1
			If UnselectedTracks(i) = "" Then
				If UnselectedTrackNames(i) = "" Then
					SelectedTracks.Add Tracks.Item(i)
				Else
					If SDB.Tools.WebSearch.NewTracks.Count - 1 < i Then
						SelectedTracks.Add Tracks.Item(i)
					Else
						SelectedTracks.Add SDB.Tools.WebSearch.NewTracks.Item(i).Title
					End If
				End If
			End If
		Next
		For i = 0 To SDB.Tools.WebSearch.NewTracks.Count -1
			SelectedSongsGlobal.Add SDB.Tools.WebSearch.NewTracks.Item(i)
		Next

		SDB.Tools.WebSearch.SmartUpdateTracks SelectedTracks

		For i = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1
			If CheckArtist Then SDB.Tools.WebSearch.NewTracks.Item(i).ArtistName = AlbumArtistTitle
			For j = 0 To Tracks.Count - 1
				If Tracks.Item(j) = SDB.Tools.WebSearch.NewTracks.Item(i).Title Then
					If UnselectedTracks(j) = "" Then
						If UnselectedTrackNames(j) = "" Then
							SDB.Tools.WebSearch.NewTracks.Item(i).Title = Tracks.Item(j)
						Else
							SDB.Tools.WebSearch.NewTracks.Item(i).Title = SDB.Tools.WebSearch.NewTracks.Item(i).Title
						End If
						If CheckArtist And ((CheckDontFillEmptyFields = True And ArtistTitles.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).ArtistName = ArtistTitles.Item(j)
						If CheckTrackNum And ((CheckDontFillEmptyFields = True And TracksNum.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).TrackOrderStr = TracksNum.Item(j)
						If CheckDiscNum And ((CheckDontFillEmptyFields = True And TracksCD.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).DiscNumberStr = TracksCD.Item(j)
						If CheckInvolved And ((CheckDontFillEmptyFields = True And InvolvedArtists.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).InvolvedPeople = InvolvedArtists.Item(j)
						If CheckGrouping And ((CheckDontFillEmptyFields = True And Grouping(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Grouping = Grouping(j)
						If CheckLyricist And ((CheckDontFillEmptyFields = True And Lyricists.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Lyricist = Lyricists.Item(j)
						If CheckComposer And ((CheckDontFillEmptyFields = True And Composers.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Author = Composers.Item(j)
						If CheckConductor And ((CheckDontFillEmptyFields = True And Conductors.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Conductor = Conductors.Item(j)
						If CheckProducer And ((CheckDontFillEmptyFields = True And Producers.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Producer = Producers.Item(j)
				 
					End If
				End If
			Next
			If CheckAlbumArtist And ((CheckDontFillEmptyFields = True And AlbumArtist <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumArtistName = AlbumArtist
			If CheckAlbum And ((CheckDontFillEmptyFields = True And AlbumTitle <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumName = AlbumTitle

			If CheckDate Then
				If Len(ReleaseDate) > 4 Then
					SDB.Tools.WebSearch.NewTracks.Item(i).Year = Mid(ReleaseDate,7,4)
					SDB.Tools.WebSearch.NewTracks.Item(i).Month = Mid(ReleaseDate,4,2)
					SDB.Tools.WebSearch.NewTracks.Item(i).Day = Mid(ReleaseDate,1,2)
				ElseIf IsNumeric(ReleaseDate) Then
					SDB.Tools.WebSearch.NewTracks.Item(i).Year = ReleaseDate
				ElseIf ReleaseDate = "" Then
					If CheckDontFillEmptyFields = False Then
						SDB.Tools.WebSearch.NewTracks.Item(i).Year = -1
					End If
				End If
			End If
			If CheckOrigDate Then
				If Len(OriginalDate) > 4 Then
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = Mid(OriginalDate,7,4)
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalMonth = Mid(OriginalDate,4,2)
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalDay = Mid(OriginalDate,1,2)
				ElseIf IsNumeric(OriginalDate) Then
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = OriginalDate
				ElseIf OriginalDate = "" Then
					If CheckDontFillEmptyFields = False Then
						SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = -1
					End If
				End If
			End If

			If CheckGenre Then
				If Not(NewGenre = "" And CheckDontFillEmptyFields = True) Then
					SDB.Tools.WebSearch.NewTracks.Item(i).Genre = NewGenre
				End If
			End If

			If CheckLabel And ((CheckDontFillEmptyFields = True And theLabels <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Publisher = theLabels

			If CheckComment And ((CheckDontFillEmptyFields = True And comment <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Comment = comment
		
			If CheckRelease And ((CheckDontFillEmptyFields = True And CurrentReleaseId <> "") Or CheckDontFillEmptyFields = False) Then
				If ReleaseTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = CurrentReleaseId
				If ReleaseTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = CurrentReleaseId
				If ReleaseTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = CurrentReleaseId
				If ReleaseTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = CurrentReleaseId
				If ReleaseTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = CurrentReleaseId
				If ReleaseTag = "Grouping" Then SDB.Tools.WebSearch.NewTracks.Item(i).Grouping = CurrentReleaseId
				If ReleaseTag = "ISRC" Then SDB.Tools.WebSearch.NewTracks.Item(i).ISRC = CurrentReleaseId
				If ReleaseTag = "Encoder" Then SDB.Tools.WebSearch.NewTracks.Item(i).Encoder = CurrentReleaseId
				If ReleaseTag = "Copyright" Then SDB.Tools.WebSearch.NewTracks.Item(i).Copyright = CurrentReleaseId
			End If

			If CheckCatalog And ((CheckDontFillEmptyFields = True And theCatalogs <> "") Or CheckDontFillEmptyFields = False) Then
				If CatalogTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCatalogs
				If CatalogTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCatalogs
				If CatalogTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCatalogs
				If CatalogTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCatalogs
				If CatalogTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCatalogs
				If CatalogTag = "ISRC" Then SDB.Tools.WebSearch.NewTracks.Item(i).ISRC = theCatalogs
			End If

			If CheckCountry And ((CheckDontFillEmptyFields = True And theCountry <> "") Or CheckDontFillEmptyFields = False) Then
				If CountryTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCountry
				If CountryTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCountry
				If CountryTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCountry
				If CountryTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCountry
				If CountryTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCountry
			End If

			If CheckFormat And ((CheckDontFillEmptyFields = True And theFormat <> "") Or CheckDontFillEmptyFields = False) Then
				If FormatTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theFormat
				If FormatTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theFormat
				If FormatTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theFormat
				If FormatTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theFormat
				If FormatTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theFormat
			End If								  
		Next
	Else
	
		If SDB.Tools.WebSearch.NewTracks.Count = 1 Then
			WriteLog "Only one track for tagging selected"
			SelectedTracks.Add SDB.Tools.WebSearch.NewTracks.Item(0).Title
			SDB.Tools.WebSearch.SmartUpdateTracks SelectedTracks
			SelectedSongsGlobal.Add SDB.Tools.WebSearch.NewTracks.Item(0)

			i = 0
			For j = 0 To Tracks.Count - 1
				If UnselectedTracks(j) = "" Then
					If UnselectedTrackNames(j) = "" Then
						SDB.Tools.WebSearch.NewTracks.Item(i).Title = Tracks.Item(j)
					Else
						SDB.Tools.WebSearch.NewTracks.Item(i).Title = SDB.Tools.WebSearch.NewTracks.Item(i).Title
					End If
					If CheckArtist And ((CheckDontFillEmptyFields = True And ArtistTitles.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).ArtistName = ArtistTitles.Item(j)
					If CheckTrackNum And ((CheckDontFillEmptyFields = True And TracksNum.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).TrackOrderStr = TracksNum.Item(j)
					If CheckDiscNum And ((CheckDontFillEmptyFields = True And TracksCD.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).DiscNumberStr = TracksCD.Item(j)
					If CheckInvolved And ((CheckDontFillEmptyFields = True And InvolvedArtists.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).InvolvedPeople = InvolvedArtists.Item(j)
					If CheckGrouping And ((CheckDontFillEmptyFields = True And Grouping(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Grouping = Grouping(j)
					If CheckLyricist And ((CheckDontFillEmptyFields = True And Lyricists.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Lyricist = Lyricists.Item(j)
					If CheckComposer And ((CheckDontFillEmptyFields = True And Composers.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Author = Composers.Item(j)
					If CheckConductor And ((CheckDontFillEmptyFields = True And Conductors.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Conductor = Conductors.Item(j)
					If CheckProducer And ((CheckDontFillEmptyFields = True And Producers.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Producer = Producers.Item(j)
					Exit For
				End If
			Next
			If CheckAlbumArtist And ((CheckDontFillEmptyFields = True And AlbumArtist <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumArtistName = AlbumArtist
			If CheckAlbum And ((CheckDontFillEmptyFields = True And AlbumTitle <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumName = AlbumTitle

			If CheckDate Then
				If Len(ReleaseDate) > 4 Then
					If QueryPage = "MetalArchives" Then
						ReleaseDate = convert_date(ReleaseDate)
					End If
					SDB.Tools.WebSearch.NewTracks.Item(i).Year = Mid(ReleaseDate,7,4)
					SDB.Tools.WebSearch.NewTracks.Item(i).Month = Mid(ReleaseDate,4,2)
					SDB.Tools.WebSearch.NewTracks.Item(i).Day = Mid(ReleaseDate,1,2)
				ElseIf IsNumeric(ReleaseDate) Then
					SDB.Tools.WebSearch.NewTracks.Item(i).Year = ReleaseDate
				ElseIf ReleaseDate = "" Then
					If CheckDontFillEmptyFields = False Then
						SDB.Tools.WebSearch.NewTracks.Item(i).Year = -1
					End If
				End If
			End If
			If CheckOrigDate And QueryPage <> "MetalArchives" Then
				If Len(OriginalDate) > 4 Then
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = Mid(OriginalDate,7,4)
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalMonth = Mid(OriginalDate,4,2)
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalDay = Mid(OriginalDate,1,2)
				ElseIf IsNumeric(OriginalDate) Then
					SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = OriginalDate
				ElseIf OriginalDate = "" Then
					If CheckDontFillEmptyFields = False Then
						SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = -1
					End If
				End If
			End If

			If CheckGenre Then
				If Not(NewGenre = "" And CheckDontFillEmptyFields = True) Then
					SDB.Tools.WebSearch.NewTracks.Item(i).Genre = NewGenre
				End If
			End If

			If CheckLabel And ((CheckDontFillEmptyFields = True And theLabels <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Publisher = theLabels

			If CheckComment And ((CheckDontFillEmptyFields = True And comment <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Comment = comment
		
			If CheckRelease And ((CheckDontFillEmptyFields = True And CurrentReleaseId <> "") Or CheckDontFillEmptyFields = False) Then
				If ReleaseTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = CurrentReleaseId
				If ReleaseTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = CurrentReleaseId
				If ReleaseTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = CurrentReleaseId
				If ReleaseTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = CurrentReleaseId
				If ReleaseTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = CurrentReleaseId
				If ReleaseTag = "Grouping" Then SDB.Tools.WebSearch.NewTracks.Item(i).Grouping = CurrentReleaseId
				If ReleaseTag = "ISRC" Then SDB.Tools.WebSearch.NewTracks.Item(i).ISRC = CurrentReleaseId
				If ReleaseTag = "Encoder" Then SDB.Tools.WebSearch.NewTracks.Item(i).Encoder = CurrentReleaseId
				If ReleaseTag = "Copyright" Then SDB.Tools.WebSearch.NewTracks.Item(i).Copyright = CurrentReleaseId
			End If

			If CheckCatalog And ((CheckDontFillEmptyFields = True And theCatalogs <> "") Or CheckDontFillEmptyFields = False) Then
				If CatalogTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCatalogs
				If CatalogTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCatalogs
				If CatalogTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCatalogs
				If CatalogTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCatalogs
				If CatalogTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCatalogs
				If CatalogTag = "ISRC" Then SDB.Tools.WebSearch.NewTracks.Item(i).ISRC = theCatalogs
			End If

			If CheckCountry And ((CheckDontFillEmptyFields = True And theCountry <> "") Or CheckDontFillEmptyFields = False) Then
				If CountryTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCountry
				If CountryTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCountry
				If CountryTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCountry
				If CountryTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCountry
				If CountryTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCountry
			End If

			If CheckFormat And ((CheckDontFillEmptyFields = True And theFormat <> "") Or CheckDontFillEmptyFields = False) Then
				If FormatTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theFormat
				If FormatTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theFormat
				If FormatTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theFormat
				If FormatTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theFormat
				If FormatTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theFormat
			End If
		Else
			WriteLog SDB.Tools.WebSearch.NewTracks.Count & " tracks for tagging selected"
			For i = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1
				WriteLog "SDB.Tools.WebSearch.NewTracks.Item(i).Title=" & SDB.Tools.WebSearch.NewTracks.Item(i).Title
			Next
			WriteLog ""
			For i = 0 To Tracks.Count - 1
				WriteLog "Tracks.Item(i)=" & Tracks.Item(i)
			Next
			WriteLog ""
			SelTracksCount = 0
			For i = 0 To Tracks.Count - 1
				If UnselectedTracks(i) = "" Then
					If UnselectedTrackNames(i) = "" Then
						SelTracksCount = SelTracksCount + 1
					Else
						SelTracksCount = SelTracksCount + 1
					End If
					If SDB.Tools.WebSearch.NewTracks.Count = SelTracksCount Then
						Exit For
					End If
				End If
			Next
			StartCount = 0
			If SelTracksCount < SDB.Tools.WebSearch.NewTracks.Count Then
				StartCount = SDB.Tools.WebSearch.NewTracks.Count - SelTracksCount
			End If
			WriteLog "StartCount=" & StartCount
			SelTracksCount = 0

			For i = 0 To Tracks.Count - 1
				If UnselectedTracks(i) = "" Then
					If UnselectedTrackNames(i) = "" Then
						SelectedTracks.Add Tracks.Item(i)
						SelTracksCount = SelTracksCount + 1
						StartCount = StartCount + 1
					Else
						SelectedTracks.Add SDB.Tools.WebSearch.NewTracks.Item(StartCount).Title
						SelTracksCount = SelTracksCount + 1
						StartCount = StartCount + 1
						WriteLog "SelTracksCount=" & SelTracksCount
						WriteLog "SDB.Tools.WebSearch.NewTracks.Count=" & SDB.Tools.WebSearch.NewTracks.Count
					End If
					If SDB.Tools.WebSearch.NewTracks.Count = SelTracksCount Then
						Exit For
					End If
				End If
			Next
			For i = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1
				SelectedSongsGlobal.Add SDB.Tools.WebSearch.NewTracks.Item(i)
			Next
			For i = 0 to SelTracksCount-1
				WriteLog "SelectedTracks=" & SelectedTracks.Item(i)
			Next

			SDB.Tools.WebSearch.SmartUpdateTracks SelectedTracks

			StartCount = 0
			If SelTracksCount < SDB.Tools.WebSearch.NewTracks.Count Then
				StartCount = SDB.Tools.WebSearch.NewTracks.Count - SelTracksCount
			End If
			WriteLog "StartCount=" & StartCount
			SelTracksCount = 0

			For i = StartCount To SDB.Tools.WebSearch.NewTracks.Count - 1
				For j = SelTracksCount To Tracks.Count - 1
					WriteLog "SelTracksCount=" & SelTracksCount
					WriteLog "j=" & j
					If UnselectedTracks(j) = "" Then
						If UnselectedTrackNames(j) = "" Then
							WriteLog "SDB.Tools.WebSearch.NewTracks.Item(i).Title=" & SDB.Tools.WebSearch.NewTracks.Item(i).Title
							SDB.Tools.WebSearch.NewTracks.Item(i).Title = Tracks.Item(j)
							WriteLog "Tracks.Item(j)=" & Tracks.Item(j) & " - j=" & j
						Else
							WriteLog "Trackname nicht gewählt=" & SDB.Tools.WebSearch.NewTracks.Item(i).Title
							SDB.Tools.WebSearch.NewTracks.Item(i).Title = SDB.Tools.WebSearch.NewTracks.Item(i).Title
						End If
						If CheckArtist And ((CheckDontFillEmptyFields = True And ArtistTitles.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).ArtistName = ArtistTitles.Item(j)
						If CheckTrackNum And ((CheckDontFillEmptyFields = True And TracksNum.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).TrackOrderStr = TracksNum.Item(j)
						If CheckDiscNum And ((CheckDontFillEmptyFields = True And TracksCD.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).DiscNumberStr = TracksCD.Item(j)
						If CheckInvolved And ((CheckDontFillEmptyFields = True And InvolvedArtists.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).InvolvedPeople = InvolvedArtists.Item(j)
						If CheckGrouping And ((CheckDontFillEmptyFields = True And Grouping(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Grouping = Grouping(j)
						If CheckLyricist And ((CheckDontFillEmptyFields = True And Lyricists.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Lyricist = Lyricists.Item(j)
						If CheckComposer And ((CheckDontFillEmptyFields = True And Composers.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Author = Composers.Item(j)
						If CheckConductor And ((CheckDontFillEmptyFields = True And Conductors.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Conductor = Conductors.Item(j)
						If CheckProducer And ((CheckDontFillEmptyFields = True And Producers.Item(j) <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Producer = Producers.Item(j)
						SelTracksCount = SelTracksCount + 1
						Exit For
					End If
					SelTracksCount = SelTracksCount + 1
				Next

				If CheckAlbumArtist And ((CheckDontFillEmptyFields = True And AlbumArtist <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumArtistName = AlbumArtist
				If CheckAlbum And ((CheckDontFillEmptyFields = True And AlbumTitle <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumName = AlbumTitle

				If CheckDate Then
					If Len(ReleaseDate) > 4 Then
						If QueryPage = "MetalArchives" Then
							ReleaseDate = convert_date(ReleaseDate)
						End If
						SDB.Tools.WebSearch.NewTracks.Item(i).Year = Mid(ReleaseDate,7,4)
						SDB.Tools.WebSearch.NewTracks.Item(i).Month = Mid(ReleaseDate,4,2)
						SDB.Tools.WebSearch.NewTracks.Item(i).Day = Mid(ReleaseDate,1,2)
					ElseIf IsNumeric(ReleaseDate) Then
						SDB.Tools.WebSearch.NewTracks.Item(i).Year = ReleaseDate
					ElseIf ReleaseDate = "" Then
						If CheckDontFillEmptyFields = False Then
							SDB.Tools.WebSearch.NewTracks.Item(i).Year = -1
						End If
					End If
				End If
				If CheckOrigDate And QueryPage <> "MetalArchives" Then
					If Len(OriginalDate) > 4 Then
						SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = Mid(OriginalDate,7,4)
						SDB.Tools.WebSearch.NewTracks.Item(i).OriginalMonth = Mid(OriginalDate,4,2)
						SDB.Tools.WebSearch.NewTracks.Item(i).OriginalDay = Mid(OriginalDate,1,2)
					ElseIf IsNumeric(OriginalDate) Then
						SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = OriginalDate
					ElseIf OriginalDate = "" Then
						If CheckDontFillEmptyFields = False Then
							SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = -1
						End If
					End If
				End If

				If CheckGenre And QueryPage <> "MetalArchives" Then
					If Not(NewGenre = "" And CheckDontFillEmptyFields = True) Then
						SDB.Tools.WebSearch.NewTracks.Item(i).Genre = NewGenre
					End If
				End If

				If CheckLabel And ((CheckDontFillEmptyFields = True And theLabels <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Publisher = theLabels

				If CheckComment And ((CheckDontFillEmptyFields = True And comment <> "") Or CheckDontFillEmptyFields = False) Then SDB.Tools.WebSearch.NewTracks.Item(i).Comment = comment
			
				If CheckRelease And QueryPage <> "MetalArchives" And ((CheckDontFillEmptyFields = True And CurrentReleaseId <> "") Or CheckDontFillEmptyFields = False) Then
					If ReleaseTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = CurrentReleaseId
					If ReleaseTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = CurrentReleaseId
					If ReleaseTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = CurrentReleaseId
					If ReleaseTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = CurrentReleaseId
					If ReleaseTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = CurrentReleaseId
					If ReleaseTag = "Grouping" Then SDB.Tools.WebSearch.NewTracks.Item(i).Grouping = CurrentReleaseId
					If ReleaseTag = "ISRC" Then SDB.Tools.WebSearch.NewTracks.Item(i).ISRC = CurrentReleaseId
					If ReleaseTag = "Encoder" Then SDB.Tools.WebSearch.NewTracks.Item(i).Encoder = CurrentReleaseId
					If ReleaseTag = "Copyright" Then SDB.Tools.WebSearch.NewTracks.Item(i).Copyright = CurrentReleaseId
				End If

				If CheckCatalog And ((CheckDontFillEmptyFields = True And theCatalogs <> "") Or CheckDontFillEmptyFields = False) Then
					If CatalogTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCatalogs
					If CatalogTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCatalogs
					If CatalogTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCatalogs
					If CatalogTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCatalogs
					If CatalogTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCatalogs
					If CatalogTag = "ISRC" Then SDB.Tools.WebSearch.NewTracks.Item(i).ISRC = theCatalogs
				End If

				If CheckCountry And QueryPage <> "MetalArchives" And ((CheckDontFillEmptyFields = True And theCountry <> "") Or CheckDontFillEmptyFields = False) Then
					If CountryTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCountry
					If CountryTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCountry
					If CountryTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCountry
					If CountryTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCountry
					If CountryTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCountry
				End If

				If CheckFormat And ((CheckDontFillEmptyFields = True And theFormat <> "") Or CheckDontFillEmptyFields = False) Then
					If FormatTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theFormat
					If FormatTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theFormat
					If FormatTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theFormat
					If FormatTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theFormat
					If FormatTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theFormat
				End If
			Next
		End If
	End If

	If CheckCover Then
		SDB.Tools.WebSearch.AlbumArtURL = AlbumArtURL
	ElseIf SmallCover Then
		SDB.Tools.WebSearch.AlbumArtURL = AlbumArtThumbNail
	Else
		SDB.Tools.WebSearch.AlbumArtURL = ""
	End If

	NewResult = False

	SDB.Tools.WebSearch.RefreshViews
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
	button.style.backgroundColor = "#FFFFFF"
	templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML=""
	WriteLog " "
	WriteLog "Stop ReloadResults"

End Sub


Function convert_date(ReleaseDate)

	Dim Day, Mon, Year, tmp, tmp2
	Day = ""
	Mon = ""
	Year = ""
	If ReleaseDate = "Unknown" Then
		ReleaseDate = ""
	Else
		If InStr(ReleaseDate, ",") = 0 And InStr(ReleaseDate, " ") > 0 Then
			tmp = InStr(ReleaseDate, " ")
			Mon = Left(ReleaseDate, tmp - 1)
			Year = Mid(ReleaseDate, tmp + 1)
		ElseIf InStr(ReleaseDate, ",") > 0 Then
			tmp = InStr(ReleaseDate, ",")
			Year = Mid(ReleaseDate, tmp + 2)
			tmp2 = Left(ReleaseDate, tmp -1)
			tmp = InStr(tmp2, " ")
			Mon = Left(tmp2, tmp - 1)
			Day = Mid(tmp2, tmp + 1)
		End If
		If Mon = "January" Then Mon = "01"
		If Mon = "February" Then Mon = "02"
		If Mon = "March" Then Mon = "03"
		If Mon = "April" Then Mon = "04"
		If Mon = "May" Then Mon = "05"
		If Mon = "June" Then Mon = "06"
		If Mon = "July" Then Mon = "07"
		If Mon = "August" Then Mon = "08"
		If Mon = "September" Then Mon = "09"
		If Mon = "October" Then Mon = "10"
		If Mon = "November" Then Mon = "11"
		If Mon = "December" Then Mon = "12"
		
		If Day <> "" Then
			tmp = Len(Day)
			Day = Left(Day, tmp - 2)
			If Int(Day) < 10 Then
				Day = "0" & Day
			End If
		End If
		If Year <> "" And Mon <> "" And Day <> "" Then
			ReleaseDate = Day & "." & Mon & "." & Year
		ElseIf Year <> "" And Mon <> "" And Day = "" Then
			ReleaseDate = "00." & Mon & "." & Year
		End If
	End If

	convert_date = ReleaseDate

End Function


Function RelationshipTypes(JSONArray)

	Dim rType, i, a, nType, found, instrument, cInstr
	ReDim role(0)
	WriteLog "Start RelTypes"
	SDB.ProcessMessages
	rType = JSONArray("type")
	nType = ""
	cInstr = 0
	
	If rType <> "other version" Then
		If JSONArray("attributes").Count > 0 Then
			For a = 0 to JSONArray("attributes").Count -1
				found = false
				For i = 0 to RelationAttrList.Count -1
					If LCase(RelationAttrList.Item(i)) = JSONArray("attributes")(a) Then
						nType = nType & RelationAttrList.Item(i) & " "
						WriteLog "Attribute found: " & RelationAttrList.Item(i)
						found = true
					End If
				Next
				If found = false Then
					If rType = "instrument" Then
						cInstr = cInstr + 1
						ReDim Preserve role(cInstr+1)
						role(0) = cInstr
						role(cInstr) = nType & UCase(Left(JSONArray("attributes")(a), 1)) & Mid(JSONArray("attributes")(a), 2)
						WriteLog cInstr & ". Instrument: " & role(cInstr)
					End If
				End If
			Next
		End If
		If rType <> "instrument" Then
			If rType = "vocal" Then rType = "vocals"
			ReDim role(2)
			role(0) = 1
			role(1) = nType & UCase(Left(rType, 1)) & Mid(rType, 2)
			WriteLog "Found Role=" & role(1)
		End If
	Else
		WriteLog "other Version found!!"
		REM SDB.MessageBox "other Version found!!", mtInformation, Array(mbOk)
		ReDim role(1)
		role(0) = 0
	End If
	RelationshipTypes = role
	WriteLog "Stop RelTypes"

End Function



Function trackNumbering(ByRef pos, byRef position, byRef TracksNum, byRef TracksCD, byRef iTrackNum, FormatCnt)

	WriteLog "start trackNumbering"
	Dim TrackPosition

	If pos > 1 And CheckNoDisc = False And FormatCnt > 1 Then ' Disc Number Included
		WriteLog "Disc Number Included"
		If CheckForceNumeric Then
			If Left(position,2) = "CD" Then
				If iAutoDiscFormat <> "CD" Then
					iAutoDiscFormat = "CD"
					iAutoTrackNumber = 1
					iAutoDiscNumber = 1
				End If
				If Mid(position,3,1) = "-" Then
					iAutoDiscNumber = 1
				Else
					If iAutoDiscNumber <> Mid(position,3,1) Then
						iAutoTrackNumber = 1
					End If
				End If
			End If

			If Left(position,3) = "DVD" Then
				If iAutoDiscFormat <> "DVD" Then
					iAutoDiscFormat = "DVD"
					iAutoDiscNumber = 1
					iAutoTrackNumber = 1
				End If
				If Mid(position,4,1) = "-" Then
					iAutoDiscNumber = 1
				Else
					If iAutoDiscNumber <> Mid(position,4,1) Then
						iAutoTrackNumber = 1
					End If
				End If
			End If

			If Left(position,3) = "VHS" Then
				If iAutoDiscFormat <> "VHS" Then
					iAutoDiscFormat = "VHS"
					iAutoTrackNumber = 1
					iAutoDiscNumber = 1
				End If
				If Mid(position,4,1) = "-" Then
					iAutoDiscNumber = 1
				Else
					If iAutoDiscNumber <> Mid(position,4,1) Then
						iAutoTrackNumber = 1
					End If
				End If
			End If
			
			If Left(position,5) = "Video" Then
				If iAutoDiscFormat <> "Video" Then
					iAutoDiscFormat = "Video"
					iAutoTrackNumber = 1
					iAutoDiscNumber = 1
				End If
				If Mid(position,6,1) = "-" Then
					iAutoDiscNumber = 1
				Else
					If iAutoDiscNumber <> Mid(position,6,1) Then
						iAutoTrackNumber = 1
					End If
				End If
			End If

			If Left(position,2) <> "CD" And Left(position,3) <> "DVD" And Left(position,3) <> "VHS" And Left(position,5) <> "Video" And IsInteger(Left(position,pos-1)) Then
				If Int(iAutoDiscNumber) <> Int(Left(position,pos-1)) Then
					iAutoTrackNumber = 1
				End If
			End If

			If UnselectedTracks(iTrackNum) <> "x" Then
				If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
					tracksNum.Add "0" & iAutoTrackNumber
					WriteLog "Tracknumber=0" & iAutoTrackNumber
				Else
					tracksNum.Add iAutoTrackNumber
					WriteLog "Tracknumber=" & iAutoTrackNumber
				End If
				iAutoTrackNumber = iAutoTrackNumber + 1
			Else
				tracksNum.Add ""
				WriteLog "Tracknumber=Nothing"
			End If
		Else
			If pos > 0 Then
				TrackPosition = Mid(position, pos+1)
				If Len(TrackPosition) > 1 Then	'minimum 2 Char after -  (1-1a, 1-II, 1-12)
					If IsInteger(Left(TrackPosition, 1)) And Not IsInteger(Right(Trackposition, 1)) Then	'Minimum first is a Number, Char at the end (1-1a, 1-1b, 1-1c,...) = maybe Sub-Track !
						tracksNum.Add TrackPosition
						WriteLog "Tracknumber=" & TrackPosition
					ElseIf IsInteger(TrackPosition) Then		'no char at all (1-01, 1-02, 1-12)
						If CheckLeadingZero = True And Int(TrackPosition) < 10 Then
							tracksNum.Add "0" & Int(TrackPosition)
							WriteLog "Tracknumber=0" & Int(TrackPosition)
						Else
							tracksNum.Add TrackPosition
							WriteLog "Tracknumber=" & TrackPosition
						End If
					Else
						tracksNum.Add TrackPosition
					End If
				ElseIf Len(TrackPosition) = 1 Then	'1 Char after -  (1-1, 1-I, 1-2)
					If IsInteger(TrackPosition) Then
						If CheckLeadingZero = True And TrackPosition < 10 Then
							tracksNum.Add "0" & TrackPosition
							WriteLog "Tracknumber=0" & TrackPosition
						Else
							tracksNum.Add TrackPosition
							WriteLog "Tracknumber=" & TrackPosition
						End If
					Else
						tracksNum.Add TrackPosition
						WriteLog "Tracknumber=" & TrackPosition
					End If
				End If
			End If
			If UnselectedTracks(iTrackNum) <> "x" Then
				If IsInteger(TrackPosition) Then
					iAutoTrackNumber = TrackPosition + 1
				Else
					iAutoTrackNumber = iAutoTrackNumber + 1
				End If
			End If
		End If
		If Left(position,2) = "CD" Then
			If iAutoDiscFormat <> "CD" Then
				iAutoDiscFormat = "CD"
				iAutoTrackNumber = 1
				iAutoDiscNumber = 1
			End If
			If Mid(position,3,1) = "-" Then
				'Or Mid(position,3,1) = "." Then
				iAutoDiscNumber = 1
			Else
				iAutoDiscNumber = Mid(position,3,1)
			End If
		End If
		If Left(position,3) = "DVD" Then
			If iAutoDiscFormat <> "DVD" Then
				iAutoDiscFormat = "DVD"
				iAutoTrackNumber = 1
				iAutoDiscNumber = 1
			End If
			If Mid(position,4,1) = "-" Then
				'Or Mid(position,3,1) = "." Then
				iAutoDiscNumber = 1
			Else
				iAutoDiscNumber = Mid(position,4,1)
			End If
		End If
		If Left(position,3) = "VHS" Then
			If iAutoDiscFormat <> "VHS" Then
				iAutoDiscFormat = "VHS"
				iAutoTrackNumber = 1
				iAutoDiscNumber = 1
			End If
			If Mid(position,4,1) = "-" Then
				'Or Mid(position,3,1) = "." Then
				iAutoDiscNumber = 1
			Else
				iAutoDiscNumber = Mid(position,4,1)
			End If
		End If
		If Left(position,5) = "Video" Then
			If iAutoDiscFormat <> "Video" Then
				iAutoDiscFormat = "Video"
				iAutoTrackNumber = 1
				iAutoDiscNumber = 1
			End If
			If Mid(position,6,1) = "-" Then
				'Or Mid(position,3,1) = "." Then
				iAutoDiscNumber = 1
			Else
				iAutoDiscNumber = Mid(position,6,1)
			End If
		End If
		If Left(position,2) <> "CD" And Left(position,3) <> "DVD" And Left(position,3) <> "VHS" And Left(position,5) <> "Video" Then iAutoDiscNumber = Left(position,pos-1)
		tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
	Else ' Apply Track Numbering Schemes
		WriteLog "Track Numbering Schemes"
		If Not CheckSidesToDisc Or IsInteger(Left(position,1)) Then
			WriteLog "Standard track numbering"
			If CheckForceNumeric Then
				If UnselectedTracks(iTrackNum) <> "x" Then
					If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
						tracksNum.Add "0" & iAutoTrackNumber
						WriteLog "Tracknumber=0" & iAutoTrackNumber
					Else
						tracksNum.Add iAutoTrackNumber
						WriteLog "Tracknumber=" & iAutoTrackNumber
					End If
					iAutoTrackNumber = iAutoTrackNumber + 1
				Else
					tracksNum.Add ""
					WriteLog "Tracknumber=Nothing"
				End If
			Else
				If CheckLeadingZero = True And IsNumeric(position) Then
					If position < 10 And position > 0 Then
						tracksNum.Add "0" & position
						WriteLog "Tracknumber=0" & position
					Else
						tracksNum.Add position
						WriteLog "Tracknumber=" & position
					End If
				Else
					tracksNum.Add position
					WriteLog "Tracknumber=" & position
				End If
				If UnselectedTracks(iTrackNum) <> "x" Then
					If IsNumeric(position) Then
						iAutoTrackNumber = position + 1
					Else
						iAutoTrackNumber = iAutoTrackNumber + 1
					End If
				End If
			End If
			If CheckForceDisc Then
				tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
			Else
				tracksCD.Add ""
			End If
		Else
			If Len(position) = 1 Then ' Only side is specified
				WriteLog "Only side is specified"
				If CheckLeadingZero = True Then
					tracksNum.Add "01"
					WriteLog "Tracknumber=01"
				Else
					tracksNum.Add "1"
					WriteLog "Tracknumber=1"
				End If
				If 	LastDisc <> position Then
					If 	LastDisc <> "" Then
						iAutoDiscNumber = iAutoDiscNumber + 1
					End If
					LastDisc = position
				End If
				If CheckForceNumeric Then
					tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
				Else
					tracksCD.Add LeadingZeroDisc(position)
				End If
			ElseIf Len(position) = 2 Then
				If IsInteger(Mid(position,2,1)) And Not IsInteger(Mid(position,1,1)) Then
					WriteLog "First is Side Second is Track"
					' First is Side Second is Track
					' WriteLog iAutoDiscNumber
					' WriteLog iAutoTrackNumber
					If 	LastDisc <> Left(position,1) Then
						If LastDisc <> "" And (LastDisc = "B" Or LastDisc = "D" Or LastDisc = "F" Or LastDisc = "H" Or LastDisc = "J") Then
							If IsInteger(iAutoDiscNumber) Then
								iAutoDiscNumber = iAutoDiscNumber + 1
							End If
							iAutoTrackNumber = 1
						End If
						LastDisc = Left(position,1)
					End If
					If UnselectedTracks(iTrackNum) <> "x" Then
						' If CheckLeadingZero = True And Mid(position,2) < 10 Then
						If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
							' tracksNum.Add "0" & Mid(position,2)
							tracksNum.Add "0" & iAutoTrackNumber
							WriteLog "Tracknumber=0" & iAutoTrackNumber
							iAutoTrackNumber = iAutoTrackNumber + 1
						Else
							' tracksNum.Add Mid(position,2)
							tracksNum.Add iAutoTrackNumber
							WriteLog "Tracknumber=" & iAutoTrackNumber
							iAutoTrackNumber = iAutoTrackNumber + 1
						End If
					Else
						tracksNum.Add ""
						WriteLog "Tracknumber=Nothing"
					End If
					If CheckForceNumeric Then
						tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
					Else
						tracksCD.Add LeadingZeroDisc(Left(position,1))
					End If
				Else ' Two byte side
					WriteLog "Two byte side"
					tracksNum.Add "1"
					WriteLog "Tracknumber=1"
					If 	LastDisc <>  position Then
						If 	LastDisc <> "" Then
							iAutoDiscNumber = iAutoDiscNumber + 1
						End If
						LastDisc = position
					End If
					If CheckForceNumeric Then
						tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
					Else
						tracksCD.Add LeadingZeroDisc(position)
					End If
				End If
			Else ' More than 2 bytes
				If IsInteger(Mid(position,2)) And CheckNoDisc = False Then
					WriteLog "First is Side Latter is Track"
				'First is Side Latter is Track
					If UnselectedTracks(iTrackNum) <> "x" Then
						tracksNum.Add Mid(position,2)
						WriteLog "Tracknumber=" & Mid(position,2)
					Else
						tracksNum.Add ""
						WriteLog "Tracknumber=Nothing"
					End If
					If 	LastDisc <> Left(position,1) Then
						If LastDisc <> "" Then
							iAutoDiscNumber = iAutoDiscNumber + 1
						End If
						LastDisc = Left(position,1)
					End If
					If CheckForceNumeric Then
						tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
					Else
						tracksCD.Add LeadingZeroDisc(Left(position,1))
					End If
				ElseIf IsInteger(Mid(position,3)) And CheckNoDisc = False Then
					WriteLog "Two Byte Side, Latter is Track"
					' Two Byte Side, Latter is Track
					If UnselectedTracks(iTrackNum) <> "x" Then
						tracksNum.Add Mid(position,3)
						WriteLog "Tracknumber=" & Mid(position,3)
					Else
						tracksNum.Add ""
						WriteLog "Tracknumber=Nothing"
					End If
					If 	LastDisc <>  Left(position,2) Then
						If 	LastDisc <> "" Then
							iAutoDiscNumber = iAutoDiscNumber + 1
						End If
						LastDisc = Left(position,2)
					End If
					If CheckForceNumeric Then
						tracksCD.Add LeadingZeroDisc(iAutoDiscNumber)
					Else
						tracksCD.Add LeadingZeroDisc(Left(position,2))
					End If
				Else ' More than two non numeric bytes!
					WriteLog "More than two non numeric bytes!"
					If CheckNoDisc = False Then
						If UnselectedTracks(iTrackNum) <> "x" Then
							tracksNum.Add position
							WriteLog "Tracknumber=" & position
							tracksCD.Add ""
						Else
							tracksNum.Add ""
							WriteLog "Tracknumber=Nothing"
							tracksCD.Add ""
						End If
					Else
						If CheckForceNumeric Then
							If UnselectedTracks(iTrackNum) <> "x" Then
								If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
									tracksNum.Add "0" & iAutoTrackNumber
									WriteLog "Tracknumber=0" & iAutoTrackNumber
								Else
									tracksNum.Add iAutoTrackNumber
									WriteLog "Tracknumber=" & iAutoTrackNumber
								End If
								iAutoTrackNumber = iAutoTrackNumber + 1
							Else
								tracksNum.Add ""
								WriteLog "Tracknumber=Nothing"
							End If
						Else
							If UnselectedTracks(iTrackNum) <> "x" Then
								If IsInteger(position) Then
									tracksNum.Add iAutoTrackNumber
									WriteLog "Tracknumber=" & iAutoTrackNumber
									iAutoTrackNumber = position + 1
								Else
									tracksNum.Add iAutoTrackNumber
									WriteLog "Tracknumber=" & iAutoTrackNumber
									iAutoTrackNumber = iAutoTrackNumber + 1
								End If
							End If
						End If
						tracksCD.Add ""
					End If
				End If
			End If
		End If
	End If
	' WriteLog "iAutoTrackNumber=" & iAutoTrackNumber
	' WriteLog "iTrackNum=" & iTrackNum
	' WriteLog "tracksnum.item(iTrackNum)=" & tracksnum.item(iTrackNum)

End Function


Function getinvolvedRole(involvedArtist, involvedRole, byRef artistList, byRef TrackFeaturing, byRef Involved_R_T, byRef TrackComposers, byRef TrackConductors, byRef TrackProducers, byRef TrackLyricists)

	WriteLog "Start getinvolvedRole"
	Dim tmp, tmp2
	If LookForFeaturing(involvedRole) Then
		If InStr(artistList, involvedArtist) = 0 Then
			If TrackFeaturing = "" Then
				If CheckFeaturingName Then
					TrackFeaturing = TxtFeaturingName & " " & involvedArtist
				Else
					TrackFeaturing = involvedRole & " " & involvedArtist
				End If
			Else
				If InStr(TrackFeaturing, involvedArtist) = 0 Then
					TrackFeaturing = TrackFeaturing & ArtistSeparator & involvedArtist
				End If
			End If
		End If
		WriteLog "TrackFeaturing=" & TrackFeaturing
	Else
		Do
			tmp = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
			If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
				TrackLyricists = tmp
				WriteLog "TrackLyricists=" & TrackLyricists
				Exit Do
			ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
				WriteLog "ALREADY_INSIDE_ROLE"
				Exit Do
			End If
			tmp = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
			If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
				TrackConductors = tmp
				WriteLog "TrackConductors=" & TrackConductors
				Exit Do
			ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
				WriteLog "ALREADY_INSIDE_ROLE"
				Exit Do
			End If
			tmp = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
			If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
				TrackProducers = tmp
				WriteLog "TrackProducers=" & TrackProducers
				Exit Do
			ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
				WriteLog "ALREADY_INSIDE_ROLE"
				Exit Do
			End If
			tmp = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
			If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
				TrackComposers = tmp
				WriteLog "TrackComposers=" & TrackComposers
				Exit Do
			ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
				WriteLog "ALREADY_INSIDE_ROLE"
				Exit Do
			End If
			tmp2 = search_involved(Involved_R_T, involvedRole)
			If tmp2 = -2 Then
				WriteLog "Ignore Role: '" & currentRole & "' (Unwanted tag)"
			ElseIf tmp2 = -1 Then
				ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
				Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & involvedArtist
				WriteLog "New Role: " & involvedRole & ": " & involvedArtist
			Else
				If InStr(Involved_R_T(tmp2), involvedArtist) = 0 Then
					Involved_R_T(tmp2) = Involved_R_T(tmp2) & ArtistSeparator & involvedArtist
					WriteLog "Role updated: " & Involved_R_T(tmp2)
				Else
					WriteLog "artist already inside role"
				End If
			End If
			Exit Do
		Loop While True
	End If
	WriteLog "Stop getinvolvedRole"

End Function

Function CheckSpecialRole(Role)

	Dim tmp, tmp2, tmp3
	WriteLog "Start CheckSpecialRole"
	ReDim SingleRole(0)
	Do While 1 = 1
		tmp = InStr(Role, ",")
		tmp2 = InStr(Role, "[")
		tmp3 = InStr(Role, "]")
		If tmp = 0 Then
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			SingleRole(UBound(SingleRole)) = Role
			Exit Do
		End If
		If tmp < tmp2 Then
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			SingleRole(UBound(SingleRole)) = Left(Role, tmp-1)
			Role = LTrim(Mid(Role, tmp+1))
		End If
		If tmp > tmp2 And tmp > tmp3 Then
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			SingleRole(UBound(SingleRole)) = Left(Role, tmp-1)
			Role = LTrim(Mid(Role, tmp+1))
		End If
		If tmp > tmp2 And tmp < tmp3 Then
			tmp = InStr(tmp3, Role, ",", 0)
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			If tmp <> 0 Then
				SingleRole(UBound(SingleRole)) = Left(Role, tmp-1)
				Role = LTrim(Mid(Role, tmp+1))
			Else
				SingleRole(UBound(SingleRole)) = Role
				Exit Do
			End If
		End If
	Loop
	WriteLog "End CheckSpecialRole"
	CheckSpecialRole = SingleRole

End Function


Function FindSubTrackSplit(position)

	Dim tmp
	For tmp = 2 To Len(position)
		If Not IsNumeric(Mid(position, tmp, 1)) Then
			FindSubTrackSplit = Left(position, tmp-1)
			Exit For
		End If
	Next

End Function


Function FindArtist(ArtistList1, ArtistList2)

	Dim tmpArtist, i
	ReDim newArtistList1(0)
	tmpArtist = Split(ArtistList1, ArtistSeparator)
	For i = 0 To UBound(tmpArtist)
		If InStr(ArtistList2,tmpArtist(i)) = 0 Then
			ReDim Preserve newArtistList1(UBound(newArtistList1)+1)
			newArtistList1(UBound(newArtistList1)) = tmpArtist(i)
		End If
	Next
	For i = 0 To UBound(newArtistList1)
		If FindArtist = "" Then
			FindArtist = newArtistList1(i)
		Else
			FindArtist = FindArtist & ArtistSeparator & newArtistList1(i)
		End If
	Next

End Function


Sub Track_from_to (currentTrack, currentArtist, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition)

	WriteLog "Start Track_from_to"
	Dim tmp, tmp4, cnt, ret
	Dim StartTrack, EndTrack
	WriteLog "currentTrack=" & currentTrack
	If InStr(currentTrack, "  ") <> 0 Then
		WriteLog "Warning: More than one space between the track positions !!!"
		Do
			If InStr(currentTrack, "  ") = 0 Then Exit Do
			currentTrack = Replace(currentTrack, "  ", " ")
		Loop While True
	End If
	tmp = Split(currentTrack, " to ")

	StartTrack = Trim(tmp(0))
	EndTrack = trim(tmp(1))
	WriteLog "StartTrack=" & StartTrack
	WriteLog "EndTrack=" & EndTrack

	tmp4 = False
	
	For cnt = 0 To UBound(Title_Position)-1
		If Title_Position(cnt) = StartTrack Then
			tmp4 = True
		End If
		If Title_Position(cnt) = EndTrack Then
			ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
			ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
			ReDim Preserve TrackPos(UBound(TrackPos)+1)
			TrackArtist2(UBound(TrackArtist2)) = currentArtist
			TrackRoles(UBound(TrackRoles)) = involvedRole
			TrackPos(UBound(TrackPos)) = Title_Position(cnt)
			WriteLog "Pos: " & Title_Position(cnt) & "  " & currentArtist & " - " & involvedRole
			Exit For
		End If
		If tmp4 = True Then
			ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
			ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
			ReDim Preserve TrackPos(UBound(TrackPos)+1)
			TrackArtist2(UBound(TrackArtist2)) = currentArtist
			TrackRoles(UBound(TrackRoles)) = involvedRole
			TrackPos(UBound(TrackPos)) = Title_Position(cnt)
			WriteLog "Pos: " & Title_Position(cnt) & "  " & currentArtist & " - " & involvedRole
		End If
	Next

	WriteLog "Stop Track_from_to"

End Sub

Function Remove_CD(TrackPos)

	If Left(TrackPos, 2) = "CD" Then
		TrackPos = Mid(TrackPos, 3)
	End If
	If Left(TrackPos, 3) = "DVD" Then
		TrackPos = Mid(TrackPos, 4)
	End If
	If UCase(Left(TrackPos, 6)) = "VIDEO " Then
		TrackPos = Mid(TrackPos, 7)
	End If
	If UCase(Left(TrackPos, 5)) = "VIDEO" Then
		TrackPos = Mid(TrackPos, 6)
	End If
	If UCase(Left(TrackPos, 3)) = "VHS" Then
		TrackPos = Mid(TrackPos, 4)
	End If
	If UCase(Left(TrackPos, 5)) = "VIDEO" Then
		TrackPos = Mid(TrackPos, 6)
	End If
	Return TrackPos

End Function


Sub Add_Track_Role(currentTrack, currentArtist, involvedRole, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition)

	Dim tmp4
	If InStr(currentTrack, "-") <> 0 Then
		tmp4 = Split(currentTrack, "-")
		If IsNumeric(tmp4(1)) = True Then
			If LeadingZeroTrackPosition = True And tmp4(1) < 10 And Left(tmp4(1), 1) <> "0" Then
				currentTrack = tmp4(0) & "-0" & tmp4(1)
			End If
		End If
	End If
	If InStr(currentTrack, ".") <> 0 Then
		tmp4 = Split(currentTrack, ".")
		If IsNumeric(tmp4(1)) = True Then
			If LeadingZeroTrackPosition = True And tmp4(1) < 10 And Left(tmp4(1), 1) <> "0" Then
				currentTrack = tmp4(0) & ".0" & tmp4(1)
			End If
		End If
	End If
	WriteLog "currentTrack=" & currentTrack
	currentTrack = exchange_roman_numbers(currentTrack)
	ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
	ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
	ReDim Preserve TrackPos(UBound(TrackPos)+1)
	TrackArtist2(UBound(TrackArtist2)) = currentArtist
	TrackRoles(UBound(TrackRoles)) = involvedRole
	TrackPos(UBound(TrackPos)) = currentTrack

End Sub


' ShowResult is called every time the search result is changed from the drop
' down at the top of the window
Sub ShowResult(ResultID)

	Dim ReleaseID, searchURL, oXMLHTTP, ResponseHTML, Title, SelectedTracks, searchURL_F, searchURL_L, i, templateHTMLDoc
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
	If Not (button is Nothing) Then
		button.style.backgroundColor = "#ED1C24"
		templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML="Please wait for collection data..."
	End If
	For i = 0 to 999
		UnselectedTracks(i) = ""
		UnselectedTrackNames(i) = ""
	Next
	SDB.Tools.WebSearch.ClearTracksData
	NewResult = True
	WebBrowser.SetHTMLDocument ""                 ' Deletes visible search result
	If ResultsReleaseID.Item(ResultID) = "" Then Exit Sub
	Dim json
	Set json = New VbsJson
	CurrentResultId = ResultID

	If QueryPage = "MetalArchives" Then
		searchURL = ResultsReleaseID.Item(ResultID)
		Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
		Call oXMLHTTP.open("GET", searchURL, False)
		Call oXMLHTTP.send()
		If oXMLHTTP.Status = 200 Then
			CurrentRelease = oXMLHTTP.responseText
			ReloadResults
		End If
	End If


	' http://musicbrainz.org/ws/2/release/e3b950f4-cc3b-3f84-b80f-2c254ffd956f?inc=recordings+recording-level-rels+work-rels+work-level-rels+artist-rels+artist-credits+media+release-group-rels+release-groups+labels&fmt=json
	If QueryPage = "MusicBrainz" Then

		WriteLog "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-"
		WriteLog "Start ShowResult MusicBrainz"
		ReleaseID = ResultsReleaseID.Item(ResultID)
		WriteLog "ReleaseID=" & ReleaseID
		If InStr(Results.Item(ResultID), "No Release found") = 0 Then
			searchURL = "http://musicbrainz.org/ws/2/release/" & ReleaseID & "?inc=recordings+recording-level-rels+work-rels+work-level-rels+artist-rels+artist-credits+media+release-group-rels+release-groups+labels&fmt=json"
			WriteLog "searchURL=" & searchURL

			Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")   
			oXMLHTTP.Open "GET", searchURL, False
			oXMLHTTP.setRequestHeader "Content-Type", "application/json"
			oXMLHTTP.setRequestHeader "User-Agent",UserAgent
			oXMLHTTP.send ()

			If oXMLHTTP.Status = 200 Then
				WriteLog "responseText=" & oXMLHTTP.responseText
				Set CurrentRelease = json.Decode(oXMLHTTP.responseText)
				CurrentReleaseId = ReleaseID
				ReloadResults
			ElseIf oXMLHTTP.Status = 503 Then
				ErrorMessage = "Status:" & oXMLHTTP.Status & " - The number of requests exceeds the limit. Please try again later"
				WriteLog "Status=" & oXMLHTTP.Status
				FormatErrorMessage ErrorMessage
			Else
				ErrorMessage = "Status:" & oXMLHTTP.Status & " - Please try again later"
				WriteLog "Status=" & oXMLHTTP.Status
				FormatErrorMessage ErrorMessage
			End If
		Else
			ErrorMessage = "Search returned no results / No Release found"
			FormatErrorMessage ErrorMessage
		End If
	End If



	If QueryPage = "Discogs" Then

		WriteLog "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-"
		WriteLog "Start ShowResult Discogs"
		ReleaseID = ResultsReleaseID.Item(ResultID)
		WriteLog "ReleaseID=" & ReleaseID

		Set GenresSelect = SDB.NewStringList

		If InStr(Results.Item(ResultID), "search returned no results") = 0 And InStr(Results.Item(ResultID), "No Release found") = 0 Then
			If InStr(Results.Item(ResultID), " * ") <> 0 Or Right(Results.Item(ResultID), 8) = "(Master)" Then  'Master-Release
				searchURL_F = "https://api.discogs.com/masters/"
				WriteLog "Show Master-Release"
			Else
				searchURL_F = "https://api.discogs.com/releases/"
			End If
			searchURL = ReleaseID
			searchURL_L = ""

			Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
			oXMLHTTP.open "POST", "https://www.germanc64.de/mm/oauth/check_new.php", False
			oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			oXMLHTTP.setRequestHeader "User-Agent",UserAgent
			WriteLog "Sending Post at=" & AccessToken & "&ats=" & AccessTokenSecret & "&searchURL=" & searchURL & "&searchURL_F=" & searchURL_F & "&searchURL_L=" & searchURL_L & "&version=" & VersionStr
			oXMLHTTP.send("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&searchURL=" & searchURL & "&searchURL_F=" & searchURL_F & "&searchURL_L=" & searchURL_L & "&version=" & VersionStr)

			If oXMLHTTP.Status = 200 Then
				WriteLog "responseText=" & oXMLHTTP.responseText
				If InStr(oXMLHTTP.responseText, "OAuth client error") <> 0 Then
					ErrorMessage = oXMLHTTP.responseText
					FormatErrorMessage ErrorMessage
				Else
					Set CurrentRelease = json.Decode(oXMLHTTP.responseText)
					CurrentReleaseId = ReleaseID
					Set oXMLHTTP = Nothing
					ReloadResults
				End If
			End If
		Else
			ErrorMessage = "Search returned no results / No Release found"
			FormatErrorMessage ErrorMessage
		End If
	End If

End Sub


' This does the final clean up, so that our script doesn't leave any unwanted traces
Sub FinishSearch(Panel)

	If isObject(WebBrowser) Then
		WebBrowser.Common.DestroyControl      ' Destroy the external control
		Set WebBrowser = Nothing              ' Release global variable
		SDB.Objects("WebBrowser") = Nothing
	End If

	Set ini = Nothing
	If isObject(ResultsReleaseID) Then
		Set ResultsReleaseID = Nothing
	End If
	Script.UnregisterAllEvents

End Sub


Sub SaveMoreImagesSub()

	Dim ret, res, RndFileName, i, itm, path, j, k, ImageSelected, SongList
	If ImageList.Count > 0 Then
		ImageSelected = False
		For i = 0 to ImageList.Count - 1
			If SaveImage.Item(i) = 1 Then ImageSelected = True
		Next
		If ImageSelected = True Then
			For i = 0 to ImageList.Count - 1
				res = 0
				If SaveImage.Item(i) = 1 Then
					Set itm = SelectedSongsGlobal.item(0)
					path = Mid(itm.Path,1,InStrRev(itm.Path,"\")-1)
					If CoverStorage = 1 Or CoverStorage = 3 Then
						If SDB.Tools.FileSystem.FileExists(path & "\" & FileNameList.Item(i)) = True Then
							res = SDB.MessageBox("The file " & FileNameList.Item(i) & " already exist. Overwrite it ?", mtConfirmation, Array(mbYes, mbNo))
							If res = 6 Then
								SDB.Tools.FileSystem.DeleteFile(path & "\" & FileNameList.Item(i))
								ret = SDB.Tools.FileSystem.MoveFile(ImageLocal.Item(i), path & "\" & FileNameList.Item(i))
								If ret = False Then
									WriteLog "ERROR:Image could not moved to : " & path & "\" & FileNameList.Item(i)
									SDB.MessageBox "ERROR:Image could not moved !", mtError, Array(mbOk)
								End If
							End If
						Else
							ret = SDB.Tools.FileSystem.MoveFile(ImageLocal.Item(i), path & "\" & FileNameList.Item(i))
							If ret = False Then
								WriteLog "ERROR:Image could not moved to : " & path & "\" & FileNameList.Item(i)
								SDB.MessageBox "ERROR:Image could not moved !", mtError, Array(mbOk)
							End If
						End If
					End If

					If res <> 7 Then 'don't overwrite file
						Set SongList = SDB.SelectedSongList
						For j = 0 To SongList.Count - 1
							Set itm = SongList.item(j)
							Dim pics : Set pics = itm.AlbumArt
							If pics Is Nothing Then
								Exit Sub
							End If
							Dim img
							Set img = pics.AddNew
							img.Description = ""

							If CoverStorage = 1 Or CoverStorage = 3 Then
								img.PicturePath = path & "\" & FileNameList.Item(i)
								img.ItemStorage = 1
							Else
								img.PicturePath = ImageLocal.Item(i)
								img.ItemStorage = 0
							End If
							For k = 0 to ImageTypeList.Count - 1
								If SaveImageType.Item(i) = ImageTypeList.Item(k) Then
									If k = 0 Then k = -2
									If k > 14 Then k = k + 1
									img.ItemType = k + 2
									pics.UpdateDB
									Exit For
								End If
							Next

							If CoverStorage = 3 Then
								Set pics = itm.AlbumArt
								Set img = pics.AddNew
								img.Description = ""
								img.PicturePath = path & "\" & FileNameList.Item(i)
								img.ItemStorage = 0
								For k = 0 to ImageTypeList.Count - 1
									If SaveImageType.Item(i) = ImageTypeList.Item(k) Then
										If k = 0 Then k = -2
										If k > 14 Then k = k + 1
										img.ItemType = k + 2
										Exit For
									End If
								Next
							End If
						Next
					End If
					If CoverStorage = 0 Then
						SDB.Tools.FileSystem.DeleteFile(ImageLocal.Item(i))
					End If
				End If
			Next
		End If
	End If

End Sub


Function GetHeader()

	Dim templateHTML, i
	templateHTML = "<HTML>"
	templateHTML = templateHTML &  "<HEAD>"
	templateHTML = templateHTML &  "<style type=""text/css"" media=""screen"">"
	templateHTML = templateHTML &  ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	templateHTML = templateHTML &  "option.tabletext{background-color:#3E7CBB;}"

	templateHTML = templateHTML &  "</style>"
	templateHTML = templateHTML &  "</HEAD>"
	templateHTML = templateHTML &  "<body bgcolor=""#FFFFFF"">"
	templateHTML = templateHTML &  "<table border=0 width=100% cellspacing=0 cellpadding=1 class=tabletext>"
	templateHTML = templateHTML &  "<tr>"
	If QueryPage = "Discogs" Then
		templateHTML = templateHTML &  "<td align=left><a href=""https://www.discogs.com"" target=""_blank""><img src=""https://www.germanc64.de/mm/i-love-discogs.png"" width=""100"" height=""60"" border=""0"" alt=""Discogs Homepage""></a><b>" & VersionStr & "</b></td>"
	End If
	If QueryPage = "MusicBrainz" Then
		templateHTML = templateHTML &  "<td align=left><a href=""https://www.musicbrainz.org"" target=""_blank""><img src=""https://wiki.musicbrainz.org/-/images/e/e4/MusicBrainz_Logo_White.png"" width=""100"" height=""70"" border=""0""/ alt=""MusicBrainz Homepage""></a><b>" & VersionStr & "</b></td>"
	End If
	If QueryPage = "MetalArchives" Then
		templateHTML = templateHTML &  "<td align=left><a href=""https://www.metal-archives.com"" target=""_blank""><img src=""https://www.metal-archives.com/css/default/images/smallerlogo.jpg"" width=""200"" height=""60"" border=""0""/ alt=""MusicBrainz Homepage""></a><b>" & VersionStr & "</b></td>"
	End If
	
	templateHTML = templateHTML &  "<td colspan=3 align=right valign=top>"

	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=2 class=tabletext>"
	REM templateHTML = templateHTML &  "<tr><td colspan=5></td><td><b>Filter Results: </b></td><td colspan=3> </td></tr>"
	If QueryPage = "MetalArchives" Then
		templateHTML = templateHTML &  "<tr><td colspan=4></td><td><button type=button class=tabletext id=""showadvancedsearch"">Manual Search</button></td><td colspan=2><b>Filter Results: Filter aren't working with MetalArchives!</b></td><td colspan=2> </td>"
	Else
		templateHTML = templateHTML &  "<tr><td colspan=2></td><td><button type=button class=tabletext style=""background-color:red;color:white"" id=""PleaseWaitButton"">Please wait for collection data..</button></td><td align=center><img src=""" & SDB.ScriptsPath & "question-mark.png"" alt=""Click here for Help !"" id=""picture""></td><td><button type=button class=tabletext id=""showadvancedsearch"">Manual Search</button></td><td><b>Filter Results: </b></td><td colspan=3> </td>"
	End If
	REM <td align=right><img src=""" & SDB.ScriptsPath & "question-mark.png"" alt=""Click here for Help !"" id=""picture""></td></tr>"
	REM templateHTML = templateHTML &  "<tr><td colspan=3></td><td colspan=2>Search for:<input type=radio id=""searchartist"" name=""SearchFor"" title=""Enter search string in upper dropdown-field and choose what to search for"" value=""Artist"">Artist<input type=radio id=""searchalbum"" name=""SearchFor"" title=""Enter search string in upper dropdown-field and choose what to search for"" value=""Album"">Album<input type=radio id=""searchrelease"" name=""SearchFor"" title=""Enter search string in upper drop-down field and choose what to search for"" value=""Release"">Release</td><td><b>Filter Results: </b></td><td colspan=3> </td></tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><b></b></td>"
	templateHTML = templateHTML &  "<td><b>                                    </b></td>"
	templateHTML = templateHTML &  "<td><b>                                    </b></td>"
	templateHTML = templateHTML &  "<td><b>Load:</b></td>"
	If QueryPage <> "MetalArchives" Then
		templateHTML = templateHTML &  "<td><b>Quick Search:</b></td>"
	Else
		templateHTML = templateHTML &  "<td></td>"
	End If
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showmediatypefilter"">Set Type Filter</button></td>"
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showmediaformatfilter"">Set Format Filter</button></td>"
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showyearfilter"">Set Year Filter</button></td>"
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showcountryfilter"">Set Country Filter</button></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td>"
	
	templateHTML = templateHTML &  "</td>"
	templateHTML = templateHTML &  "<td>                                    </td>"
	templateHTML = templateHTML &  "<td>                                    </td>"
	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""load"" class=tabletext title=""Search Result=Search with Artist and Album Title" & vbCrLf & "Master Release=Show master of selected release" & vbCrLf & "Versions of Master=Show all releases of selected master" & vbCrLf & "Releases of Artist=Show all releases of selected artist" & vbCrLf & "Releases of Label=Show all releases of selected label"">"

	For i = 0 To LoadList.Count - 1
		If LoadList.Item(i) <> CurrentLoadType Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(LoadList.Item(i)) & """>" & LoadList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(LoadList.Item(i)) & """ selected>" & LoadList.Item(i) & "</option>"
		End If
	Next
	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"
	'Alternative Searches Begin
	If QueryPage <> "MetalArchives" Then
		templateHTML = templateHTML &  "<td>"
		templateHTML = templateHTML &  "<select id=""alternative"" class=tabletext>"
		For i = 0 To AlternativeList.Count - 1
			If AlternativeList.Item(i) <> NewSearchTerm Then
				templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(AlternativeList.Item(i)) & """>" & AlternativeList.Item(i) & "</option>"
			Else
				templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(AlternativeList.Item(i)) & """ selected>" & AlternativeList.Item(i) & "</option>"
			End If
		Next
		templateHTML = templateHTML &  "</select>"
		templateHTML = templateHTML &  "</td>"
	Else
		templateHTML = templateHTML &  "<td>"
		templateHTML = templateHTML &  "</td>"
	End If
	'Alternative Searches End
	'Filters Begin
	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filtermediatype"" class=tabletext>"

	If FilterMediaType = "None" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaType Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaType Filter"">Use MediaType Filter</option>"
	ElseIf FilterMediaType = "Use MediaType Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaType Filter"">Use MediaType Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No MediaType Filter</option>"
	End If
	If FilterMediaType <> "None" And FilterMediaType <> "Use MediaType Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaType Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use MediaType Filter"">Use MediaType Filter</option>"
	End If
	For i = 1 To MediaTypeList.Count - 1
		If FilterMediaType <> MediaTypeList.Item(i) Or FilterMediaType = "None" Or FilterMediaType = "Use MediaType Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaTypeList.Item(i)) & """>" & MediaTypeList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaTypeList.Item(i)) & """ selected>" & MediaTypeList.Item(i) & "</option>"
		End If
	Next
	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"

	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filtermediaformat"" class=tabletext>"

	If FilterMediaFormat = "None" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaFormat Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
	ElseIf FilterMediaFormat = "Use MediaFormat Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No MediaFormat Filter</option>"
	End If
	If FilterMediaFormat <> "None" And FilterMediaFormat <> "Use MediaFormat Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaFormat Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
	End If
	For i = 1 To MediaFormatList.Count - 1
		If FilterMediaFormat <> MediaFormatList.Item(i) Or FilterMediaFormat = "None" Or FilterMediaFormat = "Use MediaFormat Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaFormatList.Item(i)) & """>" & MediaFormatList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaFormatList.Item(i)) & """ selected>" & MediaFormatList.Item(i) & "</option>"
		End If
	Next
	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"

	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filteryear"" class=tabletext>"

	If FilterYear = "None" Then
		templateHTML = templateHTML &  "<option value=""None"" selected>No Year Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Year Filter"">Use Year Filter</option>"
	ElseIf FilterYear = "Use Year Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Year Filter"" selected>Use Year Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No Year Filter</option>"
	End If
	If FilterYear <> "None" And FilterYear <> "Use Year Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No Year Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use Year Filter"">Use Year Filter</option>"
	End If
	For i = 1 To YearList.Count - 1
		If FilterYear <> YearList.Item(i) Or FilterYear = "None" Or FilterYear = "Use Year Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(YearList.Item(i)) & """>" & YearList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(YearList.Item(i)) & """ selected>" & YearList.Item(i) & "</option>"
		End If
	Next

	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"

	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filtercountry"" class=tabletext>"

	If FilterCountry = "None" Then
		templateHTML = templateHTML &  "<option value=""None"">No Country Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Country Filter"">Use Country Filter</option>"
	ElseIf FilterCountry = "Use Country Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Country Filter"">Use Country Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No Country Filter</option>"
	End If
	If FilterCountry <> "None" And FilterCountry <> "Use Country Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No Country Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use Country Filter"">Use Country Filter</option>"
	End If
	For i = 1 To CountryList.Count - 1
		If FilterCountry <> CountryList.Item(i) Or FilterCountry = "None" Or FilterCountry = "Use Country Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(CountryList.Item(i)) & """>" & CountryList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(CountryList.Item(i)) & """ selected=""selected"">" & CountryList.Item(i) & "</option>"
		End If
	Next

	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"
	'Filters End
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	templateHTML = templateHTML &  "</tr>"

	GetHeader = templateHTML

End Function


Function GetFooter()

	Dim templateHTML
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</body>"
	templateHTML = templateHTML &  "</HTML>"

	GetFooter = templateHTML

End Function



' We use this procedure to reformat results as soon as they are downloaded
Sub FormatSearchResultsViewer(Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, GenresList, theLabels, theCountry, AlbumArtThumbNail, releaseID, Catalog, Lyricists, Composers, Conductors, Producers, InvolvedArtists, theFormat, theMaster, comment, DiscogsTracksNum, DataQuality, Grouping)

	Dim templateHTML, checkBox, radio, text, listBox, submitButton, tmp
	Dim SelectedTracksCount, UnSelectedTracksCount
	Dim SubTrackFlag
	Dim i, theGenres
	templateHTML = ""
	templateHTML = templateHTML &  GetHeader()

	' Titles Begin
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Album Art:</b></td>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Release Information:</b></td>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Tracklisting:</b></td>"
	templateHTML = templateHTML &  "</tr>"
	' Titles End
	templateHTML = templateHTML &  "<tr>"
	' Release Cover Begin
	templateHTML = templateHTML &  "<td align=left valign=top>"
	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	If AlbumArtThumbNail <> "" Then
		If QueryPage = "Discogs" Then
			templateHTML = templateHTML &  "<tr><td colspan=2><a href=""https://www.discogs.com/viewimages?release=<!RELEASEID!>"" target=""_blank""><img src=""<!COVER!>"" border=""0""/></a></td></tr>"
		ElseIf QueryPage = "MusicBrainz" Then
			templateHTML = templateHTML &  "<tr><td colspan=2><img src=""<!COVER!>"" alt="""" border=""0""></a></td></tr>"
		ElseIf QueryPage = "MetalArchives" Then
			templateHTML = templateHTML &  "<tr><td colspan=2><img src=""<!COVER!>"" alt="""" border=""0"" width=""200"" height=""200""></a></td></tr>"
		End If
	Else
		templateHTML = templateHTML &  "<tr><td colspan=2><table width=150 height=150 border=1><tr><td><center>No Image<br>Available</center></td></tr></table></td></tr>"
	End If
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""checkcover"" >Save Cover</td></tr>"
	If ImagesCount > 1 Then
		templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""moreimages"" title=""Use this button if you like to store more images (leaflet page, media label)"">More Images</button></td></tr>"
	End If
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><br></td></tr>"
	' Release Cover End

	' Options Begin
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""saveoptions"" title=""Use this button to save the options you set"" >Save Options</button></td></tr>"
	
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""optionsform"" title=""Use this button to set tagging options"" >Tagging Options</button></td></tr>"
	
	templateHTML = templateHTML &  "<tr><td align=center colspan=2><b>Options:</b></td></tr>"
	
	REM templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""usercollection"" title=""If checked the tagged album will be added to the User Collection at Discogs"" >Add album to User Collection</td></tr>"
	If QueryPage = "Discogs" Then
		templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""usercollection"" title=""Use this button to add the selected album to your collection at Discogs"">Add to User Collection</button></td></tr>"
	End If
	If QueryPage <> "MetalArchives" Then
		templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""lyricist"" title=""If a lyricist was named in the release, it will be written into the lyricist tag"" >Save Lyricist</td></tr>"
		templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""composer"" title=""If a composer was named in the release, it will be written into the composer tag"" >Save Composer</td></tr>"
		templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""conductor"" title=""If a conductor was named in the release, it will be written into the conductor tag"" >Save Conductor</td></tr>"
		templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""producer"" title=""If a producer was named in the release, it will be written into the producer tag"" >Save Producer</td></tr>"
		templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""involved"" title=""If other involved people were named in the release, their will be written into the involved people tag"" >Save Involved People</td></tr>"
	End If
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""grouping"" title=""If tracks are grouped together with heading tracks, the name of it will be add to the grouping tag"" >Save grouping info</td></tr>"
	If trackRefresh = True Then
		templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""refresh"" title=""Use this button to start/stop refreshing the track list after de-/select a track. Use it to de-select some tracks without waiting for refresh"">Stop track refresh</button></td></tr>"
	Else
		templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""refresh"" title=""Use this button to start/stop refreshing the track list after de-/select a track. Use it to de-select some tracks without waiting for refresh"">Start track refresh</button></td></tr>"
	End If
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""deselectall"" title=""Use this button to deselect all"">Deselect all</button></td></tr>"

	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><br></td></tr>"

	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Options End
	
	' Release Information Begin
	templateHTML = templateHTML &  "<td align=left valign=top>"
	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	
	iMaxTracks = Tracks.Count
	If TracksCD.Count < iMaxTracks Then
		iMaxTracks = TracksCD.Count
	End If
	
	'Check for different Track number
	SelectedTracksCount = 0
	UnSelectedTracksCount = 0
	SubTrackFlag = False
	For i = 0 To iMaxTracks - 1
		If (UnselectedTracks(i) = "") Then
			If instr(DiscogsTracksNum.Item(i), ".") <> 0 Then
				If SubTrackFlag = False Then
					SubTrackFlag = True
					SelectedTracksCount = SelectedTracksCount + 1
				End If
			Else
				If SubTrackFlag = True Then SubTrackFlag = False
				SelectedTracksCount = SelectedTracksCount + 1
			End If
		Else
			UnSelectedTracksCount = UnSelectedTracksCount + 1
		End If
	Next
	If (iMaxTracks - UnSelectedTracksCount) <> SDB.Tools.WebSearch.NewTracks.Count Then
		DiffNumberTracks = True
		templateHTML = templateHTML &  "<tr><td colspan=3 align=center><b><span style=""color:#FF0000"">There are different numbers of tracks !</span></b></td></tr>"
		templateHTML = templateHTML &  "<tr><td colspan=3 align=center><br></td></tr>"
	Else
		DiffNumberTracks = False
	End If

	templateHTML = templateHTML &  "<tr>"
	If QueryPage <> "MetalArchives" Then
		templateHTML = templateHTML &  "<td><input type=checkbox id=""releaseid"" title=""if set, the release number will be written in a custom tag. You can change the target tag in Tools->Options->Discogs Tagger. If Release-Number saved, the Tagger automatically load next time this release from discogs for the selected files"" ></td>"
		templateHTML = templateHTML &  "<td>Release:</td>"
	Else
		templateHTML = templateHTML &  "<td></td>"
		templateHTML = templateHTML &  "<td></td>"
	End If
	If InStr(Results.Item(CurrentResultId), " * ") <> 0 Or Right(Results.Item(CurrentResultId), 8) = "(Master)" Then
		templateHTML = templateHTML &  "<td>N/A</a> (Master: <a href=""https://www.discogs.com/master/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a>)</td>"
	ElseIf (theMaster <> "") Then
		templateHTML = templateHTML &  "<td><a href=""https://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a> (Master: <a href=""https://www.discogs.com/master/<!MASTERID!>"" target=""_blank""><!MASTERID!></a>)</td>"
	Else
		If QueryPage = "Discogs" Then
			templateHTML = templateHTML &  "<td><a href=""https://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a> (Master: N/A)</td>"
		ElseIf QueryPage = "MusicBrainz" Then
			templateHTML = templateHTML &  "<td><a href=""http://www.musicbrainz.org/release/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a> (Master: N/A)</td>"
		End If
	End If
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""artist"" ></td>"
	templateHTML = templateHTML &  "<td>Artist:</td>"
	If QueryPage = "Discogs" Then
		templateHTML = templateHTML &  "<td><a href=""https://www.discogs.com/artist/" & SavedArtistID & """ target=""_blank""><!ARTIST!></a></td>"
	ElseIf QueryPage = "MusicBrainz" Then
		templateHTML = templateHTML &  "<td><a href=""http://www.musicbrainz.org/artist/" & SavedArtistID & """ target=""_blank""><!ARTIST!></a></td>"
	ElseIf QueryPage = "MetalArchives" Then
		templateHTML = templateHTML &  "<td><!ARTIST!></td>"
	End If
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""album"" title=""Store the album name in the album-tag"" ></td>"
	templateHTML = templateHTML &  "<td>Album:</td>"
	If QueryPage = "Discogs" Then
		templateHTML = templateHTML &  "<td><a href=""https://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!ALBUMTITLE!></a></td>"
	ElseIf QueryPage = "MusicBrainz" Then
		templateHTML = templateHTML &  "<td><a href=""http://www.musicbrainz.org/release/<!RELEASEID!>"" target=""_blank""><!ALBUMTITLE!></a></td>"
	ElseIf QueryPage = "MetalArchives" Then
		templateHTML = templateHTML &  "<td><!ALBUMTITLE!></td>"
	End If
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""albumartist"" title=""Store the name of the albumartist in the albumartist-tag""></td>"
	REM <input type=checkbox title=""If option set, only the name of the first artist will be assumed"" id=""albumartistfirst"" ></td>"
	templateHTML = templateHTML &  "<td>Album Artist:</td>"
	If QueryPage = "Discogs" Then
		templateHTML = templateHTML &  "<td><a href=""https://www.discogs.com/artist/" & SavedArtistID & """ target=""_blank""><!ALBUMARTIST!></a></td>"
	ElseIf QueryPage = "MusicBrainz" Then
		templateHTML = templateHTML &  "<td><a href=""http://www.musicbrainz.org/artist/" & SavedArtistID & """ target=""_blank""><!ALBUMARTIST!></a></td>"
	ElseIf QueryPage = "MetalArchives" Then
		templateHTML = templateHTML &  "<td><!ALBUMARTIST!></td>"
	End If
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""label"" title=""Store the name of the label in the label-tag"" ></td>"
	templateHTML = templateHTML &  "<td>Label:</td>"
	If QueryPage = "Discogs" Then
		templateHTML = templateHTML &  "<td><a href=""https://www.discogs.com/label/" & SavedLabelID & """ target=""_blank""><!LABEL!></a></td>"
	ElseIf QueryPage = "MusicBrainz" Then
		templateHTML = templateHTML &  "<td><a href=""http://www.musicbrainz.org/label/" & SavedLabelID & """ target=""_blank""><!LABEL!></a></td>"
	ElseIf QueryPage = "MetalArchives" Then
		templateHTML = templateHTML &  "<td><!LABEL!></td>"
	End If
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""catalog"" title=""if set, the catalog number will be written in a custom tag. You can change the target tag in Tools->Options->Discogs Tagger.""></td>"
	templateHTML = templateHTML &  "<td>Catalog#:</td>"
	templateHTML = templateHTML &  "<td><!CATALOG!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""format"" title=""if set, the format will be written in a custom tag. You can change the target tag in Tools->Options->Discogs Tagger.""></td>"
	templateHTML = templateHTML &  "<td>Format:</td>"
	templateHTML = templateHTML &  "<td><!FORMAT!></td>"
	templateHTML = templateHTML &  "</tr>"
	If QueryPage <> "MetalArchives" Then
		templateHTML = templateHTML &  "<tr>"
		templateHTML = templateHTML &  "<td><input type=checkbox id=""country"" title=""if set, the country will be written in a custom tag. You can change the target tag in Tools->Options->Discogs Tagger.""></td>"
		templateHTML = templateHTML &  "<td>Country:</td>"
		templateHTML = templateHTML &  "<td><!COUNTRY!></td>"
		templateHTML = templateHTML &  "</tr>"
	End If
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox title=""If option set, the release date of this Discogs release will be saved in the date tag"" id=""date"" ></td>"
	templateHTML = templateHTML &  "<td>Date:</td>"
	templateHTML = templateHTML &  "<td><!RELEASEDATE!></td>"
	templateHTML = templateHTML &  "</tr>"
	If QueryPage <> "MetalArchives" Then
		templateHTML = templateHTML &  "<tr>"
		templateHTML = templateHTML &  "<td><input type=checkbox title=""If option set, the release date of the Discogs master release will be saved in the original date tag"" id=""origdate"" ></td>"
		templateHTML = templateHTML &  "<td>Original Date:</td>"
		templateHTML = templateHTML &  "<td><!ORIGDATE!></td>"
		templateHTML = templateHTML &  "</tr>"
	End If
	If QueryPage = "Discogs" Then
		templateHTML = templateHTML &  "<tr>"
		templateHTML = templateHTML &  "<td><input type=checkbox id=""genre"" title=""If set, the selected genres will be written in the genre-tag""></td>"
		templateHTML = templateHTML &  "<td>Genre:</td>"
		If GenresList.Count > 0 Then
			templateHTML = templateHTML & "<td><input type=checkbox id=""genre0"" >" & GenresList.Item(0) & "</td>"
			templateHTML = templateHTML &  "</tr>"
		End If
		If GenresList.Count > 1 Then
			For i = 1 To GenresList.Count-1
				templateHTML = templateHTML &  "<tr><td></td><td></td>"
				templateHTML = templateHTML & "<td><input type=checkbox id=""genre" & i & """ >" & GenresList.Item(i) & "</td>"
				templateHTML = templateHTML &  "</tr>"
			Next
		End If
		templateHTML = templateHTML &  "<tr>"
		templateHTML = templateHTML &  "<td><input type=checkbox id=""comments"" title=""Store the comment in the comment-tag"" ></td>"
		templateHTML = templateHTML &  "<td>Comment:</td>"
		templateHTML = templateHTML &  "<td style=""width: 250px"">" & Comment & "</td>"
		templateHTML = templateHTML &  "</tr>"

		templateHTML = templateHTML &  "<tr>"
		templateHTML = templateHTML &  "<td colspan=2>Release Data Quality:</td>"
		templateHTML = templateHTML &  "<td><!DATAQUALITY!></td>"
		templateHTML = templateHTML &  "</tr>"
	End If


	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Release Information End
	' Tracklisting Begin
	templateHTML = templateHTML & "<td align=left valign=top>"
	templateHTML = templateHTML & "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	templateHTML = templateHTML & "<tr>"

	If CheckOriginalDiscogsTrack Then
		If QueryPage = "Discogs" Then
			templateHTML = templateHTML & "<td align=left><b>Discogs</b></td>"
		ElseIf QueryPage = "MusicBrainz" Then
			templateHTML = templateHTML & "<td align=left><b>MusicBrainz</b></td>"
		End If
	Else
		templateHTML = templateHTML & "<td> </td>"
	End If

	templateHTML = templateHTML & "<td><input type=checkbox id=""selectall"" title=""Select/Deselect all tracks""></td>"
	REM templateHTML = templateHTML & "<td align=center><input type=checkbox id=""discnum""></td>"
	templateHTML = templateHTML & "<td><input type=checkbox id=""discnum""><b>Disc#</b></td>"
	REM templateHTML = templateHTML & "<td><b>Disc#</b></td>"
	REM templateHTML = templateHTML & "<td align=center><input type=checkbox id=""tracknum"" title=""If option NOT set, track numbers will not set automatically (useful when you didn't select all tracks from a release""></td>"
	templateHTML = templateHTML & "<td><input type=checkbox id=""tracknum"" title=""If option NOT set, track numbers will not set automatically (useful when you didn't select all tracks from a release""><b>Track#</b></td>"
	templateHTML = templateHTML & "<td align=right><b>Artist</b></td>"
	templateHTML = templateHTML & "<td> </td>"
	templateHTML = templateHTML & "<td align=left><b>Title</b></td>"
	templateHTML = templateHTML & "<td align=right><b>Duration</b></td>"
	templateHTML = templateHTML & "</tr>"

	For i=0 To iMaxTracks - 1
		templateHTML = templateHTML &  "<tr>"
		If Grouping(i) = "|Heading|" Then
			templateHTML = templateHTML & "</tr><tr> </tr>"
			templateHTML = templateHTML & "<tr> </tr>"
		End If
		If CheckOriginalDiscogsTrack Then
			templateHTML = templateHTML & "<td align=center>" & DiscogsTracksNum.Item(i) & "</td>"
		Else
			templateHTML = templateHTML & "<td> </td>"
		End If
		If Grouping(i) = "|Heading|" Then
			REM If(UnselectedTracks(i) = "") Then
				REM templateHTML = templateHTML & "<td><input type=checkbox title=""Select/Deselect all tracks according to the heading"" id=""unselected["&i&"]"" checked></td>"
				templateHTML = templateHTML & "<td></td>"
			REM Else
				REM templateHTML = templateHTML & "<td><input type=checkbox title=""Select/Deselect all tracks according to the heading"" id=""unselected["&i&"]""></td>"
			REM End If
		Else
			If(UnselectedTracks(i) = "") Then
				templateHTML = templateHTML & "<td><input type=checkbox title=""if set, the track will be used for tagging"" id=""unselected["&i&"]"" checked></td>"
			Else
				templateHTML = templateHTML & "<td><input type=checkbox title=""if set, the track will be used for tagging"" id=""unselected["&i&"]""></td>"
			End If
		End If
		If Grouping(i) = "|Heading|" Then
			templateHTML = templateHTML & "<td align=center></td><td align=center></td><td align=right></td>"
			templateHTML = templateHTML & "<td align=center></td>"
			templateHTML = templateHTML & "<td align=left><b>Heading: " & Tracks.Item(i) & "</b></td><td align=right></td>"
		Else
			templateHTML = templateHTML & "<td align=center>" & TracksCD.Item(i) & "</td>"
			templateHTML = templateHTML & "<td align=center>" & TracksNum.Item(i) & "</td>"
			templateHTML = templateHTML & "<td align=right>" & ArtistTitles.Item(i) & "</td>"
			templateHTML = templateHTML & "<td align=center><b>-</b></td>"
			templateHTML = templateHTML & "<td align=left>"
			If(UnselectedTrackNames(i) = "") Then
				templateHTML = templateHTML & "<input type=checkbox title=""if set, the track title will be updated"" id=""unselectedtrackname["&i&"]"" checked>" & Tracks.Item(i) & "</td>"
			Else
				templateHTML = templateHTML & "<input type=checkbox title=""if set, the track title will be updated"" id=""unselectedtrackname["&i&"]"">" & Tracks.Item(i) & "</td>"
			End If
			templateHTML = templateHTML & "<td align=right>" & Durations.Item(i) & "</td>"
		End If
		templateHTML = templateHTML & "</tr>"
		
		REM If Grouping(i) = "|Heading|" Then
			REM templateHTML = templateHTML & "<tr></tr><tr></tr>"
		REM Else
			If(CheckLyricist and Lyricists.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Lyrics: "& Lyricists.Item(i) &"</td></tr>"
			If(CheckComposer and Composers.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Composer: "& Composers.Item(i) &"</td></tr>"
			If(CheckConductor and Conductors.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Conductor: "& Conductors.Item(i) &"</td></tr>"
			If(CheckProducer and Producers.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Producer: "& Producers.Item(i) &"</td></tr>"
			If(CheckGrouping and Grouping(i) <> "" and Grouping(i) <> "|Heading|") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Grouping: "& Grouping(i) &"</td></tr>"
			
			If(CheckInvolved and InvolvedArtists.Item(i) <> "") Then
				templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left><b>Involved People:</b></td></tr>"
				'SDB.Localize("Involved People")
				If CheckInvolvedPeopleSingleLine = True And InStr(InvolvedArtists.Item(i), Separator) <> 0 Then
					Dim x
					tmp = Split(InvolvedArtists.Item(i), Separator)
					For each x in tmp
						templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>"& x &"</td></tr>"
					Next
				Else
					templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>"& InvolvedArtists.Item(i) &"</td></tr>"
				End If
			End If
		REM End If
	Next

	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Tracklisting End

	templateHTML = templateHTML &  GetFooter()

	templateHTML = Replace(templateHTML, "<!RELEASEID!>", releaseID)
	templateHTML = Replace(templateHTML, "<!MASTERID!>", theMaster)
	templateHTML = Replace(templateHTML, "<!ARTIST!>", AlbumArtistTitle)
	templateHTML = Replace(templateHTML, "<!ALBUMARTIST!>",  AlbumArtist)
	templateHTML = Replace(templateHTML, "<!ALBUMTITLE!>", AlbumTitle)
	templateHTML = Replace(templateHTML, "<!RELEASEDATE!>", ReleaseDate)
	templateHTML = Replace(templateHTML, "<!ORIGDATE!>", OriginalDate)
	templateHTML = Replace(templateHTML, "<!LABEL!>", theLabels)
	templateHTML = Replace(templateHTML, "<!COUNTRY!>", theCountry)
	templateHTML = Replace(templateHTML, "<!COVER!>", AlbumArtThumbNail)
	templateHTML = Replace(templateHTML, "<!CATALOG!>", Catalog)
	templateHTML = Replace(templateHTML, "<!FORMAT!>", theFormat)
	
	If QueryPage = "Discogs" Then
		templateHTML = Replace(templateHTML, "<!DATAQUALITY!>", DataQuality)
	End If

	
	' Dim filesys, filetxt, logdatei
	' Const ForReading = 1, ForWriting = 2, ForAppending = 8
	' logdatei = SDB.ScriptsPath & "HTML.htm"
	' Set filesys = CreateObject("Scripting.FileSystemObject")
	' Set filetxt = filesys.OpenTextFile(logdatei, 2, True)
	' filetxt.WriteLine(templateHTML)
	' filetxt.Close

	WebBrowser.SetHTMLDocument templateHTML

	Dim templateHTMLDoc
	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set checkBox = templateHTMLDoc.getElementById("album")
	checkBox.Checked = CheckAlbum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("artist")
	checkBox.Checked = CheckArtist
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("albumartist")
	checkBox.Checked = CheckAlbumArtist
	Script.RegisterEvent checkBox, "onclick", "Update"
	REM Set checkBox = templateHTMLDoc.getElementById("albumartistfirst")
	REM checkBox.Checked = CheckAlbumArtistFirst
	REM Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("date")
	checkBox.Checked = CheckDate
	Script.RegisterEvent checkBox, "onclick", "Update"
	If QueryPage <> "MetalArchives" Then
		Set checkBox = templateHTMLDoc.getElementById("origdate")
		checkBox.Checked = CheckOrigDate
		Script.RegisterEvent checkBox, "onclick", "Update"
		Set checkBox = templateHTMLDoc.getElementById("country")
		checkBox.Checked = CheckCountry
		Script.RegisterEvent checkBox, "onclick", "Update"
	End If
	Set checkBox = templateHTMLDoc.getElementById("label")
	checkBox.Checked = CheckLabel
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("picture")
	Script.RegisterEvent checkBox, "onclick", "ShowHelp"
	If QueryPage = "Discogs" Then
		Set checkBox = templateHTMLDoc.getElementById("genre")
		checkBox.Checked = CheckGenre
		Script.RegisterEvent checkBox, "onclick", "Update"
		NewGenre = ""
		If GenresList.Count > 0 Then
			For i = 0 To GenresList.Count-1
				Set checkBox = templateHTMLDoc.getElementById("genre" & i)
				checkBox.Checked = GenresSelect.Item(i)
				Script.RegisterEvent checkBox, "onclick", "Update"
				If GenresSelect.Item(i) = "True" Then
					AddToField NewGenre, GenresList.Item(i)
				End If
			Next
		End If
	End If
	Set checkBox = templateHTMLDoc.getElementById("checkcover")
	checkBox.Checked = CheckCover
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("catalog")
	checkBox.Checked = CheckCatalog
	Script.RegisterEvent checkBox, "onclick", "Update"
	If QueryPage <> "MetalArchives" Then
		Set checkBox = templateHTMLDoc.getElementById("releaseid")
		checkBox.Checked = CheckRelease
		Script.RegisterEvent checkBox, "onclick", "Update"
		Set checkBox = templateHTMLDoc.getElementById("involved")
		checkBox.Checked = CheckInvolved
		Script.RegisterEvent checkBox, "onclick", "Update"
		Set checkBox = templateHTMLDoc.getElementById("lyricist")
		checkBox.Checked = CheckLyricist
		Script.RegisterEvent checkBox, "onclick", "Update"
		Set checkBox = templateHTMLDoc.getElementById("composer")
		checkBox.Checked = CheckComposer
		Script.RegisterEvent checkBox, "onclick", "Update"
		Set checkBox = templateHTMLDoc.getElementById("conductor")
		checkBox.Checked = CheckConductor
		Script.RegisterEvent checkBox, "onclick", "Update"
		Set checkBox = templateHTMLDoc.getElementById("producer")
		checkBox.Checked = CheckProducer
		Script.RegisterEvent checkBox, "onclick", "Update"
	End If
	Set checkBox = templateHTMLDoc.getElementById("grouping")
	checkBox.Checked = CheckGrouping
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set submitButton = templateHTMLDoc.getElementById("refresh")
	Script.RegisterEvent submitButton, "onclick", "trackrefreshing"
	Set submitButton = templateHTMLDoc.getElementById("deselectall")
	Script.RegisterEvent submitButton, "onclick", "deselectall"
	REM Set checkBox = templateHTMLDoc.getElementById("usercollection")
	REM checkBox.Checked = CheckUserCollection
	REM Script.RegisterEvent checkBox, "onclick", "Update"
	
	Set checkBox = templateHTMLDoc.getElementById("discnum")
	checkBox.Checked = CheckDiscNum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("tracknum")
	checkBox.Checked = CheckTrackNum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("format")
	checkBox.Checked = CheckFormat
	Script.RegisterEvent checkBox, "onclick", "Update"
	
	If QueryPage = "Discogs" Then
		Set checkBox = templateHTMLDoc.getElementById("comments")
		checkBox.Checked = CheckComment
		Script.RegisterEvent checkBox, "onclick", "Update"
	End If

	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("load")
	Script.RegisterEvent listBox, "onchange", "Filter"

	If trackRefresh = True Then
		For i=0 To iMaxTracks - 1
			Set checkBox = templateHTMLDoc.getElementById("unselected["&i&"]")
			Script.RegisterEvent checkBox, "onclick", "Unselect"
			Set checkBox = templateHTMLDoc.getElementById("unselectedtrackname["&i&"]")
			Script.RegisterEvent checkBox, "onclick", "Unselect"
		Next
	End If

	Set checkBox = templateHTMLDoc.getElementById("selectall")
	checkBox.Checked = SelectAll
	Script.RegisterEvent checkBox, "onclick", "SwitchAll"

	If QueryPage <> "MetalArchives" Then
		Set listBox = templateHTMLDoc.getElementById("alternative")
		Script.RegisterEvent listBox, "onchange", "Alternative"

		Set submitButton = templateHTMLDoc.getElementById("showcountryfilter")
		Script.RegisterEvent submitButton, "onclick", "ShowCountryFilter"

		Set submitButton = templateHTMLDoc.getElementById("showmediatypefilter")
		Script.RegisterEvent submitButton, "onclick", "ShowMediaTypeFilter"

		Set submitButton = templateHTMLDoc.getElementById("showmediaformatfilter")
		Script.RegisterEvent submitButton, "onclick", "ShowMediaFormatFilter"

		Set submitButton = templateHTMLDoc.getElementById("showyearfilter")
		Script.RegisterEvent submitButton, "onclick", "ShowYearFilter"
	End If

	Set submitButton = templateHTMLDoc.getElementById("saveoptions")
	Script.RegisterEvent submitButton, "onclick", "SaveOptions"
	
	Set submitButton = templateHTMLDoc.getElementById("optionsform")
	Script.RegisterEvent submitButton, "onclick", "ShowOptions"



	Set submitButton = templateHTMLDoc.getElementById("moreimages")
	Script.RegisterEvent submitButton, "onclick", "MoreImages"

	REM Set listBox = templateHTMLDoc.getElementById("searchpage")
	REM Script.RegisterEvent listBox, "onchange", "Filter"

	Set submitButton = templateHTMLDoc.getElementById("showadvancedsearch")
	Script.RegisterEvent submitButton, "onclick", "ShowAdvancedSearch"

	Set submitButton = templateHTMLDoc.getElementById("usercollection")
	Script.RegisterEvent submitButton, "onclick", "UserCollection"


End Sub


Sub ShowHelp()

	Dim res
	res = SDB.MessageBox("Short Help MessageBox" & vbCrLf & vbCrLf & "Hoving with the mouse over a checkbox will show a short help." & vbCrLf & "You find some more configuration settings for Discogs Tagger in MediaMonkey menu ->Tools->Options" & vbCrLf & vbCrLf & "If you want more help or information, or if you want to ask for a new feature," & vbCrLf & "please visit the Discogs Tagger on the Mediamonkey forum, we wanna try to help you" & vbCrLf & vbCrLf & "http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=72637", mtInformation, Array(mbOk))


End Sub



Sub trackRefreshing()

	If trackRefresh = True then
		trackRefresh = False
		Unselect
	Else
		trackRefresh = True
		UnSelect
	End If

End Sub


Sub deselectall()

	Dim checkbox, templateHTMLDoc
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	
	Set checkBox = templateHTMLDoc.getElementById("album")
	CheckAlbum = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("artist")
	CheckArtist = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("albumartist")
	CheckAlbumArtist = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("date")
	CheckDate = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("origdate")
	CheckOrigDate = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("label")
	CheckLabel = False
	checkBox.Checked = False
	If QueryPage = "Discogs" Then
		Set checkBox = templateHTMLDoc.getElementById("genre")
		CheckGenre = False
		checkBox.Checked = False
	End If
	Set checkBox = templateHTMLDoc.getElementById("format")
	CheckFormat = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("country")
	CheckCountry = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("checkcover")
	CheckCover = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("catalog")
	CheckCatalog = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("releaseid")
	CheckRelease = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("involved")
	CheckInvolved = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("grouping")
	CheckGrouping = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("lyricist")
	CheckLyricist = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("composer")
	CheckComposer = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("conductor")
	CheckConductor = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("producer")
	CheckProducer = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("discnum")
	CheckDiscNum = False
	checkBox.Checked = False
	Set checkBox = templateHTMLDoc.getElementById("tracknum")
	CheckTrackNum = False
	checkBox.Checked = False
	If QueryPage = "Discogs" Then
		Set checkBox = templateHTMLDoc.getElementById("comments")
		CheckComment = False
		checkBox.Checked = False
	End If
	
	OptionsChanged = True

	SDB.ProcessMessages
	ReloadResults
	
End Sub



Sub Update()

	Dim templateHTMLDoc, checkBox, text, radio, i
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
	button.style.backgroundColor = "#ED1C24"
	templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML="Please wait for collection data..."
	Set checkBox = templateHTMLDoc.getElementById("album")
	CheckAlbum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("artist")
	CheckArtist = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("albumartist")
	CheckAlbumArtist = checkBox.Checked
	REM Set checkBox = templateHTMLDoc.getElementById("albumartistfirst")
	REM CheckAlbumArtistFirst = checkBox.Checked
	If QueryPage <> "MetalArchives" Then
		Set checkBox = templateHTMLDoc.getElementById("origdate")
		CheckOrigDate = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("country")
		CheckCountry = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("releaseid")
		CheckRelease = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("involved")
		CheckInvolved = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("lyricist")
		CheckLyricist = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("composer")
		CheckComposer = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("conductor")
		CheckConductor = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("producer")
		CheckProducer = checkBox.Checked
	End If
	Set checkBox = templateHTMLDoc.getElementById("date")
	CheckDate = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("label")
	CheckLabel = checkBox.Checked
	If QueryPage = "Discogs" Then
		Set checkBox = templateHTMLDoc.getElementById("genre")
		CheckGenre = checkBox.Checked
		If GenresList.Count > 0 Then
			For i = 0 To GenresList.Count-1
				Set checkBox = templateHTMLDoc.getElementById("genre" & i)
				GenresSelect.Item(i) = checkBox.Checked
			Next
		End If
	End If
	Set checkBox = templateHTMLDoc.getElementById("format")
	CheckFormat = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("checkcover")
	CheckCover = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("catalog")
	CheckCatalog = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("grouping")
	CheckGrouping = checkBox.Checked
	REM Set checkBox = templateHTMLDoc.getElementById("usercollection")
	REM CheckUserCollection = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("discnum")
	CheckDiscNum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("tracknum")
	CheckTrackNum = checkBox.Checked
	If QueryPage = "Discogs" Then
		Set checkBox = templateHTMLDoc.getElementById("comments")
		CheckComment = checkBox.Checked
	End If

	OptionsChanged = True

	SDB.ProcessMessages
	ReloadResults

End Sub


Sub Filter()

	Dim templateHTMLDoc, listBox, ret
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	FilterMediaType = listBox.Value
	If FilterMediaType = "None" Then
		MediaTypeFilterList.Item(0) = "0"
	ElseIf FilterMediaType = "Use MediaType Filter" Then
		MediaTypeFilterList.Item(0) = "1"
	Else
		MediaTypeFilterList.Item(0) = FilterMediaType
	End If
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	FilterMediaFormat = listBox.Value
	If FilterMediaFormat = "None" Then
		MediaFormatFilterList.Item(0) = "0"
	ElseIf FilterMediaFormat = "Use MediaFormat Filter" Then
		MediaFormatFilterList.Item(0) = "1"
	Else
		MediaFormatFilterList.Item(0) = FilterMediaFormat
	End If
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	FilterCountry = listBox.Value
	If FilterCountry = "None" Then
		CountryFilterList.Item(0) = "0"
	ElseIf FilterCountry = "Use Country Filter" Then
		CountryFilterList.Item(0) = "1"
	Else
		CountryFilterList.Item(0) = FilterCountry
	End If
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	FilterYear = listBox.Value
	If FilterYear = "None" Then
		YearFilterList.Item(0) = "0"
	ElseIf FilterYear = "Use Year Filter" Then
		YearFilterList.Item(0) = "1"
	Else
		YearFilterList.Item(0) = FilterYear
	End If

	REM Set listBox = templateHTMLDoc.getElementById("searchpage")
	REM If QueryPage <> Mid(listBox.Value, 11) Then
		REM If Mid(listBox.Value, 11) = "Discogs" Then
			REM If authorize_script = False Then
				REM Exit Sub
			REM End If
		REM End If
		REM QueryPage = Mid(listBox.Value, 11)
		REM Set listBox = templateHTMLDoc.getElementById("load")
		REM listBox.Item(0).Selected = True
		REM ini.StringValue("DiscogsAutoTagWeb","QueryPage") = QueryPage
	REM End If

	Set listBox = templateHTMLDoc.getElementById("load")
	CurrentLoadType = listBox.Value

	If(CurrentLoadType = "Master Release") Then
		LoadMasterResults(SavedMasterID)
	ElseIf(CurrentLoadType = "Versions of Master") Then
		LoadVersionResults(SavedMasterID)
	ElseIf(CurrentLoadType = "Releases of Artist") Then
		LoadArtistResults(SavedArtistID)
	ElseIf(CurrentLoadType = "Releases of Label") Then
		LoadLabelResults(SavedLabelID)
	Else
		SDB.ProcessMessages
		NewSearchArtist = SavedSearchArtist
		NewSearchAlbum = SavedSearchAlbum
		FindResults NewSearchTerm, QueryPage
	End If

End Sub


Sub Alternative()

	Dim templateHTMLDoc
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
	button.style.backgroundColor = "#ED1C24"
	templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML="Please wait for collection data..."
	CurrentLoadType = "Search Results"
	NewSearchTerm = templateHTMLDoc.getElementById("alternative").Value
	NewSearchArtist = ""
	NewSearchAlbum = ""
	FindResults templateHTMLDoc.getElementById("alternative").Value, QueryPage
	
End Sub


Sub SaveOptions()

	Dim a, tmp
	' save options if ini exists
	If Not (ini Is Nothing) Then
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum") = CheckAlbum
		ini.BoolValue("DiscogsAutoTagWeb","CheckArtist") = CheckArtist
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist") = CheckAlbumArtist
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = CheckAlbumArtistFirst
		ini.BoolValue("DiscogsAutoTagWeb","CheckLabel") = CheckLabel
		ini.BoolValue("DiscogsAutoTagWeb","CheckDate") = CheckDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate") = CheckOrigDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckGenre") = CheckGenre
		ini.BoolValue("DiscogsAutoTagWeb","CheckCountry") = CheckCountry
		ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog") = CheckCatalog
		ini.BoolValue("DiscogsAutoTagWeb","CheckRelease") = CheckRelease
		ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved") = CheckInvolved
		ini.BoolValue("DiscogsAutoTagWeb","CheckGrouping") = CheckGrouping
		ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist") = CheckLyricist
		ini.BoolValue("DiscogsAutoTagWeb","CheckComposer") = CheckComposer
		ini.BoolValue("DiscogsAutoTagWeb","CheckConductor") = CheckConductor
		ini.BoolValue("DiscogsAutoTagWeb","CheckProducer") = CheckProducer
		ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum") = CheckDiscNum
		ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum") = CheckTrackNum
		ini.BoolValue("DiscogsAutoTagWeb","CheckFormat") = CheckFormat
		ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv") = CheckUseAnv
		ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = CheckYearOnlyDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric") = CheckForceNumeric
		ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc") = CheckSidesToDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc") = CheckForceDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc") = CheckNoDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero") = CheckLeadingZero
		ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZeroDisc") = CheckLeadingZeroDisc
		REM ini.BoolValue("DiscogsAutoTagWeb","CheckUserCollection") = CheckUserCollection
		ini.StringValue("DiscogsAutoTagWeb","DiscogsUsername") = DiscogsUsername
		ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = ReleaseTag
		ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = CatalogTag
		ini.StringValue("DiscogsAutoTagWeb","CountryTag") = CountryTag
		ini.StringValue("DiscogsAutoTagWeb","FormatTag") = FormatTag
		ini.BoolValue("DiscogsAutoTagWeb","CheckVarious") = CheckVarious
		ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = TxtVarious
		ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = CheckTitleFeaturing
		ini.BoolValue("DiscogsAutoTagWeb","CheckDeleteDuplicatedEntry") = CheckDeleteDuplicatedEntry
		ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = TxtFeaturingName
		ini.BoolValue("DiscogsAutoTagWeb","CheckFeaturingName") = CheckFeaturingName
		ini.BoolValue("DiscogsAutoTagWeb","CheckComment") = CheckComment
		ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection") = SubTrackNameSelection
		ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = ArtistSeparator
		ini.StringValue("DiscogsAutoTagWeb","FormatSeparator") = FormatSeparator
		ini.BoolValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack") = CheckTurnOffSubTrack
		ini.BoolValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine") = CheckInvolvedPeopleSingleLine
		REM ini.StringValue("DiscogsAutoTagWeb","QueryPage") = QueryPage
		ini.BoolValue("DiscogsAutoTagWeb","CheckTheBehindArtist") = CheckTheBehindArtist
		ini.BoolValue("DiscogsAutoTagWeb","CheckIgnoreFeatArtist") = CheckIgnoreFeatArtist
		ini.StringValue("DiscogsAutoTagWeb","SubTrackSeparator") = SubTrackSeparator
		ini.BoolValue("DiscogsAutoTagWeb","CheckNewMatchingEngine") = CheckNewMatchingEngine
		ini.BoolValue("DiscogsAutoTagWeb","CheckShowQtyFormat") = CheckShowQtyFormat
		
		tmp = CountryFilterList.Item(0)
		For a = 1 To CountryList.Count - 1
			tmp = tmp & "," & CountryFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = tmp
		tmp = MediaTypeFilterList.Item(0)
		For a = 1 To MediaTypeList.Count - 1
			tmp = tmp & "," & MediaTypeFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = tmp
		tmp = MediaFormatFilterList.Item(0)
		For a = 1 To MediaFormatList.Count - 1
			tmp = tmp & "," & MediaFormatFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = tmp
		tmp = YearFilterList.Item(0)
		For a = 1 To YearList.Count - 1
			tmp = tmp & "," & YearFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmp
	End If

End Sub

' Format Error Message
Sub FormatErrorMessage(ErrorMessage)

	Dim templateHTML, listBox, templateHTMLDoc, submitButton, res, ret, tmp
	
	WriteLog "Show ErrorMessage"
	
	If InStr(ErrorMessage, "{") > 0 Then
		tmp = InStr(ErrorMessage, "{")
		ErrorMessage = Mid(ErrorMessage, tmp + 28)
		tmp = InStr(ErrorMessage, "}")
		ErrorMessage = Left(ErrorMessage, tmp - 7)
	End If
	WriteLog "Errormessage=" & ErrorMessage
	
	templateHTML = ""
	templateHTML = templateHTML &  GetHeader()
	templateHTML = templateHTML &  "<tr></tr>"
	templateHTML = templateHTML &  "<tr></tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td colspan=4 align=center><p><b>" & ErrorMessage & "</b></p></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  GetFooter()

	Set WebBrowser = SDB.Objects("WebBrowser")
	WebBrowser.SetHTMLDocument templateHTML
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
	button.style.backgroundColor = "#FFFFFF"
	templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML=""

	WebBrowser.Interf.Visible = true
	WebBrowser.Common.BringToFront

	REM Set listBox = templateHTMLDoc.getElementById("searchpage")
	REM Script.RegisterEvent listBox, "onchange", "Filter"

	Set listBox = templateHTMLDoc.getElementById("alternative")
	Script.RegisterEvent listBox, "onchange", "Alternative"

	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("load")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set submitButton = templateHTMLDoc.getElementById("showcountryfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowCountryFilter"
	Set submitButton = templateHTMLDoc.getElementById("showmediatypefilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaTypeFilter"
	Set submitButton = templateHTMLDoc.getElementById("showmediaformatfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaFormatFilter"
	Set submitButton = templateHTMLDoc.getElementById("showyearfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowYearFilter"
	Set submitButton = templateHTMLDoc.getElementById("showadvancedsearch")
	Script.RegisterEvent submitButton, "onclick", "ShowAdvancedSearch"

	SDB.Tools.WebSearch.ClearTracksData
	
	If InStr(ErrorMessage, "You must authenticate to access this resource") > 0 Then
		res = SDB.MessageBox("OAuth Client Authentication Error. Do you want to re-authenticate Discogs Tagger ?", mtConfirmation, Array(mbYes, mbNo))
		If res = 6 Then
			AccessToken = ""
			AccessTokenSecret = ""
			ret = authorize_script(True)
			If ret = True Then
				FindResults NewSearchTerm, QueryPage
			End If
		End If
	End If

End Sub


Function JSONParser_find_result(searchURL, ArrayName, SendArtist, SendAlbum, SendTrack, SendType, SendDBSearch, SendPerPage, QueryPage, useOAuth)

	'useOAuth = True -> OAuth needed
	Dim oXMLHTTP, r, f, a, currentArtist, media, l
	Dim json
	Set json = New VbsJson

	Dim response
	Dim format, title, country, v_year, label, artist, Rtype, catNo, main_release, tmp, ReleaseDesc, FilterFound, SongCount, SongCountMax, isRelease, listCount
	Dim Page, SongPages, trackCount, currentLabel

	WriteLog " "
	WriteLog "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+"
	WriteLog "Start JSONParser_find_result"
	WriteLog "Arrayname=" & ArrayName
	WriteLog "QueryPage=" & QueryPage
	WriteLog "searchURL=" & searchURL
	ErrorMessage = ""


	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")

	If QueryPage = "MetalArchives" Then

		Dim MAReleases, MARelease, MABand, MAFormat, MALabel, MACatalogID, MADesc, MAReleasenumber, MADate, ctmp, MAURL
		Dim TXTBegin, TXTEnd, ResponseHTML, i

		Set Results = SDB.NewStringList
		Set ResultsReleaseID = SDB.NewStringList
		ErrorMessage = ""
		' searchURL = "https://www.metal-archives.com/search/ajax-advanced/searching/albums/?bandName=" & URLEncodeUTF8(CleanSearchString(NewSearchArtist)) & "&exactBandMatch=1&releaseTitle=" & URLEncodeUTF8(CleanSearchString(NewSearchAlbum)) & "&exactReleaseMatch=1&releaseYearFrom=&releaseMonthFrom=&releaseYearTo=&releaseMonthTo=&country=&location=&releaseLabelName=&genre="
		WriteLog searchURL
		oXMLHTTP.open "GET", searchURL, false
		oXMLHTTP.send()
		if oXMLHTTP.Status = 200 Then
			ResponseHTML = oXMLHTTP.responseText
			' Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
			' Dim f
			' Set f = fso.OpenTextFile(SDB.ScriptsPath&"test.log", 2, true, -1)
			' f.WriteLine ResponseHTML
			' f.Close
			TXTBegin = InStr(ResponseHTML, "iTotalDisplayRecords")
			If TXTBegin > 1 Then
				ResponseHTML = Mid(ResponseHTML, TXTBegin + 23)
				TXTEnd = InStr(ResponseHTML, ",")
				MAReleases = Left(ResponseHTML, TXTEnd -1)
				WriteLog "Anzahl=" & MAReleases
				If MAReleases = 1 Then
					TXTBegin = InStr(ResponseHTML, "\" & Chr(34) & ">")
					If TXTBegin > 1 Then
						ResponseHTML = Mid(ResponseHTML, TXTBegin + 3)
						TXTEnd = InStr(responseHTML, "</a>")
						WriteLog "Band=" & Left(ResponseHTML, TXTEnd - 1)
						MABand = Left(ResponseHTML, TXTEnd - 1)
						
						TXTBegin = InStr(ResponseHTML, "href=\")
						ResponseHTML = Mid(ResponseHTML, TXTBegin + 7)
						TXTEnd = InStr(ResponseHTML, "\" & Chr(34) & ">")
						WriteLog "Album-URL=" & Left(ResponseHTML, TXTEnd - 1)
						tmp = Left(ResponseHTML, TXTEnd - 1)
						ctmp = InStrRev(tmp, "/")
						MAReleasenumber = Mid(tmp, ctmp+1)
						WriteLog "ReleaseNumber=" & MAReleasenumber

						TXTBegin = TXTEnd + 3
						ResponseHTML = Mid(ResponseHTML, TXTBegin)
						TXTEnd = InStr(responseHTML, "</a>")
						WriteLog "Album=" & Left(ResponseHTML, TXTEnd - 1)
						MARelease = Left(ResponseHTML, TXTEnd - 1)

						ResponseHTML = Mid(ResponseHTML, TXTEnd)
						TXTBegin = InStr(ResponseHTML, ",")
						ResponseHTML = Mid(ResponseHTML, TXTBegin)
						TXTBegin = InStr(ResponseHTML, Chr(34))
						ResponseHTML = Mid(ResponseHTML, TXTBegin + 1)
						TXTEnd = InStr(responseHTML, Chr(34))
						WriteLog "Type=" & Left(ResponseHTML, TXTEnd - 1)
						
						searchURL = "https://www.metal-archives.com/release/ajax-versions/current/" & MAReleasenumber & "/parent/" & MAReleasenumber
						WriteLog searchURL
						oXMLHTTP.open "GET", searchURL, false
						oXMLHTTP.send()
						if oXMLHTTP.Status = 200 Then
							ResponseHTML = oXMLHTTP.responseText
							Do
								TXTBegin = InStr(ResponseHTML, "<td>")
								ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
								TXTBegin = InStr(ResponseHTML, "href=")
								ResponseHTML = Mid(ResponseHTML, TXTBegin + 6)
								TXTEnd = InStr(ResponseHTML, Chr(34) & ">")
								WriteLog "Album-URL=" & Left(ResponseHTML, TXTEnd - 1)
								MAURL = Left(ResponseHTML, TXTEnd - 1)
								ResponseHTML = Mid(ResponseHTML, TXTEnd + 2)
								TXTEnd = InStr(ResponseHTML, "</a>")
								MADate = Left(ResponseHTML, TXTEnd -1)
								WriteLog "Date=" & MADate
								
								ResponseHTML = Mid(ResponseHTML, TXTEnd + 4)
								TXTBegin = InStr(ResponseHTML, "<td>")
								ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
								TXTEnd = InStr(ResponseHTML, "</td>")
								MALabel = Left(ResponseHTML, TXTEnd - 1)
								WriteLog "Label=" & MALabel
								
								ResponseHTML = Mid(ResponseHTML, TXTEnd)
								TXTBegin = InStr(ResponseHTML, "<td>")
								ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
								TXTEnd = InStr(ResponseHTML, "</td>")
								If TXTEnd > 1 Then
									MACatalogID = Left(ResponseHTML, TXTEnd - 1)
								Else
									MACatalogID = ""
								End If
								WriteLog "Catalog ID=" & MACatalogID
								
								ResponseHTML = Mid(ResponseHTML, TXTEnd)
								TXTBegin = InStr(ResponseHTML, "<td>")
								ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
								TXTEnd = InStr(ResponseHTML, "</td>")
								MAFormat = Left(ResponseHTML, TXTEnd - 1)
								WriteLog "Format=" & MAFormat
								
								ResponseHTML = Mid(ResponseHTML, TXTEnd)
								TXTBegin = InStr(ResponseHTML, "<td>")
								ResponseHTML = Mid(ResponseHTML, TXTBegin + 4)
								TXTEnd = InStr(ResponseHTML, "</td>")
								If TXTEnd > 1 Then
									MADesc = Left(ResponseHTML, TXTEnd - 1)
								Else
									MADesc = ""
								End If
								WriteLog "Description=" & MADesc
								ResponseHTML = Mid(ResponseHTML, TXTEnd)
								
								Results.Add MABand & " - " & MARelease & " (" & MADate & ") " & MALabel & " / " & MAFormat & " / " & MACatalogID
								ResultsReleaseID.Add MAURL
								
							Loop While InStr(ResponseHTML, "<td>") > 0
						End If
					End If
				Else
					For i = 1 to MAReleases
						TXTBegin = InStr(ResponseHTML, "\" & Chr(34) & ">")
						If TXTBegin > 1 Then
							ResponseHTML = Mid(ResponseHTML, TXTBegin + 3)
							TXTEnd = InStr(responseHTML, "</a>")

							WriteLog "Band=" & Left(ResponseHTML, TXTEnd - 1)
							ReleaseDesc = Left(ResponseHTML, TXTEnd - 1)
							
							TXTBegin = InStr(ResponseHTML, "href=\")
							ResponseHTML = Mid(ResponseHTML, TXTBegin + 7)
							TXTEnd = InStr(ResponseHTML, "\" & Chr(34) & ">")
							WriteLog "Album-URL=" & Left(ResponseHTML, TXTEnd - 1)
							ResultsReleaseID.Add Left(ResponseHTML, TXTEnd - 1)
							
							TXTBegin = TXTEnd + 3
							ResponseHTML = Mid(ResponseHTML, TXTBegin)
							TXTEnd = InStr(responseHTML, "</a>")
							WriteLog "Album=" & Left(ResponseHTML, TXTEnd - 1)
							ReleaseDesc = ReleaseDesc & " / " & Left(ResponseHTML, TXTEnd - 1)

							ResponseHTML = Mid(ResponseHTML, TXTEnd)
							TXTBegin = InStr(ResponseHTML, ",")
							ResponseHTML = Mid(ResponseHTML, TXTBegin)
							TXTBegin = InStr(ResponseHTML, Chr(34))
							ResponseHTML = Mid(ResponseHTML, TXTBegin + 1)
							TXTEnd = InStr(responseHTML, Chr(34))
							WriteLog "Type=" & Left(ResponseHTML, TXTEnd - 1)
							ReleaseDesc = ReleaseDesc & " / " & Left(ResponseHTML, TXTEnd - 1)
							Results.Add ReleaseDesc
						End If
					Next
				End If
			End If
		End If
	End If



	If QueryPage = "MusicBrainz" Then

		oXMLHTTP.Open "GET", searchURL, False
		oXMLHTTP.setRequestHeader "Content-Type", "application/json"
		oXMLHTTP.setRequestHeader "User-Agent",UserAgent
		oXMLHTTP.send ()

		If oXMLHTTP.Status = 200 Then
			WriteLog "Musicbrainz"
			WriteLog "responseText=" & oXMLHTTP.responseText
			Set response = json.Decode(oXMLHTTP.responseText)
			If ArrayName = "Artist" Or ArrayName = "Label" Then
				ArrayName = "releases"
				SongCountMax = response("release-count")
			Else
				SongCountMax = response("count")
			End If
			'check if any results
			'and add titles to drop down
			'msgbox response(ArrayName)(0)("title")

			SongCount = 0
			
			WriteLog "SongCountMax=" & SongCountMax
			If Int(SongCountMax) = 0 Then
				WriteLog "No Release found at MusicBrainz"
				Results.Add "No Release found at MusicBrainz !! Please change the search parameter."
				ResultsReleaseID.Add 0
			Else

				isRelease = False
				If Results.Count = 1 Then isRelease = True
				For Each r In response(ArrayName) 'releases
					format = ""
					title = ""
					country = ""
					v_year = ""
					artist = ""
					label = ""
					catNo = ""
					main_release = ""
					media = ""

					SDB.ProcessMessages

					WriteLog "Release Nr." & SongCount
					title = response(ArrayName)(r)("title")
					WriteLog "title=" & title
					Set CurrentRelease = response(ArrayName)(r)
					If CurrentRelease.Exists("artist-credit") Then
						Set currentArtist = response(ArrayName)(r)("artist-credit")
						For Each f In currentArtist
							artist = artist & currentArtist(f)("artist")("name") & ", "
						Next
						If Len(artist) <> 0 Then artist = Left(artist, Len(artist)-2)
						WriteLog "artist=" & artist
					End If

					If CurrentRelease.Exists("media") Then
						For Each f In response(ArrayName)(r)("media")
							If response(ArrayName)(r)("media")(f)("format") <> "" Then
								If format = "" Then
									format = response(ArrayName)(r)("media")(f)("format")
									media = format
								Else
									If media <> response(ArrayName)(r)("media")(f)("format") Then
										format = format & " & " & response(ArrayName)(r)("media")(f)("format")
									End If
								End If
							End If
						Next
						REM If Len(format) <> 0 Then format = Left(format, Len(format)-2)
						WriteLog "format=" & format
					End If
					tmp = ""
					If CurrentRelease.Exists("release-group") Then
						tmp = CurrentRelease("release-group")("primary-type")
						If format <> "" Then
							format = format & ", " & tmp
						Else
							format = tmp
						End If
						Set tmp = CurrentRelease("release-group")
						If tmp.Exists("secondary-types") Then
							For Each f In tmp("secondary-types")
								If CurrentRelease("release-group")("secondary-types")(f) <> "" Then
									format = format & ", " & CurrentRelease("release-group")("secondary-types")(f)
								End If
							Next
							WriteLog "format=" & format
						End If
					End If

					If CurrentRelease.Exists("country") And Not IsNull(CurrentRelease("country")) Then
						country = CurrentRelease("country")
						For f = 1 To CountryCode.Count
							If country = CountryCode.Item(f) Then
								country = CountryList.Item(f)
								Exit For
							End If
						Next
						WriteLog "country=" & country
					End If
					If CurrentRelease.Exists("date") Then
						v_year = response(ArrayName)(r)("date")
						If Len(v_year) > 4 Then v_year = Left(v_year, 4)
						WriteLog "year=" & v_year
					End If

					If CurrentRelease.Exists("label-info") Then
						For Each l In CurrentRelease("label-info")
							Set currentLabel = CurrentRelease("label-info")(l)
							If currentLabel.Exists("label") Then
								If Not IsNull(currentLabel("label")("name")) Then
									If label <> "" Then
										If Left(label, Len(label)-2) <> CleanArtistName(currentLabel("label")("name")) Then
											label = label & CleanArtistName(currentLabel("label")("name")) & ", "
										End If
									Else
										label = CleanArtistName(currentLabel("label")("name")) & ", "
									End If
								End If
							End If
							If Not IsNull(currentLabel("catalog-number")) Then
								catNo = currentLabel("catalog-number")
								WriteLog "catNo=" & catNo
							End If
						Next
						If Len(label) <> 0 Then label = Left(label, Len(label)-2)
						WriteLog "label=" & label
					End If

					If CurrentRelease.Exists("track-count") And Not IsNull(CurrentRelease("track-count")) Then
						trackCount = CurrentRelease("track-count")
					Else
						trackCount = ""
					End If

					ReleaseDesc = ""
					Do
						If FilterMediaType = "Use MediaType Filter" And Format <> "" Then
							FilterFound = False
							For a = 1 To MediaTypeList.Count - 1
								If InStr(Format, MediaTypeList.Item(a)) <> 0 And MediaTypeFilterList.Item(a) = "1" Then FilterFound = True
							Next
							If FilterFound = False Then Exit Do
						End If
						If(FilterMediaType <> "None" And FilterMediaType <> "Use MediaType Filter" And InStr(format, FilterMediaType) = 0 And format <> "") Then Exit Do

						If FilterMediaFormat = "Use MediaFormat Filter" And format <> "" Then
							FilterFound = False
							For a = 1 To MediaFormatList.Count - 1
								If InStr(format, MediaFormatList.Item(a)) <> 0 And MediaFormatFilterList.Item(a) = "1" Then FilterFound = True
							Next
							If FilterFound = False Then Exit Do
						End If
						If(FilterMediaFormat <> "None" And FilterMediaFormat <> "Use MediaFormat Filter" And InStr(format, FilterMediaFormat) = 0 And Format <> "") Then Exit Do

						If FilterCountry = "Use Country Filter" And country <> "" Then
							FilterFound = False
							For a = 1 To CountryList.Count - 1
								If InStr(country, CountryList.Item(a)) <> 0 And CountryFilterList.Item(a) = "1" Then FilterFound = True
							Next
							If FilterFound = False Then Exit Do
						End If
						If(FilterCountry <> "None" And FilterCountry <> "Use Country Filter" And InStr(country, FilterCountry) = 0 And country <> "") Then Exit Do

						If FilterYear = "Use Year Filter" And v_year <> "" Then
							FilterFound = False
							For a = 1 To YearList.Count - 1
								If InStr(v_year, YearList.Item(a)) <> 0 And YearFilterList.Item(a) = "1" Then FilterFound = True
							Next
							If FilterFound = False Then Exit Do
						End If
						If(FilterYear <> "None" And FilterYear <> "Use Year Filter" And InStr(v_year, FilterYear) = 0 And v_year <> "") Then Exit Do

						If Rtype = "master" Then ReleaseDesc = ReleaseDesc & "* " End If
						If artist <> "" Then ReleaseDesc = ReleaseDesc & " " & artist End If
						If artist <> "" and title <> "" Then ReleaseDesc = ReleaseDesc & " -" End If
						If title <> "" Then ReleaseDesc = ReleaseDesc & " " & title End If
						If trackCount <> "" Then ReleaseDesc = ReleaseDesc & ", TrackCount: " & trackCount End If
						If format <> "" Then ReleaseDesc = ReleaseDesc & " [" & format & "]" End If
						If country <> "" Then ReleaseDesc = ReleaseDesc & " " & country End If
						If v_year <> "" Then ReleaseDesc = ReleaseDesc & " (" & v_year & ")" End If
						If catNo <> "" Then ReleaseDesc = ReleaseDesc & " catNo:" & catNo End If
						If ArrayName <> "Label" Then
							If label <> "" Then ReleaseDesc = ReleaseDesc & " / " & label End If
						End If

						Results.Add ReleaseDesc
						ResultsReleaseID.Add response(ArrayName)(r)("id")
					Loop While False

					SongCount = SongCount + 1
				Next
				ListCount = 1
				For r = 1 to Results.Count
					If r= 1 and isRelease = True Then
						Results.Item(0) = "(" & SongCountMax & ") " & Results.Item(0)
					Else
						If SongCount <> SongCountMax Then
							Results.Item(r-1) = "(" & ListCount & "/" & SongCount & "/" & SongCountMax & ") " & Results.Item(r-1)
						Else
							Results.Item(r-1) = "(" & ListCount & "/" & SongCountMax & ") " & Results.Item(r-1)
						End If
						ListCount = ListCount + 1
					End If
				Next
			End If
		ElseIf oXMLHTTP.Status = 503 Then
			ErrorMessage = "Status:" & oXMLHTTP.Status & " - The number of requests exceeds the limit. Please try again later"
			WriteLog "Status=" & oXMLHTTP.Status
		Else
			ErrorMessage = "Status:" & oXMLHTTP.Status & " - Please try again later"
			WriteLog "Status=" & oXMLHTTP.Status
		End If

	End If

	If QueryPage = "Discogs" Then

		' use json api with vbsjson class at start of file now

		If useOAuth = True Then
			oXMLHTTP.Open "POST", "https://www.germanc64.de/mm/oauth/check_new_v2.php", False
			oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			oXMLHTTP.setRequestHeader "User-Agent",UserAgent
			If LimitReleases = 50 Then SendPerPage = 50
			WriteLog "Sending Post at=" & AccessToken & "&ats=" & AccessTokenSecret & "&artist=" & SendArtist & "&album=" & SendAlbum & "&track=" & SendTrack & "&type=" & SendType & "&dbsearch=" & SendDBSearch & "&perpage=" & SendPerPage & "&querypage=" & QueryPage & "&page=1" & "&version=" & VersionStr
			oXMLHTTP.send ("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&artist=" & SendArtist & "&album=" & SendAlbum & "&track=" & SendTrack & "&type=" & SendType & "&dbsearch=" & SendDBSearch & "&perpage=" & SendPerPage & "&querypage=" & QueryPage & "&page=1" & "&version=" & VersionStr)

			If oXMLHTTP.Status = 200 Then
				If InStr(oXMLHTTP.responseText, "OAuth client error") <> 0 Then
					WriteLog "responseText=" & oXMLHTTP.responseText
					ErrorMessage = oXMLHTTP.responseText
					Exit Function
				Else
					WriteLog "responseText=" & oXMLHTTP.responseText
					Set response = json.Decode(oXMLHTTP.responseText)
				End If
			Else
				WriteLog "Error in JSONParser_find_result"
				WriteLog "Returned StatusCode=" & oXMLHTTP.Status
				ErrorMessage = "Problem with fetching data from Discogs  -  StatusCode=" & oXMLHTTP.Status
			End If
		Else

			oXMLHTTP.Open "GET", searchURL, False
			oXMLHTTP.setRequestHeader "Content-Type","application/json"
			oXMLHTTP.setRequestHeader "User-Agent",UserAgent
			oXMLHTTP.send()

			If oXMLHTTP.Status = 200 Then
				WriteLog "responseText=" & oXMLHTTP.responseText
				Set response = json.Decode(oXMLHTTP.responseText)
			Else
				WriteLog "Error in JSONParser_find_result"
				WriteLog "Returned StatusCode=" & oXMLHTTP.Status
				ErrorMessage = "Problem with fetching data from Discogs  -  StatusCode=" & oXMLHTTP.Status
			End If
		End If

		If ErrorMessage = "" Then
			SongCount = 0
			SongCountMax = response("pagination")("items")

			WriteLog "SongCountMax=" & SongCountMax

			If Int(SongCountMax) = 0 Then
				WriteLog "No Release found at Discogs !!"
				Results.Add "No Release found at Discogs !! Please change the search parameter in 'Manual search'."
				ResultsReleaseID.Add 0
			Else
				isRelease = False
				If Results.Count = 1 Then isRelease = True
				SongPages = response("pagination")("pages")
				WriteLog "SongPages=" & SongPages
				For Page = 1 to SongPages
					If Page <> 1 Then
						oXMLHTTP.Open "POST", "https://www.germanc64.de/mm/oauth/check_new_v2.php", False
						oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"  
						oXMLHTTP.setRequestHeader "User-Agent",UserAgent
						WriteLog "Sending Post at=" & AccessToken & "&ats=" & AccessTokenSecret & "&artist=" & SendArtist & "&album=" & SendAlbum & "&track=" & SendTrack & "&type=" & SendType & "&dbsearch=" & SendDBSearch & "&perpage=" & SendPerPage & "&querypage=" & QueryPage & "&page=" & Page & "&version=" & VersionStr
						oXMLHTTP.send ("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&artist=" & SendArtist & "&album=" & SendAlbum & "&track=" & SendTrack & "&type=" & SendType & "&dbsearch=" & SendDBSearch & "&perpage=" & SendPerPage & "&querypage=" & QueryPage & "&page=" & Page & "&version=" & VersionStr)
						If oXMLHTTP.Status = 200 Then
							If InStr(oXMLHTTP.responseText, "OAuth client error") <> 0 Then
								WriteLog "responseText=" & oXMLHTTP.responseText
								ErrorMessage = oXMLHTTP.responseText
								Exit Function
							Else
								Set response = json.Decode(oXMLHTTP.responseText)
							End If
						Else
							WriteLog "Error in JSONParser_find_result"
							ErrorMessage = "Error in JSONParser_find_result"
							Exit For
						End If
					End If

					For Each r In response(ArrayName)
						format = ""
						title = ""
						country = ""
						v_year = ""
						artist = ""
						label = ""
						Rtype = ""
						catNo = ""
						main_release = ""

						SDB.ProcessMessages

						title = response(ArrayName)(r)("title")
						Set tmp = response(ArrayName)(r)
						If tmp.Exists("artist") Then
							artist = tmp("artist")
						End If
						If tmp.Exists("main_release") Then
							main_release = tmp("main_release")
						End If
						If ArrayName = "results" Then
							If tmp.Exists("format") Then
								For Each f In response(ArrayName)(r)("format")
									AddToFieldWD format, response(ArrayName)(r)("format")(f)
									' format = format & response(ArrayName)(r)("format")(f) & ", "
								Next
								' If Len(format) <> 0 Then format = Left(format, Len(format)-2)
							End If
						Else
							If tmp.Exists("format") Then
								format = response(ArrayName)(r)("format")
							End If
						End If

						If tmp.Exists("country") Then
							country = response(ArrayName)(r)("country")
						End If
						If ArrayName = "versions" Then
							If tmp.Exists("released") Then
								v_year = response(ArrayName)(r)("released")
							End If
						Else
							If tmp.Exists("year") Then
								v_year = response(ArrayName)(r)("year")
							End If
						End If
						If tmp.Exists("catno") Then
							catNo = response(ArrayName)(r)("catno")
						End If
						If tmp.Exists("type") Then
							Rtype = response(ArrayName)(r)("type")
						End If
						If ArrayName = "results" Then
							If tmp.Exists("label") Then
								For Each f In response(ArrayName)(r)("label")
									' If label <> "" Then
										' If Left(label, Len(label)-2) <> response(ArrayName)(r)("label")(f) Then
											' label = label & response(ArrayName)(r)("label")(f) & ", "
										' End If
									' Else
										' label = response(ArrayName)(r)("label")(f) & ", "
									' End If
									AddToFieldWD label, CleanArtistName(response(ArrayName)(r)("label")(f))
								Next
								' If Len(label) <> 0 Then label = Left(label, Len(label)-2)
							End If
						Else
							If tmp.Exists("label") Then label = response(ArrayName)(r)("label")
						End If
						ReleaseDesc = ""
						Do
							If FilterMediaType = "Use MediaType Filter" And Format <> "" Then
								FilterFound = False
								For a = 1 To MediaTypeList.Count - 1
									If InStr(Format, MediaTypeList.Item(a)) <> 0 And MediaTypeFilterList.Item(a) = "1" Then FilterFound = True
								Next
								If FilterFound = False Then Exit Do
							End If
							If(FilterMediaType <> "None" And FilterMediaType <> "Use MediaType Filter" And InStr(format, FilterMediaType) = 0 And format <> "") Then Exit Do

							If FilterMediaFormat = "Use MediaFormat Filter" And format <> "" Then
								FilterFound = False
								For a = 1 To MediaFormatList.Count - 1
									If InStr(format, MediaFormatList.Item(a)) <> 0 And MediaFormatFilterList.Item(a) = "1" Then FilterFound = True
								Next
								If FilterFound = False Then Exit Do
							End If
							If(FilterMediaFormat <> "None" And FilterMediaFormat <> "Use MediaFormat Filter" And InStr(format, FilterMediaFormat) = 0 And Format <> "") Then Exit Do

							If FilterCountry = "Use Country Filter" And country <> "" Then
								FilterFound = False
								For a = 1 To CountryList.Count - 1
									If InStr(country, CountryList.Item(a)) <> 0 And CountryFilterList.Item(a) = "1" Then FilterFound = True
								Next
								If FilterFound = False Then Exit Do
							End If
							If(FilterCountry <> "None" And FilterCountry <> "Use Country Filter" And InStr(country, FilterCountry) = 0 And country <> "") Then Exit Do

							If FilterYear = "Use Year Filter" And v_year <> "" Then
								FilterFound = False
								For a = 1 To YearList.Count - 1
									If InStr(v_year, YearList.Item(a)) <> 0 And YearFilterList.Item(a) = "1" Then FilterFound = True
								Next
								If FilterFound = False Then Exit Do
							End If
							If(FilterYear <> "None" And FilterYear <> "Use Year Filter" And InStr(v_year, FilterYear) = 0 And v_year <> "") Then Exit Do

							If Rtype = "master" Then ReleaseDesc = ReleaseDesc & "* " End If
							If artist <> "" Then ReleaseDesc = ReleaseDesc & " " & artist End If
							If artist <> "" and title <> "" Then ReleaseDesc = ReleaseDesc & " -" End If
							If title <> "" Then ReleaseDesc = ReleaseDesc & " " & title End If
							If format <> "" Then ReleaseDesc = ReleaseDesc & " [" & format & "]" End If
							If country <> "" Then ReleaseDesc = ReleaseDesc & " " & country End If
							If v_year <> "" Then ReleaseDesc = ReleaseDesc & " (" & v_year & ")" End If
							If catNo <> "" Then ReleaseDesc = ReleaseDesc & " catNo:" & catNo End If
							If label <> "" Then ReleaseDesc = ReleaseDesc & " / " & label End If

							Results.Add ReleaseDesc
							ResultsReleaseID.Add response(ArrayName)(r)("id")
						Loop While False
						SongCount = SongCount + 1
						If SongCount = LimitReleases Then Exit For
					Next
					If SongCount = LimitReleases Then Exit For
				Next
				ListCount = 1
				For r = 1 to Results.Count
					If r= 1 and isRelease = True Then
						Results.Item(0) = "(" & SongCountMax & ") " & Results.Item(0)
					Else
						If SongCount <> SongCountMax Then
							Results.Item(r-1) = "(" & ListCount & "/" & SongCount & "/" & SongCountMax & ") " & Results.Item(r-1)
						Else
							Results.Item(r-1) = "(" & ListCount & "/" & SongCountMax & ") " & Results.Item(r-1)
						End IF
						ListCount = ListCount + 1
					End If
				Next
			End If
		End If
	End If
	Set oXMLHTTP = Nothing
	WriteLog "End JSONParser_find_result"
	
End Function


Function ReloadMaster(SavedMasterID)

	WriteLog "Start ReloadMaster"
	Dim oXMLHTTP, masterURL
	masterURL = "https://api.discogs.com/masters/" & SavedMasterID
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")

	Dim json
	Set json = New VbsJson
	Dim response

	oXMLHTTP.open "GET", masterURL, false
	oXMLHTTP.setRequestHeader "Content-Type","application/json"
	oXMLHTTP.setRequestHeader "User-Agent",UserAgent
	oXMLHTTP.send()

	If oXMLHTTP.Status = 200 Then
		WriteLog "responseText=" & oXMLHTTP.responseText
		If InStr(oXMLHTTP.responseText, "OAuth client error") <> 0 Then	
			ErrorMessage = oXMLHTTP.responseText
			Exit Function
		Else
			Set response = json.Decode(oXMLHTTP.responseText)
			If response.Exists("year") Then
				OriginalDate = response("year")
			Else
				OriginalDate = ""
			End If
			If OriginalDate = 0 Then OriginalDate = ""
		End If
	End If
	Set oXMLHTTP = Nothing
	ReloadMaster = OriginalDate
	WriteLog "Stop ReloadMaster"

End Function



Function get_release_ID(FirstTrack)

	CurrentReleaseId = ""
	If ReleaseTag = "Custom1" Then CurrentReleaseId = FirstTrack.Custom1
	If ReleaseTag = "Custom2" Then CurrentReleaseId = FirstTrack.Custom2
	If ReleaseTag = "Custom3" Then CurrentReleaseId = FirstTrack.Custom3
	If ReleaseTag = "Custom4" Then CurrentReleaseId = FirstTrack.Custom4
	If ReleaseTag = "Custom5" Then CurrentReleaseId = FirstTrack.Custom5
	If ReleaseTag = "Grouping" Then CurrentReleaseId = FirstTrack.Grouping
	If ReleaseTag = "ISRC" Then CurrentReleaseId = FirstTrack.ISRC
	If ReleaseTag = "Encoder" Then CurrentReleaseId = FirstTrack.Encoder
	If ReleaseTag = "Copyright" Then CurrentReleaseId = FirstTrack.Copyright

	get_release_ID = CurrentReleaseId

End Function



Sub Unselect()

	Dim templateHTMLDoc, i, checkBox

	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
	button.style.backgroundColor = "#ED1C24"
	templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML="Please wait for collection data..."
	
	For i=0 To iMaxTracks - 1
		If Grouping(i) <> "|Heading|" Then
			Set checkBox = templateHTMLDoc.getElementById("unselected["&i&"]")
			If checkBox.Checked Then
				UnselectedTracks(i) = ""
			Else
				UnselectedTracks(i) = "x"
			End If
			Set checkBox = templateHTMLDoc.getElementById("unselectedtrackname["&i&"]")
			If checkBox.Checked Then
				UnselectedTrackNames(i) = ""
			Else
				UnselectedTrackNames(i) = "x"
			End If
		End If
	Next

	ReloadResults

End Sub


Sub ShowCountryFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, countrybutton, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 875
	Form.Common.Height = 600
	Form.FormPosition = 4
	Form.Caption = "Choose the country's to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = CountryList

	If ((CountryList.Count - 1) / 65) = Int((CountryList.Count - 1) / 65) Then
		CountColumn = (CountryList.Count - 1) / 65
	Else
		CountColumn = Int((CountryList.Count - 1) / 65) + 1
	End If
	iWidth = (CountryList.Count - 1) / 65 * 200
	filterHTML = GetFilterHTML(iWidth, 64, CountColumn)

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "BtnCheckFilterClick"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "BtnUnCheckFilterClick"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To CountryList.Count - 1
		Set countrybutton = filterHTMLDoc.getElementById("Filter" & i)
		If CountryFilterList.Item(i) = "1" Then
			countrybutton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To CountryList.Count - 1
			Set countrybutton = filterHTMLDoc.getElementById("Filter" & a)
			If countrybutton.checked = True Then
				CountryFilterList.Item(a) = "1"
				FilterFound = True
			Else
				CountryFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterCountry = "None"
			CountryFilterList.Item(0) = "0"
		Else
			FilterCountry = "Use Country Filter"
			CountryFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults NewSearchTerm, QueryPage
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub ShowMediaFormatFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, MediaFormatButton, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 380
	Form.Common.Height = 700
	Form.FormPosition = 4
	Form.Caption = "Choose the MediaFormat to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = MediaFormatList
	If ((MediaFormatList.Count - 1) / 24) = Int((MediaFormatList.Count - 1) / 24) Then
		CountColumn = (MediaFormatList.Count - 1) / 24
	Else
		CountColumn = Int((MediaFormatList.Count - 1) / 24) + 1
	End If
	iWidth = (MediaFormatList.Count - 1) / 24 * 150

	filterHTML = GetFilterHTML(iWidth, 23, CountColumn)

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 20
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 5
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "BtnCheckFilterClick"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "BtnUnCheckFilterClick"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To MediaFormatList.Count - 1
		Set MediaFormatButton = filterHTMLDoc.getElementById("Filter" & i)
		If MediaFormatFilterList.Item(i) = "1" Then
			MediaFormatButton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To MediaFormatList.Count - 1
			Set MediaFormatButton = filterHTMLDoc.getElementById("Filter" & a)
			If MediaFormatButton.checked = True Then
				MediaFormatFilterList.Item(a) = "1"
				FilterFound = True
			Else
				MediaFormatFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterMediaFormat = "None"
			MediaFormatFilterList.Item(0) = "0"
		Else
			FilterMediaFormat = "Use MediaFormat Filter"
			MediaFormatFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults NewSearchTerm, QueryPage
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub ShowMediaTypeFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, MediaTypeButton, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 420
	Form.Common.Height = 600
	Form.FormPosition = 4
	Form.Caption = "Choose the MediaType to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = MediaTypeList
	If ((MediaTypeList.Count - 1) / 19) = Int((MediaTypeList.Count - 1) / 19) Then
		CountColumn = (MediaTypeList.Count - 1) / 19
	Else
		CountColumn = Int((MediaTypeList.Count - 1) / 19) + 1
	End If
	iWidth = (MediaTypeList.Count - 1) / 19 * 175
	filterHTML = GetFilterHTML(iWidth, 18, CountColumn)
	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "BtnCheckFilterClick"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "BtnUnCheckFilterClick"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To MediaTypeList.Count - 1
		Set MediaTypeButton = filterHTMLDoc.getElementById("Filter" & i)
		If MediaTypeFilterList.Item(i) = "1" Then
			MediaTypeButton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To MediaTypeList.Count - 1
			Set MediaTypeButton = filterHTMLDoc.getElementById("Filter" & a)
			If MediaTypeButton.checked = True Then
				MediaTypeFilterList.Item(a) = "1"
				FilterFound = True
			Else
				MediaTypeFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterMediaType = "None"
			MediaTypeFilterList.Item(0) = "0"
		Else
			FilterMediaType = "Use MediaType Filter"
			MediaTypeFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults NewSearchTerm, QueryPage
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub ShowYearFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, YearButton, FilterFound
	Dim i, a, row
	Set Form = UI.NewForm
	Form.Common.Width = 550
	Form.Common.Height = 550
	Form.FormPosition = 4
	Form.Caption = "Choose the Year to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = YearList
	'CountColumn = 6
	If ((YearList.Count - 1) / 6) = Int((YearList.Count - 1) / 6) Then
		iWidth = (YearList.Count - 1) / 6 * 25
		row = Int((YearList.Count - 1) / 6)
		CountColumn = 6
	Else
		row = Int((YearList.Count - 1) / 6) + 1
		iWidth = (YearList.Count - 1) / 6 * 25
		CountColumn = 6
	End If

	filterHTML = GetFilterHTML(iWidth, row-1, CountColumn)

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "BtnCheckFilterClick"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "BtnUnCheckFilterClick"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To YearList.Count - 1
		Set Yearbutton = filterHTMLDoc.getElementById("Filter" & i)
		If YearFilterList.Item(i) = "1" Then
			Yearbutton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To YearList.Count - 1
			Set Yearbutton = filterHTMLDoc.getElementById("Filter" & a)
			If Yearbutton.checked = True Then
				YearFilterList.Item(a) = "1"
				FilterFound = True
			Else
				YearFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterYear = "None"
			YearFilterList.Item(0) = "0"
		Else
			FilterYear = "Use Year Filter"
			YearFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults NewSearchTerm, QueryPage
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub BtnCheckFilterClick

	Dim WebBrowser2, FilterList, filterHTMLDoc, a, filterbutton
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set FilterList = SDB.Objects("Filter")
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	For a = 1 To FilterList.Count - 1
		Set filterbutton = filterHTMLDoc.getElementById("Filter" & a)
		filterbutton.checked = True
	Next

End Sub


Sub BtnUnCheckFilterClick

	Dim WebBrowser2, FilterList, filterHTMLDoc, a, filterbutton
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set FilterList = SDB.Objects("Filter")
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	For a = 1 To FilterList.Count - 1
		Set filterbutton = filterHTMLDoc.getElementById("Filter" & a)
		filterbutton.checked = False
	Next

End Sub


Function GetFilterHTML(Width, Row, CountColumn)

	Dim FilterList, filterHTML, i, a
	Set FilterList = SDB.Objects("Filter")
	filterHTML = "<HTML>"
	filterHTML = filterHTML & "<HEAD>"
	filterHTML = filterHTML & "<style type=""text/css"" media=""screen"">"
	filterHTML = filterHTML & ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	filterHTML = filterHTML & "</style>"
	filterHTML = filterHTML & "</HEAD>"
	filterHTML = filterHTML & "<table border=0 width=" & Width & " cellspacing=0 cellpadding=1 class=tabletext>"

	For i = 0 To Row
		filterHTML = filterHTML &  "<tr>"
		For a = 1 To CountColumn
			If FilterList.Count <= a + (i * CountColumn) Then Exit For
			filterHTML = filterHTML &  "<td><input type=checkbox id=""Filter" & a + (i * CountColumn) & """ >" & FilterList.Item(a + (i * CountColumn))
			filterHTML = filterHTML &  "</td>"
		Next
		filterHTML = filterHTML &  "</tr>"
	Next
	filterHTML = filterHTML &  "</table>"
	filterHTML = filterHTML &  "</body>"
	filterHTML = filterHTML &  "</HTML>"
	GetFilterHTML = filterHTML

End Function


Sub MoreImages()

	Dim iWidth, imageHTML, imageHTMLDoc, i, j, RndFileName, ret
	imageHTML = "<HTML>"
	imageHTML = imageHTML &  "<HEAD>"
	imageHTML = imageHTML &  "<style type=""text/css"" media=""screen"">"
	imageHTML = imageHTML &  ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	imageHTML = imageHTML &  "</style>"
	imageHTML = imageHTML &  "</HEAD>"
	imageHTML = imageHTML &  "<body bgcolor=""#FFFFFF"">"
	iWidth = (ImageList.Count - 1) * 200
	imageHTML = imageHTML &  "<table border=0 width=" & iWidth & " cellspacing=0 cellpadding=1 class=tabletext>"
	imageHTML = imageHTML &  "<tr>"

	Dim max, min
	Set ImageLocal = SDB.NewStringList
	For i = 0 To ImageList.Count - 1
		max=100000
		min=10000
		Randomize
		RndFileName = Int((max-min+1)*Rnd+min) & ".jpg"
		ret = getimages(ImageList.Item(i), sTemp & RndFileName)
		ImageLocal.Add sTemp & RndFileName
		WriteLog "Random-File = " & sTemp & RndFileName
		imageHTML = imageHTML &  "<td><img border=""0"" src=""" & sTemp & RndFileName & """ width=""180"" height=""180""></td>"
	Next
	imageHTML = imageHTML &  "</tr><tr>"
	For i = 0 To ImageList.Count - 1
		imageHTML = imageHTML &  "<td>"
		imageHTML = imageHTML &  "<select id=""ImageType" & i & """ class=tabletext>"
		For j = 0 To ImageTypeList.Count - 1
			If SaveImageType.Item(i) <> ImageTypeList.Item(j) Then
				imageHTML = imageHTML &  "<option value=""" & ImageTypeList.Item(j) & """>" & ImageTypeList.Item(j) & "</option>"
			Else
				imageHTML = imageHTML &  "<option value=""" & ImageTypeList.Item(j) & """ selected>" & ImageTypeList.Item(j) & "</option>"
			End If
		Next
		imageHTML = imageHTML &  "</select></td>"
	Next

	imageHTML = imageHTML &  "</tr><tr>"
	For i = 0 To ImageList.Count - 1
		imageHTML = imageHTML &  "<td><input type=checkbox id=""SaveImage" & i & """ >Save Image"
		imageHTML = imageHTML &  "</td>"
	Next

	imageHTML = imageHTML &  "</tr><tr>"
	For i = 0 To ImageList.Count - 1
		If CoverStorage = 1 Or CoverStorage = 3 Then
			imageHTML = imageHTML &  "<td><input type=text id=""FileName" & i & """ >"
		Else
			imageHTML = imageHTML &  "<td>Store in tag"
		End If
		imageHTML = imageHTML &  "</td>"
	Next

	imageHTML = imageHTML &  "</tr>"

	imageHTML = imageHTML &  "</table>"
	imageHTML = imageHTML &  "</body>"
	imageHTML = imageHTML &  "</HTML>"



	Dim Form
	Set Form = UI.NewForm 
	Form.Common.ClientWidth = 800
	Form.Common.ClientHeight = 400
	Form.FormPosition = 4
	Form.Caption = "Choose additional Images for the Release"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("ImageForm") = Form

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width -30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True  

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width -5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1 
	Btn2.Default = True
	
	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Common.ControlName = "Btn3"
	Btn3.Caption = SDB.Localize("Select All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = Btn2.Common.Left - Btn3.Common.Width -5
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "SelectAllClick"

	Set WebBrowser3 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser3.Common.Align = 5      ' Fill whole client rectangle
	WebBrowser3.Common.ControlName = "WebBrowser3"
	WebBrowser3.Common.Top = 100
	WebBrowser3.Common.Left = 100

	SDB.Objects("WebBrowser3") = WebBrowser3
	WebBrowser3.Interf.Visible = True
	WebBrowser3.Common.BringToFront

	WebBrowser3.SetHTMLDocument imageHTML
	Set imageHTMLDoc = WebBrowser3.Interf.Document

	Dim saveimagebutton, text
	For i = 0 To ImageList.Count - 1
		Set saveimagebutton = imageHTMLDoc.getElementById("SaveImage" & i)
		If SaveImage.Item(i) = 1 Then
			saveimagebutton.checked = 1
		End If
		If CoverStorage = 1 Or CoverStorage = 3 Then
			Set text = imageHTMLDoc.getElementById("FileName" & i)
			text.value = "folder" & i & ".jpg"
		End If
	Next

	If Form.ShowModal = 1 Then
		SetSaveImages()
	End If

	WebBrowser3.SetHTMLDocument ""
	SDB.Objects("ImageForm") = Nothing
	WebBrowser3.Common.DestroyControl      ' Destroy the external control
	Set WebBrowser3 = Nothing              ' Release global variable
	SDB.Objects("WebBrowser3") = Nothing

End Sub


Sub SelectAllClick()

	Dim imageHTMLDoc, saveimagebutton, i
	Set WebBrowser3 = SDB.Objects("WebBrowser3")
	Set imageHTMLDoc = WebBrowser3.Interf.Document

	For i = 0 To ImageList.Count - 1
		Set saveimagebutton = imageHTMLDoc.getElementById("SaveImage" & i)
		saveimagebutton.checked = 1
	Next

End Sub


Sub SetSaveImages()

	Dim imageHTMLDoc, i, checkbox, listbox, text

	Set SaveImage = SDB.NewStringList
	Set SaveImageType = SDB.NewStringList
	Set FileNameList = SDB.NewStringList

	Set imageHTMLDoc = WebBrowser3.Interf.Document
	For i = 0 To ImageList.Count - 1
		Set checkBox = imageHTMLDoc.getElementById("SaveImage" & i)
		If checkBox.Checked Then
			SaveImage.Add "1"
			Set listBox = imageHTMLDoc.getElementById("ImageType" & i)
			SaveImageType.Add listBox.Value
			If CoverStorage = 1 Or CoverStorage = 3 Then
				Set text = imageHTMLDoc.getElementById("FileName" & i)
				FileNameList.Add text.Value
			Else
				FileNameList.Add "nothing"
			End If
		Else
			SaveImage.Add "0"
			SaveImageType.Add "other"
			FileNameList.Add "nothing"
		End If
	Next

	SaveMoreImagesSub()

End Sub


Function getimages(DownloadDest, LocalFile)

	Dim oXMLHTTP, objStream
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	SDB.ProcessMessages
	WriteLog "Start getimages"
	WriteLog "LocalFile=" & LocalFile
	WriteLog "User-Agent=" & UserAgent
	oXMLHTTP.open "GET", DownloadDest, False
	oXMLHTTP.setRequestHeader "Accept", "*/*"
	oXMLHTTP.setRequestHeader "User-Agent", UserAgent
	oXMLHTTP.setRequestHeader "Host", "img.discogs.com"
	oXMLHTTP.send()
	If oXMLHTTP.Status = 200 Then
		Set objStream = CreateObject("ADODB.Stream")
		objStream.Type = 1
		objStream.Open
		objStream.Write oXMLHTTP.ResponseBody
		objStream.SavetoFile LocalFile, 2
		objStream.Close
		Set objStream = Nothing
		Set oXMLHTTP = Nothing
		getimages = LocalFile
	Else
		WriteLog "Status=" & oXMLHTTP.Status
		SDB.MessageBox "Getting cover image failed - Status-Code = " & oXMLHTTP.Status, mtInformation, Array(mbOk)
		Set oXMLHTTP = Nothing
		getimages = ""
	End If

End function


Sub WriteOptions()

	WriteLog " "
	WriteLog "Options:"
	WriteLog "CheckAlbum=" & CheckAlbum
	WriteLog "CheckArtist=" & CheckArtist
	WriteLog "CheckAlbumArtist=" & CheckAlbumArtist
	WriteLog "CheckAlbumArtistFirst=" & CheckAlbumArtistFirst
	WriteLog "CheckLabel=" & CheckLabel
	WriteLog "CheckDate=" & CheckDate
	WriteLog "CheckOrigDate=" & CheckOrigDate
	WriteLog "CheckGenre=" & CheckGenre
	WriteLog "CheckCountry=" & CheckCountry
	WriteLog "CheckSaveImage=" & CheckSaveImage
	WriteLog "CheckSmallCover=" & CheckSmallCover
	WriteLog "CheckCatalog=" & CheckCatalog
	WriteLog "CheckRelease=" & CheckRelease
	WriteLog "CheckInvolved=" & CheckInvolved
	WriteLog "CheckGrouping=" & CheckGrouping
	WriteLog "CheckLyricist=" & CheckLyricist
	WriteLog "CheckComposer=" & CheckComposer
	WriteLog "CheckConductor=" & CheckConductor
	WriteLog "CheckProducer=" & CheckProducer
	WriteLog "CheckDiscNum=" & CheckDiscNum
	WriteLog "CheckTrackNum=" & CheckTrackNum
	WriteLog "CheckFormat=" & CheckFormat
	WriteLog "CheckUseAnv=" & CheckUseAnv
	WriteLog "CheckYearOnlyDate=" & CheckYearOnlyDate
	WriteLog "CheckForceNumeric=" & CheckForceNumeric
	WriteLog "CheckSidesToDisc=" & CheckSidesToDisc
	WriteLog "CheckForceDisc=" & CheckForceDisc
	WriteLog "CheckOriginalDiscogsTrack=" & CheckOriginalDiscogsTrack
	WriteLog "CheckNoDisc=" & CheckNoDisc
	REM WriteLog "CheckUserCollection=" & CheckUserCollection
	WriteLog "DiscogsUsername=" & DiscogsUsername
	WriteLog "ReleaseTag=" & ReleaseTag
	WriteLog "CatalogTag=" & CatalogTag
	WriteLog "CountryTag=" & CountryTag
	WriteLog "FormatTag=" & FormatTag
	WriteLog "CheckLeadingZero=" & CheckLeadingZero
	WriteLog "CheckLeadingZeroDisc=" & CheckLeadingZeroDisc
	WriteLog "CheckVarious=" & CheckVarious
	WriteLog "TxtVarious=" & TxtVarious
	WriteLog "CheckTitleFeaturing=" & CheckTitleFeaturing
	WriteLog "CheckFeaturingName=" & CheckFeaturingName
	WriteLog "TxtFeaturingName=" & TxtFeaturingName
	WriteLog "CheckComment=" & CheckComment
	WriteLog "SubTrackNameSelection=" & SubTrackNameSelection
	WriteLog "Separator=" & Separator
	Set ini = SDB.IniFile
	WriteLog "CurrentCountryFilter=" & ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter")
	WriteLog "CurrentMediaTypeFilter=" & ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter")
	WriteLog "CurrentMediaFormatFilter=" & ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter")
	WriteLog "CurrentYearFilter=" & ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter")
	WriteLog "LyricistKeywords=" & LyricistKeywords
	WriteLog "ConductorKeywords=" & ConductorKeywords
	WriteLog "ProducerKeywords=" & ProducerKeywords
	WriteLog "ComposerKeywords=" & ComposerKeywords
	WriteLog "FeaturingKeywords=" & FeaturingKeywords
	WriteLog "UnwantedKeywords=" & UnwantedKeywords
	WriteLog "ArtistSeparator=" & ArtistSeparator
	WriteLog "ArtistLastSeparator=" & ArtistLastSeparator
	WriteLog "FormatSeparator=" & FormatSeparator
	WriteLog "CheckTurnOffSubTrack=" & CheckTurnOffSubTrack
	WriteLog "AccessToken=" & AccessToken
	WriteLog "AccessTokenSecret=" & AccessTokenSecret
	WriteLog "QueryPage=" & QueryPage
	WriteLog "CheckInvolvedPeopleSingleLine=" & CheckInvolvedPeopleSingleLine
	WriteLog "CheckDontFillEmptyFields=" & CheckDontFillEmptyFields
	WriteLog "CheckNewVersion=" & CheckNewVersion
	WriteLog "LastCheck=" & LastCheck
	WriteLog "CheckTheBehindArtist=" & CheckTheBehindArtist
	WriteLog "CheckDeleteDuplicatedEntry=" & CheckDeleteDuplicatedEntry
	WriteLog "StoreDate=" & StoreDate
	WriteLog "CheckLimitReleases=" & LimitReleases
	WriteLog "CheckIgnoreFeatArtist=" & CheckIgnoreFeatArtist
	WriteLog "SubTrackSeparator=" & SubTrackSeparator
	WriteLog "CheckNewMatchingEngine=" & CheckNewMatchingEngine
	WriteLog "CheckShowQtyFormat=" & CheckShowQtyFormat
	WriteLog "-+-+-+-+--+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-"

End Sub


Sub WriteLogInit

	Dim logdatei
	logdatei = SDB.ScriptsPath & "Discogs_Script.log"
	If SDB.Tools.FileSystem.FileExists(logdatei) = True Then
		SDB.Tools.FileSystem.DeleteFile(logdatei)
	End If
	WriteLog "Start DiscogsTagger " & VersionStr
	WriteOptions()

End Sub


Sub WriteLog(Text)

	Dim filesys, filetxt, logdatei, tmpText
	'Const ForReading = 1, ForWriting = 2, ForAppending = 8
	logdatei = SDB.ScriptsPath & "Discogs_Script.log"
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set filetxt = filesys.OpenTextFile(logdatei, 8, True)
	tmpText = Time & Chr(9) & SDB.ToAscii(Text)
	filetxt.WriteLine(tmpText)
	filetxt.Close

End Sub



'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-


Function searchKeyword(Keywords, Role, AlbumRole, artistName)

	Dim tmp, x, RE, searchPattern
	tmp = Split(Keywords, ",")
	Set RE = New RegExp
	RE.IgnoreCase = True
	For Each searchPattern In tmp
		If InStr(searchPattern, "*") <> 0 Then
			searchPattern = Replace(searchPattern, "*", ".*")
			RE.Pattern = "^" & searchPattern & "$"
			If RE.Test(Role) Then
				If InStr(AlbumRole, artistName) = 0 Then
					If AlbumRole = "" Then
						AlbumRole = artistName
					Else
						AlbumRole = AlbumRole & Separator & artistName
					End If
					searchKeyword = AlbumRole
				Else
					searchKeyword = "ALREADY_INSIDE_ROLE"
				End If
				WriteLog "searchPattern=" & searchPattern
				Exit For
			End If
		Else
			If Trim(LCase(Role)) = Trim(LCase(searchPattern)) Then
				If InStr(AlbumRole, artistName) = 0 Then
					If AlbumRole = "" Then
						AlbumRole = artistName
					Else
						AlbumRole = AlbumRole & Separator & artistName
					End If
					searchKeyword = AlbumRole
				Else
					searchKeyword = "ALREADY_INSIDE_ROLE"
				End If
				Exit For
			End If
		End If
	Next

End Function

Class VbsJson
	'Author: Demon
	'Date: 2012/5/3
	'Website: http://demon.tw
	Private Whitespace, NumberRegex, StringChunk
	Private b, f, r, n, t

	Private Sub Class_Initialize
		Whitespace = " " & vbTab & vbCr & vbLf
		b = ChrW(8)
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab

		Set NumberRegex = New RegExp
		NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
		NumberRegex.Global = False
		NumberRegex.MultiLine = True
		NumberRegex.IgnoreCase = True

		Set StringChunk = New RegExp
		StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
		StringChunk.Global = False
		StringChunk.MultiLine = True
		StringChunk.IgnoreCase = True
	End Sub

	'Return a JSON string representation of a VBScript data structure
	'Supports the following objects and types
	'+-------------------+---------------+
	'| VBScript          | JSON          |
	'+===================+===============+
	'| Dictionary        | object        |
	'+-------------------+---------------+
	'| Array             | array         |
	'+-------------------+---------------+
	'| String            | string        |
	'+-------------------+---------------+
	'| Number            | number        |
	'+-------------------+---------------+
	'| True              | true          |
	'+-------------------+---------------+
	'| False             | false         |
	'+-------------------+---------------+
	'| Null              | null          |
	'+-------------------+---------------+
	Public Function Encode(ByRef obj)
		Dim buf, i, c, g
		Set buf = CreateObject("Scripting.Dictionary")
		Select Case VarType(obj)
			Case vbNull
				buf.Add buf.Count, "null"
			Case vbBoolean
				If obj Then
					buf.Add buf.Count, "true"
				Else
					buf.Add buf.Count, "false"
				End If
			Case vbInteger, vbLong, vbSingle, vbDouble
				buf.Add buf.Count, obj
			Case vbString
				buf.Add buf.Count, """"
				For i = 1 To Len(obj)
					c = Mid(obj, i, 1)
					Select Case c
						Case """" buf.Add buf.Count, "\"""
						Case "\"  buf.Add buf.Count, "\\"
						Case "/"  buf.Add buf.Count, "/"
						Case b    buf.Add buf.Count, "\b"
						Case f    buf.Add buf.Count, "\f"
						Case r    buf.Add buf.Count, "\r"
						Case n    buf.Add buf.Count, "\n"
						Case t    buf.Add buf.Count, "\t"
						Case Else
							If AscW(c) >= 0 And AscW(c) <= 31 Then
								c = Right("0" & Hex(AscW(c)), 2)
								buf.Add buf.Count, "\u00" & c
							Else
								buf.Add buf.Count, c
							End If
					End Select
				Next
				buf.Add buf.Count, """"
			Case vbArray + vbVariant
				g = True
				buf.Add buf.Count, "["
				For Each i In obj
					If g Then g = False Else buf.Add buf.Count, ","
					buf.Add buf.Count, Encode(i)
				Next
				buf.Add buf.Count, "]"
			Case vbObject
				If TypeName(obj) = "Dictionary" Then
					g = True
					buf.Add buf.Count, "{"
					For Each i In obj
						If g Then g = False Else buf.Add buf.Count, ","
						buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
					Next
					buf.Add buf.Count, "}"
				Else
					Err.Raise 8732,,"None dictionary object"
				End If
			Case Else
				buf.Add buf.Count, """" & CStr(obj) & """"
		End Select
		Encode = Join(buf.Items, "")
	End Function

	'Return the VBScript representation of ``str(``
	'Performs the following translations in decoding
	'+---------------+-------------------+
	'| JSON          | VBScript          |
	'+===============+===================+
	'| object        | Dictionary        |
	'+---------------+-------------------+
	'| array         | Array             |
	'+---------------+-------------------+
	'| string        | String            |
	'+---------------+-------------------+
	'| number        | Double            |
	'+---------------+-------------------+
	'| true          | True              |
	'+---------------+-------------------+
	'| false         | False             |
	'+---------------+-------------------+
	'| null          | Null              |
	'+---------------+-------------------+
	Public Function Decode(ByRef str)
		'return base object
		Set Decode = ParseObject(str, 1)
	End Function

	Private Function ParseValue(ByRef str, ByRef idx)
		Dim c, ms

		idx = NextToken(str, idx)
		c = Mid(str, idx, 1)

		If c = "{" Then
			Set ParseValue = ParseObject(str, idx)
			Exit Function
		ElseIf c = "[" Then
			Set ParseValue = ParseArray(str, idx)
			Exit Function
		ElseIf c = """" Then
			idx = idx + 1
			ParseValue = ParseString(str, idx)
			Exit Function
		ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ParseValue = Null
			Exit Function
		ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ParseValue = True
			Exit Function
		ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
			idx = idx + 5
			ParseValue = False
			Exit Function
		Else
			Set ms = NumberRegex.Execute(Mid(str, idx))
			If ms.Count = 1 Then
				idx = idx + ms(0).Length
				SetLocale "en-US"
				ParseValue = CDbl(ms(0))
				SetLocale 0
				Exit Function
			End If
		End If

		Err.Raise 8732,,"No JSON object could be ParseValued"
	End Function

	Private Function ParseObject(ByRef str, ByRef idx)
		Dim c, key, value
		Set ParseObject = CreateObject("Scripting.Dictionary")

		idx = NextToken(str, idx)

		c = Mid(str, idx, 1)

		If c = "{" Then
			idx = NextToken(str,idx+1)
		Else
			Err.Raise 8732,,"Expected { to begin Object"
		End If

		c = Mid(str, idx, 1)

		Do
			If c <> """" And c <> "}" Then

				Err.Raise 8732,,"Expecting property name or } near: " & Mid(str,idx)

			ElseIf c = """" Then

				idx = idx + 1
				key = ParseString(str, idx)

				idx = NextToken(str, idx)
				If Mid(str, idx, 1) <> ":" Then
					Err.Raise 8732,,"Expecting : delimiter near: " & Mid(str,idx)
				End If

				' skip : and whitespace after key
				idx = NextToken(str, idx + 1)

				' check for object or array value
				If Mid(str,idx,1) = "{" Or Mid(str,idx,1) = "[" Then
					Set value = ParseValue(str, idx)
				Else
					value = ParseValue(str,idx)
				End If

				ParseObject.Add key, value
			End If

			c = Mid(str,idx,1)

			If c = "}" Then
				idx = NextToken(str,idx+1)
				Exit Function
			End If

			If c <> "," Then

				Err.Raise 8732,,"Expecting , delimiter near: " & Mid(str,idx)

			End If

			'skip , and whitespace after value
			idx = NextToken(str, idx + 1)
			c = Mid(str, idx, 1)
			If c <> """" Then
				Err.Raise 8732,,"Expecting property name"
			End If
		Loop
	End Function

	Private Function ParseArray(ByRef str, ByRef idx)
		Dim c, values, value
		Set ParseArray = CreateObject("Scripting.Dictionary")

		idx = NextToken(str, idx)
		c = Mid(str, idx, 1)

		If c = "[" Then
			idx = NextToken(str,idx+1)
		Else
			Err.Raise 8732,,"Expected [ to begin Array"
		End If

		Do
			c = Mid(str, idx, 1)

			If c = "]" Then
				idx = NextToken(str,idx+1)
				Exit Function
			End If

			ParseArray.Add ParseArray.Count, ParseValue(str, idx)

			c = Mid(str, idx, 1)

			If c = "]" Then
				idx = NextToken(str, idx+1)
				Exit Function
			End If

			If c <> "," Then
				Err.Raise 8732,,"Expecting , delimiter near: " & Mid(str,idx)
			End If

			idx = NextToken(str,idx+1)

		Loop
	End Function

	Private Function ParseString(ByRef str, ByRef idx)
		Dim chunks, content, terminator, ms, esc, char
		Set chunks = CreateObject("Scripting.Dictionary")

		Do
			Set ms = StringChunk.Execute(Mid(str, idx))
			If ms.Count = 0 Then
				Err.Raise 8732,,"Unterminated string starting"
			End If

			content = ms(0).Submatches(0)
			terminator = ms(0).Submatches(1)
			If Len(content) > 0 Then
				chunks.Add chunks.Count, content
			End If

			idx = idx + ms(0).Length

			If terminator = """" Then
				Exit Do
			ElseIf terminator <> "\" Then
				Err.Raise 8732,,"Invalid control character"
			End If

			esc = Mid(str, idx, 1)

			If esc <> "u" Then
				Select Case esc
					Case """" char = """"
					Case "\"  char = "\"
					Case "/"  char = "/"
					Case "b"  char = b
					Case "f"  char = f
					Case "n"  char = n
					Case "r"  char = r
					Case "t"  char = t
					Case Else Err.Raise 8732,,"Invalid escape"
				End Select
				idx = idx + 1
			Else
				char = ChrW("&H" & Mid(str, idx + 1, 4))
				idx = idx + 5
			End If

			chunks.Add chunks.Count, char
		Loop

		ParseString = Join(chunks.Items, "")
	End Function

	Private Function NextToken(ByRef str, ByVal idx)
		Do While idx <= Len(str) And InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		NextToken = idx
	End Function

End Class

Function AddToField(ByRef field, ByVal ftext)

	' for adding data to multi-valued fields
	If field = "" Then
		field = ftext
	Else
		field = field & Separator & ftext
	End If

End Function


Function AddToFieldWD(ByRef field, ByVal ftext)

	' for adding data to multi-valued fields without duplicated entries
	Dim tmp, x
	If field = "" Then
		field = ftext
	Else
		tmp = Split(field, Separator)
		For each x in tmp
			If LCase(ftext) = LCase(x) Then
				Exit Function
			End If
		Next
		field = field & Separator & ftext
	End If

End Function


Function LookForFeaturing(Text)

	Dim tmp, x
	tmp = Split(FeaturingKeywords, ",")
	For each x in tmp
		If LCase(Text) = LCase(x) Then
			LookForFeaturing = true
			Exit Function
		End If
	Next
	LookForFeaturing = false

End Function


Function CheckLeadingZeroTrackPosition(TrackPosition)

	Dim tmpSplit, tmpTrack, tmpPos
	If InStr(TrackPosition, "-") <> 0 Then
		tmpSplit = Split(TrackPosition, "-")
		tmpPos = tmpSplit(1)
	End If
	If InStr(TrackPosition, ".") <> 0 Then
		tmpSplit = Split(TrackPosition, ".")
		tmpPos = tmpSplit(1)
	End If
	If Left(tmpPos, 1) = "0" Then
		CheckLeadingZeroTrackPosition = True
	Else
		CheckLeadingZeroTrackPosition = False
	End If

End Function



Function exchange_roman_numbers(Text)

	If Text = "I" Then Text = 1
	If Text = "II" Then Text = 2
	If Text = "III" Then Text = 3
	If Text = "IV" Then Text = 4
	If Text = "V" Then Text = 5
	If Text = "VI" Then Text = 6
	If Text = "VII" Then Text = 7
	If Text = "VIII" Then Text = 8
	If Text = "IX" Then Text = 9
	If Text = "X" Then Text = 10
	If Text = "XI" Then Text = 11
	If Text = "XII" Then Text = 12
	If Text = "XIII" Then Text = 13
	If Text = "XIV" Then Text = 14
	If Text = "XV" Then Text = 15
	If Text = "XVI" Then Text = 16
	If Text = "XVII" Then Text = 17
	If Text = "XVIII" Then Text = 18
	If Text = "XIX" Then Text = 19
	If Text = "XX" Then Text = 20
	exchange_roman_numbers = Text

End Function

Function DecodeHtmlChars(Text)

	DecodeHtmlChars = Text
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&quot;",	"""")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&lt;",	"<")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&gt;",	">")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&amp;",	"&")

End Function


Function EncodeHtmlChars(Text)

	EncodeHtmlChars= Text
	EncodeHtmlChars= Replace(EncodeHtmlChars, "&",	"&amp;")
	EncodeHtmlChars= Replace(EncodeHtmlChars,"""",	"&quot;")
	EncodeHtmlChars= Replace(EncodeHtmlChars,"<",	"&lt;")
	EncodeHtmlChars= Replace(EncodeHtmlChars, ">",	"&gt;")

End Function


Function CleanSearchString(Text)

	CleanSearchString = Text
	CleanSearchString = Replace(CleanSearchString,")", " ")
	CleanSearchString = Replace(CleanSearchString,"(", " ")
	CleanSearchString = Replace(CleanSearchString,"[", " ")
	CleanSearchString = Replace(CleanSearchString,"]", " ")
	CleanSearchString = Replace(CleanSearchString,"@", " ")
	CleanSearchString = Replace(CleanSearchString,"_", " ")
	CleanSearchString = Replace(CleanSearchString,"?", " ")
	CleanSearchString = PackSpaces(CleanSearchString, False)

End Function


Function CleanArtistName(artistname)

	CleanArtistName = DecodeHtmlChars(artistname)
	If InStr(CleanArtistName, " (") > 0 Then CleanArtistName = Left(CleanArtistName, InStrRev(CleanArtistName, " (") - 1)
	If InStr(CleanArtistName, ", The") > 0 Then CleanArtistName = "The " & Left(CleanArtistName, InStrRev(CleanArtistName, ", The") - 1)
	CleanArtistName = PackSpaces(CleanArtistName, False)

End Function


Function AddAlternative(Alternative)

	Dim i
	Alternative = Trim(Alternative)
	If Alternative <> "" Then
		If Len(Alternative) > 50 Then Alternative = Left(Alternative, 50)
		For i = 0 To AlternativeList.Count - 1
			If AlternativeList.Item(i) = Alternative Then
				Exit Function
			End If
		Next
		AlternativeList.Add Alternative
	End If

End Function


Function AddAlternatives(Song)

	Dim SavedArtist, SavedTitle, SavedAlbum, SavedAlbumArtist, SavedFolderName, SavedFileName, Custom
	SavedArtist = Song.ArtistName
	SavedTitle = Song.Title
	SavedAlbum = Song.AlbumName
	SavedAlbumArtist = Song.AlbumArtistName
	WriteLog("Song.Path=" & Song.Path)
	If InStr(Song.Path, "\") > 0 Then
		SavedFolderName = Mid(Song.Path, 1, InStrRev(Song.Path,"\")-1)
		SavedFolderName = Mid(SavedFolderName, InStrRev(SavedFolderName,"\")+1)
		SavedFileName = Mid(Song.Path, 1, InStrRev(Song.Path,".")-1)
		SavedFileName = Mid(SavedFileName, InStrRev(SavedFileName,"\")+1)
	ElseIf InStr(Song.Path, "/") > 0 Then
		SavedFolderName = Mid(Song.Path, 1, InStrRev(Song.Path,"/")-1)
		SavedFolderName = Mid(SavedFolderName, InStrRev(SavedFolderName,"/")+1)
		SavedFileName = Mid(Song.Path, 1, InStrRev(Song.Path,".")-1)
		SavedFileName = Mid(SavedFileName, InStrRev(SavedFileName,"/")+1)
	End If
	If SavedFileName = "" Then SavedFileName = Song.Path
	If SavedFolderName = "" Then SavedFolderName = Song.Path

	AddAlternative SavedFolderName
	If(InStr(SavedFolderName,"(") > 0) Then
		Custom = Mid(SavedFolderName,1,InStr(SavedFolderName,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedFolderName,"[") > 0) Then
		Custom = Mid(SavedFolderName,1,InStr(SavedFolderName,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative SavedFileName
	If(InStr(SavedFileName,"(") > 0) Then
		Custom = Mid(SavedFileName,1,InStr(SavedFileName,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedFileName,"[") > 0) Then
		Custom = Mid(SavedFileName,1,InStr(SavedFileName,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative Custom
	AddAlternative SavedArtist
	AddAlternative SavedTitle
	AddAlternative SavedAlbum
	If SavedAlbumArtist <> "Various" And SavedAlbumArtist <> "Various Artists" And SavedAlbumArtist <> TxtVarious Then
		AddAlternative SavedAlbumArtist
	End If
	If(InStr(SavedTitle,"(") > 0) Then
		Custom = Mid(SavedTitle,1,InStr(SavedTitle,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedTitle,"[") > 0) Then
		Custom = Mid(SavedTitle,1,InStr(SavedTitle,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative SavedArtist & " " & SavedAlbum
	AddAlternative SavedAlbumArtist & " " & SavedAlbum
	AddAlternative SavedArtist & " " & SavedTitle
	AddAlternative SavedAlbumArtist & " " & SavedTitle

End Function


Function IsInteger(Str)

	Dim i, d
	IsInteger = True
	For i = 1 To Len(str)
		d = Mid(str, i, 1)
		If Asc(d) < 48 Or Asc(d) > 57 Then
			IsInteger = False
			Exit For
		End If
	Next

End Function


Function PackSpaces(Text, Report)

	PackSpaces = Text
	PackSpaces = Replace(PackSpaces,"  ", " ") 'pack spaces
	PackSpaces = Replace(PackSpaces,"  ", " ") 'pack spaces left
	If (Left(PackSpaces, 1) = " " Or Right(PackSpaces, 1) = " ") And Report = True Then
		ReportRelease()
	End If
	PackSpaces = Trim(PackSpaces) 'delete leading/trailing spaces

End Function


Function search_involved(Text, SearchText)

	WriteLog " "
	WriteLog "Start Search_involved"
	Dim tmp, x, RE, searchPattern
	Dim i
	tmp = Split(UnwantedKeywords, ",")
	Set RE = New RegExp
	RE.IgnoreCase = True
	For Each searchPattern In tmp
		If InStr(searchPattern, "*") <> 0 Then
			searchPattern = Replace(searchPattern, "*", ".*")
			RE.Pattern = "^" & searchPattern & "$"
			If RE.Test(SearchText) Then
				search_involved = -2
				Exit Function
			End If
		Else
			If Trim(LCase(SearchText)) = Trim(LCase(searchPattern)) Then
				search_involved = -2
				Exit Function
			End If
		End If
	Next

	For i = 1 To UBound(Text)
		If Left(Text(i), InStr(Text(i), ":")-1) = SearchText Then
			WriteLog "Text=" & Text(i)
			search_involved = i
			Exit Function
		End If
	Next
	search_involved = -1

End Function


Function search_involved_track(Text, SearchText)

	Dim i
	For i = 0 To UBound(Text)-1
		If Left(Text(i), Len(SearchText)) = SearchText Then
			search_involved_track = i
			Exit Function
		End If
	Next
	search_involved_track = -1

End Function


Function URLEncodeUTF8(ByRef input)

	' urlencode a string with UTF8 encoding - yes, it is cryptic but it works!
	Dim i, result, CurrentChar
	Dim FirstByte, SecondByte, ThirdByte

	result = ""
	For i = 1 To Len(input)
		CurrentChar = Mid(input, i, 1)
		CurrentChar = AscW(CurrentChar)

		If (CurrentChar < 0) Then
			CurrentChar = CurrentChar + 65536
		End If

		If (CurrentChar >= 0) And (CurrentChar < 128) Then
			' 1 byte
			If(CurrentChar = 32) Then
				' replace space with "+"
				result = result & "+"
			Else
				' replace punctuation chars with "%hex"
				result = result & Escape(Chr(CurrentChar))
			End If
		End If

		If (CurrentChar >= 128) And (CurrentChar < 2048) Then
			' 2 bytes
			FirstByte  = &HC0 Xor ((CurrentChar And &HFFFFFFC0) \ &H40&)
			SecondByte = &H80 Xor (CurrentChar And &H3F)
			result = result & "%" & Hex(FirstByte) & "%" & Hex(SecondByte)
		End If

		If (CurrentChar >= 2048) And (CurrentChar < 65536) Then
			' 3 bytes
			FirstByte  = &HE0 Xor (((CurrentChar And &HFFFFF000) \ &H1000&) And &HF)
			SecondByte = &H80 Xor (((CurrentChar And &HFFFFFFC0) \ &H40&) And &H3F)
			ThirdByte  = &H80 Xor (CurrentChar And &H3F)
			result = result & "%" & Hex(FirstByte) & "%" & Hex(SecondByte) & "%" & Hex(ThirdByte)
		End If
	Next
	URLEncodeUTF8 = result

End Function


Sub SwitchAll()

	Dim templateHTMLDoc, i, checkBox
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set checkBox = templateHTMLDoc.getElementById("selectall")
	SelectAll = checkBox.Checked

	Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
	button.style.backgroundColor = "#ED1C24"
	templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML="Please wait for collection data..."

	For i = 0 To iMaxTracks - 1
		If SelectAll Then
			UnselectedTracks(i) = ""
		Else
			UnselectedTracks(i) = "x"
		End If
	Next

	ReloadResults

End Sub


Function getArtistsName(Current, Role, QueryPage)

	REM WriteLog "Start getArtistsName"
	Dim artist, currentArtist, artistName, tmpArtistSeparator, tmp
	Dim FoundFeaturing

	Dim Artists(2)

	REM SavedArtistID = ""
	FoundFeaturing = False

	If current.Exists(Role) Then
		For Each artist in Current(Role)
			Set currentArtist = Current(Role)(artist)
			If QueryPage = "Discogs" Then
				If Not CheckUseAnv And currentArtist("anv") <> "" Then
					artistName = CleanArtistName(currentArtist("anv"))
				Else
					artistName = CleanArtistName(currentArtist("name"))
				End If
				If CheckTheBehindArtist And Left(artistName, 4) = "The " Then artistName = Mid(artistName, 5) & ", The"
				If SavedArtistID = "" Then
					SavedArtistID = currentArtist("id")
					WriteLog "SavedArtistID=" & SavedArtistID
				End If
			Else
				artistName = CleanArtistName(currentArtist("name"))
				If CheckTheBehindArtist And Left(artistName, 4) = "The " Then artistName = Mid(artistName, 5) & ", The"
				If SavedArtistID = "" Then
					SavedArtistID = currentArtist("artist")("id")
					WriteLog "SavedArtistID=" & SavedArtistID
				End If
			End If
			WriteLog "Artist found=" & artistName

			If Artists(0) = "" Then
				Artists(0) = artistName
				Artists(2) = artistName
			Else
				If FoundFeaturing = False Then
					Artists(0) = Artists(0) & tmpArtistSeparator & artistName
				Else
					If CheckIgnoreFeatArtist = False Then
						If Artists(1) = "" Then
							Artists(1) = tmpArtistSeparator & artistName
						Else
							Artists(1) = Artists(1) & ArtistSeparator & artistName
						End If
					End If
				End If
			End If

			tmp = ""
			If QueryPage = "Discogs" Then
				If currentArtist("join") <> "" Then
					tmp = Trim(currentArtist("join"))
				End If
			Else
				If currentArtist("joinphrase") <> "" Then
					tmp = Trim(currentArtist("joinphrase"))
				End If
			End If
			WriteLog "Featuring found=" & tmp

			If tmp = "," Then
				tmpArtistSeparator = ArtistSeparator
			Else
				If tmp <> "" Then
					If LookForFeaturing(tmp) = False Then
						tmpArtistSeparator = " " & tmp & " "
					End If
					If LookForFeaturing(tmp) = True And CheckIgnoreFeatArtist = False Then
						If CheckFeaturingName Then
							tmpArtistSeparator = TxtFeaturingName & " "
						Else
							tmpArtistSeparator = " " & tmp & " "
						End If
					End If
					If LookForFeaturing(tmp) = True Then
						FoundFeaturing = True
						WriteLog "FoundFeaturing=" & FoundFeaturing
					End If
				Else
					tmpArtistSeparator = ""
				End If
			End If
		Next

		If Right(Artists(0), Len(ArtistSeparator)) = ArtistSeparator Then Artists(0) = Left(Artists(0), Len(Artists(0))-Len(ArtistSeparator))
		Artists(0) = Trim(Artists(0))
		getArtistsName = Artists
		WriteLog "Artists(0)=" & Artists(0)
		WriteLog "Artists(1)=" & Artists(1)
		WriteLog "Artists(2)=" & Artists(2)
		WriteLog ""
	Else
		Artists(0) = ""
		getArtistsName = Artists
		WriteLog "No ExtraArtist found"
		WriteLog ""
	End If

End Function


Function ShowSearchFor

	Dim Form
	Set Form = UI.NewForm
	Form.Common.Width = 360
	Form.Common.Height = 90
	Form.FormPosition = 4
	Form.Caption = "What are you searching for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("ShowSearchForForm") = Form

	Dim Btn
	Set Btn = SDB.UI.NewButton(Form)
	Btn.Caption = SDB.Localize("Artist")
	Btn.Common.Width = 95
	Btn.Common.Height = 25
	Btn.Common.Left = 15
	Btn.Common.Top = 15
	Btn.Common.Anchors = 2+4
	Btn.ModalResult = 1

	Dim Btn2
	Set Btn2 = SDB.UI.NewButton(Form)
	Btn2.Caption = SDB.Localize("Album")
	Btn2.Common.Width = 95
	Btn2.Common.Height = 25
	Btn2.Common.Left = 125
	Btn2.Common.Top = 15
	Btn2.Common.Anchors = 2+4
	Btn2.ModalResult = 2

	Dim Btn3
	Set Btn3 = SDB.UI.NewButton(Form)
	If QueryPage = "Discogs" Then
		Btn3.Caption = SDB.Localize("Artist/Title/Album")
	Else
		Btn3.Caption = SDB.Localize("Artist - Album")
	End If
	Btn3.Common.Width = 95
	Btn3.Common.Height = 25
	Btn3.Common.Left = 235
	Btn3.Common.Top = 15
	Btn3.Common.Anchors = 2+4
	Btn3.ModalResult = 3

	ShowSearchFor = Form.ShowModal
	SDB.Objects("ShowSearchForForm") = Nothing
	SDB.ProcessMessages

End Function


Function ShowAdvancedSearch()

	WriteLog "Start AdvancedSearch"
	Dim Form, Label, At, Btn, Btn2, edt, searchURL
	Dim SendArtist, SendAlbum, SendTrack, SendType, SendPerPage, templateHTMLDoc

	Set Form = UI.NewForm
	Form.Common.Width = 320
	Form.Common.Height = 240
	Form.FormPosition = 4
	Form.Caption = "Advanced Search"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("ShowAdvancedSearchForm") = Form

	Set Label = UI.NewLabel(Form)
	Label.Common.SetRect 10, 10, 80, 25
	Label.Caption = "Please enter at least one field"

	Set Label = UI.NewLabel(Form)
	Label.Common.SetRect 10, 45, 60, 25
	Label.Caption = SDB.Localize("Artist")

	Set At = UI.NewEdit(Form)
	At.Common.SetRect 80, 40, 175, 35
	At.Common.ControlName = "Artist"
	At.Text = NewSearchArtist

	Set Label = UI.NewLabel(Form)
	Label.Common.SetRect 10, 80, 60, 25
	Label.Caption = SDB.Localize("Album")

	Set At = UI.NewEdit(Form)
	At.Common.SetRect 80, 75, 175, 35
	At.Common.ControlName = "Album"
	At.Text = NewSearchAlbum

	Set Label = UI.NewLabel(Form)
	Label.Common.SetRect 10, 120, 60, 25
	Label.Caption = SDB.Localize("Title")

	Set At = UI.NewEdit(Form)
	At.Common.SetRect 80, 115, 175, 35
	At.Common.ControlName = "Title"
	At.Text = NewSearchTrack

	Set Btn = SDB.UI.NewButton(Form)
	Btn.Caption = SDB.Localize("Search")
	Btn.Default = True
	Btn.Common.Width = 95
	Btn.Common.Height = 25
	Btn.Common.Left = 30
	Btn.Common.Top = 160
	Btn.Common.Anchors = 2+4
	Btn.ModalResult = 1

	Set Btn2 = SDB.UI.NewButton(Form)
	Btn2.Caption = SDB.Localize("Cancel")
	Btn2.Common.Width = 95
	Btn2.Common.Height = 25
	Btn2.Common.Left = 140
	Btn2.Common.Top = 160
	Btn2.Common.Anchors = 2+4
	Btn2.ModalResult = 2

	WriteLog "ShowForm"

	If Form.ShowModal = 1 Then
		Set edt = Form.Common.ChildControl("Artist")
		NewSearchArtist = edt.Text
		Set edt = Form.Common.ChildControl("Album")
		NewSearchAlbum = edt.Text
		Set edt = Form.Common.ChildControl("Title")
		NewSearchTrack = edt.Text

		WriteLog "NewSearchArtist=" & NewSearchArtist
		WriteLog "NewSearchAlbum=" & NewSearchAlbum
		WriteLog "NewSearchTrack=" & NewSearchTrack

		SDB.Objects("ShowAdvancedSearchForm") = Nothing
		SDB.ProcessMessages


		If NewSearchArtist <> "" Or NewSearchAlbum <> "" Or NewSearchTrack <> "" Then

			Set Results = SDB.NewStringList
			Set ResultsReleaseID = SDB.NewStringList

			SDB.Tools.WebSearch.ClearTracksData
			
			Set WebBrowser = SDB.Objects("WebBrowser")
			Set templateHTMLDoc = WebBrowser.Interf.Document
			Set button = templateHTMLDoc.getElementById("PleaseWaitButton")
			button.style.backgroundColor = "#ED1C24"
			templateHTMLDoc.getElementById("PleaseWaitButton").innerHTML="Please wait for collection data..."

			If QueryPage = "Discogs" Then
				SendType = "release"
				SendPerPage = "100"
				SendArtist = URLEncodeUTF8(CleanSearchString(NewSearchArtist))
				SendAlbum = URLEncodeUTF8(CleanSearchString(NewSearchAlbum))
				SendTrack = URLEncodeUTF8(CleanSearchString(NewSearchTrack))
				JSONParser_find_result "", "results", SendArtist, SendAlbum, SendTrack, SendType, "", SendPerPage, QueryPage, True
			End If

			If QueryPage = "MusicBrainz" Then
				searchURL = "http://musicbrainz.org/ws/2/release?query=artist:" & Chr(34) & CleanSearchString(URLEncodeUTF8(NewSearchArtist)) & Chr(34) & " AND release:" & Chr(34) & URLEncodeUTF8(CleanSearchString(NewSearchAlbum)) & Chr(34) & "&limit=50&offset=0&fmt=json"
				JSONParser_find_result searchURL, "releases", "", "", "", "", "", "", "MusicBrainz", False
			End If
			If QueryPage = "MetalArchives" Then
				searchURL = "https://www.metal-archives.com/search/ajax-advanced/searching/albums/?bandName=" & URLEncodeUTF8(CleanSearchString(NewSearchArtist)) & "&exactBandMatch=1&releaseTitle=" & URLEncodeUTF8(CleanSearchString(NewSearchAlbum)) & "&exactReleaseMatch=1&releaseYearFrom=&releaseMonthFrom=&releaseYearTo=&releaseMonthTo=&country=&location=&releaseLabelName=&genre="
				JSONParser_find_result searchURL, "", "", "", "", "", "", "", "MetalArchives", False
			End If
	
			SDB.Tools.WebSearch.SetSearchResults Results
			SDB.Tools.WebSearch.ResultIndex = 0
		End If
	Else
		SDB.Objects("ShowAdvancedSearchForm") = Nothing
		WriteLog "Advanced Search canceled"
	End If

End Function


Function UserCollection()

	WriteLog "Start UserCollection"
	Dim oXMLHTTP, response
	Dim json
	Set json = New VbsJson
	Dim Page, ReleaseCountMax, ReleasePages, r, id, ReleaseFound, ErrorFound
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	If DiscogsUsername = "" Then
		oXMLHTTP.Open "POST", "https://www.germanc64.de/mm/oauth/identity.php", False
		oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		oXMLHTTP.setRequestHeader "User-Agent",UserAgent
		WriteLog "Sending Post at=" & AccessToken & "&ats=" & AccessTokenSecret
		oXMLHTTP.send ("at=" & AccessToken & "&ats=" & AccessTokenSecret)

		If oXMLHTTP.Status = 200 Then
			If InStr(oXMLHTTP.responseText, "OAuth client error") <> 0 Then
				WriteLog "responseText=" & oXMLHTTP.responseText
				ErrorMessage = oXMLHTTP.responseText
				Exit Function
			Else
				WriteLog "responseText=" & oXMLHTTP.responseText
				Set response = json.Decode(oXMLHTTP.responseText)
				DiscogsUsername = response("username")
				ini.StringValue("DiscogsAutoTagWeb", "DiscogsUsername") = DiscogsUsername
			End If
		End If
	End If

	
	oXMLHTTP.Open "POST", "https://www.germanc64.de/mm/oauth/get_release_new.php", False
	oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXMLHTTP.setRequestHeader "User-Agent", UserAgent
	WriteLog "Sending Post at=" & AccessToken & "&ats=" & AccessTokenSecret & "&username=" & DiscogsUsername & "&release=" & CurrentReleaseID
	oXMLHTTP.send("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&username=" & DiscogsUsername & "&release=" & CurrentReleaseID)
	If oXMLHTTP.Status = 200 Then
		If InStr(oXMLHTTP.responseText, "OAuth client error") = 0 Then
			WriteLog "Response=" & oXMLHTTP.responseText
			WriteLog "User Collection received"
			Set response = json.Decode(oXMLHTTP.responseText)
			ReleaseCountMax = response("pagination")("items")
			WriteLog "ReleaseCountMax=" & ReleaseCountMax
			If Int(ReleaseCountMax) = 0 Then
				oXMLHTTP.Open "POST", "https://www.germanc64.de/mm/oauth/add_release.php", False
				oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				oXMLHTTP.setRequestHeader "User-Agent", UserAgent
				WriteLog "Sending Post at=" & AccessToken & "&ats=" & AccessTokenSecret & "&username=" & DiscogsUsername & "&release=" & CurrentReleaseID
				oXMLHTTP.send("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&username=" & DiscogsUsername & "&release=" & CurrentReleaseID)
				If oXMLHTTP.Status = 200 Then
					If InStr(oXMLHTTP.responseText, "OAuth client error") = 0 Then
						WriteLog "Response=" & oXMLHTTP.responseText
						WriteLog "Release added to User Collection"
						SDB.MessageBox "Release added to your User Collection", mtInformation, Array(mbOk)
					End If
				End If
			Else
				WriteLog "Release already exists"
				SDB.MessageBox "Release already exists in User Collection", mtInformation, Array(mbOk)
			End If
		Else
			WriteLog "responseText=" & oXMLHTTP.responseText
			WriteLog "Error getting user collection"
			SDB.MessageBox "OAuth client error", mtError, Array(mbOk)
		End If
	End If
	Set oXMLHTTP = Nothing

End Function


Sub ReportRelease()

	'If Trackname has leading and/or trailing spaces the release-id will be stored in database to fix the entry at discogs
	Dim oXMLHTTP
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	oXMLHTTP.Open "POST", "https://www.germanc64.de/mm/oauth/report_release.php", False
	oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXMLHTTP.setRequestHeader "User-Agent", UserAgent
	WriteLog "Sending release=" & CurrentReleaseID
	oXMLHTTP.send("release=" & CurrentReleaseID)
	Set oXMLHTTP = Nothing

End Sub


Function LeadingZeroDisc(TestNumber)

	If CheckLeadingZeroDisc = False Then
		LeadingZeroDisc = TestNumber
	Else
		If IsInteger(TestNumber) Then
			If Left(cStr(TestNumber), 1) <> "0" And TestNumber < 10 Then
				LeadingZeroDisc = "0" & TestNumber
			Else
				LeadingZeroDisc = TestNumber
			End If
		Else
			LeadingZeroDisc = TestNumber
		End If
	End If

End Function


Function bubblesort(Trackorder_org, byRef SortedTrackList)

	Dim TempValue, i, j, TempDisc, start, TempValue2, found
	' Set TrackOrderTable = SDB.NewStringList

	found = True
	For i = 0 to Trackorder_org.Count - 1
		SortedTrackList.add i
		If Trackorder_org.item(i).DiscNumberStr = "" Or Not IsNumeric(Trackorder_org.item(i).DiscNumberStr) Then found = False
	Next
	If found = True Then
		For i = 0 to Trackorder_org.Count - 1
			For j = 0 to Trackorder_org.Count - 2
				If j <> Int(Trackorder_org.Count) - 2 Then
					If Int(Trackorder_org.item(j).DiscNumberStr) > Int(Trackorder_org.item(j+1).DiscNumberStr) Then
						TempValue = Trackorder_org.item(j+1).DiscNumberStr
						Trackorder_org.item(j+1).DiscNumberStr = Trackorder_org.item(j).DiscNumberStr
						Trackorder_org.item(j).DiscNumberStr = TempValue
						TempValue2 = SortedTrackList.item(j+1)
						SortedTrackList.item(j+1) = SortedTrackList.item(j)
						SortedTrackList.item(j) = TempValue2
					End If
				End If
			Next
		Next
		For i = 0 to SortedTrackList.Count - 1
			WriteLog "disc=" & Trackorder_org.item(SortedTrackList.item(i)).DiscNumberStr & " track=" & Trackorder_org.item(SortedTrackList.item(i)).TrackOrderStr
		Next
	End If
	start = 0
	Do While start < Int(Trackorder_org.Count)
		If found = True Then
			TempDisc = Trackorder_org.item(start).DiscNumberStr
			For i = start to Int(Trackorder_org.Count) - 1
				If Int(TempDisc) <> Int(Trackorder_org.item(i).DiscNumberStr) Then
					start = i
					WriteLog "TempDisc=" & TempDisc & " New:" & Trackorder_org.item(i).DiscNumberStr & " i=" & i
					Exit For
				End If
				For j = start to Int(Trackorder_org.Count) - 2
					If j <> Int(Trackorder_org.Count) - 2 Then
						If Int(TempDisc) <> Int(Trackorder_org.item(j+1).DiscNumberStr) Then
							WriteLog "2. TempDisc=" & TempDisc & " New:" & Trackorder_org.item(j+1).DiscNumberStr & " j=" & j
							Exit For
						End If
						If Int(Trackorder_org.item(j).TrackOrderStr) > Int(Trackorder_org.item(j+1).TrackOrderStr) Then
							TempValue = Trackorder_org.item(j+1).TrackOrderStr
							Trackorder_org.item(j+1).TrackOrderStr = Trackorder_org.item(j).TrackOrderStr
							Trackorder_org.item(j).TrackOrderStr = TempValue
							TempValue2 = SortedTrackList.item(j+1)
							SortedTrackList.item(j+1) = SortedTrackList.item(j)
							SortedTrackList.item(j) = TempValue2
						End If
					End If
				Next
			Next
		Else
			For i = 0 to Int(Trackorder_org.Count) - 1
				For j = 0 to Int(Trackorder_org.Count) - 2
					If j <> Int(Trackorder_org.Count) - 2 Then
						If Trackorder_org.item(j).TrackOrderStr > Trackorder_org.item(j+1).TrackOrderStr Then
							TempValue = Trackorder_org.item(j+1).TrackOrderStr
							Trackorder_org.item(j+1).TrackOrderStr = Trackorder_org.item(j).TrackOrderStr
							Trackorder_org.item(j).TrackOrderStr = TempValue
							TempValue2 = SortedTrackList.item(j+1)
							SortedTrackList.item(j+1) = SortedTrackList.item(j)
							SortedTrackList.item(j) = TempValue2
						End If
					End If
				Next
			Next
		End If
		If i = Int(Trackorder_org.Count) Then Exit Do
	Loop
	For i = 0 to SortedTrackList.Count - 1
		WriteLog "SortedTrackList=" & SortedTrackList.Item(i)
	Next
	For i = 0 to SortedTrackList.Count - 1
		WriteLog "disc=" & Trackorder_org.item(SortedTrackList.item(i)).DiscNumberStr & " track=" & Trackorder_org.item(SortedTrackList.item(i)).TrackOrderStr
	Next

End Function


Function authorize_script(AuthFailed)

	WriteLog "Start Authorize Function"

	Dim authHTML, TypeLib
	Dim IEobj, oXMLHTTP
	Dim retIE, retryCnt, start, a
	Dim Form, WebBrowser2, Btn1, Lbl1, Btn2, Btn4
	
	If AuthFailed = False And AccessToken <> "" And AccessTokenSecret <> "" Then
		authorize_script = True
		Exit Function
	End If
	
	If AuthFailed = True Then
		SDB.MessageBox "Authorize failed (Err=1)! You have to authorize Discogs Tagger with your Discogs account !", mtError, Array(mbOk)
	End If
	
	Set Form = UI.NewForm
	Form.Common.Width = 300
	Form.Common.Height = 280
	Form.FormPosition = 4
	Form.Caption = "Authorize Discogs Tagger"
	Form.BorderStyle = 3
	Form.StayOnTop = True

	Set Lbl1 = UI.NewLabel(Form)
    Lbl1.Common.ControlName = "Lbl1"
    Lbl1.Common.SetRect 10, 10, 100, 16
    Lbl1.Caption = "In order to use the Discogs Tagger, you must"

	Set Btn1 = UI.NewButton(Form)
    Btn1.Common.ControlName = "Btn1"
    Btn1.Common.SetRect 10, 40, 250, 25
    Btn1.Caption = "1. Create an account on the Discogs website"
    Script.RegisterEvent Btn1, "OnClick", "Btn1Click"

	Set Btn2 = UI.NewButton(Form)
    Btn2.Common.ControlName = "Btn2"
    Btn2.Common.SetRect 10, 80, 250, 25
    Btn2.Caption = "2. Authorize the Discogs Tagger"
    Script.RegisterEvent Btn2, "OnClick", "Btn2Click"

	Set Lbl1 = UI.NewLabel(Form)
    Lbl1.Common.ControlName = "Lbl1"
    Lbl1.Common.SetRect 10, 120, 100, 16
    Lbl1.Caption = "to use your newly-created account."
	
	Set Btn4 = UI.NewButton(Form)
    Btn4.Common.ControlName = "Btn4"
    Btn4.Common.SetRect 10, 160, 250, 25
    Btn4.Caption = "Read this PDF if you need help"
    Script.RegisterEvent Btn4, "OnClick", "Btn4Click"

	Set Lbl1 = UI.NewLabel(Form)
    Lbl1.Common.ControlName = "Lbl1"
    Lbl1.Common.SetRect 10, 200, 100, 16
    Lbl1.Caption = "Press 'OK', once you've authorized the Discogs Tagger."

	Set TypeLib = CreateObject("Scriptlet.TypeLib")
	
	GUID = Mid(TypeLib.Guid, 2, 36)
	WriteLog "GUID=" & GUID

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("Ok")
	Btn3.Common.Width = 75
	Btn3.Common.Height = 25
	Btn3.Common.Left = 110
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Btn3.UseScript = Script.ScriptPath
	Btn3.ModalResult = 1
	Btn3.Default = True
	
	If Form.ShowModal = 1 Then
		Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
		oXMLHTTP.open "GET", "https://www.germanc64.de/mm/oauth/get_oauth_guid.php?f=" & GUID, false
		oXMLHTTP.send()
		If oXMLHTTP.Status = 200 Then
			retIE = oXMLHTTP.responseText

			If InStr(retIE, "AccessToken=") <> 0 Then
				start = InStr(retIE, "AccessToken=")
				retIE = Mid(retIE, start + 12)
				AccessToken = Left(retIE, 40)
				ini.StringValue("DiscogsAutoTagWeb","AccessToken") = AccessToken
				WriteLog "AccessToken=" & AccessToken
				start = InStr(retIE, "AccessTokenSecret=")
				retIE = Mid(retIE, start + 18)
				AccessTokenSecret = Left(retIE, 40)
				ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = AccessTokenSecret
				WriteLog "AccessTokenSecret=" & AccessTokenSecret
				SDB.MessageBox "Congratulations! You can now use Discogs Tagger!", mtInformation, Array(mbOk)
				authorize_script = True
			End If
		Else
			WriteLog "Authorize failed (Err=1)!"
			Set oXMLHTTP = Nothing
			authorize_script(False)
		End If

		Set oXMLHTTP = Nothing
	Else
		WriteLog "AccessToken found in ini = " & AccessToken
		WriteLog "AccessTokenSecret found in ini = " & AccessTokenSecret
		authorize_script = True
	End If
	WriteLog "End Authorize Function"
	

End Function


Sub Btn1Click

	Dim objShell
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "explorer.exe", "https://www.discogs.com"

End Sub

Sub Btn2Click

	Dim objShell, link
	Set objShell = CreateObject("Shell.Application")
	link = chr(34) & "https://www.germanc64.de/mm/oauth/oauth_guid.php?f=" & GUID & chr(34)
	objShell.ShellExecute "explorer.exe", link

End Sub

Sub Btn4Click

	Dim objShell, link
	Set objShell = CreateObject("Shell.Application")
	link = chr(34) & "https://www.germanc64.de/mm/oauth/HowTo_authorize_with_discogs.pdf" & chr(34)
	objShell.ShellExecute "explorer.exe", link

End Sub

Sub showOptions()

	WriteLog "Start OptionsForm"

	Dim Form, WebBrowser2, optionsHTML, optionsHTMLDoc, checkbox, text
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 275
	Form.Common.Height = 620
	Form.FormPosition = 4
	Form.Caption = "Tagging Options"
	Form.BorderStyle = 2
	Form.StayOnTop = True
	
	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = 170
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = 15
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	optionsHTML = "<HTML>"
	optionsHTML = optionsHTML &  "<HEAD>"
	optionsHTML = optionsHTML &  "<style type=""text/css"" media=""screen"">"
	optionsHTML = optionsHTML &  ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	optionsHTML = optionsHTML &  "option.tabletext{background-color:#3E7CBB;}"

	optionsHTML = optionsHTML &  "</style>"
	optionsHTML = optionsHTML &  "</HEAD>"
	optionsHTML = optionsHTML &  "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"

	optionsHTML = optionsHTML &  "<tr><td align=center><b>Tagging options:</b></td></tr>"
	If QueryPage = "Discogs" Then
		optionsHTML = optionsHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""useanv"" title=""Artist Name Variation - Using no name variation (e.g. nickname)"" >Don't Use ANV's</td></tr>"
	End If
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""yearonlydate"" title=""If checked only the Year will be saved (e.g. 14.01.1982 -> 1982)"" >Only Year Of Date</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""deleteduplicatedentry"" title=""If checked the duplicated entries in the label-tag or the catalog-tag will be deleted (e.g. Label-tag: Roadrunner Records; Roadrunner Records  ->  Roadrunner Records"" >Delete duplicated Entries</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""titlefeaturing"" title=""If checked the feat. artist appears in the title tag (e.g. Aaliyah ft. Timbaland - We Need a Resolution  ->  Aaliyah - We Need a Resolution (ft. Timbaland) )"" >feat. Artist behind Title</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""ignorefeaturing"" title=""If checked the feat. artist will be ignored"" >Ignore feat. artist</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""TheBehindArtist"" title=""If checked the ArtistName will be 'Beatles, The' instead of 'The Beatles'"" >Move 'The' in Artist to the end</td></tr>"

	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""FeaturingName"" title=""Rename 'feat.' to the given word"" >Rename 'feat.' to:</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=text id=""TxtFeaturingName"" ></td></tr>"

	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""various"" title=""Rename 'Various' Artist to the given word"" >Rename 'Various' Artist to:</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=text id=""txtvarious"" ></td></tr>"

	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""involvedpeoplesingle"" title=""Print every involved people on individual lines"" >Involved people on individual lines</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""showqtyformat"" title=""Show qty of format in format tag, if there is more than one"" >Show quantity of the format (i.e. 2 x CD)</td></tr>"

	optionsHTML = optionsHTML &  "<tr><td align=center><br></td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=center><b>Disc/Track Numbering:</b></td></tr>"

	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""TurnOffSubTrack"" title=""If checked the Sub-Track detection is turned off"" >No Sub-Track detection</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""SubTrackNameSelection"" title=""If checked the sub-track will be named like 'Sub-Track 1, Sub-Track 2, Sub Track 3'  if not checked the Sub-Tracks will be named like 'Track Name (Sub-Track 1, Sub-Track 2, Sub Track 3)'"" >Other Sub-Track Naming</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=text id=""SubTrackSeparator"" title=""Choose the separator used for sub-tracks"" size=""5""> Separator for sub-tracks</td></tr>"

	If QueryPage = "Discogs" Then
		optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""forcenumeric"" title=""Always use numbers instead of letters (Vinyl-releases use A1, A2,..., B1, B2 as track numbering)"" >Force To Numeric</td></tr>"
		optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""sidestodisc"" title=""Save the Vinyl sides to the disc tag"" >Sides To Disc</td></tr>"
	End If
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""forcedisc"" title=""Always add a disc-number"" >Force Disc Usage</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""nodisc"" title=""Prevent the script from interpret sub tracks as disc-numbers"" >Force NO Disc Usage</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""leadingzero"" title=""Track Position: 1 -> 01   2 -> 02 ..."" >Add Leading Zero (Track#)</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""leadingzeroDisc"" title=""Disc-number: 1 -> 01   2 -> 02 ..."" >Add Leading Zero (Disc#)</td></tr>"
	optionsHTML = optionsHTML &  "<tr><td align=left><input type=checkbox id=""NewMatchingEngine"" title=""If you use the Discogs Tagger for tagging complete albums, use the new Matching Engine. Uncheck this if you only tagging single tracks"" >Use new Matching Engine</td></tr>"

	optionsHTML = optionsHTML &  "</table>"
	
	optionsHTML = optionsHTML &  "</body>"
	optionsHTML = optionsHTML &  "</HTML>"


	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument optionsHTML
	Set optionsHTMLDoc = WebBrowser2.Interf.Document

	If QueryPage = "Discogs" Then
		Set checkBox = optionsHTMLDoc.getElementById("useanv")
		checkBox.Checked = CheckUseAnv
	End If
	Set checkBox = optionsHTMLDoc.getElementById("yearonlydate")
	checkBox.Checked = CheckYearOnlyDate
	If QueryPage = "Discogs" Then
		Set checkBox = optionsHTMLDoc.getElementById("forcenumeric")
		checkBox.Checked = CheckForceNumeric
		Set checkBox = optionsHTMLDoc.getElementById("sidestodisc")
		checkBox.Checked = CheckSidesToDisc
	End If
	Set checkBox = optionsHTMLDoc.getElementById("forcedisc")
	checkBox.Checked = CheckForceDisc
	Set checkBox = optionsHTMLDoc.getElementById("nodisc")
	checkBox.Checked = CheckNoDisc
	Set checkBox = optionsHTMLDoc.getElementById("leadingzero")
	checkBox.Checked = CheckLeadingZero
	Set checkBox = optionsHTMLDoc.getElementById("leadingzeroDisc")
	checkBox.Checked = CheckLeadingZeroDisc
	Set checkBox = optionsHTMLDoc.getElementById("titlefeaturing")
	checkBox.Checked = CheckTitleFeaturing
	Set checkBox = optionsHTMLDoc.getElementById("deleteduplicatedentry")
	checkBox.Checked = CheckDeleteDuplicatedEntry
	Set text = optionsHTMLDoc.getElementById("TxtFeaturingName")
	text.value = TxtFeaturingName
	Set checkbox = optionsHTMLDoc.getElementById("FeaturingName")
	checkBox.Checked = CheckFeaturingName
	Set checkbox = optionsHTMLDoc.getElementById("NewMatchingEngine")
	checkBox.Checked = CheckNewMatchingEngine
	
	Set text = optionsHTMLDoc.getElementById("txtvarious")
	text.value = TxtVarious
	Set checkBox = optionsHTMLDoc.getElementById("various")
	checkBox.Checked = CheckVarious
	Set checkBox = optionsHTMLDoc.getElementById("SubTrackNameSelection")
	checkBox.Checked = SubTrackNameSelection
	Set checkBox = optionsHTMLDoc.getElementById("TurnOffSubTrack")
	checkBox.Checked = CheckTurnOffSubTrack
	Set checkBox = optionsHTMLDoc.getElementById("involvedpeoplesingle")
	checkBox.Checked = CheckInvolvedPeopleSingleLine
	Set checkBox = optionsHTMLDoc.getElementById("showqtyformat")
	checkBox.Checked = CheckShowQtyFormat
	Set checkBox = optionsHTMLDoc.getElementById("TheBehindArtist")
	checkBox.Checked = CheckTheBehindArtist
	Set text = optionsHTMLDoc.getElementById("SubTrackSeparator")
	text.value = SubTrackSeparator


	If Form.ShowModal = 1 Then
		
		If QueryPage = "Discogs" Then
			Set checkBox = optionsHTMLDoc.getElementById("useanv")
			CheckUseAnv = checkBox.Checked
		End If
		Set checkBox = optionsHTMLDoc.getElementById("yearonlydate")
		CheckYearOnlyDate = checkBox.Checked
		If QueryPage = "Discogs" Then
			Set checkBox = optionsHTMLDoc.getElementById("forcenumeric")
			CheckForceNumeric = checkBox.Checked
			Set checkBox = optionsHTMLDoc.getElementById("sidestodisc")
			CheckSidesToDisc = checkBox.Checked
		End If
		Set checkBox = optionsHTMLDoc.getElementById("forcedisc")
		If Not CheckForceDisc And checkBox.Checked Then
			CheckNoDisc = False
			CheckForceDisc = checkBox.Checked
		Else
			CheckForceDisc = checkBox.Checked
			Set checkBox = optionsHTMLDoc.getElementById("nodisc")
			If Not CheckNoDisc And checkBox.Checked Then
				CheckForceDisc = False
			End If
			CheckNoDisc = checkBox.Checked
		End If
		Set checkBox = optionsHTMLDoc.getElementById("leadingzero")
		CheckLeadingZero = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("leadingzeroDisc")
		CheckLeadingZeroDisc = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("titlefeaturing")
		CheckTitleFeaturing = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("deleteduplicatedentry")
		CheckDeleteDuplicatedEntry = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("FeaturingName")
		CheckFeaturingName = checkBox.Checked
		Set text = optionsHTMLDoc.getElementById("TxtFeaturingName")
		TxtFeaturingName = text.Value

		Set checkBox = optionsHTMLDoc.getElementById("various")
		CheckVarious = checkBox.Checked
		Set text = optionsHTMLDoc.getElementById("txtvarious")
		TxtVarious = text.Value
		Set checkBox = optionsHTMLDoc.getElementById("SubTrackNameSelection")
		SubTrackNameSelection = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("TurnOffSubTrack")
		CheckTurnOffSubTrack = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("involvedpeoplesingle")
		CheckInvolvedPeopleSingleLine = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("showqtyformat")
		CheckShowQtyFormat = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("TheBehindArtist")
		CheckTheBehindArtist = checkBox.Checked
		Set checkBox = optionsHTMLDoc.getElementById("ignorefeaturing")
		CheckIgnoreFeatArtist = checkBox.Checked
		Set text = optionsHTMLDoc.getElementById("SubTrackSeparator")
		SubTrackSeparator = text.Value
		Set checkbox = optionsHTMLDoc.getElementById("NewMatchingEngine")
		CheckNewMatchingEngine = checkBox.Checked


		SDB.Objects("WebBrowser2") = Nothing
		ReloadResults
	Else
		SDB.Objects("WebBrowser2") = Nothing
	End If


End Sub
