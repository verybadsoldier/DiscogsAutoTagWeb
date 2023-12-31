'
' MediaMonkey Script
'
' NAME: Discogs Tagger Options 3.8
'
' AUTHOR: crap_inhuman
' DATE : 13/09/2017
'
'
' INSTALL: Automatic installation during Discogs Tagger install
'
'Changes from 3.7 to 3.8
'Added option to change search-mode (Using album name for initial search Yes/no)

'Changes from 3.6 to 3.7
'Removed option for Style tag

'Changes from 3.5 to 3.6
'Added option to replace a � or ` with a '

'Changes from 3.4 to 3.5
'Added "FormatSeparator" option

'Changes from 3.3 to 3.4
'Added new 'where to store date' option

'Changes from 3.2 to 3.3
'Changed the 'where to store date' option

'Changes from 3.1 to 3.2
'Added/Changed tooltips

'Changes from 3.0 to 3.1
'Added option for limiting releases

'Changes from 2.9 to 3.0
'Added option for Date tag storing

'Changes from 2.8 to 2.9
'Added option to turn 'check if it's already in Discogs Collection' off

'Changes from 2.7 to 2.8
'"Save small cover" option corrected

'Changes from 2.6 to 2.7
'Removed CheckImmedSaveImage

'Changes from 2.5 to 2.6
'Added check for new version once a day

'Changes from 2.4 to 2.5
'Adding ISRC to CatalogTag

'Changes from 2.3 to 2.4
'Removed a bug with empty keyword fields

'Changes from 2.2 to 2.3
'Wrong option-name fixed

'Changes from 2.1 to 2.2
'Added option to enter unwanted tags in involved people
'Added option to save selected "More images" after closing the popup
'Added option "Don't copy empty values to non-empty fields"

'Changes from 2.0 to 2.1
'Removed metal-archives.com for release search

'Changes from 1.9 to 2.0
'Added text fields to manually copy/paste the access token and access token secret

'Changes from 1.8 to 1.9
'Added metal-archives.com for release search instead of discogs(BETA)

'Changes from 1.7 to 1.8
'Split the options in 2 parts
'
'Changes from 1.6 to 1.7
'Added the option for switching the last artist separator ("&" or "chosen separator")
'
'Changes from 1.5 to 1.6
'Added the option for changing the artist separator
'
'Changes from 1.4 to 1.5
'Added the "featuring" keywords
'
'Changes from 1.3 to 1.4
'Changed the separator from '|' to ','
'
'Changes from 1.2 to 1.3
'Added 'Don't save' and 4 more fields for saving release-number
'
'Changes from 1.1 to 1.2
'Added 3 new options
'
'Changes from 1.0 to 1.1
'Added option for changing keywords
'




Sub OnStartup

	DiscogsOptions = SDB.UI.AddOptionSheet( "Discogs Tagger", Script.ScriptPath, "InitSheet", "SaveSheet", -3)
	Call SDB.UI.AddOptionSheet("Keywords",Script.ScriptPath,"InitSheet2","SaveSheet2", DiscogsOptions)

End Sub

Sub InitSheet(Sheet)

	Dim UI : Set UI = SDB.UI
	Set ini = SDB.IniFile
	If Not (ini Is Nothing) Then
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

		If ini.StringValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = true
		End If
		' If ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "" Then
			' ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "Default (stored with Genre)"
		' End If

		If ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = ", "
		End If
		If ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","FormatSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FormatSeparator") = ", "
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = "" Then
			If ini.ValueExists("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") Then
				If ini.BoolValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") = false Then
					ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 0
				Else
					ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 1
				End If
				ini.DeleteKey "DiscogsAutoTagWeb","CheckNotAlwaysSaveimage"
				
			Else
				ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 1
			End If
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSmallCover") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover") = False
		End If
		If ini.ValueExists("DiscogsAutoTagWeb","CheckCover") Then
			ini.DeleteKey "DiscogsAutoTagWeb","CheckCover"
		End If
		If ini.ValueExists("DiscogsAutoTagWeb","UseMetalArchives") Then
			ini.DeleteKey "DiscogsAutoTagWeb","UseMetalArchives"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","AccessToken") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","AccessToken") = ""
		End If
		If ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = ""
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDontFillEmptyFields") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDontFillEmptyFields") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckNewVersion") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckNewVersion") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDiscogsCollectionOff") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDiscogsCollectionOff") = True
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
		If ini.ValueExists("DiscogsAutoTagWeb","ImmedSaveImage") Then
			ini.DeleteKey "DiscogsAutoTagWeb","ImmedSaveImage"
		End If
		If ini.ValueExists("DiscogsAutoTagWeb","CheckImmedSaveImage") Then
			ini.DeleteKey "DiscogsAutoTagWeb","CheckImmedSaveImage"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckReplaceAccentChar") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckReplaceAccentChar") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSearchWithAlbum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSearchWithAlbum") = True
		End If
	End If


	ReleaseTag = ini.StringValue("DiscogsAutoTagWeb","ReleaseTag")
	CatalogTag = ini.StringValue("DiscogsAutoTagWeb","CatalogTag")
	CountryTag = ini.StringValue("DiscogsAutoTagWeb","CountryTag")
	FormatTag = ini.StringValue("DiscogsAutoTagWeb","FormatTag")
	CheckOriginalDiscogsTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack")
		ArtistSeparator = ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator")
	FormatSeparator = ini.StringValue("DiscogsAutoTagWeb","FormatSeparator")
	ArtistLastSeparator = ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator")
	CheckSaveImage = ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage")			'0 = Always save - 1 = Only when no image found - 2 = always don't save
	CheckSmallCover = ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover")
	AccessToken = ini.StringValue("DiscogsAutoTagWeb","AccessToken")
	AccessTokenSecret = ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret")
	CheckDontFillEmptyFields = ini.BoolValue("DiscogsAutoTagWeb","CheckDontFillEmptyFields")
	CheckNewVersion = ini.BoolValue("DiscogsAutoTagWeb","CheckNewVersion")
	CheckDiscogsCollectionOff = ini.BoolValue("DiscogsAutoTagWeb","CheckDiscogsCollectionOff")
	StoreDate = ini.StringValue("DiscogsAutoTagWeb","StoreDate")
	CheckLimitReleases = ini.StringValue("DiscogsAutoTagWeb","LimitReleases")
	CheckReplaceAccentChar = ini.BoolValue("DiscogsAutoTagWeb","CheckReplaceAccentChar")
	CheckSearchWithAlbum = ini.BoolValue("DiscogsAutoTagWeb","CheckSearchWithAlbum")

	CustomField1 = "Custom1 (" & ini.StringValue("CustomFields","Fld1Name") & ")"
	CustomField2 = "Custom2 (" & ini.StringValue("CustomFields","Fld2Name") & ")"
	CustomField3 = "Custom3 (" & ini.StringValue("CustomFields","Fld3Name") & ")"
	CustomField4 = "Custom4 (" & ini.StringValue("CustomFields","Fld4Name") & ")"
	CustomField5 = "Custom5 (" & ini.StringValue("CustomFields","Fld5Name") & ")"


	Dim GroupBox0
	Set GroupBox0 = UI.NewGroupBox(Sheet)
	GroupBox0.Caption = "Please choose the custom fields, where Discogs Tagger save the information"
	GroupBox0.Common.SetRect 10, 10, 500, 250

	Dim Label1
	Set Label1 = UI.NewLabel(GroupBox0)
	Label1.Common.SetRect 65, 15, 150, 25
	Label1.Caption = "Don't choose a Custom Field more than once !!"
	Dim DD1
	Set DD1 = UI.NewDropDown(GroupBox0)
	DD1.Common.SetRect 240, 40, 200, 25
	DD1.Style = 2
	DD1.AddItem (CustomField1)
	DD1.AddItem (CustomField2)
	DD1.AddItem (CustomField3)
	DD1.AddItem (CustomField4)
	DD1.AddItem (CustomField5)
	DD1.AddItem ("Don't save")
	DD1.AddItem (SDB.Localize("ISRC"))
	DD1.AddItem (SDB.Localize("Grouping"))
	DD1.AddItem (SDB.Localize("Encoder"))
	DD1.AddItem (SDB.Localize("Copyright"))

	DD1.Common.ControlName = "ReleaseTag"
	If ReleaseTag = "Custom1"Then
		DD1.ItemIndex = 0
	End If
	If ReleaseTag = "Custom2" Then
		DD1.ItemIndex = 1
	End If
	If ReleaseTag = "Custom3" Then
		DD1.ItemIndex = 2
	End If
	If ReleaseTag = "Custom4" Then
		DD1.ItemIndex = 3
	End If
	If ReleaseTag = "Custom5" Then
		DD1.ItemIndex = 4
	End If
	If ReleaseTag = "Don't save" Then
		DD1.ItemIndex = 5
	End If
		If ReleaseTag = "ISRC" Then
		DD1.ItemIndex = 6
	End If
	If ReleaseTag = "Grouping" Then
		DD1.ItemIndex = 7
	End If
	If ReleaseTag = "Encoder" Then
		DD1.ItemIndex = 8
	End If
	If ReleaseTag = "Copyright" Then
		DD1.ItemIndex = 9
	End If
	Set Label1 = UI.NewLabel(GroupBox0)
	Label1.Common.SetRect 20, 45, 150, 25
	Label1.Caption = "Choose field for saving release-number"
	Label1.Common.Hint = "In brackets you see the name you choose for the custom tag"

	Dim DD2
	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 70, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)
	DD2.AddItem ("Don't save")
	DD2.AddItem (SDB.Localize("ISRC"))

	DD2.Common.ControlName = "CatalogTag"
	If CatalogTag = "Custom1" Then
		DD2.ItemIndex = 0
	End If
	If CatalogTag = "Custom2" Then
		DD2.ItemIndex = 1
	End If
	If CatalogTag = "Custom3" Then
		DD2.ItemIndex = 2
	End If
	If CatalogTag = "Custom4" Then
		DD2.ItemIndex = 3
	End If
	If CatalogTag = "Custom5" Then
		DD2.ItemIndex = 4
	End If
	If CatalogTag = "Don't save" Then
		DD2.ItemIndex = 5
	End If
	If CatalogTag = "ISRC" Then
		DD2.ItemIndex = 6
	End If
	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 75, 150, 25
	Label2.Caption = "Choose field for saving catalog number"
	Label2.Common.Hint = "In brackets you see the name you choose for the custom tag"

	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 100, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)
	DD2.AddItem ("Don't save")

	DD2.Common.ControlName = "CountryTag"
	If CountryTag = "Custom1" Then
		DD2.ItemIndex = 0
	End If
	If CountryTag = "Custom2" Then
		DD2.ItemIndex = 1
	End If
	If CountryTag = "Custom3" Then
		DD2.ItemIndex = 2
	End If
	If CountryTag = "Custom4" Then
		DD2.ItemIndex = 3
	End If
	If CountryTag = "Custom5" Then
		DD2.ItemIndex = 4
	End If
	If CountryTag = "Don't save" Then
		DD2.ItemIndex = 5
	End If

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 105, 150, 25
	Label2.Caption = "Choose field for saving release country"
	Label2.Common.Hint = "In brackets you see the name you choose for the custom tag"

	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 130, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)
	DD2.AddItem ("Don't save")

	DD2.Common.ControlName = "FormatTag"
	If FormatTag = "Custom1" Then
		DD2.ItemIndex = 0
	End If
	If FormatTag = "Custom2" Then
		DD2.ItemIndex = 1
	End If
	If FormatTag = "Custom3" Then
		DD2.ItemIndex = 2
	End If
	If FormatTag = "Custom4" Then
		DD2.ItemIndex = 3
	End If
	If FormatTag = "Custom5" Then
		DD2.ItemIndex = 4
	End If
	If FormatTag = "Don't save" Then
		DD2.ItemIndex = 5
	End If

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 135, 150, 25
	Label2.Caption = "Choose field for saving media format"
	Label2.Common.Hint = "In brackets you see the name you choose for the custom tag"

	' Dim Combo
	' Set Combo = UI.NewDropDown(GroupBox0)
	' Combo.Common.SetRect 240, 160, 200, 25
	' Combo.Style = 2     ' List
	' Combo.Common.ControlName = "CheckStyleField"

	' Combo.AddItem ("Default (stored with Genre)")
	' Combo.AddItem (CustomField1)
	' Combo.AddItem (CustomField2)
	' Combo.AddItem (CustomField3)
	' Combo.AddItem (CustomField4)
	' Combo.AddItem (CustomField5)

	' If CheckStyleField = "Default (stored with Genre)" Then
		' Combo.ItemIndex = 0
	' End If
	' If CheckStyleField = "Custom1" Then
		' Combo.ItemIndex = 1
	' End If
	' If CheckStyleField = "Custom2" Then
		' Combo.ItemIndex = 2
	' End If
	' If CheckStyleField = "Custom3" Then
		' Combo.ItemIndex = 3
	' End If
	' If CheckStyleField = "Custom4" Then
		' Combo.ItemIndex = 4
	' End If
	' If CheckStyleField = "Custom5" Then
		' Combo.ItemIndex = 5
	' End If

	' Set Label2 = UI.NewLabel(GroupBox0)
	' Label2.Common.SetRect 20, 165, 50, 25
	' Label2.Caption = "Choose field for saving Style"
	' Label2.Common.Hint = "In brackets you see the name you choose for the custom tag"

	Set Combo = UI.NewDropDown(GroupBox0)
	Combo.Common.SetRect 220, 220, 220, 25
	Combo.Style = 2     ' List
	Combo.Common.ControlName = "StoreDate"

	Combo.AddItem ("No (Default)")
	Combo.AddItem ("Yes, store release date into both date fields")
	Combo.AddItem ("Yes, store original date into both date fields")

	Combo.ItemIndex = StoreDate

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 225, 50, 25
	Label2.Caption = "Use only one release date?"
	Label2.Common.Hint = "If you choose no, the release date will be stored in the date field and the master release date in the original date field (only with Discogs)" & vbCrLf & "If you choose yes, the release or original date will be stored into both date fields in mm."

	
	Dim GroupBox1
	Set GroupBox1 = UI.NewGroupBox(Sheet)
	GroupBox1.Caption = "Cover-Images"
	GroupBox1.Common.SetRect 10, 270, 500, 80

	Dim Checkbox1
	Set Checkbox1 = UI.NewCheckBox(GroupBox1)
	Checkbox1.Common.SetRect 20, 20, 250, 15
	Checkbox1.Common.ControlName = "ControlSaveImage1"
	Checkbox1.Caption = "Always set option for saving Cover-Images"
	Checkbox1.Common.Hint = "The script always set the option to save the Cover-Image."
	Set SDB.Objects("CoverSaveOn") = Checkbox1
	If CheckSaveImage = 0 or CheckSaveImage = 1 Then
		Checkbox1.checked = True
	Else
		Checkbox1.checked = False
	End If

	Dim Checkbox12
	Set Checkbox12 = UI.NewCheckBox(GroupBox1)
	Checkbox12.Common.SetRect 40, 40, 250, 15
	Checkbox12.Common.ControlName = "ControlSaveImage12"
	Checkbox12.Caption = "Only if no image already exists"
	Checkbox12.Common.Hint = "If option set the script only mark covers for save when no image already exists."
	Set SDB.Objects("CoverSaveIfEmpty") = Checkbox12
	If CheckSaveImage = 0 Then
		Checkbox12.checked = False
		Checkbox12.Common.Enabled = True
	ElseIf CheckSaveImage = 1 Then
		Checkbox12.checked = True
		Checkbox12.Common.Enabled = True
	Else
		Checkbox12.checked = False
		Checkbox12.Common.Enabled = False
	End If

	Dim Checkbox13
	Set Checkbox13 = UI.NewCheckBox(GroupBox1)
	Checkbox13.Common.SetRect 20, 60, 250, 15
	Checkbox13.Common.ControlName = "ControlSaveImage13"
	Checkbox13.Caption = "Small Cover (150x150)"
	Checkbox13.Common.Hint = "If option not set the script get the large cover images."
	Set SDB.Objects("SmallCoverSave") = Checkbox13
	If CheckSmallCover = False Then
		Checkbox13.checked = False
	Else
		Checkbox13.checked = True
	End If
	
	Script.RegisterEvent Checkbox1.Common, "OnClick", "ChBClick"

	Dim GroupBox2
	Set GroupBox2 = UI.NewGroupBox(Sheet)
	GroupBox2.Caption = "Misc"
	GroupBox2.Common.SetRect 10, 360, 500, 250


	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 40, 20, 50, 25
	Label2.Caption = "Show the original Discogs track position"

	Dim Checkbox2
	Set Checkbox2 = UI.NewCheckBox(GroupBox2)
	Checkbox2.Common.SetRect 20, 20, 15, 15
	Checkbox2.Common.ControlName = "CheckOriginalDiscogsTrack"
	If CheckOriginalDiscogsTrack = true Then Checkbox2.Checked = true

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 40, 40, 50, 25
	Label2.Caption = "Don't copy empty values to non-empty fields"

	Set Checkbox2 = UI.NewCheckBox(GroupBox2)
	Checkbox2.Common.SetRect 20, 40, 15, 15
	Checkbox2.Common.ControlName = "CheckDontFillEmptyFields"
	Checkbox2.Common.Hint = "If checked, the script write blank value if the release at discogs have blank value too"
	If CheckDontFillEmptyFields = true Then Checkbox2.Checked = true

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 40, 60, 50, 25
	Label2.Caption = "Check for new version"

	Set Checkbox2 = UI.NewCheckBox(GroupBox2)
	Checkbox2.Common.SetRect 20, 60, 15, 15
	Checkbox2.Common.ControlName = "CheckNewVersion"
	Checkbox2.Common.Hint = "If checked, the script check for new version once a day"
	If CheckNewVersion = true Then Checkbox2.Checked = true

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 20, 90, 50, 25
	Label2.Caption = SDB.Localize("Artist Separator")
	Label2.Common.Hint = "Standard is ', ' without apostrophe" & vbCrLf & "The artist separator will only be used in the artist tag and the involved people tag"

	Set EditArtistSep = UI.NewEdit(GroupBox2)
	EditArtistSep.Common.SetRect 20, 105, 50, 35
	EditArtistSep.Common.ControlName = "ArtistSeparator"
	EditArtistSep.Text = ArtistSeparator
	EditArtistSep.Common.Hint = "Standard is ', ' without apostrophe" & vbCrLf & "The artist separator will only be used in the artist tag and the involved people tag"

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 320, 90, 50, 25
	Label2.Caption = SDB.Localize("Format Separator")
	Label2.Common.Hint = "Standard is ', ' without apostrophe" & vbCrLf & "The format separator will only be used in the media format tag"

	Set EditFormatSep = UI.NewEdit(GroupBox2)
	EditFormatSep.Common.SetRect 320, 105, 50, 35
	EditFormatSep.Common.ControlName = "FormatSeparator"
	EditFormatSep.Text = FormatSeparator
	EditFormatSep.Common.Hint = "Standard is ', ' without apostrophe" & vbCrLf & "The format separator will only be used in the media format tag"

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 165, 107, 125, 25
	Label2.Caption = "Artist Last Separator = &&"
	Label2.Common.Hint = "If checked artist list will be Artist1" & ArtistSeparator & "Artist2 & Artist3" & vbCrLf & "If not checked it will be Artist1" & ArtistSeparator & "Artist2" & ArtistSeparator & "Artist3"

	Dim Checkbox3
	Set Checkbox3 = UI.NewCheckBox(GroupBox2)
	Checkbox3.Common.SetRect 145, 107, 15, 15
	Checkbox3.Common.ControlName = "EditArtistLastSep"
	Checkbox3.Common.Hint = "If checked artist list will be Artist1" & ArtistSeparator & "Artist2 & Artist3" & vbCrLf & "If not checked it will be Artist1" & ArtistSeparator & "Artist2" & ArtistSeparator & "Artist3"
	If ArtistLastSeparator = true Then Checkbox3.Checked = true

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 40, 145, 125, 25
	Label2.Caption = "Don't check if it's already in Discogs Collection"
	Label2.Common.Hint = "If checked the script will not check if the album is already in the Discogs Collection"

	Dim Checkbox4
	Set Checkbox4 = UI.NewCheckBox(GroupBox2)
	Checkbox4.Common.SetRect 20, 145, 15, 15
	Checkbox4.Common.ControlName = "CheckDCoff"
	Checkbox4.Common.Hint = "If checked the script will not check if the album is already in the Discogs Collection"
	If CheckDiscogsCollectionOff = true Then Checkbox4.Checked = true

	Set Label1 = UI.NewLabel(GroupBox2)
	Label1.Common.SetRect 20, 170, 150, 25
	Label1.Caption = "Limit amount of releases"
	Label1.Common.Hint = "Limit the amount of fetched releases - The script need more time if you choose more releases!"

	Set Combo = UI.NewDropDown(GroupBox2)
	Combo.Common.SetRect 240, 165, 200, 25
	Combo.Style = 2     ' List
	Combo.Common.ControlName = "LimitReleases"

	Combo.AddItem ("50")
	Combo.AddItem ("100")
	Combo.AddItem ("200")
	Combo.AddItem ("400")

	Combo.ItemIndex = CheckLimitReleases

	Dim Checkbox14
	Set Checkbox14 = UI.NewCheckBox(GroupBox2)
	Checkbox14.Common.SetRect 20, 200, 250, 15
	Checkbox14.Common.ControlName = "ControlReplaceAccentChar"
	Checkbox14.Caption = "Replace ` or � with a '"
	Checkbox14.Common.Hint = "In songnames the char �or ` will be replaced with a '"
	If CheckReplaceAccentChar = 0 Then
		Checkbox14.checked = False
	Else
		Checkbox14.checked = True
	End If

	Dim Checkbox15
	Set Checkbox15 = UI.NewCheckBox(GroupBox2)
	Checkbox15.Common.SetRect 20, 230, 250, 15
	Checkbox15.Common.ControlName = "UseAlbumForInitialSearch"
	Checkbox15.Caption = "Use album name for initial search"
	Checkbox15.Common.Hint = "If checked the album name of the song will be used for searching the release. (standard)" & vbCrLf & "If not checked, only the artist and the song name will be used for searching"
	If CheckSearchWithAlbum = True Then
		Checkbox15.checked = True
	Else
		Checkbox15.checked = False
	End If

	Dim GroupBox4
	Set GroupBox4 = UI.NewGroupBox(Sheet)
	GroupBox4.Caption = "Discogs Access Token"
	GroupBox4.Common.SetRect 10, 620, 500, 90

	Set Label2 = UI.NewLabel(GroupBox4)
	Label2.Common.SetRect 300, 25, 100, 25
	Label2.Caption = "Access Token"

	Dim At
	Set At = UI.NewEdit(GroupBox4)
	At.Common.SetRect 20, 20, 275, 35
	At.Common.ControlName = "AccessToken"
	At.Text = AccessToken

	Set Label2 = UI.NewLabel(GroupBox4)
	Label2.Common.SetRect 300, 55, 100, 25
	Label2.Caption = "Access Token Secret"

	Dim Ats
	Set Ats = UI.NewEdit(GroupBox4)
	Ats.Common.SetRect 20, 50, 275, 35
	Ats.Common.ControlName = "AccessTokenSecret"
	Ats.Text = AccessTokenSecret

End Sub

Sub ChBClick(CheckBox1)

	Set CB1 = SDB.Objects("CoverSaveOn")
	Set CB12 = SDB.Objects("CoverSaveIfEmpty")
	Set CB13 = SDB.Objects("SmallCoverSave")
	CB12.Common.Enabled = CB1.checked
	CB13.Common.Enabled = CB1.checked

End Sub

Sub SaveSheet(Sheet)

	Dim ini
	Set ini = SDB.IniFile
	Dim edt
	Set edt = Sheet.Common.ChildControl("ReleaseTag")
	ini.StringValue("DiscogsAutoTagWeb", "ReleaseTag") = GetCustom(edt.ItemIndex)
	Set edt = Sheet.Common.ChildControl("CatalogTag")
	ini.StringValue("DiscogsAutoTagWeb", "CatalogTag") = GetCustom(edt.ItemIndex)
	Set edt = Sheet.Common.ChildControl("CountryTag")
	ini.StringValue("DiscogsAutoTagWeb", "CountryTag") = GetCustom(edt.ItemIndex)
	Set edt = Sheet.Common.ChildControl("FormatTag")
	ini.StringValue("DiscogsAutoTagWeb", "FormatTag") = GetCustom(edt.ItemIndex)
	' Set edt = Sheet.Common.ChildControl("CheckStyleField")
	' ini.StringValue("DiscogsAutoTagWeb", "CheckStyleField") = GetCustom(edt.ItemIndex - 1)
	Set edt = Sheet.Common.ChildControl("StoreDate")
	ini.StringValue("DiscogsAutoTagWeb", "StoreDate") = edt.ItemIndex
	
	If Sheet.Common.ChildControl("ControlSaveImage1").Checked = False Then
		ini.StringValue("DiscogsAutoTagWeb", "CheckSaveImage") = 2
	Else
		If Sheet.Common.ChildControl("ControlSaveImage12").Checked = False Then
			ini.StringValue("DiscogsAutoTagWeb", "CheckSaveImage") = 0
		Else
			ini.StringValue("DiscogsAutoTagWeb", "CheckSaveImage") = 1
		End If
	End If

	If Sheet.Common.ChildControl("ControlSaveImage13").Checked = True Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckSmallCover") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckSmallCover") = false
	End If

	If Sheet.Common.ChildControl("CheckDontFillEmptyFields").Checked = True Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckDontFillEmptyFields") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckDontFillEmptyFields") = false
	End If

	Set checkbox = Sheet.Common.ChildControl("CheckOriginalDiscogsTrack")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckOriginalDiscogsTrack") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckOriginalDiscogsTrack") = false
	End If

	Set checkbox = Sheet.Common.ChildControl("CheckNewVersion")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckNewVersion") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckNewVersion") = false
	End If

	Set edt = Sheet.Common.ChildControl("ArtistSeparator")
	ini.StringValue("DiscogsAutoTagWeb", "ArtistSeparator") = edt.Text

	Set edt = Sheet.Common.ChildControl("FormatSeparator")
	ini.StringValue("DiscogsAutoTagWeb", "FormatSeparator") = edt.Text

	Set checkbox = Sheet.Common.ChildControl("CheckDCoff")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckDiscogsCollectionOff") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckDiscogsCollectionOff") = false
	End If

	Set checkbox = Sheet.Common.ChildControl("EditArtistLastSep")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "ArtistLastSeparator") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "ArtistLastSeparator") = false
	End If

	Set edt = Sheet.Common.ChildControl("AccessToken")
	ini.StringValue("DiscogsAutoTagWeb", "AccessToken") = edt.Text

	Set edt = Sheet.Common.ChildControl("AccessTokenSecret")
	ini.StringValue("DiscogsAutoTagWeb", "AccessTokenSecret") = edt.Text

	Set edt = Sheet.Common.ChildControl("LimitReleases")
	ini.StringValue("DiscogsAutoTagWeb", "LimitReleases") = edt.ItemIndex
	
	Set checkbox = Sheet.Common.ChildControl("ControlReplaceAccentChar")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckReplaceAccentChar") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckReplaceAccentChar") = false
	End If

	Set checkbox = Sheet.Common.ChildControl("UseAlbumForInitialSearch")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckSearchWithAlbum") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckSearchWithAlbum") = false
	End If

	Script.UnregisterAllEvents

End Sub

Sub InitSheet2(Sheet)

	Dim UI : Set UI = SDB.UI
	Set ini = SDB.IniFile
	If Not (ini Is Nothing) Then
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

		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords"), "|", ",")
	End If

	LyricistKeywords = ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords")
	ConductorKeywords = ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords")
	ProducerKeywords = ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords")
	ComposerKeywords = ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords")
	FeaturingKeywords = ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords")
	UnwantedKeywords = ini.StringValue("DiscogsAutoTagWeb","UnwantedKeywords")

	Dim GroupBox0
	Set GroupBox0 = UI.NewGroupBox(Sheet)
	GroupBox0.Caption = "Enter the keywords for linking with discogs"
	GroupBox0.Common.Hint = " - Separate the keywords with ','" & vbCrLf & " - You can use * as a wildcard" & vbCrLf & " - If you don't know what to enter here, let the keywords as is !!"
	GroupBox0.Common.SetRect 10, 10, 500, 275

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 20, 50, 25
	Label2.Caption = SDB.Localize("Lyricist")
	Set EditLyricist = UI.NewEdit(GroupBox0)
	EditLyricist.Common.SetRect 20, 35, 450, 35
	EditLyricist.Common.ControlName = "LyricistKeywords"
	EditLyricist.Text = LyricistKeywords


	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 60, 50, 25
	Label2.Caption = SDB.Localize("Conductor")
	Set EditConductor = UI.NewEdit(GroupBox0)
	EditConductor.Common.SetRect 20, 75, 450, 35
	EditConductor.Common.ControlName = "ConductorKeywords"
	EditConductor.Text = ConductorKeywords


	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 100, 50, 25
	Label2.Caption = SDB.Localize("Producer")
	Set EditProducer = UI.NewEdit(GroupBox0)
	EditProducer.Common.SetRect 20, 115, 450, 35
	EditProducer.Common.ControlName = "ProducerKeywords"
	EditProducer.Text = ProducerKeywords


	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 140, 50, 25
	Label2.Caption = SDB.Localize("Composer")
	Set EditComposer = UI.NewEdit(GroupBox0)
	EditComposer.Common.SetRect 20, 155, 450, 35
	EditComposer.Common.ControlName = "ComposerKeywords"
	EditComposer.Text = ComposerKeywords

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 180, 50, 25
	Label2.Caption = SDB.Localize("Featuring")
	Set EditFeaturing = UI.NewEdit(GroupBox0)
	EditFeaturing.Common.SetRect 20, 195, 450, 35
	EditFeaturing.Common.ControlName = "FeaturingKeywords"
	EditFeaturing.Text = FeaturingKeywords

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 220, 50, 25
	Label2.Caption = SDB.Localize("Ignore this tags")
	Set EditUnwanted = UI.NewEdit(GroupBox0)
	EditUnwanted.Common.SetRect 20, 235, 450, 35
	EditUnwanted.Common.ControlName = "UnwantedKeywords"
	EditUnwanted.Common.Hint = "Enter the keywords you don't like to store in involved people" & vbCrLf & "You can use * as a wildcard"
	EditUnwanted.Text = UnwantedKeywords

End Sub

Sub SaveSheet2(Sheet)

	Dim ini
	Set ini = SDB.IniFile
	Dim edt
	Dim tmp, x, editText

	Set edt = Sheet.Common.ChildControl("LyricistKeywords")
	If edt.Text <> "" Then
		tmp = Split(edt.Text, ",")
		editText = ""
		For each x in tmp
			editText = editText & Trim(x) & ","
		Next
		ini.StringValue("DiscogsAutoTagWeb", "LyricistKeywords") = Left(editText, Len(editText)-1)
	Else
		ini.StringValue("DiscogsAutoTagWeb", "LyricistKeywords") = ""
	End If

	Set edt = Sheet.Common.ChildControl("ConductorKeywords")
	If edt.Text <> "" Then
		tmp = Split(edt.Text, ",")
		editText = ""
		For each x in tmp
			editText = editText & Trim(x) & ","
		Next
		ini.StringValue("DiscogsAutoTagWeb", "ConductorKeywords") = Left(editText, Len(editText)-1)
	Else
		ini.StringValue("DiscogsAutoTagWeb", "ConductorKeywords") = ""
	End If

	Set edt = Sheet.Common.ChildControl("ProducerKeywords")
	If edt.Text <> "" Then
		tmp = Split(edt.Text, ",")
		editText = ""
		For each x in tmp
			editText = editText & Trim(x) & ","
		Next
		ini.StringValue("DiscogsAutoTagWeb", "ProducerKeywords") = Left(editText, Len(editText)-1)
	Else
		ini.StringValue("DiscogsAutoTagWeb", "ProducerKeywords") = ""
	End If

	Set edt = Sheet.Common.ChildControl("ComposerKeywords")
	If edt.Text <> "" Then
		tmp = Split(edt.Text, ",")
		editText = ""
		For each x in tmp
			editText = editText & Trim(x) & ","
		Next
		ini.StringValue("DiscogsAutoTagWeb", "ComposerKeywords") = Left(editText, Len(editText)-1)
	Else
		ini.StringValue("DiscogsAutoTagWeb", "ComposerKeywords") = ""
	End If

	Set edt = Sheet.Common.ChildControl("FeaturingKeywords")
	If edt.Text <> "" Then
		tmp = Split(edt.Text, ",")
		editText = ""
		For each x in tmp
			editText = editText & Trim(x) & ","
		Next
		ini.StringValue("DiscogsAutoTagWeb", "FeaturingKeywords") = Left(editText, Len(editText)-1)
	Else
		ini.StringValue("DiscogsAutoTagWeb", "FeaturingKeywords") = ""
	End If

	Set edt = Sheet.Common.ChildControl("UnwantedKeywords")
	If edt.Text <> "" Then
		tmp = Split(edt.Text, ",")
		editText = ""
		For each x in tmp
			editText = editText & Trim(x) & ","
		Next
		ini.StringValue("DiscogsAutoTagWeb", "UnwantedKeywords") = Left(editText, Len(editText)-1)
	Else
		ini.StringValue("DiscogsAutoTagWeb", "UnwantedKeywords") = ""
	End If

End Sub

Function GetCustom(Index)

	Dim ini : Set ini = SDB.IniFile
	If Index = 0 Then GetCustom = "Custom1"
	If Index = 1 Then GetCustom = "Custom2"
	If Index = 2 Then GetCustom = "Custom3"
	If Index = 3 Then GetCustom = "Custom4"
	If Index = 4 Then GetCustom = "Custom5"
	If Index = 5 Then GetCustom = "Don't save"
	If Index = -1 Then GetCustom = "Default (stored with Genre)"
	If Index = 6 Then GetCustom = "ISRC"
	If Index = 7 Then GetCustom = "Grouping"
	If Index = 8 Then GetCustom = "Encoder"
	If Index = 9 Then GetCustom = "Copyright"

End Function
