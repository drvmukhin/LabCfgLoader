// Constants
var MINIMUM_CLIENT_VERSION = 14;
var REDIRECT_TO_REACH_SL_OVERRIDE = 'sl=';

// This is CU2_HF3 build # as this is when OC / AOC plugin dlls were updated in the build.
var MINOR_CLIENT_VERSION_FOR_CU2 = "7577.280";
var INTENT_BASED_CHROME_VERSION = 25;
var CHROME_CANARY_VERSION = 42;

// Plugin Loaders for each client
var pluginLoaderOC = null;
var pluginLoaderSamara = null;
var pluginLoaderAOC = null;

// Plugin Objects for each client
var pluginObjectOC = null;
var pluginObjectSamara = null;
var pluginObjectAOC = null;

// Version info for each client
var majorVersionOC = null;
var majorVersionSamara = null;
var majorVersionAOC = null;

var minorVersionOC = null;
var minorVersionAOC = null;
var minorVersionSamara = null;

var majorVersionOCCapability = null;
var majorVersionAOCCapability = null;
var majorVersionSamaraCapability = null;

var defaultExperienceVersion = "1400";
var newExperienceVersion = "1500";
var invalidConfErrorCode = "6";
var serverBusyErrorCode = "7";
var failoverOrFailback = "8";

var ResourceURL = "";

var isMobileDevice = false;
var isNokiaDevice = false;
var isAndroidDevice = false;
var isWinPhoneDevice = false;
var isIPhoneDevice = false;
var isIPadDevice = false;

var chromeBrowserVersion = "";

// Initialize with some default values
var lyncJoinConferenceUrl = "lync://";
var lync15JoinConferenceUrl = "lync15:";
var lync15DesktopJoinConfUrl = "lync15classic:";
var lync15MobileJoinConfUrl = "lync15://";
var mlxJoinConfUrl = "lync15mlx:";

var confJoinParams = "confjoin?url=";

var isImmersiveIE = false;
var noAppTimeout = 1500;

var loading = "true";

var joinOptionUsingMacClient = 2;

//
// Initialize resource strings before we try to use them.
//
var txt_languageSettingsLabel = "";
var txt_launchRichClientHeaderLabel = "";
var txt_launchRichClientTextLabel = "";
var txt_unableToJoinLabel = "";
var txt_onlineHelpLink = "";
var txt_copyRightTextLabel = "";
var textDirection = "";
var txt_joinUsingReachLink = "";
var txt_connecting = "";
var txt_unableToLaunchLyncMobile = "";
var txt_unableToLaunchLyncMobile2 = "";
var txt_unsupportedMobileDevice = "";

var txt_immersiveIESwitch = "";
var txt_64bitbrowserUnsupportedLabel = "";
var txt_64bitUnsupportedText1 = "";
var txt_64bitUnsupportedOption1 = "";
var txt_64bitJoinUsingLync = "";
var txt_64bitUnsupportedOption2 = "";

// Add any new resource strings here.

var requestArray = new Array(10);


function MainForm()
{
    this.connection = new ConnectionObject();
}

MainForm.prototype.GetMinorVersion = function (fullVersion) {
    if (fullVersion == null) {
        return null;
    }

    var minorVersion = 0.0;

    // Full version number string format: "4.0.1234.5678"
    // Minor version number string format: "1234.5678"
    // Ignore any non-numeric characters preceding the version number start
    if (/(\D*)(\d\.\d\.)(\d{4}\.\d+)/.test(fullVersion)) {
        var minorVersionString = RegExp.$3;
        minorVersion = parseFloat(minorVersionString);

    }
    return minorVersion;
}

MainForm.prototype.GetMajorVersion = function(fullVersion)
{
    if (fullVersion == null)
    {
        return null;
    }

    var majorVersion = "";

    // Full version number string format: "4.0.1234.5678"
    // Major version number string format: "4.0"
    // Ignore any non-numeric characters preceding the version number start
    if (/(\D*)(\d\.\d)(\.\d{4}\.\d+)/.test(fullVersion)) {
        majorVersion = RegExp.$2;
    }
    return majorVersion;
}

MainForm.prototype.listener = function (event) {
    // Register an listener to close the window when other iframe
    // send "closeWindow" message.
    if (event.data == "closeWindow") {
        MainForm.prototype.PrepareToClosePage();
    }
}

MainForm.prototype.GetClientBasedOnCapability = function(capabilities)
{
    if (capabilities == null)
    {
        return null;
    }

    //assign it as 14 by default
    majorVersionCapability = 14;

    var clientVersion = "";
    if (/(([a-z\-]*):(\d+\.\d+)(,*))*/.test(capabilities)) {
        majorVersionEncoded = RegExp.$2;
        minorVersionEncoded = RegExp.$3;
        // We dont look at the minorVersion yet since we dont need to block any CUs in 15 yet.

        if (majorVersionEncoded == "ms-launch-lync")
        {
            majorVersionCapability = 15;
        }

        if (majorVersionEncoded == "ms-launch-ocsmeet")
        {
            majorVersionCapability = 14;
        }

        if (majorVersionEncoded == "ms-launch-conf")
        {
            majorVersionCapability = 13;
        }

    }
    return majorVersionCapability;
}

MainForm.prototype.OnLoad = function () {

    // un-hide the Language dropdown.
    // hide the 15 experience table for now
    document.getElementById("mainTable15").style.display = "none";
    document.getElementById("languageSettingsDiv").style.display = "block";
    document.getElementById("languageSettingsDiv15").style.display = "block";
    document.getElementById("languageSettingsDivMobile").style.display = "block";
    document.getElementById("errorLogo").style.display = "none";
    document.getElementById("footer").className = "";

    this.UpdateSelectedLanguage(currentLanguage);

    // Fetch the current language resources.
    this.GetUpdatedResources(currentLanguage);
}

MainForm.prototype.OnLoadComplete = function () {
    loading = "false";

    lyncJoinConferenceUrl = mobileW1ProtocolHandler.toLowerCase() + confJoinParams;
    lync15MobileJoinConfUrl = mobileW2ProtocolHandler.toLowerCase() + confJoinParams;
    lync15DesktopJoinConfUrl = lync15ClassicProtocolHandler.toLowerCase() + confJoinParams;
    lync15JoinConferenceUrl = lync15CommonProtocolHandler.toLowerCase() + confJoinParams;
    mlxJoinConfUrl = mlxProtocolHandler.toLowerCase() + confJoinParams;

    isMobileDevice = isMobile.toLowerCase();

    var unsupportedDevice = isUnsupported.toLowerCase();
    if (unsupportedDevice == "true") {

        if (userExperience.toLowerCase() != defaultExperienceVersion) {
            document.getElementById("UnknownMobileDeviceDiv15").style.display = "block";
            document.getElementById("unknownMobileDeviceLabel15").innerHTML = txt_unsupportedMobileDevice;
        }
        else {
            document.getElementById("UnknownMobileDeviceDiv").style.display = "block";
            document.getElementById("unknownMobileDeviceLabel").innerHTML = txt_unsupportedMobileDevice;
        }
        document.getElementById("launchReachLink").style.display = "none";
        return;
    }
    // Launching a protocol (lync://) from inside iFrames no longer works on iOS9 Safari, 
    // so handling it on the main window
    if ((isMobileDevice == "true") && (isIPhone.toLowerCase() == "true") && isSafari()) {
        this.LaunchOnSafariForIOS();
        return;
    }

    //
    // Make sure Meeting is valid
    //
    var valid = validMeeting.toLowerCase();
    if (valid == "false") {
        this.ShowError();
        return;
    }
    var fwdurl = domainOwnerJoinLauncherUrl.toLowerCase();
    if (fwdurl != "") {
        if (isIE8OrLater()) {
            if (window.addEventListener) {
                window.addEventListener("message", MainForm.prototype.listener, false);
            } else {
                window.attachEvent("onmessage", MainForm.prototype.listener);
            }
        }

        domainOwnerJoinLauncherUrl = domainOwnerJoinLauncherUrl + "&iframe=1";

        // Redirect the user to the correct join launcher
        this.RedirectToReach(domainOwnerJoinLauncherUrl);
        return;
    }

    //
    // Do something here as cracking was received
    //
    var reachRequested = reachClientRequested.toLowerCase();
    if (reachRequested == "true") {
        this.RedirectToReach(reachURL);
        return;
    }

    var lwaRequested = htmlLwaClientRequested.toLowerCase();
    if (lwaRequested == "true") {
        // TODO: revert change to display message post LWA DF
        var resp = confirm("Welcome! You are participating in Lync 15 Technical Preview and today is Lync Web App day (LWA Day).\n\n" +
            "Every LWA Day, you'll be joining meetings using Lync Web App by default instead of the installed Lync client. " +
            "Please take this opportunity to put our web client through its paces 'cause we want to know what you think.\n\n " +
            "If you need to, you can switch to the installed Lync client by clicking Cancel button.");


        if (resp) {
            this.RedirectToReach(reachURL);
            return;
        }
    }

    if (isMobileDevice == "true") {
        isNokiaDevice = isNokia.toLowerCase();
        isAndroidDevice = isAndroid.toLowerCase();
        isWinPhoneDevice = isWinPhone.toLowerCase();
        isIPhoneDevice = isIPhone.toLowerCase();
        isIPadDevice = isIPad.toLowerCase();

        chromeBrowserVersion = chromeVersion.toLowerCase();

        lyncJoinConferenceUrl = lyncJoinConferenceUrl + document.URL;
        lync15MobileJoinConfUrl = lync15MobileJoinConfUrl + document.URL;

        if (isNokiaDevice == "true" || isIPhoneDevice == "true" || isIPadDevice == "true" || isWinPhoneDevice == "true" || isAndroidDevice == "true") {

            var lyncDownloadFromOviUrl = "";


            if (isNokiaDevice == "true") {
                lyncDownloadFromOviUrl = "http://go.microsoft.com/fwlink/?LinkId=217188";
            }

            if (isIPhoneDevice == "true") {
                lyncDownloadFromOviUrl = "http://go.microsoft.com/fwlink/?LinkId=217185";
            }

            if (isIPadDevice == "true") {
                lyncDownloadFromOviUrl = "http://go.microsoft.com/fwlink/?LinkId=217187";
            }

            if (isWinPhoneDevice == "true") {
                lyncDownloadFromOviUrl = "http://go.microsoft.com/fwlink/?LinkId=217184";
            }

            if (isAndroidDevice == "true") {
                lyncDownloadFromOviUrl = "http://go.microsoft.com/fwlink/?LinkId=217189";
            }

            // Try to launch lync with the lyncJoinLaunchUrl. Start the timeout in parallel,
            // After timeout, redirect to install page.
            document.getElementById("launchReachLink").style.display = "none";

            if (userExperience.toLowerCase() != defaultExperienceVersion) {
                document.getElementById("mobileinterimDiv15").style.display = "block";
                document.getElementById("connectingLabel15").innerHTML = txt_connecting;
            }
            else {
                document.getElementById("mobileinterimDiv").style.display = "block";
                document.getElementById("connectingLabel").innerHTML = txt_connecting;
            }

            // If device is Android and Chrome browser major version is greater than 25 then we have to use Intent based URLs
            if ((isAndroidDevice == "true") && (chromeBrowserVersion !== "") &&
                (parseInt(chromeBrowserVersion, 10) > INTENT_BASED_CHROME_VERSION)) {
                // TODO: think if we need to serve this url from service side
                window.top.location = "intent://" + confJoinParams + document.URL + "#Intent;scheme=lync;package=com.microsoft.office.lync15;end";
            } else {
                window.frames["launchReachFrame"].location = lyncJoinConferenceUrl;
            }

            setTimeout(function () {
                if (userExperience.toLowerCase() != defaultExperienceVersion) {
                    document.getElementById("mobileappstoreDiv15").style.display = "block";
                    //document.getElementById("copyRightInfoDiv").style.display = "block";

                    document.getElementById("mobileinterimDiv15").style.display = "none";
                    document.getElementById("mobileappstoreLabel15").innerHTML = txt_unableToLaunchLyncMobile;

                    document.getElementById("mobileappstoreLabel315").innerHTML = txt_unableToLaunchLyncMobile3;

                    document.getElementById("mobileappstoreLabel215").setAttribute('href', lyncDownloadFromOviUrl);
                    document.getElementById("mobileappstoreLabel215").innerHTML = txt_unableToLaunchSfbMobile2;
                }
                else {
                    document.getElementById("mobileappstoreDiv").style.display = "block";
                    document.getElementById("mobileinterimDiv").style.display = "none";
                    document.getElementById("mobileappstoreLabel").innerHTML = txt_unableToLaunchLyncMobile;

                    document.getElementById("mobileappstoreLabel2").setAttribute('href', lyncDownloadFromOviUrl);
                    document.getElementById("mobileappstoreLabel2").innerHTML = txt_unableToLaunchLyncMobile2;
                }
            }, noAppTimeout);
        }
        return;
    }

    //
    // Initialize version info of all clients
    //
    this.InitializeVersionInformationOfAllClients();

    // MLX DCR: Check if MLX has requested an escalation to Lync desktop.
    var desktopescalate = escalateToDesktop.toLowerCase();

    // Before checking for the plugin, now we should look for the Lync15:// protocol handler
    // only on Windows >= Win8 or ARM devices or Chrome Canary Browser versions 
    var tryLyncProtocolHandler = isWin8OrLater() || isArm() || (parseInt(chromeVersion, 10) >= CHROME_CANARY_VERSION); 
    if (tryLyncProtocolHandler) {

        // Check if the new API msLaunchUri is available or not.
        if (navigator.msLaunchUri) {

            // If its available, look if Escalation is ON from MLX
            if (desktopescalate == "true") {
                // Yes, launch using Lync15Desktop to ensure to NOT launch MLX
                // if the detection/launch was successful, close the page.
                // if it was unsuccessful, try to launch a client usig Active X next.
                navigator.msLaunchUri(lync15DesktopJoinConfUrl + document.URL, this.PrepareToClosePage, this.LaunchUsingActiveX);
            }
            else {
                // No, launch using common Lync15:// to ensure we launch whichever is default MLX, W15 or W14 CU7
                navigator.msLaunchUri(lync15JoinConferenceUrl + document.URL, this.PrepareToClosePage, this.LaunchUsingActiveX);
            }
            return;
        }

        // To Do: Remove after IE removes the support for deprecated msProtocols API.
        if (navigator.msProtocols) {
            // This means the protocol hanlder detection API is available.
            // we must be on IE10 or mo-bro or higher
            //       if (lync15mlx:// not found) {
            //          if (can use ActiveX) // IE 10 classic
            //              use ActiveX to launch
            //          else // IE 10 Windows Store Mode
            //              use Protocol Handler
            //          }

            // References:http://msdn.microsoft.com/en-us/library/hh772146(v=vs.85).aspx 
            if (navigator.msProtocols["lync15mlx"] == undefined || desktopescalate == "true") {
                // => This means MLX is NOT installed or is installed but we have to pretend as if it isn't (desktopescalation scenario)

                // Check if conf: is Registered
                if (navigator.msProtocols["conf"] != undefined) {
                    // => Conf: is registered so W14 or above client is installed.

                    // Check if the plugin can be loaded.
                    if (majorVersionAOC == null && majorVersionOC == null && majorVersionSamara == null) {
                        // => Plugin info can not be found/loaded

                        //References: http://blogs.msdn.com/b/ie/archive/2011/05/02/activex-filtering-for-developers.aspx
                        if (typeof window.external.msActiveXFilteringEnabled != "undefined") {
                            if (window.external.msActiveXFilteringEnabled() == false) {
                                // => We're in immersive mode as activeXFiltering is false but plugin is NOT found.

                                // Check if Lync 15 is installed using the Lync15 specific protocol handler
                                if (navigator.msProtocols["lync15desktop"] != undefined) {
                                    window.frames["launchReachFrame"].location = lync15DesktopJoinConfUrl + encodeURIComponent(document.URL);
                                    return;
                                }
                            } // msActiveXFilteringCheck
                        }
                    } // Plugin can be loaded or not
                } // conf check
            }
            else {
                // MLX is installed and desktop escalation is false
                // Launch Lync15:// blindly as MLX and all other clients are expected to register for Lync15://
                window.frames["launchReachFrame"].location = lync15JoinConferenceUrl + encodeURIComponent(document.URL);
                return;
            }
        }

        if (isArm()) {
            this.RedirectToReach(reachURL);
            return;
        }

        if (parseInt(chromeVersion, 10) >= CHROME_CANARY_VERSION) {
            window.frames["launchReachFrame"].location = lync15JoinConferenceUrl + encodeURIComponent(document.URL);
            MainForm.prototype.DisplayInstalledClientLaunchedPage("oc");
            return;
        }

    }

    this.LaunchUsingActiveX();
}


MainForm.prototype.LaunchOnSafariForIOS = function () {
    lyncJoinConferenceUrl = lyncJoinConferenceUrl + document.URL;
    lync15MobileJoinConfUrl = lync15MobileJoinConfUrl + document.URL;

    // Try to launch lync with the lyncJoinLaunchUrl. Start the timeout in parallel,
    // After timeout, redirect to store page.
    document.getElementById("launchReachLink").style.display = "none";

    if (userExperience.toLowerCase() != defaultExperienceVersion) {
        document.getElementById("mobileinterimDiv15").style.display = "block";
        document.getElementById("connectingLabel15").innerHTML = txt_connecting;
    }
    else {
        document.getElementById("mobileinterimDiv").style.display = "block";
        document.getElementById("connectingLabel").innerHTML = txt_connecting;
    }

    console.log("Launching in the mainframe IOS Safari");
    window.top.location = lyncJoinConferenceUrl;
    setTimeout(function () {
        window.top.location = "http://go.microsoft.com/fwlink/?LinkId=217185";
    }, 0);
}

MainForm.prototype.LaunchUsingActiveX = function()
{
    var isHtmlLwaEnabled = isLwaEnabled.toLowerCase();

    // Treat MAC OS separately. We need to take the users to LWA if the 15 HTML LWA is enabled and meeting join
    // preference is not MacClient
    if (isMacOS() && isHtmlLwaEnabled == "true") {
        var joinUsingLWA = true;

        try
        {
            joinUsingLWA = (pluginObjectOC.object.GetMeetingJoinPreference() != joinOptionUsingMacClient);
        }
        catch(ex)
        {
            // Unable to get preference from plugin, using LWA.
            joinUsingLWA = true;
        }

        if (joinUsingLWA)
        {
            MainForm.prototype.RedirectToReach(reachURL);
            return;
        }
    }

    //
    // Now that we have the installed clients and
    // their version info, figure out which client
    // to launch and try to launch it.
    //
    var launched = "false";

    var blockprecu2clients = blockPreCU2Clients.toLowerCase();

    try {
        if ((majorVersionOC != null) && (majorVersionOCCapability >= MINIMUM_CLIENT_VERSION)) {
            if (blockprecu2clients == "true" && minorVersionOC < MINOR_CLIENT_VERSION_FOR_CU2) {
                // do not launch OC here as we need to block these.
                // instead show error message and an option to launch LWA
                MainForm.prototype.ShowClientBlockScreen();
                return;
            }
            else {
                // Launch using OC plugin
                // Note: OC Plugin launches whichever client ran last (OC/Samara)
                pluginObjectOC.object.LaunchUCClient(escapedXML);
                launched = "true";
                MainForm.prototype.DisplayInstalledClientLaunchedPage("oc");
            }
        }
        else if ((majorVersionSamara != null) && (majorVersionSamaraCapability >= MINIMUM_CLIENT_VERSION)) {
            // Launch using Samara plugin
            // Note: Samara Plugin launches whichever client ran last (OC/Samara)
            pluginObjectSamara.object.LaunchUCClient(escapedXML);
            launched = "true";
            MainForm.prototype.DisplayInstalledClientLaunchedPage("oc");
        }
        // make sure to launch AOC only when HTML LWA is NOT enabled
        else if ((isHtmlLwaEnabled == "false") && ((majorVersionAOC != null) && (majorVersionAOCCapability >= MINIMUM_CLIENT_VERSION))) {
            if (blockprecu2clients == "true" && minorVersionAOC < MINOR_CLIENT_VERSION_FOR_CU2) {
                // do not launch OC here as we need to block these.
                // instead show error message and an option to launch LWA
                MainForm.prototype.ShowClientBlockScreen();
                return;
            }
            else {
                // Launch using AOC plugin
                // Note: AOC Plugin launches only AOC
                pluginObjectAOC.object.LaunchUCClient(escapedXML);
                launched = "true";
                MainForm.prototype.DisplayInstalledClientLaunchedPage("aoc");
            }
        }
    } catch (ex) {
        // Failed to launch client
        launched = "false";
    }

    //
    // Unload all Plugins
    //
    MainForm.prototype.UnloadAllPlugins();

    if (launched == "false") {
        // Either we failed to detect an installed client OR
        // the installed client is not the right version OR
        // launching the installed client failed.
        // In any case, redirect to Reach.

        MainForm.prototype.RedirectToReach(reachURL);
    }
    else {
        // We managed to successfully launch the client to join
        // the meeting, try to close the browser window on the 
        // browsers that support closing it silently
        MainForm.prototype.PrepareToClosePage();
    }
}

MainForm.prototype.PrepareToClosePage = function () {

    // We only close the page on selected IE browsers.
    // We dont do it on FF/Chrome etc other browsers.
    // So, if we're not on IE, lets not reload the page with ?closeme=1 in the first place

    if (!isIE()) {
        return;
    }

    var currentUrl = window.location.href;

    if (currentUrl.indexOf("&iframe=1") < 0) {
        // Directly close page from top-level window
        MainForm.prototype.ClosePage();
        return;
    }
    else {
        // Post message to the parent window to close the page
        if (isIE8OrLater()) {
            window.parent.postMessage("closeWindow", "*");
        }
    }
}

MainForm.prototype.InitializeVersionInformationOfAllClients = function()
{
    try
    {
        navigator.plugins.refresh();
    } catch (ex) {
        // no need to do anything here
    }

    // First determine the browser tag
    // needed to look up the config for
    // the clients
    var configTag = GetBrowserTag();

    this.InitializeVersionInformationForOC(configTag);
    this.InitializeVersionInformationForSamara(configTag);
    this.InitializeVersionInformationForAOC(configTag);
}

MainForm.prototype.InitializeVersionInformationForOC = function(configTag)
{
    var pluginConfig = GetConfigForClient(InstalledClient.OC, configTag);
    if (!pluginConfig)
    {
        return;
    }

    pluginLoaderOC = new PluginLoader();
    pluginLoaderOC.Initialize(
                "OC",
                pluginConfig.Version_Name,
                pluginConfig.Version_CLSID,
                null);

    pluginObjectOC = pluginLoaderOC.LoadPlugin();
    if (!pluginObjectOC.object)
    {
        return;
    }

    try
    {
        var fullVersionOC = pluginObjectOC.object.GetVersionString();
        majorVersionOC = this.GetMajorVersion(fullVersionOC);
        minorVersionOC = this.GetMinorVersion(fullVersionOC);

        if (isMacOS())
        {
            // Bug 3305806: On Mac (if execution gets to this point without throwing),
            // GetSupportedProtocolVersionString also throws an exception; when it does, the plugin
            // itself enters an unusable state and launch of the client will fail. Work around this
            // by using the default major version for Mac.

            majorVersionOCCapability = 14;
        }
        else
        {
            var clientCapabilities = pluginObjectOC.object.GetSupportedProtocolVersionString();
            majorVersionOCCapability = this.GetClientBasedOnCapability(clientCapabilities);
        }
    } catch (ex) {
        // Unable to get version details, continue anyway
        // Bug # 3108649: For MAC, 14.0.2 version plugin has a regression that
        // there is no version API provided. So work around this
        // issue by using the default version only for MAC.
        if (isMacOS()) {
            majorVersionOC = '4.0';
            minorVersionOC = 7577.280;
            majorVersionOCCapability = 14;
        }
    }
}

MainForm.prototype.InitializeVersionInformationForSamara = function(configTag)
{
    var pluginConfig = GetConfigForClient(InstalledClient.Samara, configTag);
    if (!pluginConfig)
    {
        return;
    }

    pluginLoaderSamara = new PluginLoader();
    pluginLoaderSamara.Initialize(
                "Samara",
                pluginConfig.Version_Name,
                pluginConfig.Version_CLSID,
                null);

    pluginObjectSamara = pluginLoaderSamara.LoadPlugin();
    if (!pluginObjectSamara.object)
    {
        return;
    }

    try
    {
        var fullVersionSamara = pluginObjectSamara.object.GetVersionString();
        majorVersionSamara = this.GetMajorVersion(fullVersionSamara);

        var clientCapabilities = pluginObjectSamara.object.GetSupportedProtocolVersionString();
        majorVersionSamaraCapability = this.GetClientBasedOnCapability(clientCapabilities);

    } catch (ex) {
        // Unable to get version details, continue anyway
    }
}

MainForm.prototype.InitializeVersionInformationForAOC = function(configTag)
{
    var pluginConfig = GetConfigForClient(InstalledClient.AOC, configTag);
    if (!pluginConfig)
    {
        return;
    }

    pluginLoaderAOC = new PluginLoader();
    pluginLoaderAOC.Initialize(
                "AOC",
                pluginConfig.Version_Name,
                pluginConfig.Version_CLSID,
                null);

    pluginObjectAOC = pluginLoaderAOC.LoadPlugin();
    if (!pluginObjectAOC.object)
    {
        return;
    }

    try
    {
        var fullVersionAOC = pluginObjectAOC.object.GetVersionString();
        majorVersionAOC = this.GetMajorVersion(fullVersionAOC);
        minorVersionAOC = this.GetMinorVersion(fullVersionAOC);

        var clientCapabilities = pluginObjectAOC.object.GetSupportedProtocolVersionString();
        majorVersionAOCCapability = this.GetClientBasedOnCapability(clientCapabilities);

    } catch (ex) {
        // Unable to get version details, continue anyway
    }
}

MainForm.prototype.UnloadAllPlugins = function()
{
    // Unload all Plugins so that references to the DLLs are released

    // Unload OC Plugin
    if (pluginLoaderOC != null)
    {
        pluginLoaderOC.UnloadPlugin();
    }

    // Unload Samara Plugin
    if (pluginLoaderSamara != null)
    {
        pluginLoaderSamara.UnloadPlugin();
    }

    // Unload AOC Plugin
    if (pluginLoaderAOC != null)
    {
        pluginLoaderAOC.UnloadPlugin();
    }
}

MainForm.prototype.ShowClientBlockScreen = function () {
    if (userExperience.toLowerCase() != defaultExperienceVersion) {
        document.getElementById("UnsupportedClientBlockDiv15").style.display = "block";

        document.getElementById("UnsupportedClientBlockMessageLabel15").innerHTML = txt_UnsupportedSfbVersion1;
        document.getElementById("JoinUsingLWAOption115").innerHTML = txt_UnsupportedSfbVersion2;
        document.getElementById("JoinMeetingUsingLWA_Link15").innerHTML = txt_UnsupportedSfbVersion3;
        document.getElementById("JoinMeetingUsingLWA_Link15").href = reachURL;
    }
    else {
        document.getElementById("UnsupportedClientBlockDiv").style.display = "block";

        document.getElementById("UnsupportedClientBlockMessageLabel").innerHTML = txt_UnsupportedLyncVersion1;
        document.getElementById("JoinUsingLWAOption1").innerHTML = txt_UnsupportedLyncVersion2;
        document.getElementById("JoinMeetingUsingLWA_Link").innerHTML = txt_UnsupportedLyncVersion3;
        document.getElementById("JoinMeetingUsingLWA_Link").href = reachURL;

    }
}

MainForm.prototype.SelectConferenceError2 = function () {
    if (errorCode == invalidConfErrorCode) {
        return conferenceErrorExpiredMeeting;
    }
    else if (errorCode == failoverOrFailback) {
        return conferenceErrorFailoverOrFailback;
    }
    else {
        return conferenceError2;
    }
}

MainForm.prototype.ShowError = function () {

    if (userExperience.toLowerCase() != defaultExperienceVersion) {
        document.getElementById("joinLauncherErrorDiv15").style.display = "block";

        if (errorCode == serverBusyErrorCode) {
            document.getElementById("errorTextLabel15").innerHTML = conferenceErrorServerBusy1;
            document.getElementById("checkUrlLabel15").innerHTML = conferenceErrorServerBusy2;
        }
        else {
            document.getElementById("errorTextLabel15").innerHTML = conferenceError1;
            document.getElementById("checkUrlLabel15").innerHTML = MainForm.prototype.SelectConferenceError2();
        }

        document.getElementById("errorTextLabel15").style.display = "block";
        document.getElementById("checkUrlLabel15").style.display = "block";

        document.getElementById("diagLabel15").innerHTML = diag;
        document.getElementById("diagLabel15").style.display = "block";
        var diagInfoTextString = diagInfo.toString();
        document.getElementById("diagInfoText15").value = diagInfoTextString;
        document.getElementById("diagInfoText15").style.display = "none";
    }
    else { // 1400
        if (isMobileDevice == "true") {
            document.getElementById("joinLauncherErrorDivMobile").style.display = "block";

            if (errorCode == serverBusyErrorCode) {
                document.getElementById("errorTextLabelMobile").innerHTML = conferenceErrorServerBusy1;
                document.getElementById("checkUrlLabelMobile").innerHTML = conferenceErrorServerBusy2;
            }
            else {
                document.getElementById("errorTextLabelMobile").innerHTML = conferenceError1;
                document.getElementById("checkUrlLabelMobile").innerHTML = MainForm.prototype.SelectConferenceError2();
            }

            document.getElementById("checkUrlLabelMobile").style.display = "block";
            document.getElementById("errorTextLabelMobile").style.display = "block";

            document.getElementById("diagLabelMobile").innerHTML = diag;
            document.getElementById("diagLabelMobile").style.display = "block";
            var diagInfoTextString = diagInfo.toString();
            document.getElementById("diagInfoTextMobile").value = diagInfoTextString;
            document.getElementById("diagInfoTextMobile").style.display = "none";
        }
        else { // non-mobile
            document.getElementById("joinLauncherErrorDiv").style.display = "block";

            if (errorCode == serverBusyErrorCode) {
                document.getElementById("errorTextLabel").innerHTML = conferenceErrorServerBusy1;
                document.getElementById("checkUrlLabel").innerHTML = conferenceErrorServerBusy2;
            }
            else {
                document.getElementById("errorTextLabel").innerHTML = conferenceError1;
                document.getElementById("checkUrlLabel").innerHTML = MainForm.prototype.SelectConferenceError2();
            }

            document.getElementById("errorTextLabel").style.display = "block";
            document.getElementById("checkUrlLabel").style.display = "block";

            document.getElementById("diagLabel").innerHTML = diag;
            document.getElementById("diagLabel").style.display = "block";
            var diagInfoTextString = diagInfo.toString();
            document.getElementById("diagInfoText").value = diagInfoTextString;
            document.getElementById("diagInfoText").style.display = "none";
        }
    }
}

MainForm.prototype.DisplayInstalledClientLaunchedPage = function (launchedClient) {
    if (userExperience.toLowerCase() != defaultExperienceVersion) {

        // Show the Rich client launched text, etc.
        document.getElementById("launchRichClientDiv15").style.display = "block";

        // Tell Reach which client was launched so the Reach Landing
        // Page can be more intelligent about the links it displays
        // to the user.
        var fullReachURL = reachURL + "&launched=" + launchedClient;

        document.getElementById("JoinUsingLWAOptionHelp115").innerHTML = txt_UnsupportedSfbVersion2;
        document.getElementById("JoinMeetingUsingLWA2_Link15").innerHTML = txt_UnsupportedSfbVersion3;

        // Update the link for launching Reach.
        document.getElementById("JoinMeetingUsingLWA2_Link15").href = fullReachURL;
    }
    else {

        // Show the Rich client launched text, etc.
        document.getElementById("launchRichClientDiv").style.display = "block";

        // Show the Contact Support text with Reach option as well as the link to
        // launch Reach.
        document.getElementById("contactSupportLabelWithReachOption").style.display = "block";
        document.getElementById("launchReachLink").style.display = "block";

        // Tell Reach which client was launched so the Reach Landing
        // Page can be more intelligent about the links it displays
        // to the user.
        var fullReachURL = reachURL + "&launched=" + launchedClient;

        // Update the link for launching Reach.
        document.getElementById("launchReachLink").href = fullReachURL;
    }
    // Hide the iFrame we use to launch Reach and expand it to 100%
    // so it occupies the entire area of the page.
    document.getElementById("launchReachDiv").style.display = "none";
    document.getElementById("launchReachFrame").style.width = "0px";
    document.getElementById("launchReachFrame").style.height = "0px";
    document.getElementById("launchReachFrame").style.display = "";

}

MainForm.prototype.RedirectToReach = function (url) {
    // Launch Reach from an iFrame, so that the
    // URL does not change in the Browser window.

    // Hide the main table that has all the page content.
    document.getElementById("mainTable").style.display = "none";
    document.getElementById("mainTable15").style.display = "none";
    document.getElementById("mainTablemobile").style.display = "none";

    // Un-hide the iFrame we use to launch Reach.
    document.getElementById("launchReachDiv").style.display = "block";
    document.getElementById("launchReachFrame").style.width = "100%";
    document.getElementById("launchReachFrame").style.height = "100%";
    document.getElementById("launchReachFrame").style.display = "block";

    // Hide the scrollbar for the main window, since the iFrame
    // will have its own scrollbar and that is the only one that
    // is relevant.
    window.document.body.style.overflow = "hidden";

    // Launch Reach by updating the src of the iFrame.
    document.getElementById("launchReachFrame").src = url;
}

MainForm.prototype.ClosePage = function () {
    // Inspect browser version and take appropriate close action

    if (top == self) {
        if (isIE7OrLater()) {
            // set opener as self by invoking open - for IE7+ browsers
            window.open("", "_self");

            // close the browser window
            window.close();
        }
        else if (isIE6()) {
            // set the opener as self by directly setting the value of 
            // self.opener - for IE6 browser only
            self.opener = this;

            // close the browser window
            window.close();
        }
    }
    else {
        if (isIE6OrLater()) {
            // set opener as self by invoking open - for IE7,IE8 and IE9 browsers
            top.window.opener = top;
            top.window.open("", "_parent");
            // close the browser window
            top.window.close();
        }
    }
}

MainForm.prototype.LanguageSelectionChanged = function () {
    var languageSelector;

    isMobileDevice = isMobile.toLowerCase();

    if (userExperience.toLowerCase() != defaultExperienceVersion) {
        languageSelector = document.getElementById("languageSelectCmb15");
    }
    else {
        if (isMobileDevice == "true") {
            languageSelector = document.getElementById("languageSelectCmbMobile");
        }
        else {
            languageSelector = document.getElementById("languageSelectCmb");
        }
    }

    if (languageSelector.selectedIndex != -1) {
        var newLanguage = languageSelector.options[languageSelector.selectedIndex].value;
        if (newLanguage.toLowerCase() != currentLanguage.toLowerCase()) {
            this.UpdateSelectedLanguage(newLanguage);
            this.GetUpdatedResources(newLanguage);
            currentLanguage = newLanguage;
        }
    }
}

MainForm.prototype.UpdateSelectedLanguage = function (language) {
    var languageSelector;

    isMobileDevice = isMobile.toLowerCase();
    if (userExperience.toLowerCase() != defaultExperienceVersion) {
        var languageSelector = document.getElementById("languageSelectCmb15");
    }
    else {
        if (isMobileDevice == "true") {
            languageSelector = document.getElementById("languageSelectCmbMobile");
        }
        else {
            var languageSelector = document.getElementById("languageSelectCmb");
        }
    }

    var index = -1;
    for (i = 0; i < languageSelector.options.length; i++) {
        if (languageSelector.options[i].value.toLowerCase() == language.toLowerCase()) {
            index = i;
            break;
        }
    }

    languageSelector.selectedIndex = index;
}

MainForm.prototype.GetUpdatedResources = function (language) {
    ResourceURL = resourceUrl.toLowerCase();
    var url = ResourceURL + language;
    var languageSelector;

    try {
        window.setTimeout(TimerHandler(this.connection, this.connection.SendHttpRequest, "GET", url), 0);

        // Disable the language selector (so the user cannot keep changing
        // the language) until we receive the response with language resources
        // for the currently selected language.

        isMobileDevice = isMobile.toLowerCase();
        if (userExperience.toLowerCase() != defaultExperienceVersion) {
            languageSelector = document.getElementById("languageSelectCmb15");
        }
        else {
            if (isMobileDevice == "true") {
                languageSelector = document.getElementById("languageSelectCmbMobile");
            }
            else {
                languageSelector = document.getElementById("languageSelectCmb");
            }
        }
        languageSelector.disabled = true;
    } catch (e) {
        // Ignore language related failures
    }
}

MainForm.prototype.OnUpdatedResourcesCallback = function () {
    var XMLHTTPREQUEST_COMPLETE = 4;
    var XMLHTTPREQUEST_OK = 200;

    var currentState = null;
    var httpCode = null;

    try {
        currentState = this.connection._httpRequest.readyState;
    } catch (e) {
        // Ignore language related failures
        return;
    }

    try {
        // For Safari 10.1.3 the end status is 0
        if (currentState == 0 || currentState == XMLHTTPREQUEST_COMPLETE) {

            // Enable the language selector again.

            var languageSelector;

            isMobileDevice = isMobile.toLowerCase();
            if (userExperience.toLowerCase() != defaultExperienceVersion) {
                languageSelector = document.getElementById("languageSelectCmb15");
            }
            else {
                if (isMobileDevice == "true") {
                    languageSelector = document.getElementById("languageSelectCmbMobile");
                }
                else {
                    languageSelector = document.getElementById("languageSelectCmb");
                }
            }
            languageSelector.disabled = false;

            var httpCode = this.connection._httpRequest.status;

            if (httpCode == XMLHTTPREQUEST_OK) {
                // Get the text that the Server sent back
                // for the request we made.
                var text = this.connection._httpRequest.responseText;
                eval(text);

                this.UpdateUI();

                if (loading == "true") {
                    this.OnLoadComplete();
                }
            } else if (currentState != 0) {
                //Safari cannot get httpcode if network is down.

                // Ignore language related failures
            }
        }
    } catch (e) {
        // Ignore language related failures
    }
}

MainForm.prototype.UpdateHelpImage15 = function (isFlip) {
    if (isFlip) {
        document.getElementById("helpImage15").className = "helpImageFlip";

        // Bug 3246825: Ensure that the "?" image is flipped properly. The CSS contains the
        // correct styling for IE9 (-ms-transform) and IE10+ (transform), but IE versions
        // <= 8 do not support these properties. Simply adding the older -ms-filter or filter
        // to the CSS will not work, as IE9+ will apply both the transform and filter: hence
        // the special casing here. tridentVersion < 4 implies <= IE7 without also including
        // IE9 compatibility mode.
        if (tridentVersion < 4 || isIE8()) {
            document.getElementById("helpImage15").style.filter = "progid:DXImageTransform.Microsoft.BasicImage(mirror=1)";
        }
    }
    else {
        document.getElementById("helpImage15").className = "helpImage";

        if (tridentVersion < 4 || isIE8()) {
            document.getElementById("helpImage15").style.filter = "none";
        }
    }
}

MainForm.prototype.UpdateUI = function () {

    isMobileDevice = isMobile.toLowerCase();

    // Only localize the Skype brand title
    var isLegacy = isLegacyWebExperience.toLowerCase();
    if (isLegacy == "false") {
        window.document.title = SFB_AppName;
    }

    if (textDirection) {
        window.document.dir = textDirection;
    }

    if (window.document.dir == "rtl") {
        document.getElementById("copyright").className = "copyrightFlip";
        document.getElementById("content").className = "helpFlip";
    }
    else {
        document.getElementById("copyright").className = "copyright";
        document.getElementById("content").className = "help";
    }

    // Only flip the help image when it's RTL and not Hebrew.
    this.UpdateHelpImage15(window.document.dir == "rtl" && !isHebrew(currentLanguage));

    if (userExperience.toLowerCase() != defaultExperienceVersion) {

        document.getElementById("mainTablemobile").style.display = "none";
        document.getElementById("mainTable").style.display = "none";
        document.getElementById("mainTable15").style.display = "block";
        document.getElementById("mainTable15").style.backgroundColor = "#FFFFFF";
        document.getElementById("sfbProductTitleLabel").innerHTML = txt_SfbBrandName;

        if (isMobileDevice == "true") {

            document.getElementById("languageSettingsLabel15").style.display = "none";

            document.getElementById("languageSelectCmb15").style.width = "100px";
            document.getElementById("errorTextLabel15").className = "errorregularMobile";
            document.getElementById("checkUrlLabel15").className = "errorlowMobile";
            document.getElementById("errorHeader2").className = "errorHeader2Mobile";
            document.getElementById("diagLabel15").className = "bulletpointMobile";
            document.getElementById("errorBulletText").className = "bulletTextErrorMobile";
            document.getElementById("diagLabel215").className = "errorverylowMobile";
            document.getElementById("diagInfoText15").style.width = "60%";

            if (errorCode == "-1") {
                document.getElementById("copyright").innerHTML = txt_copyRightTextLabel15;
                document.getElementById("copyright").style.display = "block";
                document.getElementById("copyright").className = isRtl() ? "copyrightMobileFlip" : "copyrightMobile";
                document.getElementById("lynclogo").style.display = "block";
                document.getElementById("lynclogo").className = isRtl() ? "logospaceMobileFlip" : "logospaceMobile";
                document.getElementById("officeLogoColumn").className = isRtl() ? "officelogoColumnSuccessMobileFlip" : "officelogoColumnSuccessMobile";
                document.getElementById("contentRow15").className = isRtl() ? "contentClassSuccessMobileFlip" : "contentClassSuccessMobile";
                document.getElementById("successLogoMobile").style.display = "block";
                document.getElementById("footer").className = "footerSuccessMobile";
                document.getElementById("errorLogoMobile").style.display = "none";
                document.getElementById("content").className = isRtl() ? "helpSuccessMobileFlip" : "helpSuccessMobile";
            }
            else {
                document.getElementById("lynclogo").style.display = "none";
                document.getElementById("officeLogoColumn").className = "officelogoColumnErrorMobile";
                document.getElementById("contentRow15").className = isRtl() ? "contentClassErrorMobileFlip" : "contentClassErrorMobile";
                document.getElementById("errorLogoMobile").style.display = "block";
                document.getElementById("footer").className = "footerErrorMobile";
                document.getElementById("content").className = isRtl() ? "helpErrorMobileFlip " : "helpErrorMobile";
            }
        }
        else {
            if (errorCode == "-1") {
                document.getElementById("lynclogo").style.display = "block";
                document.getElementById("officeLogoColumn").className = "officelogoColumnSuccess";
                document.getElementById("successLogo").style.display = "block";
                document.getElementById("sfbBrandNameUnderSuccessLogoLine1").innerHTML = txt_SfbBrandNameMultiLine1;
                document.getElementById("sfbBrandNameUnderSuccessLogoLine2").innerHTML = txt_SfbBrandNameMultiLine2;
                //document.getElementById("spacer").className = "spacerSuccess";
                document.getElementById("firstRow").className = "firstRowSuccess";
                document.getElementById("copyright").style.display = "block";
                document.getElementById("copyright").innerHTML = txt_copyRightTextLabel15;
                document.getElementById("errorLogo").style.display = "none";

                if (window.document.dir == "rtl") {
                    document.getElementById("footer").className = "footerSuccess1280Flip";
                    document.getElementById("contentRow15").className = "contentClassSuccessFlip";
                    document.getElementById("lynclogo").className = "logospaceFlip";
                }
                else
                {
                    document.getElementById("footer").className = "footerSuccess1280";
                    document.getElementById("contentRow15").className = "contentClassSuccess";
                    document.getElementById("lynclogo").className = "logospace";
                }
            }
            else {
                document.getElementById("lynclogo").style.display = "none";
                document.getElementById("officeLogoColumn").className = "officelogoColumnError";
                document.getElementById("errorLogo").style.display = "block";
                document.getElementById("sfbBrandNameUnderFailureLogoLine1").innerHTML = txt_SfbBrandNameMultiLine1;
                document.getElementById("sfbBrandNameUnderFailureLogoLine2").innerHTML = txt_SfbBrandNameMultiLine2;
                //document.getElementById("spacer").className = "spacerError";
                document.getElementById("firstRow").className = "firstRowError";
                document.getElementById("footer").className = "footerError";
                document.getElementById("copyright").style.display = "none";

                if (window.document.dir == "rtl") {
                    document.getElementById("contentRow15").className = "contentClassErrorFlip";
                    document.getElementById("lynclogo").className = "logospaceFlip";
                }
                else
                {
                    document.getElementById("contentRow15").className = "contentClassError";
                    document.getElementById("lynclogo").className = "logospace";

                }

            }
        }

        document.getElementById("languageSettingsLabel15").innerHTML = txt_languageSettingsLabel;
        document.getElementById("launchRichClientHeaderLabel15").innerHTML = txt_launchRichClientHeaderLabel;
        document.getElementById("launchRichClientTextLabel15").innerHTML = txt_launchRichClientTextLabel;
        document.getElementById("contactSupportLabelWithReachOption").innerHTML = txt_unableToJoinLabel;
        document.getElementById("JoinUsingLWAOptionHelp115").innerHTML = txt_UnsupportedSfbVersion2;
        document.getElementById("JoinMeetingUsingLWA2_Link15").innerHTML = txt_UnsupportedSfbVersion3;

        var launchReachText = txt_joinUsingReachLink.replace(/%ReachClientProductName%/g, reachClientProductName);
        document.getElementById("launchReachLink").innerHTML = launchReachText;
        document.getElementById("launchReachLink").title = launchReachText;

        document.getElementById("errorTextLabel15").innerHTML = conferenceError1;
        document.getElementById("checkUrlLabel15").innerHTML = MainForm.prototype.SelectConferenceError2();
        document.getElementById("diagLabel15").innerHTML = diag;

        var diagInfoTextString = diagInfo.toString();
        document.getElementById("diagInfoText15").value = diagInfoTextString;
        document.getElementById("diagInfoText15").style.display = "none";
        document.getElementById("diagLabel215").style.display = "none";
        document.getElementById("diagLabel215").innerHTML = diag2;

        document.getElementById("onlineHelpLink").innerHTML = txt_onlineHelpLink;
        document.getElementById("onlineHelpLink").title = txt_onlineHelpLink;
        document.getElementById("helpImage15").title = txt_onlineHelpLink;
        //document.getElementById("copyRightTextLabel15").innerHTML = txt_copyRightTextLabel15;

        document.getElementById("UnsupportedClientBlockMessageLabel15").innerHTML = txt_UnsupportedSfbVersion1;
        document.getElementById("JoinUsingLWAOption115").innerHTML = txt_UnsupportedSfbVersion2;
        document.getElementById("JoinMeetingUsingLWA_Link15").innerHTML = txt_UnsupportedSfbVersion3;

        document.getElementById("connectingLabel15").innerHTML = txt_connecting;
        document.getElementById("mobileappstoreLabel15").innerHTML = txt_unableToLaunchLyncMobile;
        document.getElementById("mobileappstoreLabel215").innerHTML = txt_unableToLaunchSfbMobile2;
        document.getElementById("mobileappstoreLabel315").innerHTML = txt_unableToLaunchLyncMobile3;
        document.getElementById("unknownMobileDeviceLabel15").innerHTML = txt_unsupportedMobileDevice;

    }
    else {

        if (isMobileDevice == "true") {
            document.getElementById("mainTable").style.display = "none";
            document.getElementById("mainTable15").style.display = "none";
            document.getElementById("mainTablemobile").style.display = "block";

            document.getElementById("languageSettingsLabelMobile").innerHTML = txt_languageSettingsLabel;

            document.getElementById("errorTextLabelMobile").innerHTML = conferenceError1;
            document.getElementById("checkUrlLabelMobile").innerHTML = MainForm.prototype.SelectConferenceError2();
            document.getElementById("diagLabelMobile").innerHTML = diag;

            var diagInfoTextString = diagInfo.toString();
            document.getElementById("diagInfoTextMobile").value = diagInfoTextString;
            document.getElementById("diagInfoTextMobile").style.display = "none";
            document.getElementById("diagLabel2Mobile").style.display = "none";
            document.getElementById("diagLabel2Mobile").innerHTML = diag2;

            document.getElementById("onlineHelpLinkMobile").innerHTML = txt_onlineHelpLink;
            document.getElementById("copyRightTextLabelMobile").innerHTML = txt_copyRightTextLabel;

            document.getElementById("connectingLabel").innerHTML = txt_connecting;
            document.getElementById("mobileappstoreLabel").innerHTML = txt_unableToLaunchLyncMobile;
            document.getElementById("mobileappstoreLabel2").innerHTML = txt_unableToLaunchLyncMobile2;
            document.getElementById("unknownMobileDeviceLabel").innerHTML = txt_unsupportedMobileDevice;
            return;
        }

        // non-mobile
        document.getElementById("mainTable15").style.display = "none";
        document.getElementById("mainTablemobile").style.display = "none";
        document.getElementById("mainTable").style.display = "block";

        document.getElementById("languageSettingsLabel").innerHTML = txt_languageSettingsLabel;
        document.getElementById("launchRichClientHeaderLabel").innerHTML = txt_launchRichClientHeaderLabel;
        document.getElementById("launchRichClientTextLabel").innerHTML = txt_launchRichClientTextLabel;
        document.getElementById("contactSupportLabelWithReachOption").innerHTML = txt_unableToJoinLabel;

        var launchReachText = txt_joinUsingReachLink.replace(/%ReachClientProductName%/g, reachClientProductName);
        document.getElementById("launchReachLink").innerHTML = launchReachText;
        document.getElementById("launchReachLink").title = launchReachText;

        document.getElementById("errorTextLabel").innerHTML = conferenceError1;
        document.getElementById("checkUrlLabel").innerHTML = MainForm.prototype.SelectConferenceError2();
        document.getElementById("diagLabel").innerHTML = diag;

        var diagInfoTextString = diagInfo.toString();
        document.getElementById("diagInfoText").value = diagInfoTextString;
        document.getElementById("diagInfoText").style.display = "none";
        document.getElementById("diagLabel2").style.display = "none";
        document.getElementById("diagLabel2").innerHTML = diag2;

        document.getElementById("onlineHelpLink").innerHTML = txt_onlineHelpLink;
        document.getElementById("copyRightTextLabel").innerHTML = txt_copyRightTextLabel;

        document.getElementById("64bitUnSupportedMessage").innerHTML = txt_64bitUnsupportedText1;
        document.getElementById("64bitUnSupportedOption1").innerHTML = txt_64bitUnsupportedOption1;
        document.getElementById("JoinMeetingUsingLync_Link").innerHTML = txt_64bitJoinUsingLync;
        document.getElementById("64bitUnSupportedOption2").innerHTML = txt_64bitUnsupportedOption2;
        document.getElementById("64bitUnSupportedMessageLabel").innerHTML = txt_64bitbrowserUnsupportedLabel;

        document.getElementById("UnsupportedClientBlockMessageLabel").innerHTML = txt_UnsupportedLyncVersion1;
        document.getElementById("JoinUsingLWAOption1").innerHTML = txt_UnsupportedLyncVersion2;
        document.getElementById("JoinMeetingUsingLWA_Link").innerHTML = txt_UnsupportedLyncVersion3;
    }
}

//
// Connection object - BEGIN
//
function ConnectionObject()
{
    // Properties of the object
    this._httpRequest = this._CreateXMLHttpRequestObject( );
}

ConnectionObject.prototype._CreateXMLHttpRequestObject = function()
{
    var httpRequest = null;

    try
    {
        if( window.XMLHttpRequest ) // for none IE browsers
        {
            httpRequest = new XMLHttpRequest();
        }
        else if( window.ActiveXObject ) // for IE
        {
            var MSXML_XMLHTTP_PROGIDS = new Array(
                            'Microsoft.XMLHTTP',
                            'MSXML2.XMLHTTP.5.0',
                            'MSXML2.XMLHTTP.4.0',
                            'MSXML2.XMLHTTP.3.0',
                            'MSXML2.XMLHTTP'
                            );

            for (var i=0; i < MSXML_XMLHTTP_PROGIDS.length; i++)
            {
                httpRequest = new ActiveXObject(MSXML_XMLHTTP_PROGIDS[i]);

                if( httpRequest != null )
                {
                    break;
                }
            }
        }
    } catch (e) {
        // Ignore any exception since
        // we are just initializing.
    }

    return httpRequest;
}

ConnectionObject.prototype.SendHttpRequest = function(type, url) {
    try {
        // Set up a new request to the Server.
        // Request is async.
        this._httpRequest.open(type, url, true);

        this._httpRequest.onreadystatechange = Delegate(mainWindow, mainWindow.OnUpdatedResourcesCallback);

        // Send the request to the Server.
        // No data to send, so pass null.
        this._httpRequest.send(null);
    } catch (e) {
        // Ignore language related failures
    }
}
//
// Connection object - END
//
