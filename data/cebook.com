<!--NOTE: If DOCTYPE element is present, it causes the iFrame to be displayed in a small-->
<!--portion of the browser window instead of occupying the full browser window.-->
<html dir="LTR" xmlns="http://www.w3.org/1999/xhtml" class="reachJoinHtml"><head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=10; IE=9; IE=8;">
    <meta name="viewport" content="width=device-width; initial-scale=1.0; maximum-scale=1.0; user-scalable=0;">
    <title>Skype for Business Web App</title>
    <script type="text/javascript">
        var reachURL = "https://sfbwebext.integrationpartners.com/lwa/WebPages/LwaClient.aspx?legacy=RmFsc2U!&xml=PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48Y29uZi1pbmZvIHhtbG5zOnhzZD0iaHR0cDovL3d3dy53My5vcmcvMjAwMS9YTUxTY2hlbWEiIHhtbG5zOnhzaT0iaHR0cDovL3d3dy53My5vcmcvMjAwMS9YTUxTY2hlbWEtaW5zdGFuY2UiIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL3J0Yy8yMDA5LzA1L3NpbXBsZWpvaW5jb25mZG9jIj48Y29uZi11cmk-c2lwOnNoYXJnaXNAaW50ZWdyYXRpb25wYXJ0bmVycy5jb207Z3J1dTtvcGFxdWU9YXBwOmNvbmY6Zm9jdXM6aWQ6UFNERk9aQzQ8L2NvbmYtdXJpPjxzZXJ2ZXItdGltZT43Mi4wMDI8L3NlcnZlci10aW1lPjxvcmlnaW5hbC1pbmNvbWluZy11cmw-aHR0cHM6Ly9tZWV0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3NoYXJnaXMvUFNERk9aQzQ8L29yaWdpbmFsLWluY29taW5nLXVybD48Y29uZi1rZXk-UFNERk9aQzQ8L2NvbmYta2V5PjxmYWxsYmFjay11cmw-aHR0cHM6Ly9tZWV0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3NoYXJnaXMvUFNERk9aQzQ_c2w9PC9mYWxsYmFjay11cmw-PHVjd2EtdXJsPmh0dHBzOi8vc2Zid2ViZXh0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3Vjd2EvdjEvYXBwbGljYXRpb25zPC91Y3dhLXVybD48dWN3YS1leHQtdXJsPmh0dHBzOi8vc2Zid2ViZXh0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3Vjd2EvdjEvYXBwbGljYXRpb25zPC91Y3dhLWV4dC11cmw-PHVjd2EtaW50LXVybD5odHRwczovL3NmYndlYmludC5pcGMubG9jYWwvdWN3YS92MS9hcHBsaWNhdGlvbnM8L3Vjd2EtaW50LXVybD48dGVsZW1ldHJ5LWlkPjJiZWI1NDA4LTRkOWEtNDczOS04MjFlLWI1YzFiM2E4YTE2OTwvdGVsZW1ldHJ5LWlkPjwvY29uZi1pbmZvPg!!";
        var escapedXML = "\x3c\x3fxml version\x3d\x221.0\x22 encoding\x3d\x22utf-8\x22\x3f\x3e\x3cconf-info xmlns\x3axsd\x3d\x22http\x3a\x2f\x2fwww.w3.org\x2f2001\x2fXMLSchema\x22 xmlns\x3axsi\x3d\x22http\x3a\x2f\x2fwww.w3.org\x2f2001\x2fXMLSchema-instance\x22 xmlns\x3d\x22http\x3a\x2f\x2fschemas.microsoft.com\x2frtc\x2f2009\x2f05\x2fsimplejoinconfdoc\x22\x3e\x3cconf-uri\x3esip\x3ashargis\x40integrationpartners.com\x3bgruu\x3bopaque\x3dapp\x3aconf\x3afocus\x3aid\x3aPSDFOZC4\x3c\x2fconf-uri\x3e\x3cserver-time\x3e72.002\x3c\x2fserver-time\x3e\x3coriginal-incoming-url\x3ehttps\x3a\x2f\x2fmeet.integrationpartners.com\x2fshargis\x2fPSDFOZC4\x3c\x2foriginal-incoming-url\x3e\x3cconf-key\x3ePSDFOZC4\x3c\x2fconf-key\x3e\x3cfallback-url\x3ehttps\x3a\x2f\x2fmeet.integrationpartners.com\x2fshargis\x2fPSDFOZC4\x3fsl\x3d\x3c\x2ffallback-url\x3e\x3cucwa-url\x3ehttps\x3a\x2f\x2fsfbwebext.integrationpartners.com\x2fucwa\x2fv1\x2fapplications\x3c\x2fucwa-url\x3e\x3cucwa-ext-url\x3ehttps\x3a\x2f\x2fsfbwebext.integrationpartners.com\x2fucwa\x2fv1\x2fapplications\x3c\x2fucwa-ext-url\x3e\x3cucwa-int-url\x3ehttps\x3a\x2f\x2fsfbwebint.ipc.local\x2fucwa\x2fv1\x2fapplications\x3c\x2fucwa-int-url\x3e\x3ctelemetry-id\x3e2beb5408-4d9a-4739-821e-b5c1b3a8a169\x3c\x2ftelemetry-id\x3e\x3c\x2fconf-info\x3e";
        var validMeeting = "True";
        var reachClientRequested = "False";
        var htmlLwaClientRequested = "False";
        var currentLanguage = "en-US";
        var reachClientProductName = "商務用 Skype Web App";
        var blockPreCU2Clients = "False";
        var isNokia = "False";
        var isAndroid = "False";
        var isWinPhone = "False";
        var isIPhone = "False";
        var isIPad = "False";
        var isMobile = "False";
        var isUnsupported = "False";
        var isLegacyWebExperience = "False";
        var domainOwnerJoinLauncherUrl = "";
        var lyncLaunchLink = "conf:sip:shargis@integrationpartners.com;gruu;opaque=app:conf:focus:id:PSDFOZC4%3Frequired-media=audio";
        var diagInfo = "Machine\x3aDC01-SFBFE01Join attempted at\x28UTC\x29\x3a11\x2f19\x2f2015 5\x3a52\x3a51 PMCorrelationId\x3a746613143User Agent\x3aMozilla\x2f5.0 \x28Windows NT 6.1\x3b WOW64\x3b rv\x3a42.0\x29 Gecko\x2f20100101 Firefox\x2f42.0Accept Header\x3atext\x2fhtmlapplication\x2fxhtml\x2bxmlapplication\x2fxml\x3bq\x3d0.9\x2a\x2f\x2a\x3bq\x3d0.8Incoming URL\x3ahttps\x3a\x2f\x2fmeet.integrationpartners.com\x2fshargis\x2fPSDFOZC4TelemetryId\x3a2beb5408-4d9a-4739-821e-b5c1b3a8a169";
        var userExperience = "1500";
        var isLwaEnabled = "True";
        var escalateToDesktop = "False";
        var resourceUrl = "/meet/JavaScriptResourceHandler.ashx?lcs_se_w16_msrc6.0.9319.88&language=";
        var telemetryId = "2beb5408-4d9a-4739-821e-b5c1b3a8a169";
        var errorCode = "-1";
        var reachClientTitleString = "商務用 Skype Web App";
        var mobileW1ProtocolHandler = "lync://";
        var mobileW2ProtocolHandler = "lync15://";
        var lync15CommonProtocolHandler = "lync15:";
        var lync15ClassicProtocolHandler = "lync15classic:";
        var mlxProtocolHandler = "lync15mlx:";
        var parentOriginWhitelist = "[]";
        var parentOriginJsCheckEnabled = "False".toLowerCase() === 'true';
        var chromeVersion = "";

        togglediag = function () {
            if (userExperience.toLowerCase() != defaultExperienceVersion) {
                if (document.getElementById("diagInfoText15").style.display == "none") {
                    document.getElementById("diagInfoText15").style.display = "block";
                    document.getElementById("diagLabel215").style.display = "block";
                }
                else {
                    document.getElementById("diagLabel215").style.display = "none";
                    document.getElementById("diagInfoText15").style.display = "none";
                }
            }
            else { // 1400 beahavior
                if (isMobile.toLowerCase() == "true") {
                    if (document.getElementById("diagInfoTextMobile").style.display == "none") {
                        document.getElementById("diagInfoTextMobile").style.display = "block";
                        document.getElementById("diagLabel2Mobile").style.display = "block";
                    }
                    else {
                        document.getElementById("diagLabel2Mobile").style.display = "none";
                        document.getElementById("diagInfoTextMobile").style.display = "none";
                    }
                }
                else {
                    if (document.getElementById("diagInfoText").style.display == "none") {
                        document.getElementById("diagInfoText").style.display = "block";
                        document.getElementById("diagLabel2").style.display = "block";
                    }
                    else {
                        document.getElementById("diagLabel2").style.display = "none";
                        document.getElementById("diagInfoText").style.display = "none";
                    }
                }
            }
        }
    </script>
    <script type="text/javascript" src="cebook_files/Utilities.js"></script>
    <script type="text/javascript" src="cebook_files/PluginLoader.js"></script>
    <script type="text/javascript" src="cebook_files/Launch.js"></script>
    <link rel="Stylesheet" type="text/css" href="cebook_files/ReachClient.css">

    
		<link rel="shortcut icon" href="https://meet.integrationpartners.com/meet/Resources/SFBFavIcon.ico" type="image/x-icon"> 
		<link rel="icon" href="https://meet.integrationpartners.com/meet/Resources/SFBFavIcon.ico" type="image/ico">
		<link rel="icon" href="https://meet.integrationpartners.com/meet/Resources/SFBFavIcon.ico" type="image/x-icon">
    

</head>
<body class="reachJoinBody">
    <div id="launchReachDiv" style="display: none;" class="launchReachDiv">
        <!--NOTE: IE6 displays a mixed content warning if iFrame source is set to an empty string initially and later set dynamically, so instead set it to blank page here.-->
        <iframe id="launchReachFrame" name="launchReachFrame" src="cebook_files/blank.htm" scrolling="auto" style="overflow: hidden; width: 0px; height: 0px;" allowfullscreen="true" frameborder="0"></iframe>
    </div>
    <table id="mainTable" class="mainTable" style="display:none" align="center" cellpadding="0" cellspacing="0">
        <tbody><tr id="languageSettingsRow">
            <td id="languageSettingsCell" colspan="3" style="text-align: right;">
                <div id="languageSettingsDiv" style="display: block;">
                    <label id="languageSettingsLabel" class="smallerBlurredBodyText" style="text-align: center;">
                        Language:
                    </label>
                    <select id="languageSelectCmb" onchange="mainWindow.LanguageSelectionChanged();" class="smallerBlurredBodyText" style="margin-right: 15px;" tabindex="1">
                        <option selected="selected" value="ar">العربية‏</option><option value="az">Azərbaycan­ılı‎</option><option value="be">Беларуская‎</option><option value="bg">български‎</option><option value="ca">Català‎</option><option value="cs">čeština‎</option><option value="da">dansk‎</option><option value="de">Deutsch‎</option><option value="el">Ελληνικά‎</option><option value="en-US">English (United States)‎</option><option value="es">español‎</option><option value="et">eesti‎</option><option value="eu">euskara‎</option><option value="fa">فارسى‏</option><option value="fi">suomi‎</option><option value="fil">Filipino‎</option><option value="fr">français‎</option><option value="gl">galego‎</option><option value="he">עברית‏</option><option value="hi">हिंदी‎</option><option value="hr">hrvatski‎</option><option value="hu">magyar‎</option><option value="id">Bahasa Indonesia‎</option><option value="it">italiano‎</option><option value="ja">日本語‎</option><option value="kk">Қазақ‎</option><option value="ko">한국어‎</option><option value="lt">lietuvių‎</option><option value="lv">latviešu‎</option><option value="mk">македонски јазик‎</option><option value="ms">Bahasa Melayu‎</option><option value="nl">Nederlands‎</option><option value="no">norsk‎</option><option value="pl">polski‎</option><option value="pt-BR">português (Brasil)‎</option><option value="pt-PT">português (Portugal)‎</option><option value="ro">română‎</option><option value="ru">русский‎</option><option value="sk">slovenčina‎</option><option value="sl">slovenščina‎</option><option value="sq">Shqip‎</option><option value="sr-Cyrl">српски‎</option><option value="sr">srpski‎</option><option value="sv">svenska‎</option><option value="th">ไทย‎</option><option value="tr">Türkçe‎</option><option value="uk">українська‎</option><option value="uz">O'zbekcha‎</option><option value="vi">Tiếng Việt‎</option><option value="zh-Hans">中文(简体)‎</option><option value="zh-Hant">中文(繁體)‎</option>
                    </select>
                </div>
            </td>
        </tr>
        <tr id="headerRow">
            <td id="headerCell" colspan="3">
                <table id="headerLogoTable" class="headerLogoTable" cellpadding="0" cellspacing="0">
                    <tbody><tr id="headerLogoTableRow">
                        <td id="headerLogoTableCell">
                            <img src="cebook_files/mainbg_top.png" style="margin-left: 2px; vertical-align: bottom;" alt="">
                        </td>
                    </tr>
                </tbody></table>
            </td>
        </tr>
        <tr id="contentRow">
            <td id="leftBorderCell" class="leftBorderCell">
                &nbsp;
            </td>
            <td id="ContentCell" class="mainContentCell">
                <table id="mainContentTable" class="mainContentTable">
                    <tbody><tr id="topLogoRow">
                        <td id="topLogoCell" class="topLogoCell">
                            <img id="communicatorLogoImage" name="communicatorLogoImage" src="cebook_files/CommunicatorLogoType.png" alt="">
                        </td>
                    </tr>
                    <tr id="landingPageTextContentRow">
                        <td id="landingPageTextContentCell" class="landingPageTextContentCell">
                            <noscript id="noScriptContent">
                                <div id="javaScriptDisabledDiv">
                                    <table cellpadding="0" cellspacing="0" class="contentHeaderTable">
                                        <tr>
                                            <td class="contentHeaderTableCell">
                                                <label id="javaScriptDisabledHeaderLabel" class="bodyBoldText">
                                                    Never want to see this screen again?
                                                </label>
                                                <br id="lineBreakJavaScriptDisabledHeader1" />
                                                <label id="javaScriptDisabledTextLabel" class="bodyText">
                                                    Just enable JavaScript in your browser settings, and you can skip this screen and join meetings automatically.
                                                </label>
                                            </td>
                                        </tr>
                                    </table>
                                    <br />
                                    <table cellspacing="0" cellpadding="0">
                                        <tr>
                                            <td colspan="3">
                                                <img id="JSDisabledOptionTopImage" src="/meet/Resources/Button_normal_top.png"
                                                    style="vertical-align: bottom;" alt="" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td id="JSDisabledOptionLeftImage" class="leftNormalButtonCell">
                                                &nbsp;
                                            </td>
                                            <td class="reachButtonContentCell">
                                                <a id="javaScriptDisabledLaunchRichClientLink" name="javaScriptDisabledLaunchRichClientLink"
                                                    href="conf:sip:shargis@integrationpartners.com;gruu;opaque=app:conf:focus:id:PSDFOZC4%3Frequired-media=audio" class="bodyLinkText" tabindex="2" style="text-decoration: none;"
                                                    title="Forget about the JavaScript and join now.">
                                                    Forget about the JavaScript and join now.
                                                </a>
                                            </td>
                                            <td id="JSDisabledOptionRightImage" class="rightNormalButtonCell">
                                                &nbsp;
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <img id="JSDisabledOptionBottomImage" src="/meet/Resources/Button_normal_bottom.png"
                                                    style="vertical-align: top;" alt="" />
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </noscript>
                            
                            <script type="text/javascript">
                                var mainWindow = new MainForm();
                            </script>
                            
                            <div id="joinLauncherErrorDiv" style="display: none;">
                                <label id="errorTextLabel" style="color: #FF0000; font-weight: bold;">
                                    Sorry, something went wrong, and we can't get you into the meeting.
                                </label>
                                <br id="lineBreakJoinLauncherError2">
                                <label id="checkUrlLabel" style="font-weight: normal;font-size: 12px;">
                                It's possible you're using a bad URL. 
Try calling into the meeting using the phone number on the invite, or 
ask the organizer to drag you in.
                                </label>
                                <br id="lineBreakJoinLauncherError5">
                                <label id="diagLabel" style="font-weight: normal;font-size: 11px; text-decoration: underline; color: #0000FF;" onmouseover="this.style.cursor='hand'" onclick="togglediag()">
                                <br id="lineBreakJoinLauncherError6">
                                More about this error
                                </label>
                                <br>
                                <label id="diagLabel2" style="font-weight: normal;font-size: 11px; display:none">
                                Copy this diagnostic info and send it to your admin.
                                </label>
                                <textarea id="diagInfoText" cols="70" rows="5" style="font-size: 8px; overflow:auto" readonly="readonly">                                </textarea>

                            </div>
                            <div id="64bitUnsupportedDiv" style="display: none">
                                   <table cellpadding="0" cellspacing="0">
                                        <tbody><tr>
                                            <td colspan="3">
                                              <label id="64bitUnSupportedMessageLabel" style="display:block;" class="headerText">
                                                Your 64-bit browser is preventing you from joining the meeting.
                                              </label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                              <label id="64bitUnSupportedMessage" style="display:block;" class="bodyText">
                                               To join the meeting, do one of the following:
                                             </label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <img id="OCJoinOptionNoScrTopImage" src="cebook_files/Button_normal_top.png" style="vertical-align: bottom;" alt="">
                                            </td>
                                        </tr>
                                        <tr>
                                           <td id="Td5" class="leftNormalButtonCell">
                                              &nbsp;
                                           </td>
                                           <td class="reachButtonContentCell">
                                              <label id="64bitUnSupportedOption1" style="display:block;" class="bodyText">
                                                  If you have Lync installed, click to join.
                                              </label>
                                              <a id="JoinMeetingUsingLync_Link" class="bodyLinkText" href="conf:sip:shargis@integrationpartners.com;gruu;opaque=app:conf:focus:id:PSDFOZC4%3Frequired-media=audio" title="">
                                                   Join Using Lync
                                              </a>
                                           </td>
                                           <td id="Td6" class="rightNormalButtonCell">
                                              &nbsp;
                                           </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <img id="OCJoinOptionNoScrBottomImage" src="cebook_files/Button_normal_bottom.png" style="vertical-align: top;" alt="">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <img id="OCJoinOptionNoScrBottomImage" src="cebook_files/Button_normal_top.png" style="vertical-align: bottom;" alt="">
                                            </td>
                                        </tr>
                                        <tr>
                                          <td id="Td3" class="leftNormalButtonCell">
                                              &nbsp;
                                          </td>
                                          <td class="reachButtonContentCell">
                                              <label id="64bitUnSupportedOption2" style="display:block;" class="bodyText">
                                                  Copy the meeting URL 
into the address bar of a 32-bit browser. (Find a browser on your start 
menu.)
                                              </label>
                                          </td>
                                          <td id="Td4" class="rightNormalButtonCell">
                                              &nbsp;
                                          </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <img id="OCJoinOptionNoScrBottomImage" src="cebook_files/Button_normal_bottom.png" style="vertical-align: top;" alt="">
                                            </td>
                                        </tr>
                                </tbody></table>

                            </div>
                            <div id="UnsupportedClientBlockDiv" style="display: none">
                                   <table cellpadding="0" cellspacing="0">
                                        <tbody><tr>
                                            <td colspan="3">
                                              <label id="UnsupportedClientBlockMessageLabel" style="display:block;" class="headerText">
                                                You have a version of 
Lync that doesn't support joining this online meeting. Contact your 
support team to upgrade to a newer version of Lync.
                                              </label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <img id="OCJoinOptionNoScrTopImage" src="cebook_files/Button_normal_top.png" style="vertical-align: bottom;" alt="">
                                            </td>
                                        </tr>
                                        <tr>
                                           <td id="Td5" class="leftNormalButtonCell">
                                              &nbsp;
                                           </td>
                                           <td class="reachButtonContentCell">
                                              <label id="JoinUsingLWAOption1" style="display:block;" class="bodyText">
                                                  You can join the meeting now using Lync Web App.
                                              </label>
                                              <a id="JoinMeetingUsingLWA_Link" class="bodyLinkText" href="" target="_blank" title="">
                                                   Join Using Lync Web App
                                              </a>
                                           </td>
                                           <td id="Td6" class="rightNormalButtonCell">
                                              &nbsp;
                                           </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <img id="OCJoinOptionNoScrBottomImage" src="cebook_files/Button_normal_bottom.png" style="vertical-align: top;" alt="">
                                            </td>
                                        </tr>
                                </tbody></table>
                            </div>

                            <div id="launchRichClientDiv" style="display: none;">
                                <br id="lineBreakLaunchRichClientDiv">
                                <label id="launchRichClientHeaderLabel" class="headerText">
                                    You are joining the meeting.
                                </label>
                                <br id="lineBreakLaunchRichClientHeader1">
                                <label id="launchRichClientTextLabel" class="bodyText">
                                    You may close this web page now.
                                </label>
                                <br id="lineBreakLaunchRichClientText1">
                                <br id="lineBreakLaunchRichClientText2">
                                <br id="lineBreakLaunchRichClientText3">
                                <br id="lineBreakLaunchRichClientText4">
                                <br id="lineBreakLaunchRichClientText5">
                            </div>
                        </td>
                    </tr>
                    <tr id="bottomHelpRow">
                        <td class="landingPageTextContentCell" style="padding-top:0px;">
                            <label id="contactSupportLabelWithReachOption" class="bodyText" style="display: none;">Having trouble joining the meeting?</label>
                            <a id="launchReachLink" name="launchReachLink" href="" class="bodyLinkText" onclick="mainWindow.RedirectToReach();" target="launchReachFrame" tabindex="3" style="display: none;" title="Try 商務用 Skype Web App">Try 商務用 Skype Web App</a>
                            <br id="lineBreakContactSupport1">
                            <a id="onlineHelpLink" name="onlineHelpLink" href="http://go.microsoft.com/fwlink/?LinkId=190632" class="bodyLinkText" target="_blank" tabindex="4" title="Online Help">Online Help</a>
                        </td>
                    </tr>
                </tbody></table>
            </td>
            <td id="rightBorderCell" class="rightBorderCell">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <table id="bottomCenteredTable" class="footerTable" cellpadding="0" cellspacing="0">
                    <tbody><tr id="bottomCenteredRow">
                        <td id="bottomCenteredCell" class="footerTableCell">
                            <img id="officeLogoImage" name="officeLogoImage" src="cebook_files/OfficeLogo.png" alt="">
                            <br id="lineBreakOfficeLogoImage">
                            <label id="copyRightTextLabel" class="smallerBlurredBodyText">
                                © 2010 Microsoft Corporation. All rights reserved.
                            </label>
                        </td>
                    </tr>
                </tbody></table>
            </td>
        </tr>
    </tbody></table>
    <table id="mainTable15" class="maintable15" style="background-color: rgb(255, 255, 255); width: 100%; height: 100%; display: block;" border="0" cellpadding="0" cellspacing="0">
        <tbody><tr class="firstRowSuccess" id="firstRow">
            <td id="officeLogoColumn" class="officelogoColumnSuccess">
                <div id="successLogo" class="centeredImageSuccess" style="display: block;">
                    <div style="margin-left:auto; margin-right:auto;" align="center">
                        <img src="cebook_files/SfB_Logo_Vertical.png" style="margin-left:auto; margin-right:auto;" alt="">
                    </div>
                    <div style="margin-left:auto; margin-right:auto;" align="center">
                        <label id="sfbBrandNameUnderSuccessLogoLine1" class="sfbBrandUnderLogo">Skype</label>
                    </div>
                    <div style="margin-left:auto; margin-right:auto;" align="center">
                        <label id="sfbBrandNameUnderSuccessLogoLine2" class="sfbBrandUnderLogo">for Business</label>
                    </div>
                </div>
                <div style="display: none;" id="errorLogo" class="centeredImageError">
                    <div style="margin-left:auto; margin-right:auto;" align="center">
                        <img src="cebook_files/SfB_Logo_Vertical.png" style="margin-left:auto; margin-right:auto;" alt="">
                    </div>
                    <div style="margin-left:auto; margin-right:auto;" align="center">
                        <label id="sfbBrandNameUnderFailureLogoLine1" class="sfbBrandUnderLogo">Skype</label>
                    </div>
                    <div style="margin-left:auto; margin-right:auto;" align="center">
                        <label id="sfbBrandNameUnderFailureLogoLine2" class="sfbBrandUnderLogo">for Business</label>
                    </div>
                 </div>
                <div id="successLogoMobile" class="centeredImageSuccessMobile" style="display:none"><img src="cebook_files/SfB_Logo_Vertical_Mobile.png" alt=""> </div>
                <div id="errorLogoMobile" class="centeredImageErrorMobile" style="display:none"><img src="cebook_files/SfB_Logo_Vertical_Mobile.png" alt=""> </div>
            </td>
            <td id="content" class="help" colspan="2">
                <div style="display: block;" id="languageSettingsDiv15" class="languageHelp">
                    <label id="languageSettingsLabel15" class="smallerBlurredBodyText15" style="vertical-align:middle">Language:</label>
                    <select id="languageSelectCmb15" onchange="mainWindow.LanguageSelectionChanged();" class="smallerBlurredBodyText15" style="vertical-align:middle; width:120px">
                        <option value="ar">العربية‏</option><option value="az">Azərbaycan­ılı‎</option><option value="be">Беларуская‎</option><option value="bg">български‎</option><option value="ca">Català‎</option><option value="cs">čeština‎</option><option value="da">dansk‎</option><option value="de">Deutsch‎</option><option value="el">Ελληνικά‎</option><option selected="selected" value="en-US">English (United States)‎</option><option value="es">español‎</option><option value="et">eesti‎</option><option value="eu">euskara‎</option><option value="fa">فارسى‏</option><option value="fi">suomi‎</option><option value="fil">Filipino‎</option><option value="fr">français‎</option><option value="gl">galego‎</option><option value="he">עברית‏</option><option value="hi">हिंदी‎</option><option value="hr">hrvatski‎</option><option value="hu">magyar‎</option><option value="id">Bahasa Indonesia‎</option><option value="it">italiano‎</option><option value="ja">日本語‎</option><option value="kk">Қазақ‎</option><option value="ko">한국어‎</option><option value="lt">lietuvių‎</option><option value="lv">latviešu‎</option><option value="mk">македонски јазик‎</option><option value="ms">Bahasa Melayu‎</option><option value="nl">Nederlands‎</option><option value="no">norsk‎</option><option value="pl">polski‎</option><option value="pt-BR">português (Brasil)‎</option><option value="pt-PT">português (Portugal)‎</option><option value="ro">română‎</option><option value="ru">русский‎</option><option value="sk">slovenčina‎</option><option value="sl">slovenščina‎</option><option value="sq">Shqip‎</option><option value="sr-Cyrl">српски‎</option><option value="sr">srpski‎</option><option value="sv">svenska‎</option><option value="th">ไทย‎</option><option value="tr">Türkçe‎</option><option value="uk">українська‎</option><option value="uz">O'zbekcha‎</option><option value="vi">Tiếng Việt‎</option><option value="zh-Hans">中文(简体)‎</option><option value="zh-Hant">中文(繁體)‎</option>
                    </select>
                    <a id="onlineHelpLink15" name="onlineHelpLink" href="http://go.microsoft.com/fwlink/?LinkId=262053" class="bodyLinkText" target="_blank" tabindex="4" title="Online Help">
                        <img title="Online Help" style="filter: none;" id="helpImage15" src="cebook_files/Help_19x30.png" class="helpImage" alt="Online Help">
                   </a>
                </div>
            </td>
        </tr>
        <tr style="display: block;" id="lynclogo" class="logospace">
            <td style="vertical-align:top">
                <label id="sfbProductTitleLabel" class="errorbold">Skype for Business</label>
            </td>
            <td></td>
        </tr>
        <tr id="contentRow15" class="contentClassSuccess">
            <td id="landingPageTextContentCell15" class="landingPageTextContentCell15">
                <noscript id="noScriptContent15">
                    <div id="javaScriptDisabledDiv15">
                        <table cellspacing="0" cellpadding="0">
                            <tr>
                                <td colspan = "2">
                                    <label id="javaScriptDisabledHeaderLabel15" class="errorbold">
                                        Never want to see this screen again?
                                    </label>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding-top:29px" colspan = "2">
                                    <label id="javaScriptDisabledTextLabel15" class="errorregular">
                                        Just enable JavaScript in your browser settings, and you can skip this screen and join meetings automatically.
                                    </label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bulletTextError">
                                        <a id="javaScriptDisabledLaunchRichClientLink15" name="javaScriptDisabledLaunchRichClientLink15"
                                            href="conf:sip:shargis@integrationpartners.com;gruu;opaque=app:conf:focus:id:PSDFOZC4%3Frequired-media=audio" class = "bulletpoint" style="vertical-align:top; text-align:left; cursor:pointer; text-decoration:none"
                                            title="Forget about the JavaScript and join now.">
                                            Forget about the JavaScript and join now.
                                        </a>
                                </td>
                            </tr>
                        </table>
                    </div>
                </noscript>

                <div id="joinLauncherErrorDiv15" style="display: none; vertical-align:top;">
                <table cellpadding="0" cellspacing="0">
                    <tbody><tr>
                        <td colspan="2">
                            <label id="errorTextLabel15" class="errorbold">We're having trouble getting you into the meeting.</label>
                        </td>
                    </tr>
                    <tr>
                        <td id="errorHeader2" colspan="2" class="errorHeader2">
                            <label id="checkUrlLabel15" class="errorregular">It's
 possible you're using a bad URL. Try calling into the meeting using the
 phone number on the invite, or ask the organizer to drag you into the 
meeting from the Contacts list.</label>
                        </td>
                    </tr>
                    <tr>
                        <td id="errorBulletText" class="bulletTextError">
                            <label id="diagLabel15" class="bulletpoint" style="vertical-align:top; cursor:pointer" onclick="togglediag()">Share this error with your admin</label>
                        </td>
                    </tr>
                    <tr style="padding-bottom:6px">
                        <td colspan="2">
                            <label id="diagLabel215" class="errorlow" style="display:none">Copy the diagnostic info and send it to your admin.</label>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <textarea id="diagInfoText15" cols="70" rows="3" style="font-size: 8px; overflow: auto; display: none;" readonly="readonly" draggable="false">Machine:DC01-SFBFE01Join
 attempted at(UTC):11/19/2015 5:52:51 PMCorrelationId:746613143User 
Agent:Mozilla/5.0 (Windows NT 6.1; WOW64; rv:42.0) Gecko/20100101 
Firefox/42.0Accept 
Header:text/htmlapplication/xhtml+xmlapplication/xml;q=0.9*/*;q=0.8Incoming
 
URL:https://meet.integrationpartners.com/shargis/PSDFOZC4TelemetryId:2beb5408-4d9a-4739-821e-b5c1b3a8a169</textarea>
                        </td>
                    </tr>
                </tbody></table>
                </div>
                <div id="UnsupportedClientBlockDiv15" style="display: none">
                <table cellpadding="0" cellspacing="0">
                    <tbody><tr>
                        <td colspan="2">
                            <label id="UnsupportedClientBlockMessageLabel15" style="display:block;" class="errorregular">You
 have a version of Skype for Business that doesn't support joining this 
online meeting. Contact your support team to upgrade to a newer version 
of Skype for Business.</label>
                        </td>
                    </tr>
                    <tr>
                        <td class="bulletTextError">
                            <a id="JoinMeetingUsingLWA_Link15" class="bulletpoint" style="text-align:left; cursor:pointer; text-decoration:none" href="" target="_blank" title="">Join Using Skype for Business Web App instead</a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <label id="JoinUsingLWAOption115" style="display:block;" class="errorlow">Skype for Business Web App</label>
                        </td>
                    </tr>

                </tbody></table>
                </div>

                <div id="launchRichClientDiv15" style="display: block; vertical-align: top;">
                <table cellpadding="0" cellspacing="0">
                    <tbody><tr>
                        <td colspan="2">
                            <label id="launchRichClientHeaderLabel15" style="display:block;" class="errorregular">You are joining the meeting.</label>
                            <label id="launchRichClientTextLabel15" style="display:block;" class="errorlow">You may close this web page now.</label>
                        </td>
                    </tr>
                    <tr>
                        <td class="bulletTextError">
                            <a id="JoinMeetingUsingLWA2_Link15" class="bulletpoint" style="text-align:left; cursor:pointer; text-decoration:none" href="https://sfbwebext.integrationpartners.com/lwa/WebPages/LwaClient.aspx?legacy=RmFsc2U%21&amp;xml=PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48Y29uZi1pbmZvIHhtbG5zOnhzZD0iaHR0cDovL3d3dy53My5vcmcvMjAwMS9YTUxTY2hlbWEiIHhtbG5zOnhzaT0iaHR0cDovL3d3dy53My5vcmcvMjAwMS9YTUxTY2hlbWEtaW5zdGFuY2UiIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL3J0Yy8yMDA5LzA1L3NpbXBsZWpvaW5jb25mZG9jIj48Y29uZi11cmk-c2lwOnNoYXJnaXNAaW50ZWdyYXRpb25wYXJ0bmVycy5jb207Z3J1dTtvcGFxdWU9YXBwOmNvbmY6Zm9jdXM6aWQ6UFNERk9aQzQ8L2NvbmYtdXJpPjxzZXJ2ZXItdGltZT43Mi4wMDI8L3NlcnZlci10aW1lPjxvcmlnaW5hbC1pbmNvbWluZy11cmw-aHR0cHM6Ly9tZWV0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3NoYXJnaXMvUFNERk9aQzQ8L29yaWdpbmFsLWluY29taW5nLXVybD48Y29uZi1rZXk-UFNERk9aQzQ8L2NvbmYta2V5PjxmYWxsYmFjay11cmw-aHR0cHM6Ly9tZWV0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3NoYXJnaXMvUFNERk9aQzQ_c2w9PC9mYWxsYmFjay11cmw-PHVjd2EtdXJsPmh0dHBzOi8vc2Zid2ViZXh0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3Vjd2EvdjEvYXBwbGljYXRpb25zPC91Y3dhLXVybD48dWN3YS1leHQtdXJsPmh0dHBzOi8vc2Zid2ViZXh0LmludGVncmF0aW9ucGFydG5lcnMuY29tL3Vjd2EvdjEvYXBwbGljYXRpb25zPC91Y3dhLWV4dC11cmw-PHVjd2EtaW50LXVybD5odHRwczovL3NmYndlYmludC5pcGMubG9jYWwvdWN3YS92MS9hcHBsaWNhdGlvbnM8L3Vjd2EtaW50LXVybD48dGVsZW1ldHJ5LWlkPjJiZWI1NDA4LTRkOWEtNDczOS04MjFlLWI1YzFiM2E4YTE2OTwvdGVsZW1ldHJ5LWlkPjwvY29uZi1pbmZvPg%21%21&amp;launched=oc" target="_blank" title="">Join Using Skype for Business Web App instead</a>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <label id="JoinUsingLWAOptionHelp115" style="display:block;" class="errorlow">Skype for Business Web App</label>
                        </td>
                    </tr>

                </tbody></table>
                </div>

                <div id="mobileinterimDiv15" style="display: none;">
                <table cellpadding="0" cellspacing="0">
                    <tbody><tr>
                        <td colspan="2">
                            <label id="connectingLabel15" style="display:block;" class="errorregular">Connecting...</label>
                        </td>
                    </tr>
                </tbody></table>
                </div>

                <div id="mobileappstoreDiv15" style="display: none;">
                <table cellpadding="0" cellspacing="0">
                    <tbody><tr>
                        <td colspan="2">
                            <label id="mobileappstoreLabel15" style="display:block;" class="errorregular">Have you successfully joined the meeting? Feel free to close this page.</label>
                        </td>
                    </tr>
                    <tr height="10px"> </tr>
                    <tr>
                        <td colspan="2">
                            <label id="mobileappstoreLabel315" style="display:block;" class="errorregular">OR</label>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding-top:25px; vertical-align:top; text-align:left">
                            <a id="mobileappstoreLabel215" class="bulletpointMobile" style="vertical-align:top; text-align:left; cursor:pointer; text-decoration:none" href="" target="_blank" title="">Don’t have Skype for Business app? Install now</a>
                        </td>
                    </tr>

                </tbody></table>
                </div>

                <div id="UnknownMobileDeviceDiv15" style="display: none;">
                <table cellpadding="0" cellspacing="0">
                    <tbody><tr>
                        <td colspan="2">
                            <label id="unknownMobileDeviceLabel15" style="display:block;">We’re sorry, but you can’t join the meeting using this device. Please try using the phone numbers in the invitation.</label>
                        </td>
                    </tr>
                </tbody></table>
                </div>

            </td>
            <td>
        </td></tr>
        <tr class="footerSuccess1280" id="footer">
            <td id="copyright" class="copyright" style="display: block;">© 2014 Microsoft Corporation. All rights reserved.</td>
        </tr>
    </tbody></table>
    <table id="mainTablemobile" style="background-color:#0A92B9; width:100%; height:100%; display:none" align="center" cellpadding="0" cellspacing="0">
        <tbody><tr id="contentRowMobile">
            <td id="ContentCellMobile" class="mainContentCellMobile">
                <table id="mainContentTableMobile" class="mainContentTableMobile" border="0">
                    <tbody><tr id="topLogoRowMobile">
                        <td id="topLogoCellMobile" class="topLogoCellMobile">
                            <img id="communicatorLogoImageMobile" name="communicatorLogoImage" src="cebook_files/LyncLogo2011.png" alt="">
                        </td>
                    </tr>
                    <tr id="landingPageTextContentRowMobile">
                        <td id="landingPageTextContentCellMobile">
                            <noscript id="noScriptContentMobile">
                                <div id="javaScriptDisabledDivMobile">
                                    <table cellpadding="0" cellspacing="0" align="center" class="contentHeaderTableMobile">
                                        <tr>
                                            <td>
                                                <label id="javaScriptDisabledHeaderLabelMobile" class="bodyBoldTextMobile">
                                                    Never want to see this screen again?
                                                </label>
                                                <br id="Br1" />
                                                <label id="javaScriptDisabledTextLabelMobile" class="bodyTextMobile">
                                                    Just enable JavaScript in your browser settings, and you can skip this screen and join meetings automatically.
                                                </label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                              <br/>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <a id="javaScriptDisabledLaunchRichClientLinkMobile" name="javaScriptDisabledLaunchRichClientLink"
                                                    href="conf:sip:shargis@integrationpartners.com;gruu;opaque=app:conf:focus:id:PSDFOZC4%3Frequired-media=audio" class="bodyTextMobile" tabindex="2"
                                                    title="Forget about the JavaScript and join now.">
                                                    Forget about the JavaScript and join now.
                                                </a>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </noscript>
                            
                            <div id="joinLauncherErrorDivMobile" style="display: none;">
                                <label id="errorTextLabelMobile" style="color: #FF0000; font-weight: bold;">
                                    Sorry, something went wrong, and we can't get you into the meeting.
                                </label>
                                <br id="Br2">
                                <label id="checkUrlLabelMobile" style="font-weight: normal;font-size: 12px;">
                                It's possible you're using a bad URL. 
Try calling into the meeting using the phone number on the invite, or 
ask the organizer to drag you into the meeting from the Contacts list.
                                </label>
                                <br id="Br3">
                                <label id="diagLabelMobile" style="font-weight: normal;font-size: 11px; text-decoration: underline; color: #0000FF;" onmouseover="this.style.cursor='hand'" onclick="togglediag()">
                                <br id="Br4">
                                More about this error
                                </label>
                                <br>
                                <label id="diagLabel2Mobile" style="font-weight: normal;font-size: 11px; display:none">
                                Copy this diagnostic info and send it to your admin.
                                </label>
                                <textarea id="diagInfoTextMobile" cols="70" rows="5" style="font-size: 8px; overflow:auto" readonly="readonly">                                </textarea>
                            </div>

                            <div id="mobileinterimDiv" style="display: none;">
                                <label id="connectingLabel" style="display:block;" class="bodyTextMobile">
                                Connecting...
                                </label>
                            </div>
                            <div id="mobileappstoreDiv" style="display: none;" class="bodyTextMobile">
                                <label id="mobileappstoreLabel" style="display:block;" class="bodyBoldTextMobile">
                                In the meeting? You can go ahead and close this page.
                                </label>
                                <br>
                                <br>
                                <a id="mobileappstoreLabel2" style="display:block;">
                                Install Skype for Business and enjoy one-click join from your phone!
                                </a>
                            </div>
                            <div id="UnknownMobileDeviceDiv" style="display: none;">
                                <label id="unknownMobileDeviceLabel" style="display:block;">
                                Cannot join this meeting using this 
device. Please use any alternate mechanisms like phone numbers 
associated with the meeting to join the meeting
                                </label>
                            </div>
                        </td>
                    </tr>
                </tbody></table>
            </td>
        </tr>
        <tr>
            <td>
                <table id="Table1" class="footerTableMobile" align="center" cellpadding="0" cellspacing="0">
                    <tbody><tr id="Tr2">
                        <td style="padding-top:0px;">
                            <br id="Br6">
                            <a id="onlineHelpLinkMobile" name="onlineHelpLink" href="http://go.microsoft.com/fwlink/?LinkId=190632" class="bodyTextMobile" target="_blank" tabindex="4" title="Online Help">
                                Online Help
                            </a>
                        </td>
                    </tr>
                    <tr>
                       <td>
                           <br>
                       </td>
                    </tr>
                    <tr id="Tr3">
                        <td id="Td1" class="footerTableCellMobile">
                            <img id="Img2" name="officeLogoImage" src="cebook_files/OfficeLogo2011.png" alt="">
                            <br>
                            <label id="copyRightTextLabelMobile" class="smallerBodyTextMobile">
                                © 2010 Microsoft Corporation. All rights reserved.
                            </label>
                        </td>
                    </tr>
                </tbody></table>
            </td>
        </tr>
        <tr id="Tr4" align="center">
            <td id="Td2" class="bodyTextMobile">
                <div id="languageSettingsDivMobile" style="display: block;">
                    <label id="languageSettingsLabelMobile">
                        Language: &nbsp;
                    </label>
                    <select id="languageSelectCmbMobile" onchange="mainWindow.LanguageSelectionChanged();" tabindex="1">
                        <option selected="selected" value="ar">العربية‏</option><option value="az">Azərbaycan­ılı‎</option><option value="be">Беларуская‎</option><option value="bg">български‎</option><option value="ca">Català‎</option><option value="cs">čeština‎</option><option value="da">dansk‎</option><option value="de">Deutsch‎</option><option value="el">Ελληνικά‎</option><option value="en-US">English (United States)‎</option><option value="es">español‎</option><option value="et">eesti‎</option><option value="eu">euskara‎</option><option value="fa">فارسى‏</option><option value="fi">suomi‎</option><option value="fil">Filipino‎</option><option value="fr">français‎</option><option value="gl">galego‎</option><option value="he">עברית‏</option><option value="hi">हिंदी‎</option><option value="hr">hrvatski‎</option><option value="hu">magyar‎</option><option value="id">Bahasa Indonesia‎</option><option value="it">italiano‎</option><option value="ja">日本語‎</option><option value="kk">Қазақ‎</option><option value="ko">한국어‎</option><option value="lt">lietuvių‎</option><option value="lv">latviešu‎</option><option value="mk">македонски јазик‎</option><option value="ms">Bahasa Melayu‎</option><option value="nl">Nederlands‎</option><option value="no">norsk‎</option><option value="pl">polski‎</option><option value="pt-BR">português (Brasil)‎</option><option value="pt-PT">português (Portugal)‎</option><option value="ro">română‎</option><option value="ru">русский‎</option><option value="sk">slovenčina‎</option><option value="sl">slovenščina‎</option><option value="sq">Shqip‎</option><option value="sr-Cyrl">српски‎</option><option value="sr">srpski‎</option><option value="sv">svenska‎</option><option value="th">ไทย‎</option><option value="tr">Türkçe‎</option><option value="uk">українська‎</option><option value="uz">O'zbekcha‎</option><option value="vi">Tiếng Việt‎</option><option value="zh-Hans">中文(简体)‎</option><option value="zh-Hant">中文(繁體)‎</option>
                    </select>
                </div>
            </td>
        </tr>
    </tbody></table>
    
    <script>
        (function () {

            function isParentOriginSame() {
                try {
                    return location.host === parent.location.host;
                } catch (e) {
                    return false;
                }
            }

            function contains(whitelist, origin) {
                var i;
                for (i = 0; i < whitelist.length; i += 1) {
                    if (whitelist[i] === origin) {
                        return true;
                    }
                }
                return false;
            }

            function disablePage() {
                window.onmessage = null;
                document.body.innerHTML = '';
            }

            function initPage() {
                mainWindow = new MainForm();
                mainWindow.OnLoad();
            }

            function initOriginCheck() {
                var timeoutId = setTimeout(disablePage, 500),
                    messageId = String(Math.random() * 0x80000000 | 0), // a relatively unique uint31
                    origins = eval(parentOriginWhitelist); // IE8 quirks mode - no JSON.parse()

                window.onmessage = function (event) {
                    event = event || window.event;
                    if (event.data === messageId) {
                        clearTimeout(timeoutId);
                        window.onmessage = null;
                        if (contains(origins, event.origin)) {
                            initPage();
                        } else {
                            disablePage();
                        }
                    }
                };

                parent.postMessage(messageId, '*');
            }

            if (!parentOriginJsCheckEnabled || parent === window || isParentOriginSame()) {
                initPage();
            } else {
                reachClientRequested = 'True'; // force web app when embedded
                initOriginCheck();
            }
        }());
    </script>


<div style="position: absolute; width: 1px; height: 1px; left: -100px; top: 0px; overflow: hidden; visibility: hidden;"><embed id="_ucclient_plugin_OC" type="application/vnd.microsoft.communicator.ocsmeeting"></div></body></html>