/* Browser Related Constants */
var INTERNET_EXPLORER = "Microsoft Internet Explorer";
var FIREFOX = "FireFox";
var CHROME = "Chrome"
var SAFARI = "Safari";

/* OS Related Constants */
var WINDOWS = "Windows";
var MAC = "Mac";
var LINUX = "Linux";

/* Variables */
var browserName = ""; // INTERNET_EXPLORER, FIREFOX, CHROME, SAFARI
var browserVersion = 0.0;
var tridentVersion = 0.0; // Trident Token for IE
var osName = ""; // WINDOWS, MAC, LINUX
var osVersion = 0.0;
var platform = ""; // 'win32', 'win64', 'mac', 'linux', 'other'
var browserArch = "";

/* Browser Detection */
if (/MSIE (\d+\.\d+)/.test(navigator.userAgent)) {
    browserName = INTERNET_EXPLORER;
    browserVersion = RegExp.$1;
    var regularExpression = new RegExp(/Trident\/(\d+\.\d+)/);
    if (regularExpression.test(navigator.userAgent)) {
        tridentVersion = RegExp.$1;
    }
} else if (/Trident\/\d+\.\d+/.test(navigator.userAgent) && /\brv:(\d+\.\d+)/.test(navigator.userAgent)) {
    // IE11 removed MSIE from the userAgent and added rv:11.0
    browserName = INTERNET_EXPLORER;
    browserVersion = RegExp.$1;
    var regularExpression = new RegExp(/Trident\/(\d+\.\d+)/);
    if (regularExpression.test(navigator.userAgent)) {
        tridentVersion = RegExp.$1;
    }
} else if (/Firefox[\/\s](\d+\.\d+)/.test(navigator.userAgent)) {
    browserName = FIREFOX;
    browserVersion = RegExp.$1;
} else if (/Chrome[\/\s](\d+\.\d+)/.test(navigator.userAgent)) {
    browserName = CHROME;
    browserVersion = RegExp.$1;
} else if (/Safari/.test(navigator.userAgent) && !/Chrome/.test(navigator.userAgent)) {
    browserName = SAFARI;
    if (/Version[\/\s](\d+\.\d+)/.test(navigator.userAgent)) {
        browserVersion = RegExp.$1;
    }
}

/* OS Detection */
if (/Windows NT (\d+\.\d+)/.test(navigator.userAgent)) {
    osName = WINDOWS;
    osVersion = RegExp.$1;
} else if (/Intel Mac OS X/.test(navigator.userAgent)) {
    osName = MAC;
    if(/Intel Mac OS X (\d+\.\d+)/.test(navigator.userAgent)){
        osVersion = RegExp.$1;
        if(/Intel Mac OS X (\d+\.\d+\.\d+)/.test(navigator.userAgent)){
            osVersion = RegExp.$1;
        }
    } else if(/Intel Mac OS X (\d+\_\d+)/.test(navigator.userAgent)){
        osVersion = (RegExp.$1).replace('_', '.');
        if(/Intel Mac OS X (\d+\_\d+\_\d+)/.test(navigator.userAgent)) {
            osVersion = RegExp.$1.replace('_', '.');
        }
    }
} else if (navigator.userAgent.indexOf('Linux') >= 0) {
    osName = LINUX;
    //TODO: version
}

/* Platform Detection */
var info = navigator.platform.toLowerCase();
if (info.match(/win32/)) {
    platform = "win32";
} else if (info.match(/win64/)) {
    platform = "win64";
} else if (info.match(/mac/)) {
    platform = "mac";
} else if (info.match(/linux/)) {
    platform = "linux";
} else {
    // Do we need to identify hp, and sun platform?
    platform = "other";
}

if (isPlatformWin32() || isPlatformWin64()) {
    var info = navigator.userAgent.toLowerCase();

    if (matchResult = info.match(/win64; x64/)) {
        browserArch = "amd64";
    } else if (matchResult = info.match(/win64; ia64/)) {
        browserArch = "ia64";
    } else {
        browserArch = "x86";
    }
}

/* OS Methods */
function isWindowsOS()          { return (WINDOWS == osName); }
function isMacOS()              { return (MAC == osName); }
function isLinuxOS()            { return (LINUX == osName); }
function isWindows2k()          { return osVersion == 5.0; }
function isWindowsXP()          { return osVersion == 5.1; }
function isWindowsXPx64Or2k3()  { return osVersion == 5.2; }
function isWindowsVista()       { return osVersion == 6.0; }
function isWin7Or2k8()          { return osVersion == 6.1; }
function isWin8()               { return osVersion == 6.2; }
function isWin8OrLater()        { return isWindowsOS() && osVersion >= 6.2; } 
function isIntelBasedMacOs10x() { return isMacOS() && osVersion >= '10.4.8';}
function getOSVersion()         { return osVersion; }
function isArm()                { return navigator.userAgent.toLowerCase().indexOf("arm") >= 0; }

/* Browser Methods */
function isIE()                 { return browserName === INTERNET_EXPLORER; }
function isFF()                 { return FIREFOX == browserName; }
function isChrome()             { return CHROME == browserName; }
function isSafari()             { return SAFARI == browserName; }
function isIE8()                { return isIE() && browserVersion == 8.0; }
function isIE7()                { return isIE() && browserVersion == 7.0; }
function isIE6()                { return isIE() && browserVersion == 6.0; }

// IE9/8 send the “MSIE 7.0” version information when viewing sites with
// Compatibility View enabled.  We need to rely on the “Trident” token
// in the User-Agent string to detect IE9/8 clients when they are using
// the Compatibility View feature.
function isIE9TridentVersion()  { return (isIE() && tridentVersion == 5.0); }
function isIE8TridentVersion()  { return (isIE() && tridentVersion == 4.0); }
function isIE10TridentVersion() { return (isIE() && tridentVersion == 6.0); }

function isIE8OrLater() { return (isIE() && tridentVersion >= 4.0); }
function isIE7OrLater() { return isIE7() || (isIE() && tridentVersion >= 4.0); }
function isIE6OrLater() { return isIE6() || isIE7OrLater(); }

function isFF3x()               { return (FIREFOX == browserName && browserVersion >= 3.0 && browserVersion < 4.0); }
function isSafari4x()           { return (SAFARI == browserName && browserVersion >= 4.0 && browserVersion < 5.0); }
function isSafari5x()           { return (SAFARI == browserName && browserVersion >= 5.0 && browserVersion < 6.0); }
function getBrowserVersion()    { return browserVersion; }
function getBrowserArch()       { return browserArch; }

/* Platform Methods */
function isPlatformWin32()                 { return (platform == "win32"); }
function isPlatformWin64()                 { return (platform == "win64"); }

/* Is RTL */
function isRtl() { return window.document.dir == "rtl"; }

/* Is Hebrew */
function isHebrew(languageName) {  return languageName.indexOf("he") == 0; }

function isSupportedOSAndBrowserVersion() {
    if (isWin7Or2k8() && (isIE8() || isIE7() || isFF3x())) {
        return true;
    } else if (isWindowsVista() && (isIE8() || isIE7() || isFF3x())) {
        return true;
    } else if (isWindowsXP() && (isIE8() || isIE7() || isIE6() || isFF3x())) {
        return true;
    } else if (isWindowsXPx64Or2k3() && (isIE8() || isIE7() || isIE6() || isFF3x())) {
        return true;
    } else if (isWindows2k() && isIE6()) {
        return true;
    } else if (isIntelBasedMacOs10x() && (isFF3x() || isSafari4x() || isSafari5x())) {
        return true;
    } else if (isMacOS() && osVersion >= '10.4' && isFF3x()) {
        return true;
    }
    return false;
}
function isSupportedOSAndBrowser() {
    return isWindowsOS() || isMacOS();
}

function isBlockedPlatform() {
    return (isWindowsOS() && isIE() && getBrowserArch() !== "x86") ||
           (isMacOS() && isChrome());
}

/* Cookie Related Methoda */
function isCookieEnabled() {return navigator.cookieEnabled;}

function createCookie(cookieName, cookieValue, expiryDays) {
    var expires = "";
    if (expiryDays) {
        var currDate = new Date();
        currDate.setTime(currDate.getTime() + (expiryDays * 24 * 60 * 60 * 1000));
        expires = "; expires=" + currDate.toGMTString();
    }
    document.cookie = cookieName + "=" + cookieValue + expires + "; path=/";
}
function readCookie(cookieName) {
    var nameEQ = cookieName + "=";
    var cookieAttributes = document.cookie.split(';');
    for (var i = 0; i < cookieAttributes.length; i++) {
        var c = cookieAttributes[i];
        while (c.charAt(0) == ' ') {
            c = c.substring(1, c.length);
        }
        if (c.indexOf(nameEQ) == 0) {
            return c.substring(nameEQ.length, c.length);
        }
    }
    return null;
}

function eraseCookie(cookieName) {
    createCookie(cookieName, "", -1);
}

/* Helper Methods */
function getUrlParameters() {
    var indexOfParam = window.location.href.indexOf("?");
    if (indexOfParam != -1) {
        return window.location.href.substring(indexOfParam, window.location.href.length);
    } else {
        return "";
    }
}

function ResizeTo(width, height) {
    try {
        window.resizeTo(width, height);
    } catch (e) {
        return "Name: " + e.name + " Message: " + e.Message;
    }
    return "";
}

function ResizeBy(width, height) {
    try {
        window.resizeBy(width, height);
    } catch (e) {
        return "Name: " + e.name + " Message: " + e.Message;
    }
    return "";
}

function IsInPopup() {
    if (window.opener && window.opener != null) {
        return true;
    }
}

function GetScreenWidth() {
    return window.screen.width;
}

function GetScreenHeight() {
    return window.screen.height;
}

function GetLoginScreenHeight() {
    var minScreenHeight = 560;
    var maxScreenHeight = 650;
    var screenHeight = minScreenHeight;
    try {
        screenHeight = 0.85 * window.screen.height;
        if (screenHeight < minScreenHeight)
            screenHeight = minScreenHeight;
        else if (screenHeight > maxScreenHeight)
            screenHeight = maxScreenHeight;
    } catch (e) { screenHeight = minScreenHeight; }

    return screenHeight;
}

function ResizeAndMove(width, height, x, y) {
    try {
        window.moveTo(x, y);
        window.resizeTo(width, height);
    } catch (e) {
        return "Name: " + e.name + " Message: " + e.Message;
    }
    return "";
}

function ResizeToMaxAvail() {
    try {
        window.moveTo(0, 0);

        var availWidth = window.screen.availWidth;
        var availHeight = window.screen.availHeight;

        window.resizeTo(availWidth, availHeight);
    } catch (e) {
        return "Name: " + e.name + " Message: " + e.Message;
    }
    return "";
}

function GoToMaxSizeOnNonIEBrowsers() {
    if (!isIE()) {
        return ResizeToMaxAvail();
    }
    return "";
}

// This function is similar to Delegate. But it is used in Timer.
// For Firefox/Mozilla, a weird issue is that upon timeout,
// the first parameter sent to handler is always timer ID.
// The handler created by this function will deal with this issue.
function TimerHandler(contextObj, funcObj)
{
    var _arguments = null;

    if( arguments.length > 2)
    {
        _arguments = new Array();

        for( var i = 2; i < arguments.length; i ++ )
        {
            _arguments[i - 2] = arguments[i];
        }
    }

    var _delegate =  function()
    {
        if( _arguments )
        {
            funcObj.apply( contextObj, _arguments );
        }
        else
        {
            funcObj.apply( contextObj );
        }
    };
    return _delegate;
}

//=================================================
// Delegate
//  Unlike C#, javascript has the first-class function
//  object and it is weak-typed.
//  The problem of javascript function object is that:
//  it does not include the execution context information.
//  The function 'delegate' is to encapsulate an
//  execution context with a function object, and hence
//  make a delegate object.
//=================================================

// delegate function, it needs that the function object should have build-in 'call' method,
// which is not supported on some down-level browsers like IE5.01, but all of our supported
// browsers will have this feature.
function Delegate(contextObj, funcObj) {
    var _arguments = null;

    if (arguments.length > 2) {
        _arguments = new Array();
        for (var i = 2; i < arguments.length; i++) {
            _arguments[i - 2] = arguments[i];
        }
    }

    var _delegate = function () {
        if (_arguments) {
            var args;
            if (arguments.length > 0) {
                args = new Array();
                args.push.apply(args, arguments);
                args.push.apply(args, _arguments);
            }
            else {
                args = _arguments;
            }
        }
        else {
            args = arguments;
        }

        return funcObj.apply(contextObj, args);
    };

    return _delegate;
}
