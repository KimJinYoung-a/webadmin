function BrowserInfo() 
{
    this.isMSIE = false;
    this.isFirefox = false;
    this.isChrome = false;
    this.isSafari = false;
    this.etc = false; // IE, FF를 제외한 브라우저

    try {
        var x = ActiveXObject;
        this.isMSIE = true;
    } catch (e) {
        this.isMSIE = false;
    }
    var userAgentString = navigator.userAgent.toLowerCase();
    if (userAgentString.indexOf("chrome") > -1) this.isChrome = true;
    if (userAgentString.indexOf("firefox") > -1) this.isFirefox = true;
    if (userAgentString.indexOf("safari") > -1) this.isSafari = true;
    if (this.isFirefox) this.isWindowlessSupported = "true";
    if (!this.isMSIE && !this.isFirefox) this.etc = true;
}

var GBrowser = new BrowserInfo();
var licensekey = 'hTBMSylD3xqJrfOFCJI5EtLWiMbyWH9k6s1uJuRbDxI=';	// for localhost

function ExistPlugin() 
{
    var mimetype = navigator.mimeTypes["application/x-tabsfileup"];
    if (mimetype) 
	{
        var enablePlugin = mimetype.enabledPlugin;
        if (enablePlugin)
            return true;
        else
            return false;
    }
    else 
	{
        return false;
    }
}