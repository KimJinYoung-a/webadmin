
function BrowserNavigationBar() {
};

BrowserNavigationBar.prototype.init = function(url, options) {
	cordova.exec(this._onEvent, this._onError, "BrowserNavigationBar", "init", []);
};

if(!window.plugins) {
    window.plugins = {};
}

if (!window.plugins.browserNavigationBar) {
	window.plugins.browserNavigationBar = new BrowserNavigationBar();
}
