var ImageupCommand = function() { }
ImageupCommand.GetState = function() {
    return FCK_TRISTATE_OFF;
}

// window.open
ImageupCommand.prototype.Execute = function() { }
ImageupCommand.Execute = function() {
 //   window.open(FCKPlugins.Items['Imageup'].Path + 'popup.html', 'tabsimageup', 'width=400,height=120,scrollbars=no,scrolling=no,location=no,toolbar=no');
      window.open(FCKPlugins.Items['Imageup'].Path + 'popup.asp', 'tabsimageup', 'width=400,height=120,scrollbars=no,scrolling=no,location=no,toolbar=no');
}
FCKCommands.RegisterCommand('tabsimageup', ImageupCommand);

var ImageupButton = new FCKToolbarButton('tabsimageup', FCKLang.ImageupBtn, null, null, false, true);
//var ImageupButton = new FCKToolbarButton('tabsimageup', FCKLang.ImageupBtn, null, FCK_TOOLBARITEM_ICONTEXT, true, true);

if ( /\/editor\/skins\/(.*)\//.test(FCKConfig.SkinPath) )
	ImageupButton.IconPath = FCKPlugins.Items['Imageup'].Path + 'images/icon_' + RegExp.$1 + '.gif';
else
	ImageupButton.IconPath = FCKPlugins.Items['Imageup'].Path + 'images/editor_icon.gif';

FCKToolbarItems.RegisterItem('tabsimageup', ImageupButton);