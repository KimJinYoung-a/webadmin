/*
 * TABSLAB- http://www.tabslab.com
 * Copyright 2003~2007 TABS LABORATORIES CORPORATION. All rights reserved. 
 *  
 * basePath:                   http://localhost/WebEditor/smart/
 * FCKConfig.BasePath:     http://localhost/WebEditor/fckeditor/editor/ 
 */
 
//FCKConfig.ImageUploadURL = '/WebEditor2/smart/Fileup/upload.asp';
//FCKConfig.ImageUploadURL = FCKConfig.BasePath + 'filemanager/connectors/asp/upload.asp';
 
FCKConfig.AutoDetectLanguage	= false ;
FCKConfig.DefaultLanguage		= 'ko' ;

//FCKConfig.StylesXmlPath           = '/tabsstyles.xml';
//FCKConfig.EditorAreaCSS           = '/tabsstyles.css';
//FCKConfig.TemplatesXmlPath       = '/tabstemplate.xml';

//FCKConfig.SpellChecker             = 'ieSpell';	// 'ieSpell' | 'SpellerPages'

FCKConfig.ToolbarCanCollapse = false;

//Formats
//FCKConfig.FontFormats = 'p;h4';
FCKConfig.EnterMode = 'br' ;			// p | div | br
FCKConfig.ShiftEnterMode = 'p' ;	// p | div | br
FCKConfig.FontNames = '돋움;돋움체;굴림;굴림체;바탕;바탕체;궁서;Arial;Tahoma;Times New Roman;Verdana' ;
FCKConfig.FontSizes = '7pt;8pt;9pt;10pt;11pt;12pt;13pt;14pt;18pt;24pt;36pt' ;

//Add Plugins. 
var basePath = FCKConfig.BasePath.substr(0, FCKConfig.BasePath.length - (('editor/').length + ('fckeditor/').length)) + 'smart/';
FCKConfig.Plugins.Add('Imageup', 'ko', basePath);

//Skin
//FCKConfig.SkinPath = FCKConfig.BasePath + 'skins/silver/' ;

// ToolbarSets
FCKConfig.ToolbarSets["TABSWebEditor"] = [
    ['Source','-','Preview','-','Templates','Cut','Copy','Paste','PasteText','PasteWord','Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat','Link','Unlink','Image','tabsimageup','Flash','Table','Rule','Smiley','SpecialChar','PageBreak','Print']
    ,'/',
    ['FontName','FontSize','TextColor','BGColor','Bold','Italic','Underline','StrikeThrough','-','Subscript','Superscript','OrderedList','UnorderedList','-','Outdent','Indent','Blockquote','-','JustifyLeft','JustifyCenter','JustifyRight','JustifyFull','ShowBlocks']
] ;
FCKConfig.ToolbarSets["Lite"] = [
    ['FontName','FontSize'],
    ['Bold','Italic','Underline','StrikeThrough','TextColor','BGColor'],
    ['JustifyLeft','JustifyCenter','JustifyRight','JustifyFull','OrderedList','UnorderedList','Outdent','Indent'],
    ['Link','Smiley','Table'],
    ['tabsimageup'],
    ['Source']
];

// localhost를 위한 라이센스키
FCKConfig.LicenseKey = '6DA46E01BB55F4254591DBC25DD46646';
