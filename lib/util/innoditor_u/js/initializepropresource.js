﻿if("" == g_strCustomLanguageSetAndLoad){if("ko" == g_browserCHK.language){indr010541536 = g_browserCHK.language;}else if("ko-kr" == g_browserCHK.language){indr010541536 = "ko";}else if("en" == g_browserCHK.language){indr010541536 = g_browserCHK.language;}else if("ja" == g_browserCHK.language){indr010541536 = g_browserCHK.language;}else if("zh-cn" == g_browserCHK.language){indr010541536 = g_browserCHK.language;}else{indr010541536 = "en";}}else{indr010541536 = g_strCustomLanguageSetAndLoad.toLowerCase();}document.write('<link rel="stylesheet" href="' + g_strPath_CSS + indr010541536 + '/colorwin.css" type="text/css">');document.write('<link rel="stylesheet" href="' + g_strPath_CSS + indr010541536 + '/propwin.css" type="text/css">');document.write('<script type="text/javascript" src="' + g_strPath_JS + 'res/' + indr010541536 + '/resource_prop.js"></scrip' +'t>');if(!g_bCustomColorTableUseYN){if(1 == g_nCustomColorBasicDetailUseType){document.write('<script type="text/javascript" src="' + g_strPath_JS + 'layer/lycr00000prop.js"></scrip' +'t>');}else if(2 == g_nCustomColorBasicDetailUseType){document.write('<script type="text/javascript" src="' + g_strPath_JS + 'layer/lycr10000prop.js"></scrip' +'t>');}}