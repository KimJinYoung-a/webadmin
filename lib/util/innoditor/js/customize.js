/////////////////////////////////////////////////////////////////////////////////////////////////////
// editor ����(1��~20��) �� editor�� �ε��� ������ ID ����
// ���������� editor�� ������ �����ϰ��� �ϴ� ���� ���������� ������ ��
// ����Ʈ ��ü�� �ϰ������� 1�� �Ǵ� N�� �ε��ϴ� ���� �Ʒ��κ��� �ּ��� �����ϰ� ����

//var g_arrSetEditorArea = new Array();
//g_arrSetEditorArea[0] = "EDITOR_AREA_CONTAINER";// �̳���͸� ��ġ��ų ������ ID�� ����
/////////////////////////////////////////////////////////////////////////////////////////////////////



// skin ����(0~9������ skin ����)
var g_nSkinNumber = 0;

var g_strPath_Image = "/lib/util/innoditor/image/";
var g_strPath_JS = "/lib/util/innoditor/js/";
var g_strPath_CSS = "/lib/util/innoditor/css/";
var g_strPath_Property = "/lib/util/innoditor/";
var g_strPath_License = "/lib/util/innoditor/";

var g_nEditorWidth = 670;
var g_nEditorHeight = 600;


// ��ǰ ���� URL
var g_strHelpPageURL = "/lib/util/innoditor/pop_help.html";


// ��ǰ���� URL
var g_strProductInfoURL = "/lib/util/innoditor/pop_productinfo.html";
//var g_strProductInfoURL = "/lib/util/innoditor/pop_productinfo_en.html";// ������� ����


// Image ���ε� Page URL
//var g_strUploadImageURL = "/lib/util/innoditor/pop_simple_img.html";//����
var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img.asp"; //2013-01-17 ���� ����
//var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img.html";
//var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img_en.html";// ������� ����
//var g_strUploadImageURL = "/lib/util/innoditor/pop_link_img.html";// �ܺθ�ũ�� ����(�ѱ�)
//var g_strUploadImageURL = "/lib/util/innoditor/pop_link_img_en.html";// �ܺθ�ũ�� ����(����)


// Flash ���ε� Page URL
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_simple_flash.html";//����
var g_strUploadFlashURL = "/lib/util/innoditor/pop_upload_flash.html";
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_upload_flash_en.html";// ������� ����
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_link_flash.html";// �ܺθ�ũ�� ����(�ѱ�)
//var g_strUploadFlashURL = "/lib/util/innoditor/pop_link_flash_en.html";// �ܺθ�ũ�� ����(����)


// Media ���ε� Page URL
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_simple_media.html";//����
var g_strUploadMediaURL = "/lib/util/innoditor/pop_upload_media.html";
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_upload_media_en.html";// ������� ����
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_link_media.html";// �ܺθ�ũ�� ����(�ѱ�)
//var g_strUploadMediaURL = "/lib/util/innoditor/pop_link_media_en.html";// �ܺθ�ũ�� ����(����)


// ��� Image ���ε� �� ���� Page URL
//var g_strUploadBackgroundImageURL = "/lib/util/innoditor/pop_simple_img_bg.html";//����
var g_strUploadBackgroundImageURL = "/lib/util/innoditor/pop_upload_img_bg.html";
//var g_strUploadBackgroundImageURL = "/lib/util/innoditor/pop_upload_img_bg_en.html";// ������� ����


// ���� Templete ���� Page URL
var g_strInsertDocTempleteURL = "/lib/util/innoditor/pop_doc_templete.html";
//var g_strInsertDocTempleteURL = "/lib/util/innoditor/pop_doc_templete_en.html";// ������� ����


// �Ӽ� Page URL
var g_strPropertyPageURL = "/lib/util/innoditor/pop_property.html";
//var g_strPropertyPageURL = "/lib/util/innoditor/pop_property_en.html";// ������� ����


// �̸����� Page URL
//var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview.html";
//var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview_en.html";// ������� ����
var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview_x.html";// XHTML ��¹������ ������ ���
//var g_strPreviewPageURL = "/lib/util/innoditor/pop_preview_x_en.html";// XHTML ��¹������ ������ ���(����)



// ���̼���
var g_arrDomainName = new Array();
g_arrDomainName[0] = "localhost";
g_arrDomainName[1] = "webadmin.10x10.co.kr";
g_arrDomainName[2] = "testwebadmin.10x10.co.kr";

var g_arrLicenseKey = new Array();
g_arrLicenseKey[0] = "Mv5Oi$BZ+q3Pm/Lq4h@MX4Nh#AYs26EYo&@Tbq3Pm/Lhf+Ap7Y{D";
g_arrLicenseKey[1] = "Ak1Rs9ap!5FUpl<nFwO#!FgKkuP*a:-{U/d>qK.DuM~V.a7:-2Ol.Kh&j4g?pJ%BYr*f*Kk0";
g_arrLicenseKey[2] = "y@b%Km4VxAl1DUhy1MPq7XyAb*p7Y{D4Kl2S#41@3^5hbO&_9oJxT.c<pJ}W**QwDjLO&_9oJapJ}W1P[4Yj+If$Eb}A1VT%BYrqE";



// �޴��� show �Ǵ� hidden �� ����(�޴����̾�)
var g_bCustomize_MenuBar_Display = true;

// Bottom Tab�� show �Ǵ� hidden �� ����(�̸�����,����â,�ҽ�â ��ư)
var g_bCustomize_TabBar_Display = true;

// ù��° ���� show �Ǵ� hidden ����(���� �׸��� ���� ����)
var g_bCustomize_ToolBar1_Display = true;

// �ι�° ���� show �Ǵ� hidden ����(���� �׸��� ���� ����)
var g_bCustomize_ToolBar2_Display = true;

// ����° ���� show �Ǵ� hidden ����(���� �׸��� ���� ����)
var g_bCustomize_ToolBar3_Display = true;

// ����� ���� ����(�̳���Ϳ��� �����Ǵ� ����� ������� ���� ��� �̳���� Interface �� ����)
var g_bCustomize_CustomToolbar_Display = false;
var g_bCustomize_CustomToolbar_Layout = 0;// 0 - �ش���� ����, 1 - Top(���� ��� ����), 2 - Bottom(���� �ϴ� ����)
var g_bCustomize_CustomToolbar_HTML = "";// ��������� ���ٿ� �� HTML ����(<table> ~ </table> : table�� �����Ͽ� table�� ������ ��)


// ���� ���ÿ� ���� ���� (���� ���� ���� �׸� ������ customize_ui.js ����)
var g_arrCustomToolbar1 = new Array();
var g_arrCustomToolbar2 = new Array();
var g_arrCustomToolbar3 = new Array();


document.write('<script type="text/javascript" src="' + g_strPath_JS + 'browser.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr489343715.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr670454868.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr873475877.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr528318566.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr696495397.js"></scrip' +'t>');
document.write('<script type="text/javascript" src="' + g_strPath_JS + 'indr988789177.js"></scrip' +'t>');
