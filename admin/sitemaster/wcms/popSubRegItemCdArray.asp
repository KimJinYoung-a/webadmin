<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mainWCMSCls.asp" -->
<%
'###############################################
' PageName : popSubItemEdit.asp
' Discription : 서브컨텐츠 상품코드 일괄 등록
' History : 2013.05.14 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim mainIdx, i
Dim oTemplate, oMain
Dim tplIdx, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse
Dim isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isExtDataUse, isImgDescUse, tplinfoDesc, tplSortNo
Dim itemname, smallImage

Dim mainStartDate, mainEndDate, mainTitle, mainSubNum

Dim subImage1, subImage2, subLinkUrl, subText1, subText2, subItemid, subVideoUrl, subBGColor, subImageDesc
Dim subSortNo, subRegUserid, subRegDate, subLastModiUserid, subLastModiDate, subIsUsing

'// 파라메터 접수
mainIdx = request("mainIdx")

if mainIdx="" then
	Call Alert_Close("파라메터 오류(Err:01)")
	dbget.Close: Response.End
end if

'// 메인 내용
set oMain = new CCMSContent
	oMain.FRectMainIdx = MainIdx
	oMain.GetOneMainPage
	if oMain.FResultCount>0 then
		tplIdx = oMain.FOneItem.FtplIdx
		mainStartDate = oMain.FOneItem.FmainStartDate
		mainEndDate = oMain.FOneItem.FmainEndDate
		mainTitle = oMain.FOneItem.FmainTitle
		mainSubNum = oMain.FOneItem.FmainSubNum
	end if
set oMain = Nothing

if tplIdx="" then
	Call Alert_Close("존재하지 않거나 삭제된 내용입니다. (Err:02)")
	dbget.Close: Response.End
end if

'// 템플릿 내용
set oTemplate = new CCMSContent
oTemplate.FRectTplIdx = tplIdx
oTemplate.GetOneTemplate
if oTemplate.FResultCount>0 then
	tplType			= oTemplate.FOneItem.FtplType
	tplName			= oTemplate.FOneItem.FtplName
	isTimeUse		= oTemplate.FOneItem.FisTimeUse
	isIconUse		= oTemplate.FOneItem.FisIconUse
	isSubNumUse		= oTemplate.FOneItem.FisSubNumUse
	isTopImgUse		= oTemplate.FOneItem.FisTopImgUse
	isTopLinkUse	= oTemplate.FOneItem.FisTopLinkUse
	isImageUse		= oTemplate.FOneItem.FisImageUse
	isTextUse		= oTemplate.FOneItem.FisTextUse
	isLinkUse		= oTemplate.FOneItem.FisLinkUse
	isItemUse		= oTemplate.FOneItem.FisItemUse
	isVideoUse		= oTemplate.FOneItem.FisVideoUse
	isBGColorUse	= oTemplate.FOneItem.FisBGColorUse
	isExtDataUse	= oTemplate.FOneItem.FisExtDataUse
	isImgDescUse	= oTemplate.FOneItem.FisImgDescUse
	tplinfoDesc		= oTemplate.FOneItem.FtplinfoDesc
	tplSortNo		= oTemplate.FOneItem.FtplSortNo
end if
set oTemplate = Nothing

if isExtDataUse="Y" then
	Call Alert_Close("외부데이터를 사용하는 소재 템플릿입니다.\n소재를 등록할 수 없습니다.")
	dbget.Close: Response.End
end if

if Not(isItemUse="Y") then
	Call Alert_Close("상품코드를 사용하지 않는 소재 템플릿입니다.\n소재를 등록할 수 없습니다.")
	dbget.Close: Response.End
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
$(function(){
	//라디오버튼
	$("#rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});

// 폼검사
function SaveForm(frm) {
	var selChk=true;
	if(frm.subItemidArray.value=="") {
		alert("일괄 등록하실 상품코드를 입력해주세요");
		frm.subItemidArray.focus();
		return;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<center>
<form name="frmSub" method="post" action="doSubRegItemCdArray.asp" style="margin:0px;">
<input type="hidden" name="mainIdx" value="<%=mainIdx%>" />
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>소재 정보 - 상품코드 일괄 등록</b></td>
</tr>
<colgroup>
	<col width="100" />
	<col width="*" />
	<col width="100" />
	<col width="*" />
</colgroup>
<tr height="26" bgcolor="#FFFFFF">
    <td rowspan="2" bgcolor="#DDDDFF">템플릿</td>
    <td colspan="3">
        [<%=tplName %>]
        <b>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="3" style="padding:5px;">
        <%=nl2br(tplinfoDesc)%>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">상품코드</td>
    <td colspan="3">
        <textarea name="subItemidArray" class="textarea" title="상품코드" style="width:100%; height:80px;"></textarea>
        <p>※ 상품코드를 쉼표(,) 또는 엔터로 구분하여 입력</p>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">정렬순서</td>
    <td>
        <input type="text" name="subSortNo" class="text" size="4" value="0" />
    </td>
    <td bgcolor="#DDDDFF">사용여부</td>
    <td>
		<span id="rdoUsing">
		<input type="radio" name="subIsUsing" id="rdoUsing1" value="Y" checked /><label for="rdoUsing1">사용</label>
		<input type="radio" name="subIsUsing" id="rdoUsing2" value="N" /><label for="rdoUsing2">삭제</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" 저 장 " onClick="SaveForm(this.form);"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->