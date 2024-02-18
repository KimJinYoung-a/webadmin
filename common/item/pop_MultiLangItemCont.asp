<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
response.charSet = "utf-8"
%>
<%
'####################################################
' Description :  온라인 다국어 상품 설명 입력 팝업
' History : 2013.07.10 허진원 생성
'			2016.08.30 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/itemMultiLangCls.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<%
Dim vItemID, vCountryCd, vCountryName, oitem, cOverSeas, vItemName, vItemContent, vItemCopy, vOriginListImage, vOriginItemName
dim vOriginMakerID, vOriginSellCash, vItemSource, vItemSize, vMakerName, vSourceArea, useyn, makerid
dim vItemNameKr,vItemCopyKr,vItemSourceKr,vItemSizeKr,vSourceAreaKr,vMakerNameKr,vKeywordsKr
dim orgprice, oitemoption, keywords, i, bChkReg

	vItemID = requestCheckVar(Request("itemid"),10)
    vCountryCd = requestCheckVar(Request("lang"),32)
    if vCountryCd="" then vCountryCd="EN"		'기본값 "영어"
    vCountryName = getCountryCdName(vCountryCd)

If vItemID <> "" Then
	set cOverSeas = new CMultiLang
		cOverSeas.FRectItemId = vItemID
		cOverSeas.FRectCountryCd = vCountryCd
		cOverSeas.GetMultiLangItemInfo

		vItemName = cOverSeas.FOneItem.Fitemname
		vItemContent = cOverSeas.FOneItem.Fitemcontent
		vItemCopy = cOverSeas.FOneItem.Fitemcopy
		vItemSource = cOverSeas.FOneItem.Fitemsource
		vItemSize = cOverSeas.FOneItem.Fitemsize
		vMakerName = cOverSeas.FOneItem.Fmakername
		vSourceArea = cOverSeas.FOneItem.Fsourcearea
		useyn = cOverSeas.FOneItem.fuseyn
		keywords = cOverSeas.FOneItem.fkeywords
		bChkReg = cOverSeas.FOneItem.FchkMultiLang

		vItemNameKr = cOverSeas.FOneItem.Fitemname_kr
		vItemCopyKr = cOverSeas.FOneItem.Fitemcopy_kr
		vItemSourceKr = cOverSeas.FOneItem.Fitemsource_kr
		vItemSizeKr = cOverSeas.FOneItem.Fitemsize_kr
		vSourceAreaKr = cOverSeas.FOneItem.Fsourcearea_kr
		vMakerNameKr = cOverSeas.FOneItem.Fmakername_kr
		vKeywordsKr = cOverSeas.FOneItem.fkeywords_kr

	set cOverSeas = Nothing

	set oitem = new CItemInfo
		oitem.FRectItemId = vItemID
		oitem.GetOneItemInfo
		
		vOriginListImage = oitem.FOneItem.FListImage
		vOriginItemName = oitem.FOneItem.FItemName
		vOriginMakerID = oitem.FOneItem.FMakerid
		makerid = oitem.FOneItem.fmakerid
		vOriginSellCash = FormatNumber(oitem.FOneItem.FSellcash,0)
	set oitem = Nothing
Else
	Response.Write "<script>alert('잘못된 경로입니다.');window.close()</script>"
	session.codePage = 949 : dbget.close() : Response.End
End IF

'/옵션
set oitemoption = new CMultiLang
oitemoption.FRectItemId = vItemID
oitemoption.frectCountryCd = vCountryCd

if vItemID<>"" then
	oitemoption.GetItemOptionMultiLang
end if

if useyn="" or isnull(useyn) then useyn="Y"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/common.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm_utf8.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">

//저장
function goSubmit(){
	if(document.frmreg.itemname.value == ""){
		alert("상품명을 입력하세요.");
		document.frmreg.itemname.focus();
		return;
	}
//	if(document.frmreg.itemsource.value == ""){
//		alert("재료를 입력하세요.");
//		document.frmreg.itemsource.focus();
//		return;
//	}
	if(document.frmreg.sourcearea.value == ""){
		alert("원산지를 입력하세요.");
		document.frmreg.sourcearea.focus();
		return;
	}

	if(confirm("언어(<%=vCountryName%>) 내용을 <%=chkIIF(bChkReg,"수정","신규 등록")%> 하시겠습니까?")){
		document.frmreg.submit();
	}
}

$(function() {
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
  	$(".rdoUsing > input[type='radio']").click(function(){
  		$("input[name='optisusing']").eq($(this).attr("chk")).val($(this).val());
  	});
  	$("input[name='optiontypename']").keyup(function(){
  		$("input[name='optiontypename']").not(this).val($(this).val());
  	});
});

</script>
</head>
<body bgcolor="#F4F4F4">
<form name="frmreg" method="post" action="MultiLangItemContentProc.asp" style="margin:0px;">
<input type="hidden" name="itemid" value="<%=vItemID%>">
<input type="hidden" name="CountryCd" value="<%= vCountryCd %>">
<input type="hidden" name="mode" value="<%=chkIIF(bChkReg,"modi","new")%>">

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="#F8F8F8" colspan="2">
		<strong style="font-size:16px;">다국어 상품정보 입력</strong>
		<div style="padding-top:3px;">- 상품정보를 아래 [언어SET]에 지정된 언어로 입력해주세요.</div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" border="0" class="a">
		<tr>
			<td width="100"><img src="<%=vOriginListImage%>" width="100" height="100"></td>
			<td valign="top">
				<table width="100%" border="0" class="a">
				<tr>
					<td height="23">상품명 : <%=vOriginItemName%></td>
				</tr>
				<tr>
					<td height="23">상품코드 : <%=vItemID%> - [<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=vItemID%>" target="_blank">상품상세보기페이지</a>]</td>
				</tr>
				<tr>
					<td height="23">브랜드ID : <%=vOriginMakerID%></td>
				</tr>
				<tr>
					<td height="23">판매가 : <%=vOriginSellCash%></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width="10%">언어 SET</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= vCountryName %>
		/ <%=chkIIF(bChkReg,"<font color=darkblue>등록됨</font>","<font color=darkred>미등록</font>")%>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>* 상품명</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="itemname" value="<%=vItemName%>" maxlength="60" id="[on,off,off,off][상품명]" style="width:100%;" />
		<div style="color:#BBB;padding-top:2px;"><%=vItemNameKr%></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>상품설명<br />(간략히 서술)</td>
	<td bgcolor="#FFFFFF" align="left"><textarea name="itemcontent" rows="14" id="[on,off,off,off][상품설명]" style="width:100%;" ><%=vItemContent%></textarea></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>상품카피</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="itemcopy" value="<%=vItemCopy%>" maxlength="250" style="width:100%;" />
		<% if Not(vItemCopyKr="" or isNull(vItemCopyKr)) then %><div style="color:#BBB;padding-top:2px;"><%=vItemCopyKr%></div><% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>재료</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="itemsource" value="<%=vItemSource%>" maxlength="128" id="[on,off,off,off][재료]" style="width:100%;" />
		<div style="color:#BBB;padding-top:2px;"><%=vItemSourceKr%></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>크기</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="itemsize" value="<%=vItemSize%>" maxlength="128" id="[on,off,off,off][크기]" style="width:100%;" />
		<div style="color:#BBB;padding-top:2px;"><%=vItemSizeKr%></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>제조사</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="makername" value="<%=vMakerName%>" maxlength="64" id="[on,off,off,off][제조사]" style="width:100%;" />
		<div style="color:#BBB;padding-top:2px;"><%=vMakerNameKr%></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>* 원산지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="sourcearea" value="<%=vSourceArea%>" maxlength="128" id="[on,off,off,off][원산지]" style="width:100%;" />
		<div style="color:#BBB;padding-top:2px;"><%=vSourceAreaKr%></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>키워드</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="keywords" value="<%=keywords%>" maxlength="128" style="width:100%;" />
		<div style="color:#BBB;padding-top:2px;"><%=vKeywordsKr%></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" width=100>옵션</td>
	<td valign="top" bgcolor="#FFFFFF">
		<table width="100%" cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tabletop") %>">
		<%
			If oitemoption.FResultCount > 0 Then
				For i=0 To oitemoption.FResultCount - 1
		%>
		<tr>
			<td bgcolor="#FFFFFF" align="center">
				<table border="0" cellpadding="2" cellspacing="0" class="a">
				<tr>
					<td>
						<input type="hidden" name="itemoption" value="<%= oitemoption.FITemList(i).FItemOption %>" /><%= oitemoption.FITemList(i).FItemOption %>
						<input type="hidden" name="optisusing" value="<%= oitemoption.FITemList(i).FOptIsUsing %>" />
					</td>
				<% if oitemoption.FItemList(i).Fitemoption="0000" then %>
					<td>
						* 옵션없음
						<input type="hidden" name="optiontypename" value="<%= oitemoption.FITemList(i).FOptionTypeName %>">
						<input type="hidden" name="optionname" value="<%= oitemoption.FITemList(i).FOptionName %>">
					</td>
				<% else %>
					<td>
						<input type="text" name="optiontypename" value="<%=chkIIF(left(oitemoption.FITemList(i).FItemOption,1)="Z","Multiple",oitemoption.FITemList(i).FOptionTypeName) %>" size="10" <%=chkIIF(left(oitemoption.FITemList(i).FItemOption,1)="Z","readonly","")%> class="<%=chkIIF(left(oitemoption.FITemList(i).FItemOption,1)="Z","text_ro","text")%>" id="[on,off,off,off][옵션구분명]">
						<div style="color:#BBB;padding-top:2px;"><%=oitemoption.FITemList(i).FOptionTypeName_kr%></div>
					</td>
					<td>
						<input type="text" name="optionname" value="<%= oitemoption.FITemList(i).FOptionName %>" size="30" class="text" id="[on,off,off,off][옵션명]" >
						<div style="color:#BBB;padding-top:2px;"><%=oitemoption.FITemList(i).FOptionName_kr%></div>
					</td>
					<td>
						<span class="rdoUsing">
							<input type="radio" name="optisusing<%=i%>" id="rdoUsing<%=i%>_1" chk="<%=i%>" value="Y" <%= CHKIIF(oitemoption.FITemList(i).FOptIsUsing="Y","checked","") %> /><label for="rdoUsing<%=i%>_1">사용</label>
							<input type="radio" name="optisusing<%=i%>" id="rdoUsing<%=i%>_2" chk="<%=i%>" value="N" <%= CHKIIF(oitemoption.FITemList(i).FOptIsUsing="N","checked","") %> /><label for="rdoUsing<%=i%>_2">사용안함</label>
						</span>
					</td>
				<% end if %>
				</tr>
				</table>
			</td>
		</tr>
		<%
				Next
			else
				Response.Write "<tr><td bgcolor='#FFFFFF'>* 옵션없음</td></tr>"
			End if
		%>
		</table>
	</td>
</tr>
</table>

<input type="hidden" name="useyn" value="<%=useyn%>">
</form>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="window.close();" style="cursor:pointer">
    </td>
    <td align="right">
    	<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="goSubmit();" style="cursor:pointer">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	set oitemoption = Nothing
	session.codePage = 949
%>