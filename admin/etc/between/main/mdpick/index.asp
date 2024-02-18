<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메인페이지
' History : 2014.04.01 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/mainCls.asp"-->
<!-- #include virtual="/admin/etc/between/main/inc_mainhead.asp"-->
<%
Dim page, i
Dim omdpick, gender

page	= request("page")
gender	= request("gender")

If page = "" Then page=1
SET omdpick = new cMain
	omdpick.FPageSize		= 20
	omdpick.FCurrPage		= page
	omdpick.FRectGender		= gender
	omdpick.GetMdpickList()
%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type='text/javascript'>
<!--
function RefreshCaFavKeyWordRec(){
	var frm;
	frm = document.frmSvArr;

	var chkSel=0;
	var sValue;
	sValue = "";
	try {
		if(frm.cksel.length>1) {
			for(var i=0;i<frm.cksel.length;i++) {
				if(frm.cksel[i].checked) chkSel++;
			}
		} else {
			if(frm.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	if (frm.cksel.length > 1){
		for (var i=0;i<frm.cksel.length;i++){
			if (frm.cksel[i].checked){
				sValue = sValue + frm.cksel[i].value + ",";
			}
		}
	}else{
		if (frm.cksel.checked){
			sValue = sValue + frm.cksel.value + ",";
		}
	}
	var AssignReal;
	AssignReal = window.open("<%=mobileUrl%>/chtml/between/make_mdpick_xml.asp?idxarr=" +sValue+"&gender=<%=gender%>", "AssignReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignReal.focus();
}

// 상품 순서 일괄 저장
function jsSortSize() {
	var frm;
	var sValue, sSort
	frm = document.frmSvArr;
	sValue = "";
	sSort = "";

	var chkSel=0;
	try {
		if(frm.cksel.length>1) {
			for(var i=0;i<frm.cksel.length;i++) {
				if(frm.cksel[i].checked) chkSel++;
			}
		} else {
			if(frm.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	if (frm.cksel.length > 1){
		for (var i=0;i<frm.cksel.length;i++){
			if(!IsDigit(frm.sSort[i].value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sSort[i].focus();
				return;
			}

			if (sValue==""){
				sValue = frm.cksel[i].value;		
			}else{
				sValue =sValue+","+frm.cksel[i].value;		
			}	
			
			// 정렬순서
			if (sSort==""){
				sSort = frm.sSort[i].value;		
			}else{
				sSort =sSort+","+frm.sSort[i].value;		
			}
		}
	}else{
		sValue = frm.cksel.value;
		if(!IsDigit(frm.sSort.value)){
			alert("순서지정은 숫자만 가능합니다.");
			frm.sSort.focus();
			return;
		}
		sSort =  frm.sSort.value; 
	}
	document.frmSortSize.idxarr.value = sValue;
	document.frmSortSize.sortarr.value = sSort;
	document.frmSortSize.submit();
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<div style="padding-bottom:10px;">
		* 성별 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<select name="gender" class="select">
			<option value="">-Choice-</option>
			<option value="M" <%= Chkiif(gender="M", "selected", "") %> >남자</option>
			<option value="F" <%= Chkiif(gender="F", "selected", "") %> >여자</option>
		</select>
		</div>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:submit();">
	</td>
</tr>
</form>	
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<% If gender <> "" Then %>
	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>XML Real 적용</a></td>
	<% Else %>
	<td>&nbsp;</td>
	<% End If %>
	<td align="right"><input type="button" value="순서 저장" onClick="jsSortSize();" class="button"></td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		총 등록수 : <b><%=omdpick.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=omdpick.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td>상품코드</td>
	<td>브랜드<br>상품명</td>
	<td>판매가</td>
	<td>마지막 <br/>real 적용시간</td>
	<td>성별</td>
	<td>이미지</td>
	<td>품절여부</td>
	<td>정렬번호</td>
</tr>
<% 
	For i = 0 to omdpick.FResultCount - 1 
%>
<tr height="30" align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= omdpick.FItemList(i).FIdx %>"></td>
    <td><%= omdpick.FItemList(i).FItemid %></td>
    <td align="left"><%= omdpick.FItemList(i).FMakerid %> <%= omdpick.FItemList(i).getDeliverytypeName %><br><%= omdpick.FItemList(i).FItemName %></td>
	<td align="right">
		<% If omdpick.FItemList(i).FSaleYn = "Y" Then %>
		<strike><%= FormatNumber(omdpick.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(omdpick.FItemList(i).FSellcash,0) %></font>
		<% Else %>
		<%= FormatNumber(omdpick.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td>
		<%
			If omdpick.FItemList(i).FMainMdpickXMLRegdate <> "" then
				Response.Write replace(left(omdpick.FItemList(i).FMainMdpickXMLRegdate,10),"-",".") & " <br/> " & Num2Str(hour(omdpick.FItemList(i).FMainMdpickXMLRegdate),2,"0","R") & ":" &Num2Str(minute(omdpick.FItemList(i).FMainMdpickXMLRegdate),2,"0","R")
			End If 
		%>
	</td>
	<td>
	<%
		If omdpick.FItemList(i).FGender = "M" Then
			response.write "<font Color='BLUE'>남자</font>"
		Else
			response.write "<font Color='PINK'>여자</font>"
		End If
	%>
	</td>
    <td><img src="<%= omdpick.FItemList(i).Fsmallimage %>" width="50"> </td>
	<td align="center">
	    <% If omdpick.FItemList(i).IsSoldOut Then %>
	        <% If omdpick.FItemList(i).FSellyn = "N" Then %>
	    	    <font color="red">품절</font>
	        <% Else %>
	    	    <font color="red">일시<br>품절</font>
	        <% End If %>
	    <% End If %>
	</td>
	<td><input type="text" name="sSort" value="<%=omdpick.FItemList(i).FMainMdpickSortNo%>" size="4" style="text-align:right;"></td>
</tr>
<% Next %>
</form>
</table>
<!-- paging -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="20" bgcolor="#FFFFFF">
	<td colspan="18" align="center" bgcolor="#FFFFFF">
	    <% if omdpick.HasPreScroll then %>
		<a href="javascript:goPage('<%= omdpick.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + omdpick.StartScrollPage to omdpick.FScrollCount + omdpick.StartScrollPage - 1 %>
			<% if i>omdpick.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if omdpick.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<%
Set omdpick = Nothing
%>
<form name="refreshFrm" method="post">
<input type="hidden" name="gender" value="<%= gender %>">
<input type="hidden" name="idxarr" value="">
</form>
<!-- 순서 변경--->
<form name="frmSortSize" method="post" action="mdpickitem_process.asp">
<input type="hidden" name="mode" value="S">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="gender" value="<%=gender%>">
<input type="hidden" name="sortarr" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->