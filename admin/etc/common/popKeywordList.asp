<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/common/commonCls.asp"-->
<%
Dim itemid, i, mallgubun, oCommon, arrRows, arrRows2
mallgubun	= request("mallgubun")
itemid		= request("itemid")

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

SET oCommon = new CCommon
	oCommon.FRectMallGubun = mallgubun
	oCommon.FRectItemid = itemid
	arrRows = oCommon.getTenKeyWordsList
SET oCommon = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function keyWordsProcess() {
	var chkSel=0;
	var keywords = "";
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) {
					chkSel++;
					keywords = keywords + frmSvArr.keywords[i].value + "*(^!";
				}
			}
		} else {
			if(frmSvArr.cksel.checked){
				 chkSel++;
				 keywords = frmSvArr.keywords.value;
			}
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

    if (confirm(chkSel + '개의 상품 키워드 변경을 적용하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.mode.value = "REG";
		document.frmSvArr.arrkeywords.value = keywords;
		document.frmSvArr.action = "/admin/etc/common/procKeywords.asp"
		document.frmSvArr.submit();
	}
}
function goPage(pg){
    frm2.page.value = pg;
    frm2.submit();
}
</script>

<!-- #include virtual="/admin/etc/common/inc_tabkeyword.asp"-->

<div style="width:49%;float:left;">
	<span align="center"><h4>텐바이텐 키워드 목록</h4></span>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="mallgubun" value="<%= mallgubun %>" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</table>
	</form>
	<br />
	<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
	<input type="hidden" name="arrkeywords" value="" />
	<input type="hidden" name="mode" value="" />
	<input type="hidden" name="mallgubun" value="<%= mallgubun %>" />
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<th width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
		<th width="70">상품코드</td>
		<th>키워드</td>
	</tr>
<%
If IsArray(arrRows) Then
	For i = 0 To Ubound(arrRows, 2)
%>
	<tr align="center"  bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= arrRows(0,i) %>"></td>
		<td><%= arrRows(0,i) %></td>
		<td align="LEFT">
			10x10 : <%= arrRows(1,i) %>
			<br /><%= mallgubun %> :
			<%
				If arrRows(2,i)="" Then
					response.write "<font color='RED'>등록전</font>"
				End If
			%>
			<input type="text" class="text" name="keywords" size="60" value="<%= CHKIIF(arrRows(2,i)="", arrRows(1,i), arrRows(2,i))  %>" />
		</td>
	</tr>
<%
	Next
%>
	<tr align="center"  bgcolor="#FFFFFF">
		<td colspan="3">
			<input type="button" class="button" value="저장" onclick="keyWordsProcess();" />
		</td>
	</tr>
<%
Else
%>
	<tr align="center" height="50" bgcolor="#FFFFFF">
		<td colspan="3">검색 된 데이터가 없습니다.</td>
	</tr>
<%
End If
%>
	</table>
	</form>
</div>
<%
Dim sItemid
sItemid = request("sItemid")

If sItemid<>"" then
	Dim iA2, arrTemp2, arrItemid2
	sItemid = replace(sItemid,",",chr(10))
	sItemid = replace(sItemid,chr(13),"")
	arrTemp2 = Split(sItemid,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid2 = arrItemid2 & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	sItemid = left(arrItemid2,len(arrItemid2)-1)
End If

Dim page, k
page = request("page")
If page = "" Then page = 1
SET oCommon = new CCommon
	oCommon.FCurrPage		= page
	oCommon.FPageSize		= 20
	oCommon.FRectMallGubun	= mallgubun
	oCommon.FRectSItemid	= sItemid
	oCommon.getOutmallKeyWordsList
%>
<div style="width:49%;float:right;">
	<span align="center"><h4><%= mallgubun %> 키워드 목록</h4></span>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm2" method="get" action="">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			몰 구분 :
			<select name="mallgubun" class="select" onchange="location.replace('/admin/etc/common/popKeywordList.asp?mallgubun='+this.value);">
				<option value="ssg" <%= CHKIIF(mallgubun="ssg", "selected", "") %> >SSG</option>
				<option value="WMP" <%= CHKIIF(mallgubun="WMP", "selected", "") %> >위메프</option>
				<option value="WMP" <%= CHKIIF(mallgubun="coupang", "selected", "") %> >쿠팡</option>
			</select>&nbsp;
			상품코드 : <textarea rows="2" cols="20" name="sItemid" id="sItemid"><%=replace(sItemid,",",chr(10))%></textarea>
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm2.submit();">
		</td>
	</table>
	</form>
	<br />
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<th width="70">상품코드</td>
		<th>키워드</td>
	</tr>
<%
If oCommon.FResultCount > 0 Then
	For k=0 to oCommon.FResultCount - 1
%>
	<tr align="center" bgcolor="#FFFFFF" <%= Chkiif(Trim(itemid) = Trim(oCommon.FItemList(k).FItemID), "class='H'", "") %> >
		<td style="cursor:pointer;" onclick="location.replace('/admin/etc/common/popKeywordList.asp?mallgubun=<%= mallgubun %>&itemid=<%= oCommon.FItemList(k).FItemid %>');"><%= oCommon.FItemList(k).FItemid %></td>
		<td align="left"><%= oCommon.FItemList(k).FKeywords %></td>
	</tr>
<%
	Next
%>
	<tr height="20">
		<td colspan="19" align="center" bgcolor="#FFFFFF">
		<% If oCommon.HasPreScroll Then %>
			<a href="javascript:goPage('<%= oCommon.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
			[pre]
		<%
			End If
			For i = 0 + oCommon.StartScrollPage to oCommon.FScrollCount + oCommon.StartScrollPage - 1
				If i > oCommon.FTotalpage Then Exit For
				If CStr(page)=CStr(i) Then
		%>
				<font color="red">[<%= i %>]</font>
		<%		Else %>
				<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<%		End If
			Next
			If oCommon.HasNextScroll Then
		%>
				<a href="javascript:goPage('<%= i %>');">[next]</a>
		<%	Else %>
				[next]
		<%	End If %>
		</td>
	</tr>
<%
Else
%>
	<tr align="center" height="50" bgcolor="#FFFFFF">
		<td colspan="3">등록 된 데이터가 없습니다.</td>
	</tr>
<%
End If
%>
	</table>
</div>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<% SET oCommon = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->