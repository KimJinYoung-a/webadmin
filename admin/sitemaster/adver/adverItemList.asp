<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/sitemaster/adver/adverCls.asp"-->
<%
Dim page, i, research
Dim itemid, oAdver, alarmyn
Dim mode, idx, strSql
page    			= request("page")
research			= request("research")
itemid  			= request("itemid")
alarmyn				= request("alarmyn")
mode				= request("mode")
idx					= request("idx")

If mode = "D" Then
	strSql = ""
	strSql = strSql & " DELETE FROM db_sitemaster.[dbo].[tbl_adver_item] WHERE idx = '"& idx &"' "
	dbget.execute strSql
	response.redirect("/admin/sitemaster/adver/adveritemList.asp?menupos="&menupos&"")
End If

If page = "" Then page = 1

'텐바이텐 상품코드 엔터키로 검색되게
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

Set oAdver = new CAdver
	oAdver.FCurrPage			= page
	oAdver.FPageSize			= 50
	oAdver.FRectItemid			= itemid
	oAdver.FRectalarmyn			= alarmyn
	oAdver.getAdverItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function fnAdverManager(){
    var popwin = window.open("/admin/sitemaster/adver/popAdverManager.asp","popAdverManager","width=600,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}
function fnAdverAddItem(){
	var popwin;
	popwin = window.open("/admin/sitemaster/adver/popRegFile.asp", "popup_item", "width=500,height=230,scrollbars=yes,resizable=yes");
	popwin.focus();	
}
function fnDelitem(v){
	if(confirm("삭제 하시겠습니까?")) {
		$("#mode").val('D');
		$("#idx").val(v);	
		document.afrm.submit();
	}
}
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
</script>
<form name="afrm" action="adverItemList.asp">
<input type="hidden" name="menupos" value="<%= menupos %>"> 
<input type="hidden" name="mode" id="mode"> 
<input type="hidden" name="idx"  id="idx">
</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>&nbsp;
		알람여부 : 
		<select name="alarmyn" class="select">
			<option value="">-Choice-</option>
			<option value="Y" <%= Chkiif(alarmyn = "Y", "selected", "") %> >Y</option>
			<option value="N" <%= Chkiif(alarmyn = "N", "selected", "") %>>N</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p />
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="상품등록" onclick="fnAdverAddItem();">
			</td>
		</tr>
		</table>
	</td>
</tr>	
</table>
<p />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oAdver.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oAdver.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="10%">상품코드</td>
	<td>상품명</td>
	<td width="10%">시작일</td>
	<td width="10%">종료일</td>
	<td width="10%">알람여부</td>
	<td width="10%">알람전송일</td>
	<td width="5%">관리</td>
</tr>
<% For i = 0 To oAdver.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= oAdver.FItemList(i).FItemid %></td>
	<td align="LEFT"><%= oAdver.FItemList(i).FItemname %></td>
	<td><%= oAdver.FItemList(i).FStartdate %></td>
	<td><%= oAdver.FItemList(i).FEnddate %></td>
	<td>
		<%
			Select Case oAdver.FItemList(i).FAlarmyn
				Case "Y"		response.write "<font color='RED'>Y</font>"
				Case "N"		response.write "<font color='BLUE'>N</font>"
			End Select
		%>
	</td>
	<td><%= oAdver.FItemList(i).FAlarmdate %></td>
	<td>
		<input type="button" onclick="fnDelitem('<%= oAdver.FItemList(i).FIdx %>')" value="삭제" class="button">
	</td>
	</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oAdver.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAdver.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oAdver.StartScrollPage to oAdver.FScrollCount + oAdver.StartScrollPage - 1 %>
    		<% if i>oAdver.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oAdver.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% Set oAdver = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->