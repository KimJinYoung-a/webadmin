<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/etc/gmarket/gmarketcls.asp"-->
<%
Dim i, strSql, arrList, vAction, vCurrPage, vItemID, vItemName, vMakerID, vIsUsing
Dim tmpItemID, arrTemp, arrItemid
vAction 		= Request("action")
vCurrPage		= NullFillWith(Request("cp"),1)
vItemID			= Request("itemid")
vItemName		= Request("itemname")
vMakerID		= Request("makerid")
vIsUsing		= Request("isUsing")

If vCurrPage = "" Then vCurrPage = 1
If vIsUsing = "" Then vIsUsing = "Y"

If vItemID<>"" then
	tmpItemID = vItemid
	tmpItemID = replace(tmpItemID,",",chr(10))
	tmpItemID = replace(tmpItemID,chr(13),"")
	arrTemp = Split(tmpItemID,chr(10))
	i = 0
	Do While i <= ubound(arrTemp)
		If Trim(arrTemp(i))<>"" then
			If Not(isNumeric(trim(arrTemp(i)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(i)) & ","
			End If
		End If
		i = i + 1
	Loop
	vItemID = left(arrItemid,len(arrItemid)-1)
end if

If vAction = "insert" OR vAction = "delete" Then
	Call Proc()
End If

Dim oGmarket
%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function insert_id()
{
	var i;
	if(frm.in_id.value == "")
	{
		alert("ID를 입력하세요.");
		frm.in_id.focus();
		return;
	}
	frm.action.value = "insert";
	frm.submit();
}
function delete_id()
{
	frm.action.value = "delete";
	frm.submit();
}

function jsGoPage(iP){
	document.frmsearch.cp.value = iP;
	document.frmsearch.submit();
}
function jsSubmit() {
	var frm = document.frmsearch;
	frm.submit();
}
window.onload = function() {
	window.resizeTo(600, 770);
}
</script>
<br>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="cp" value="<%=vCurrPage%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td width="15%">브랜드ID :</td>
			<td><input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="20"></td>
			<td rowspan="4" width="10%"><input type="button" value="검 색" style="width:50px;height:50px;" onClick="jsSubmit()"></td>
		</tr>
		<tr>
			<td>상품ID :</td>
			<td><textarea class="textarea" name="itemid" rows="2" cols="16"><%=replace(vItemid,",",chr(10))%></textarea></td>
		</tr>
		<tr>
			<td>상품명 :</td>
			<td><input type="text" class="text" name="itemname" value="<%=vItemName%>" size="30"></td>
		</tr>
		<tr>
			<td>사용여부 :</td>
			<td>
				<label><input type="radio" class="radio" name="isUsing" value="Y" <%= CHKIIF(vIsUsing="Y", "checked", "") %> >Y</label>&nbsp;
				<label><input type="radio" class="radio" name="isUsing" value="N" <%= CHKIIF(vIsUsing="N", "checked", "") %>>N</label>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<p />
<font color="red">G마켓에 등록된 상품만 저장 가능합니다.</font>
<form name="frm" action="g9SpecialItem.asp" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="cp" value="<%=vCurrPage%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td>제외 상품ID :</td>
			<td width="85%">
				<textarea class="textarea" name="in_id" rows="2" cols="16"></textarea>
				<input type="button" class="button" value="저 장" onClick="insert_id()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
	Set oGmarket = new CGmarket
		oGmarket.FCurrPage			= vCurrPage
		oGmarket.FPageSize			= 15
		oGmarket.FRectmakerid		= vMakerID
		oGmarket.FRectItemid		= vItemID
		oGmarket.FRectItemName		= vItemName
		oGmarket.FRectIsUsing		= vIsUsing
		oGmarket.getG9SpecialItemList
%>

<p />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
			</td>
			<td width="20%" align="right">검색결과 : <b><%= FormatNumber(oGmarket.FTotalCount,0) %></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="10%">상품ID</td>
	<td width="50%">상품명</td>
	<td width="30%">
		<input type="button" class="button" value="사용안함처리" onClick="delete_id()">
		<input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frm.del_id);">
	</td>
</tr>
<%
	IF oGmarket.FResultCount > 0 THEN
		For i=0 to oGmarket.FResultCount - 1
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center"><%= oGmarket.FItemList(i).FItemid %></td>
		<td>
		<%
			rw oGmarket.FItemList(i).FItemname
		%>
		</td>
		<td align="center"><input type="checkbox" name="del_id" value="<%= oGmarket.FItemList(i).Fidx %>"></td>
	</tr>
<%
		Next
	Else
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="4" align="center" class="page_link">[데이터가 없습니다.]</td>
	</tr>
<%
	End If
%>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oGmarket.HasPreScroll then %>
		<a href="javascript:jsGoPage('<%= oGmarket.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oGmarket.StartScrollPage to oGmarket.FScrollCount + oGmarket.StartScrollPage - 1 %>
    		<% if i>oGmarket.FTotalpage then Exit for %>
    		<% if CStr(vCurrPage)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:jsGoPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oGmarket.HasNextScroll then %>
    		<a href="javascript:jsGoPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>

<%
Function Proc()
	Dim strSql, vAction, vItemid, vCurrPage, arrList, j, k
	Dim iA, tmpItemID, arrTemp, arrItemid, isRegGmarket
	vAction = Request("action")
	vCurrPage = NullFillWith(Request("cp"),1)
	''response.end
	If vAction = "insert" Then
		vItemid = Request("in_id")
		If vItemid<>"" then
			tmpItemID = vItemid
			tmpItemID = replace(tmpItemID,",",chr(10))
			tmpItemID = replace(tmpItemID,chr(13),"")
			arrTemp = Split(tmpItemID,chr(10))
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
			arrItemid = left(arrItemid,len(arrItemid)-1)
		End If

		arrItemid = Split(arrItemid, ",")
		for j = 0 to UBound(arrItemid)
			isRegGmarket = "Y"
			if Trim(arrItemid(j)) <> "" then
				vItemid = Trim(arrItemid(j))
				strSql = ""
				strSql = strSql & " SELECT COUNT(*) as cnt "
				strSql = strSql & " FROM db_etcmall.dbo.tbl_gmarket_regItem "
				strSql = strSql & " WHERE itemid in ("& vItemid &") "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If rsget("cnt") = 0 Then
					isRegGmarket = "N"
				End If
				rsget.Close

				If isRegGmarket = "N" Then
					Response.Write "<script language=javascript>alert('[" & vItemid & "]은 G마켓에 등록된 상품코드가 아닙니다.');location.href='g9SpecialItem.asp?cp=" & vCurrPage & "';</script>"
					dbget.close()	:	response.End
				Else
					strSql = "	DECLARE @Temp CHAR(1) " & _
								"	If NOT EXISTS(SELECT * FROM db_etcmall.[dbo].[tbl_G9_Special_itemid] Where itemid = '" & vItemid & "' and isUsing = 'N') " & _
								"		BEGIN " & _
								"			INSERT INTO db_etcmall.[dbo].[tbl_G9_Special_itemid] (itemid, regdate, regUserId, isUsing) VALUES ('" & vItemid & "', getdate(), '"& session("ssBctID") &"', 'Y') " & _
								"		END	" & _
								"	ELSE " & _
								"		BEGIN " & _
								"			UPDATE db_etcmall.[dbo].[tbl_G9_Special_itemid] " & _
								"			SET isUsing = 'Y' " & _
								"			WHERE itemid in ('" & vItemid & "') " & _
								"		END "
					dbget.execute strSql
					response.write strSql & "<br />"
				End If
			end if
		Next
		vItemid = Request("in_id")
	ElseIf vAction = "delete" Then
		vItemid = Replace(Request("del_id")," ","")
		vItemid = "'" & Replace(vItemid,",","','") & "'"
		strSql = "UPDATE db_etcmall.[dbo].[tbl_G9_Special_itemid] SET isUsing = 'N' WHERE idx IN (" & vItemid & ")"
		dbget.execute strSql
	End IF

	Response.Write "<script>alert('처리되었습니다.');location.href='g9SpecialItem.asp?cp=" & vCurrPage & "';</script>"
	Response.End
End Function

SET oGmarket = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
