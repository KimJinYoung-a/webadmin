<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/common/commonCls.asp"-->
<%
Dim vMallgubun, vMakerID, vItemID, vItemName, vAction, i, tmpItemID, arrItemid, arrTemp, iA, j
Dim page, oNoitemid, oSQL, in_id, del_id, vBigo, vBigoText
vMallgubun	= request("Mallgubun")
vMakerID	= request("makerid")
vItemID		= request("itemid")
vItemName	= request("itemname")
page    	= request("page")
vAction		= request("action")
vBigo			= Request("bigo")
vBigoText		= Request("bigoText")

If page = "" Then page = 1
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

If vAction = "insert" Then
	Dim bigo
	'in_id		= Trim(request("in_id"))
	vItemid = Request("in_id")
	bigo = NullFillWith(Trim(requestCheckVar(request("bigo"),300)),"")
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


	' oSQL = ""
	' oSQL = oSQL & " SELECT COUNT(*) as cnt FROM db_Appwish.dbo.tbl_item where itemid = '"& in_id &"' "
	' rsCTget.Open oSQL,dbCTget,1
	' If rsCTget("cnt") = 0 Then
	' 	Call Alert_move("해당상품이 없기에 등록할 수 없습니다.","targetMall_Not_In_Itemid.asp?mallgubun="&vMallgubun&"")
	' End If
	' rsCTget.Close

	arrItemid = Split(arrItemid, ",")
	for j = 0 to UBound(arrItemid)
		if Trim(arrItemid(j)) <> "" then
			vItemid = Trim(arrItemid(j))
			oSQL = " DECLARE @Temp CHAR(1) " & _
					"	If NOT EXISTS(SELECT * FROM [db_etcmall].dbo.tbl_targetMall_Not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
					"		BEGIN " & _
					"			INSERT INTO [db_etcmall].dbo.tbl_targetMall_Not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
					"		END	"
			dbget.execute oSQL
			''response.write oSQL & "<br />"
			oSQL = " DECLARE @Temp CHAR(1) " & _
					"	If NOT EXISTS(SELECT * FROM db_outmall.dbo.tbl_targetMall_Not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
					"		BEGIN " & _
					"			INSERT INTO db_outmall.dbo.tbl_targetMall_Not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
					"		END	"
			dbCTget.execute oSQL
		end if
	Next
	vItemid = Request("in_id")
	Response.Write "<script>alert('처리되었습니다.');location.href='targetMall_Not_In_Itemid.asp?mallgubun=" & vMallGubun & "&page=" & page & "';</script>"
	Response.End
ElseIf vAction = "delete" Then
	del_id = Replace(Request("del_id")," ","")
	del_id = "'" & Replace(del_id,",","','") & "'"

	oSQL = "DELETE [db_etcmall].dbo.tbl_targetMall_Not_in_itemid  WHERE mallgubun = '" & vMallGubun & "' AND itemid IN(" & del_id & ")"
	dbget.execute oSQL

	oSQL = "DELETE db_outmall.dbo.tbl_targetMall_Not_in_itemid WHERE mallgubun = '" & vMallGubun & "' AND itemid IN(" & del_id & ")"
	dbCTget.execute oSQL
	Response.Write "<script>alert('처리되었습니다.');location.href='targetMall_Not_In_Itemid.asp?mallgubun=" & vMallGubun & "&page=" & page & "';</script>"
	Response.End
End If

Set oNoitemid = new CCommon
	oNoitemid.FPageSize 		= 20
	oNoitemid.FCurrPage			= page
	oNoitemid.FRectMallgubun	= vMallgubun
	oNoitemid.FRectMakerID		= vMakerID
	oNoitemid.FRectItemID		= vItemID
	oNoitemid.FRectItemName		= vItemName
	oNoitemid.FRectBigo			= vBigo
	oNoitemid.FRectBigoText		= vBigoText
	oNoitemid.getTargetMall_Not_In_itemid_List
%>
<script language="javascript">
function insert_id()
{
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
	var chkSel=0;
	try {
		if(frm.del_id.length>1) {
			for(var i=0;i<frm.del_id.length;i++) {
				if(frm.del_id[i].checked) chkSel++;
			}
		} else {
			if(frm.del_id.checked) chkSel++;
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
	frm.action.value = "delete";
	frm.submit();
}

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
</script>
<center>
Mall 구분 : <b><%=vMallGubun%></b>
</center>
<br>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">

		<tr>
			<td width="90%">브랜드ID : <input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="20"></td>
			<td rowspan="3" width="10%"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td width="90%">상품ID : <textarea class="textarea" name="itemid" rows="2" cols="16"><%=replace(vItemid,",",chr(10))%></textarea></td>
		</tr>
		<tr>
			<td>상품명 : <input type="text" class="text" name="itemname" value="<%=vItemName%>" size="30"></td>
		</tr>
		<tr>
			<td>코맨트여부 :
				<Select name="bigo" class="select">
					<option value="">-전체-
					<option value="Y" <%= Chkiif(vBigo="Y", "selected", "") %> >Y
					<option value="N" <%= Chkiif(vBigo="N", "selected", "") %> >N
				</select>
				<input type="text" name="bigoText" value="<%=vBigoText%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
<form name="frm" action="targetMall_Not_In_Itemid.asp" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td>제외 상품ID :</td>
			<td width="90%">
				<textarea class="textarea" name="in_id" rows="2" cols="16"></textarea>
			</td>
		</tr>
		<tr>
			<td>코맨트 :</td>
			<td width="90%">
				<input type="text" class="text" name="bigo" size="40">
				<input type="button" class="button" value="저 장" onClick="insert_id()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
			</td>
			<td width="20%" align="right">검색결과 : <b><%= FormatNumber(oNoitemid.FTotalCount,0) %></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="10%">상품ID</td>
	<td width="60%">상품명</td>
	<td width="30%"><input type="button" value="선택 상품ID 삭제" onClick="delete_id()"></td>
</tr>
<% If oNoitemid.FTotalCount > 0 Then %>
<% For i=0 to oNoitemid.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center"><%= oNoitemid.FItemList(i).FItemid %></td>
	<td>
		<%
			rw oNoitemid.FItemList(i).FItemname
			If oNoitemid.FItemList(i).FBigo <> "" Then
				response.write "<font color='blue'>코멘트 : " & oNoitemid.FItemList(i).FBigo & "</font>"
			End If
		%>
	</td>
	<td align="center"><input type="checkbox" name="del_id" value="<%= oNoitemid.FItemList(i).FItemid %>"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oNoitemid.HasPreScroll then %>
		<a href="javascript:goPage('<%= oNoitemid.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oNoitemid.StartScrollPage to oNoitemid.FScrollCount + oNoitemid.StartScrollPage - 1 %>
    		<% if i>oNoitemid.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oNoitemid.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF" height="30">
	<td colspan="3" align="center" class="page_link">[데이터가 없습니다.]</td>
</tr>
<% End If %>
</table>
</form>
<% Set oNoitemid = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->