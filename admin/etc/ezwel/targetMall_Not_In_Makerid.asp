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
Dim vMallgubun, sMakerID, vAction, i, vMakerid, in_id, del_id
Dim page, oSQL, oNoMakerid
vMallgubun	= request("Mallgubun")
sMakerID	= request("smakerid")
page    	= request("page")
vAction		= request("action")

If page = "" Then page = 1

If vAction = "insert" Then
	in_id		= Trim(request("makerid"))

	oSQL = ""
	oSQL = oSQL & " SELECT COUNT(*) as cnt FROM db_Appwish.dbo.tbl_user_c where userid = '"& in_id &"' "
	rsCTget.Open oSQL,dbCTget,1
	If rsCTget("cnt") = 0 Then
		Call Alert_move("해당 브랜드가 없기에 등록할 수 없습니다.","targetMall_Not_In_Makerid.asp?mallgubun="&vMallgubun&"") 
	End If
	rsCTget.Close

	oSQL = " DECLARE @Temp CHAR(1) " & _
			 "	If NOT EXISTS(SELECT * FROM db_outmall.dbo.tbl_targetMall_Not_in_makerid Where mallgubun = '" & vMallGubun & "' AND makerid = '" & in_id & "') " & _
			 "		BEGIN " & _
			 "			INSERT INTO db_outmall.dbo.tbl_targetMall_Not_in_makerid(makerid,mallgubun,reguserid) VALUES('" & in_id & "','" & vMallGubun & "', '"&session("ssBctID")&"') " & _
			 "		END	"
	dbCTget.execute oSQL

	oSQL = " DECLARE @Temp CHAR(1) " & _
			 "	If NOT EXISTS(SELECT * FROM db_etcmall.dbo.tbl_targetMall_Not_in_makerid Where mallgubun = '" & vMallGubun & "' AND makerid = '" & in_id & "') " & _
			 "		BEGIN " & _
			 "			INSERT INTO db_etcmall.dbo.tbl_targetMall_Not_in_makerid(makerid,mallgubun,reguserid) VALUES('" & in_id & "','" & vMallGubun & "', '"&session("ssBctID")&"') " & _
			 "		END	"
	dbget.execute oSQL

	Response.Write "<script>alert('처리되었습니다.');location.href='targetMall_Not_In_Makerid.asp?mallgubun=" & vMallGubun & "&page=" & page & "';</script>"
	Response.End
ElseIf vAction = "delete" Then
	del_id = Replace(Request("del_id")," ","")
	del_id = "'" & Replace(del_id,",","','") & "'"
	oSQL = "DELETE db_outmall.dbo.tbl_targetMall_Not_in_makerid WHERE mallgubun = '" & vMallGubun & "' AND makerid IN(" & del_id & ")"
	dbCTget.execute oSQL

	oSQL = "DELETE db_etcmall.dbo.tbl_targetMall_Not_in_makerid WHERE mallgubun = '" & vMallGubun & "' AND makerid IN(" & del_id & ")"
	dbget.execute oSQL
	Response.Write "<script>alert('처리되었습니다.');location.href='targetMall_Not_In_Makerid.asp?mallgubun=" & vMallGubun & "&page=" & page & "';</script>"
	Response.End
End If

Set oNoMakerid = new CCommon
	oNoMakerid.FPageSize 		= 20
	oNoMakerid.FCurrPage		= page
	oNoMakerid.FRectMallgubun	= vMallgubun
	oNoMakerid.FRectMakerID		= sMakerID
	oNoMakerid.getTargetMall_Not_In_makerid_List
%>
<script language="javascript">
function insert_id()
{
	if(frm.makerid.value == "")
	{
		alert("ID를 입력하세요.");
		frm.makerid.focus();
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
			alert("선택한 브랜드가 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("브랜드가 없습니다.");
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
			<td width="90%">브랜드ID : <input type="text" class="text" name="smakerid" value="<%=sMakerID%>" size="20"></td>
			<td rowspan="3" width="10%"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
<form name="frm" action="targetMall_Not_In_Makerid.asp" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
				브랜드ID : <input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
				<input type="button" value="저 장" onClick="insert_id()">
			</td>
			<td width="20%" align="right">상품수 : <b><%=oNoMakerid.FTotalCount%></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="10%">브랜드ID</td>
	<td width="60%">브랜드명</td>
	<td width="30%"><input type="button" value="선택 브랜드ID 삭제" onClick="delete_id()"></td>
</tr>
<% If oNoMakerid.FTotalCount > 0 Then %>
<% For i=0 to oNoMakerid.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center"><%= oNoMakerid.FItemList(i).FMakerid %></td>
	<td> <%=oNoMakerid.FItemList(i).FSocname_kor%></td>
	<td align="center"><input type="checkbox" name="del_id" value="<%= oNoMakerid.FItemList(i).FMakerid %>"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oNoMakerid.HasPreScroll then %>
		<a href="javascript:goPage('<%= oNoMakerid.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oNoMakerid.StartScrollPage to oNoMakerid.FScrollCount + oNoMakerid.StartScrollPage - 1 %>
    		<% if i>oNoMakerid.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oNoMakerid.HasNextScroll then %>
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
<% Set oNoMakerid = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->