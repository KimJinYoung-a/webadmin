<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<%
Dim vAction
vAction = Request("action")
If vAction = "proc" Then
	Call Proc()
End If

If vAction = "chgCate" Then
	Call chgCateProc()
End If

dim sRect, mode, iPageSize, iNotIn, iPage, sSub, sOrderBy, iTotCnt, vTopCate, vOption, vCode, vCateName, vCateUseYN
sRect = requestCheckVar(request("sRect"),32)
mode = requestCheckVar(request("mode"),32)
vCode = requestCheckVar(request("code"),32)
vCateUseYN = "Y"

Dim vSearchCode, vSearchCodeName, vSearchUseYN
vSearchCode 	= requestCheckVar(request("search_code"),32)
vSearchCodeName = requestCheckVar(request("search_codename"),100)
vSearchUseYN	= NullFillWith(request("search_useyn"),"")

iPage = NullFillWith(request("ipage"),"1")
iPageSize = "30"
iNotIn = (iPage - 1) * iPageSize

dim sqlStr
dim iRowsData

'sSub = " dispyn='Y' "
sSub = " 1=1 "
If vSearchCode <> "" Then
    sSub = sSub & " AND DispCateCode = '" & vSearchCode & "' "
End If

If vSearchCodeName <> "" Then
    sSub = sSub & " AND dispcatename LIKE '%" & vSearchCodeName & "%' "
End If

If vSearchUseYN <> "" Then
    sSub = sSub & " AND dispyn = '" & vSearchUseYN & "' "
End If

sOrderBy = " order by dispcatecode "

sqlStr = sqlStr + "	Select Top " & iPageSize & " * From [db_temp].dbo.tbl_interpark_Tmp_DispCategory"
sqlStr = sqlStr + " 	Where " & sSub & ""
sqlStr = sqlStr + "			AND DispCateCode Not In(Select Top " & iNotIn & " DispCateCode From [db_temp].dbo.tbl_interpark_Tmp_DispCategory Where " & sSub & sOrderBy & ") "
sqlStr = sqlStr + "		" & sOrderBy & ""

rsget.Open sqlStr,dbget,1
if Not Rsget.Eof then
    iRowsData = rsget.GetRows
end if
rsget.close
sqlStr = "Select COUNT(DispCateCode) From [db_temp].dbo.tbl_interpark_Tmp_DispCategory Where " & sSub & ""
rsget.Open sqlStr,dbget,1
iTotCnt = rsget(0)
rsget.close

If vCode <> "" Then
	rsget.Open "SELECT DispCatename, dispyn From [db_temp].dbo.tbl_interpark_Tmp_DispCategory Where DispCateCode = '" & vCode & "'",dbget,1
	If Not rsget.Eof Then
		vCateName 	= rsget("DispCatename")
		vCateUseYN	= rsget("dispyn")
	End IF
	rsget.close
End IF

dim i,RowCnt

IF IsArray(iRowsData) then
    RowCnt = UBound(iRowsData,2)
else
    RowCnt = -1
End if
%>
<script language="javascript">
function edit_cate(code)
{
	location.href = "Pop_InterPark_Category.asp?ipage=<%=iPage%>&code="+code+"&search_code=<%=vSearchCode%>&search_codename=<%=vSearchCodeName%>&search_useyn=<%=vSearchUseYN%>";
}

function jsGoPage(iP){
	document.frmpage.ipage.value = iP;
	document.frmpage.submit();
}

function goSubmit()
{
	if(frm.catecode.value == "")
	{
		alert("카테고리 코드를 입력하세요.");
		frm.catecode.focus();
		return;
	}
	if(frm.catename.value == "")
	{
		alert("카테고리명을 입력하세요.");
		frm.catename.focus();
		return;
	}
	frm.submit();
}

function cateItemBatchChange(){
    var frm = document.frm1;
    
    if(frm.orgcate.value == "")	{
		alert("변경이전 카테고리 코드를 입력하세요.");
		frm.orgcate.focus();
		return;
	}
	if(frm.nextcate.value == "")	{
		alert("변경이후 카테고리 코드를 입력하세요.");
		frm.nextcate.focus();
		return;
	}
	
	if (confirm('카테고리 코드를 일괄 변경 하시겠습니까?')){
	    frm.submit();
	}
	
}

function reFresh()
{
	location.href = "Pop_InterPark_Category.asp?ipage=<%=iPage%>&search_code=<%=vSearchCode%>&search_codename=<%=vSearchCodeName%>&search_useyn=<%=vSearchUseYN%>";
}

function search()
{
	var c = frm.search_code.value;
	var n = frm.search_codename.value;
	var u = frm.search_useyn.value;
	location.href = "Pop_InterPark_Category.asp?search_code="+c+"&search_codename="+n+"&search_useyn="+u+"";
}

function iParkCateReceive(){
    var popwin = window.open('<%=apiURL%>/outmall/interpark/actInterparkReq.asp?cmdparam=cateRcv','iParkAPI_Process','width=100,height=100,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<form name="frm" action="Pop_InterPark_Category.asp" method="post">
<input type="hidden" name="action" value="proc">
<input type="hidden" name="realcatecode" value="<%=vCode%>">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		카테고리코드 : <input type="text" name="search_code" value="<%=vSearchCode%>">
		카테고리명 : <input type="text" name="search_codename" value="<%=vSearchCodeName%>">
		사용여부 : 	<select name="search_useyn">
						<option value="">전체</option>
						<option value="Y" <% If vSearchUseYN = "Y" Then %>selected<% End If %>>Y</option>
						<option value="N" <% If vSearchUseYN = "N" Then %>selected<% End If %>>N</option>
					</select>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="검 색" onClick="search()">
		
		
	</td>
	<td align="right"><!-- 관리자 메뉴 -->
		<input type="button" class="button" value="카테고리 땡겨오기" onClick="iParkCateReceive()"></td>
</tr>
</table>
<br>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE" class="a">
<tr bgcolor="#FFFFFF">
	<td width="120"><input type="text" name="catecode" value="<%=vCode%>" size="15" <% If vCode <> "" Then %>readonly<% End If %>></td>
	<td><input type="text" name="catename" value="<%=vCateName%>" size="67"></td>
	<td width="50">
		<input type="radio" name="cateuseyn" value="Y" <% If vCateUseYN = "Y" Then %>checked<% End If %>> Y<br>
		<input type="radio" name="cateuseyn" value="N" <% If vCateUseYN = "N" Then %>checked<% End If %>> N
	</td>
	<td width="167"><input type="button" class="button" value="저 장" onClick="goSubmit()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" class="button" value="취 소" onClick="reFresh()"></td>
</tr>
</table>
</form>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE" class="a">
<form name="frm1" action="Pop_InterPark_Category.asp" method="post">
<input type="hidden" name="action" value="chgCate">
<tr>
    <td colspan="2">카테고리 일괄변경</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td>
        <input type="text" name="orgcate" value="" size="15">
        ==&gt;<input type="text" name="nextcate" value="" size="15">
    </td>
    <td width="167"><input type="button" class="button" value="일괄변경" onClick="cateItemBatchChange();"></td>
</tr>
</table>
</form>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE" class="a">
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="120">카테코드</td>
	<td>카테고리명</td>
	<td width="50">사용여부</td>
	<td width="120">마지막등록일</td>
	<td width="25">&nbsp;</td>
</tr>
<%
For i=0 To RowCnt

	If Trim(vTopCate) <> Trim(Split(iRowsData(1,i),">")(0)) Then
		vTopCate = Trim(Split(iRowsData(1,i),">")(0))
		vOption = vOption & "<option value='" & Split(iRowsData(1,i),">")(0) & "'>" & Split(iRowsData(1,i),">")(0) & "</option>"
	End IF
	
	Response.Write "<tr bgcolor='#FFFFFF' height='20'>" & vbCrLf
	Response.Write "	<td align='center'>" & iRowsData(0,i) & "</td><td>" & iRowsData(1,i) & "</td><td align='center'>" & iRowsData(2,i) & "</td><td align='center'>" & iRowsData(3,i) & "</td>" & vbCrLf
	Response.Write "	<td><input type='button' value='수정' onClick=""edit_cate('"&iRowsData(0,i)&"')"" class='button'></td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf

Next
%>
<tr bgcolor="#FFFFFF">
	<td colspan="5" align="center">
		<!-- 페이징처리 -->
		<%
		Dim iStartPage, iEndPage, iTotalPage, iPerCnt, ix
		iPerCnt = 10
		iStartPage = (Int((iPage-1)/iPerCnt)*iPerCnt) + 1
		iTotalPage = int((iTotCnt-1)/iPageSize) +1
		
		If (iPage mod iPerCnt) = 0 Then
			iEndPage = iPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<form name="frmpage" method="post" action="Pop_InterPark_Category.asp">
		<input type="hidden" name="ipage" value="<%=iPage%>">
		<input type="hidden" name="search_code" value="<%=vSearchCode%>">
		<input type="hidden" name="search_codename" value="<%=vSearchCodeName%>">
		<input type="hidden" name="search_useyn" value="<%=vSearchUseYN%>">
		    <tr valign="bottom" height="25">        
		        <td valign="bottom" align="center">
		         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
				<% else %>[pre]<% end if %>
		        <%
					for ix = iStartPage  to iEndPage
						if (ix > iTotalPage) then Exit for
						if Cint(ix) = Cint(iPage) then
				%>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
				<%		else %>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
				<%
						end if
					next
				%>
		    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
				<% else %>[next]<% end if %>
		        </td>        
		    </tr>    
		    </form>
		</table>
	</td>
</tr>
</table>

<%
Function Proc()
	Dim vCode, vRealCode, vCodeName, vCodeUseYN
	vCode 		= requestCheckVar(Request("catecode"),32)
	vRealCode	= NullFillWith(Request("realcatecode"),"")
	vCodeName	= NullFillWith(Request("catename"),"")
	vCodeUseYN	= NullFillWith(Request("cateuseyn"),"Y")
	
	If vRealCode = "" Then
		rsget.Open "SELECT COUNT(*) From [db_temp].dbo.tbl_interpark_Tmp_DispCategory WHERE DispCateCode = '" & vCode & "'",dbget,1
		If rsget(0) > 0 Then
			rsget.close()
			dbget.close()
			Response.Write "<script>alert('동일한 카테고리코드가 존재합니다.');location.href='Pop_InterPark_Category.asp';</script>"
			Response.End
		End If
		rsget.close()
		dbget.execute "INSERT INTO [db_temp].dbo.tbl_interpark_Tmp_DispCategory(DispCateCode, DispCatename, dispyn) VALUES('" & vCode & "', '" & vCodeName & "', '" & vCodeUseYN & "')"
	Else
		dbget.execute "UPDATE [db_temp].dbo.tbl_interpark_Tmp_DispCategory SET DispCatename = '" & vCodeName & "', dispyn = '" & vCodeUseYN & "' WHERE DispCateCode = '" & vCode & "'"
	End If
	
	dbget.close()
	Response.Write "<script>alert('저장하였습니다.');location.href='Pop_InterPark_Category.asp';</script>"
	Response.End
End Function

Function chgCateProc()
    dim orgcate : orgcate = requestCheckVar(Request("orgcate"),32)
    dim nextcate : nextcate = requestCheckVar(Request("nextcate"),32)
    dim AssignedRow : AssignedRow = 0
    Dim sqlStr
    
    rsget.Open "SELECT COUNT(*) From [db_temp].dbo.tbl_interpark_Tmp_DispCategory WHERE DispCateCode = '" & orgcate & "'",dbget,1
	If rsget(0) <1 Then
		rsget.close()
		dbget.close()
		Response.Write "<script>alert('"&orgcate&" 카테고리 코드가 존재 하지 않습니다.');location.href='Pop_InterPark_Category.asp';</script>"
		Response.End
	End If
	rsget.close()
	
	rsget.Open "SELECT COUNT(*) From [db_temp].dbo.tbl_interpark_Tmp_DispCategory WHERE DispCateCode = '" & nextcate & "'",dbget,1
	If rsget(0) <1 Then
		rsget.close()
		dbget.close()
		Response.Write "<script>alert('"&nextcate&" 카테고리 코드가 존재 하지 않습니다.');location.href='Pop_InterPark_Category.asp';</script>"
		Response.End
	End If
	rsget.close()
	
	
	sqlStr = " update R"
    sqlStr = sqlStr & " set interparklastupdate='2007-01-01'" '' ---변경되게.. 날짜강제변경
    sqlStr = sqlStr & " ,Pinterparkdispcategory=(CASE WHEN Pinterparkdispcategory is Not NULL THEN '"&nextcate&"' ELSE NULL END)"
    sqlStr = sqlStr & " From db_item.dbo.tbl_interpark_reg_item r"
    sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i"
    sqlStr = sqlStr & " 	on i.itemid=r.itemid"
    sqlStr = sqlStr & " 	Join  db_item.dbo.tbl_interpark_dspcategory_mapping M"
    sqlStr = sqlStr & " 	on i.cate_large=M.tencdl"
    sqlStr = sqlStr & " 	and i.cate_mid=M.tencdm"
    sqlStr = sqlStr & " 	and i.cate_small=M.tencdn"
    sqlStr = sqlStr & " where M.interparkdispcategory='"&orgcate&"'"
    
    dbget.Execute sqlStr,AssignedRow
    
    sqlStr = " update db_item.dbo.tbl_interpark_dspcategory_mapping"
    sqlStr = sqlStr & " set interparkdispcategory='"&nextcate&"'"
    sqlStr = sqlStr & " where interparkdispcategory='"&orgcate&"'"
    dbget.Execute sqlStr
	
	dbget.close()
	Response.Write "<script>alert('"&AssignedRow&" 건 반영되었습니다.');location.href='Pop_InterPark_Category.asp';</script>"
	Response.End
end function
%>