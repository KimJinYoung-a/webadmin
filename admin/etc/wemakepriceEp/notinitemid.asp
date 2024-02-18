<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wemakepriceEp/epShopCls.asp"-->
<%
Dim mode, NOTitemid, sqlStr, itemid, SavearrCnt
Dim nItem, page, itemidarr, isusingarr
page		= request("page")
mode		= request("mode")
NOTitemid	= Trim(request("NOTitemid"))
If Right(NOTitemid,1) = "," Then NOTitemid = Left(NOTitemid, Len(NOTitemid) - 1)

itemidarr	= request("itemidarr")
isusingarr	= request("isusingarr")
itemid		= request("itemid")
SavearrCnt = Ubound(Split(NOTitemid,",")) + 1
If page = "" Then page = 1
If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT count(i.itemid) as userCNT, count(ni.itemid) as notItemCNT "
	sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_item as i "
	sqlStr = sqlStr & " LEFT JOIN db_outmall.dbo.tbl_EpShop_not_in_itemid as ni on i.itemid = ni.itemid AND ni.mallgubun = 'wemakepriceEP' "
	sqlStr = sqlStr & " WHERE i.itemid in ("&NOTitemid&") "
	rsCTget.Open sqlStr,dbCTget,1
	If rsCTget("userCNT") <> SavearrCnt Then
		response.write "<script language='javascript'>alert('텐바이텐에 등록된 상품코드가 아닙니다.');location.href='/admin/etc/wemakepriceEp/notinitemid.asp?menupos="&menupos&"';</script>"
	ElseIf rsCTget("notItemCNT") > 0 Then
		response.write "<script language='javascript'>alert('이미 등록된 상품코드 입니다.');location.href='/admin/etc/wemakepriceEp/notinitemid.asp?menupos="&menupos&"';</script>"
	Else
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_EpShop_not_in_itemid (itemid, mallgubun, isusing, regdate, regid) "
		sqlStr = sqlStr & " SELECT itemid, 'wemakepriceEP', 'Y' ,getdate(), '"&session("ssBctID")&"' FROM db_AppWish.dbo.tbl_item "
		sqlStr = sqlStr & " WHERE itemid in ("&NOTitemid&") "
		dbCTget.Execute sqlStr
		response.write "<script language='javascript'>alert('저장하였습니다.');location.href='/admin/etc/wemakepriceEp/notinitemid.asp?menupos="&menupos&"';</script>"
	End If
	rsCTget.Close
ElseIf mode = "U" Then
	Dim cnt, tmpIsusing
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)
	
	isusingarr	=  split(isusingarr,",")
	For i = 0 to cnt
		tmpIsusing = isusingarr(i)
		sqlStr = "UPDATE db_outmall.dbo.tbl_EpShop_not_in_itemid SET "
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,updateid = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)
		dbCTget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('수정하였습니다.');location.href='/admin/etc/wemakepriceEp/notinitemid.asp?menupos="&menupos&"';</script>"
End If

SET nItem = new epShop
	nItem.FCurrPage					= page
	nItem.FPageSize					= 20
	nItem.FItemid					= itemid
    nItem.EpshopnotinitemidList
%>
<script language='javascript'>
var ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

//제외 브랜드 수정하기
function jsIsusing() {
	var frm;
	var sValue, sortNo, isusing;
	frm = document.fitem;
	sValue = "";
	isusing = "";
	chkSel	= 0;

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked) chkSel++;
			if (frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}

				// 사용여부
				if (isusing==""){
					isusing = frm.isusing[i].value;
				}else{
					isusing =isusing+","+frm.isusing[i].value;
				}
			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
			isusing =  frm.isusing.value;
		}
	}
	if(chkSel<=0) {
		alert("선택한 상품이 없습니다.");
		return;
	}
	document.frmIsusing.itemidarr.value = sValue;
	document.frmIsusing.isusingarr.value = isusing;
	document.frmIsusing.mode.value = "U";
	document.frmIsusing.submit();
}

function insert_makerid()
{
	if(document.frm.NOTitemid.value == "")
	{
		alert("브랜드ID를 입력하세요.");
		document.frm.NOTitemid.focus();
		return;
	}
	document.frm.mode.value = "I";
	document.frm.submit();
}

function goPage(page){
    var frm = document.frmSearch;
    frm.page.value=page;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/etc/wemakepriceEP/inc_wemakepriceHead.asp" -->
<!-- 검색 시작 -->
<form name="frmsearch" method="get" action="notinitemid.asp" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall 구분 : 위메프EP</td>
		    <td rowspan="4" width="10%"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			상품코드 : <input type="text" class="text" name="itemid" value="<%=itemid%>" size="20">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<form name="frmIsusing" method="post" action="notinitemid.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<form name="frm" action="notinitemid.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<tr>
	<td>
		제외 상품코드 <input type="text" class="text" name="NOTitemid">&nbsp;&nbsp;<input type="button" class="button" value="저장" onClick="insert_makerid()">
	</td>
	<td align="right">
		<% If nItem.fresultcount >0 then %>	
			<input class="button" type="button" id="btnEditSel" value="제외상품Y/N 수정" onClick="jsIsusing();">
	    <% End If %>
	</td>
</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
    <td>몰구분</td>
	<td>상품코드</td>
	<td>등록일</td>
	<td>등록자</td>
	<td>최종수정일</td>
	<td>최종수정자</td>
	<td>제외상품Y/N</td>
</tr>
<% If nItem.FResultCount > 0 Then %>
<% For i = 0 To nItem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nItem.FItemlist(i).FIdx %>"></td>
	<td><%=nItem.FItemList(i).FMallgubun%></td>
	<td><%=nItem.FItemList(i).FItemid%></td>
	<td><%=nItem.FItemList(i).FRegdate%></td>
	<td><%=nItem.FItemList(i).FRegid%></td>
	<td><%=nItem.FItemList(i).FLastupdate%></td>
	<td><%=nItem.FItemList(i).FUpdateid%></td>
	<td>
		<select name="isusing" class="select">
			<option value="Y" <%=Chkiif(nItem.FItemList(i).FIsusing = "Y","selected","")%> >Y</option>
			<option value="N" <%=Chkiif(nItem.FItemList(i).FIsusing = "N","selected","")%> >N</option>
		</select>
	</td>
</tr>
<% Next %>
<tr height="30">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
	<% If nItem.HasPreScroll Then %>
		<a href="javascript:goPage('<%= nItem.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + nItem.StartScrollPage To nItem.FScrollCount + nItem.StartScrollPage - 1 %>
		<% If i>nItem.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If nItem.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
		등록된 상품코드가 없습니다
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->