<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim mode, chgNmitemid, depthname, sqlStr, itemid, updatedepthnamearr
Dim nItem, page, itemidarr, isusingarr
page				= request("page")
mode				= request("mode")
chgNmitemid			= Trim(request("chgNmitemid"))
depthname			= Trim(request("depthname"))
itemidarr			= request("itemidarr")
isusingarr			= request("isusingarr")
itemid				= Trim(request("itemid"))
updatedepthnamearr	= Trim(request("updatedepthnamearr"))
If page = "" Then page = 1
If mode = "I" Then
	If isnumeric(chgNmitemid) = "False" Then
		response.write "<script language='javascript'>alert('텐바이텐에 등록된 상품이 아닙니다.');location.href='/admin/etc/naverEp/3depthitemid.asp?menupos="&menupos&"';</script>"
	Else
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as userCNT, count(m.itemid) as notitemCNT "
		sqlStr = sqlStr & " FROM db_AppWish.[dbo].[tbl_item] as i "
		sqlStr = sqlStr & " LEFT JOIN db_outmall.dbo.tbl_EpShop_itemid_3depthName as m on m.itemid = i.itemid AND m.mallgubun = 'naverEP' "
		sqlStr = sqlStr & " WHERE i.itemid = '"&chgNmitemid&"' "
		rsCTget.Open sqlStr,dbCTget,1
		If rsCTget("userCNT") = 0 Then
			response.write "<script language='javascript'>alert('텐바이텐에 등록된 상품이 아닙니다.');location.href='/admin/etc/naverEp/3depthitemid.asp?menupos="&menupos&"';</script>"
		ElseIf rsCTget("notitemCNT") > 0 Then
			response.write "<script language='javascript'>alert('이미 등록된 상품 입니다.');location.href='/admin/etc/naverEp/3depthitemid.asp?menupos="&menupos&"';</script>"
		Else
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_EpShop_itemid_3depthName (itemid, mallgubun, depthname, isusing, regdate, regid) VALUES "
			sqlStr = sqlStr & " ('"&chgNmitemid&"', 'naverep', '"&depthname&"', 'Y' ,getdate(), '"&session("ssBctID")&"') "
			sqlStr = sqlStr & vbCRLF & "; update I SET lastupdate=getdate() from [db_AppWish].dbo.tbl_item I where I.itemid="&chgNmitemid&""
			dbCTget.Execute sqlStr

			response.write "<script language='javascript'>alert('저장하였습니다.');location.href='/admin/etc/naverEp/3depthitemid.asp?menupos="&menupos&"';</script>"
		End If
		rsCTget.Close
	End If
ElseIf mode = "U" Then
	Dim cnt, tmpIsusing
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)

	updatedepthnamearr = split(updatedepthnamearr,",")
	isusingarr	=  split(isusingarr,",")
	For i = 0 to cnt
		tmpIsusing = isusingarr(i)
		sqlStr = "UPDATE db_outmall.dbo.tbl_EpShop_itemid_3depthName SET "
		sqlStr = sqlStr & " depthname = '"&html2db(updatedepthnamearr(i))&"'"
		sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,updateid = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)

		sqlStr = sqlStr & ";"&VbCRLF

		sqlStr = sqlStr & " update I SET lastupdate=getdate() "
		sqlStr = sqlStr & " from [db_AppWish].dbo.tbl_item I "
		sqlStr = sqlStr & " INNER JOIN db_outmall.dbo.tbl_EpShop_itemid_3depthName T "
		sqlStr = sqlStr & " ON T.idx="&itemidarr(i)
		sqlStr = sqlStr & " and T.itemid=I.itemid"
		dbCTget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('수정하였습니다.');location.href='/admin/etc/naverEp/3depthitemid.asp?menupos="&menupos&"';</script>"
End If

SET nItem = new epShop
	nItem.FCurrPage					= page
	nItem.FPageSize					= 20
	nItem.FRectItemid				= itemid
    nItem.EpshopChgItemid3depthList
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

function jsSearchEpItemID(v) {
	var pEPI = window.open("/admin/itemmaster/itemlist.asp?itemid="+v,"pop_EPItemID","width=1400,height=600,scrollbars=yes,resizable=yes");
	pEPI.focus();
}
function popItemManage() {
	var pItemManage = window.open("/admin/etc/naverEP/pop_ItemManage.asp", "pop_ItemManage", "width=1800,height=700,scrollbars=yes,resizable=yes");
	pItemManage.focus();
}

//제외 상품 수정하기
function jsIsusing() {
	var frm;
	var sValue, sortNo, isusing;
	var chg3depthnm;
	frm = document.fitem;
	sValue = "";
	isusing = "";
	updatedepthname = "";
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

				// 3depthNm
				if (updatedepthname==""){
					updatedepthname = frm.updatedepthname[i].value;
				}else{
					updatedepthname =updatedepthname+","+frm.updatedepthname[i].value;
				}

			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
			isusing =  frm.isusing.value;
			updatedepthname=  frm.updatedepthname.value;
		}
	}
	if(chkSel<=0) {
		alert("선택한 상품이 없습니다.");
		return;
	}
	document.frmIsusing.itemidarr.value = sValue;
	document.frmIsusing.isusingarr.value = isusing;
	document.frmIsusing.updatedepthnamearr.value = updatedepthname;
	document.frmIsusing.mode.value = "U";
	document.frmIsusing.submit();
}

function chk3depthform()
{
	if(document.frm.chgNmitemid.value == "")
	{
		alert("상품코드를 입력하세요.");
		document.frm.chgNmitemid.focus();
		return;
	}
	if(document.frm.depthname.value == "")
	{
		alert("상품 3Depth명을 입력하세요.");
		document.frm.depthname.focus();
		return;
	}
	document.frm.mode.value = "I";
	document.frm.submit();
}

function goPage(page){
    var frm = document.frmsearch;
    frm.page.value=page;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
<!-- 검색 시작 -->
<form name="frmsearch" method="get" action="3depthitemid.asp" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall 구분 : 네이버EP</td>
		    <td rowspan="4" width="10%"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			상품코드 : <input type="text" class="text" name="itemid" value="<%=itemidid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchEpItemID(document.frmsearch.itemid.value);" >
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<form name="frmIsusing" method="post" action="3depthitemid.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="updatedepthnamearr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<form name="frm" action="3depthitemid.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<tr>
	<td>
		상품코드 <input type="text" class="text" name="chgNmitemid">&nbsp;&nbsp;
		상품코드 3Depth명 <input type="text" class="text" name="depthname">&nbsp;&nbsp;<input type="button" class="button" value="저장" onClick="chk3depthform();">
	</td>
	<td align="right">
		<input class="button" type="button" id="btnPopAdd" value="미등록 상품" onClick="popItemManage();">
		<% If nItem.fresultcount >0 then %>
			<input class="button" type="button" id="btnEditSel" value="3Depth명/사용YN 수정" onClick="jsIsusing();">
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
	<td>브랜드ID</td>
	<td>상품코드</td>
	<td>특정 상품 3Depth명</td>
	<td>등록일</td>
	<td>등록자</td>
	<td>최종수정일</td>
	<td>최종수정자</td>
	<td>사용Y/N</td>
</tr>
<% If nItem.FResultCount > 0 Then %>
<% For i = 0 To nItem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nItem.FItemlist(i).FIdx %>"></td>
	<td><%=nItem.FItemList(i).FMallgubun%></td>
	<td><%=nItem.FItemList(i).FMakerid%></td>
	<td><%=nItem.FItemList(i).FItemid%></td>
	<td>
		<%=nItem.FItemList(i).FItemname%>_<input type="text" name="updatedepthname" value="<%=nItem.FItemList(i).FDepthname%>">
	</td>
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
		등록된 상품이 없습니다
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->