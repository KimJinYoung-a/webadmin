<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim mode, chgNmmakerid, depthname, sqlStr, makerid, updatedepthnamearr
Dim nMaker, page, itemidarr, isusingarr
page			= request("page")
mode			= request("mode")
chgNmmakerid	= Trim(request("chgNmmakerid"))
depthname		= Trim(request("depthname"))
itemidarr		= request("itemidarr")
isusingarr		= request("isusingarr")
makerid			= Trim(request("makerid"))
updatedepthnamearr	= Trim(request("updatedepthnamearr"))
If page = "" Then page = 1
If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT count(c.userid) as userCNT, count(m.makerid) as notmakerCNT "
	sqlStr = sqlStr & " FROM db_AppWish.dbo.tbl_user_c as c "
	sqlStr = sqlStr & " LEFT JOIN db_outmall.dbo.tbl_EpShop_makerid_3depthName as m on m.makerid = c.userid AND m.mallgubun = 'naverEP' "
	sqlStr = sqlStr & " WHERE c.userid = '"&chgNmmakerid&"' "
	rsCTget.Open sqlStr,dbCTget,1
	If rsCTget("userCNT") = 0 Then
		response.write "<script language='javascript'>alert('텐바이텐에 등록된 브랜드가 아닙니다.');location.href='/admin/etc/naverEp/3depthmakerid.asp?menupos="&menupos&"';</script>"
	ElseIf rsCTget("notmakerCNT") > 0 Then
		response.write "<script language='javascript'>alert('이미 등록된 브랜드 입니다.');location.href='/admin/etc/naverEp/3depthmakerid.asp?menupos="&menupos&"';</script>"
	Else
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_EpShop_makerid_3depthName (makerid, mallgubun, depthname, isusing, regdate, regid) VALUES "
		sqlStr = sqlStr & " ('"&chgNmmakerid&"', 'naverep', '"&depthname&"', 'Y' ,getdate(), '"&session("ssBctID")&"') "
		dbget.Execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_EpShop_makerid_3depthName (makerid, mallgubun, depthname, isusing, regdate, regid) VALUES "
		sqlStr = sqlStr & " ('"&chgNmmakerid&"', 'naverep', '"&depthname&"', 'Y' ,getdate(), '"&session("ssBctID")&"') "
		sqlStr = sqlStr & vbCRLF & "; update I SET lastupdate=getdate() from [db_AppWish].dbo.tbl_item I where I.makerid='"&chgNmmakerid&"'"
		dbCTget.Execute sqlStr
		response.write "<script language='javascript'>alert('저장하였습니다.');location.href='/admin/etc/naverEp/3depthmakerid.asp?menupos="&menupos&"';</script>"
	End If
	rsCTget.Close
ElseIf mode = "U" Then
	Dim cnt, tmpIsusing
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)

	updatedepthnamearr = split(updatedepthnamearr,",")
	isusingarr	=  split(isusingarr,",")
	For i = 0 to cnt
		tmpIsusing = isusingarr(i)
		sqlStr = "UPDATE db_temp.dbo.tbl_EpShop_makerid_3depthName SET "
		sqlStr = sqlStr & " depthname = '"&html2db(updatedepthnamearr(i))&"'"
		sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,updateid = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)
		dbget.execute sqlStr

		sqlStr = "UPDATE db_outmall.dbo.tbl_EpShop_makerid_3depthName SET "
		sqlStr = sqlStr & " depthname = '"&html2db(updatedepthnamearr(i))&"'"
		sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,updateid = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)
		sqlStr = sqlStr & " ;" &vbCRLF
		sqlStr = sqlStr & " update I SET lastupdate=getdate() "
		sqlStr = sqlStr & " from [db_AppWish].dbo.tbl_item I "
		sqlStr = sqlStr & " INNER JOIN db_outmall.dbo.tbl_EpShop_makerid_3depthName T "
		sqlStr = sqlStr & " ON T.idx="&itemidarr(i)
		sqlStr = sqlStr & " and T.makerid=I.makerid and i.sellyn='Y'"
		dbCTget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('수정하였습니다.');location.href='/admin/etc/naverEp/3depthmakerid.asp?menupos="&menupos&"';</script>"
End If

SET nMaker = new epShop
	nMaker.FCurrPage					= page
	nMaker.FPageSize					= 20
	nMaker.FMakerId						= makerid
    nMaker.EpshopChgMakerid3depthList
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
		alert("선택한 브랜드가 없습니다.");
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
	if(document.frm.chgNmmakerid.value == "")
	{
		alert("브랜드ID를 입력하세요.");
		document.frm.chgNmmakerid.focus();
		return;
	}
	if(document.frm.depthname.value == "")
	{
		alert("브랜드 3Depth명을 입력하세요.");
		document.frm.depthname.focus();
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
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
<!-- 검색 시작 -->

<form name="frmsearch" method="get" action="3depthmakerid.asp" style="margin:0px;">
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
			브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<form name="frmIsusing" method="post" action="3depthmakerid.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="updatedepthnamearr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<form name="frm" action="3depthmakerid.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<tr>
	<td>
		브랜드 <input type="text" class="text" name="chgNmmakerid">&nbsp;&nbsp;
		브랜드 3Depth명 <input type="text" class="text" name="depthname">&nbsp;&nbsp;<input type="button" class="button" value="저장" onClick="chk3depthform();">
	</td>
	<td align="right">
		<% If nMaker.fresultcount >0 then %>
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
	<td>특정 브랜드 3Depth명</td>
	<td>등록일</td>
	<td>등록자</td>
	<td>최종수정일</td>
	<td>최종수정자</td>
	<td>사용Y/N</td>
</tr>
<% If nMaker.FResultCount > 0 Then %>
<% For i = 0 To nMaker.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nMaker.FItemlist(i).FIdx %>"></td>
	<td><%=nMaker.FItemList(i).FMallgubun%></td>
	<td><%=nMaker.FItemList(i).FMakerid%></td>
	<td>
		<input type="text" name="updatedepthname" value="<%=nMaker.FItemList(i).FDepthname%>">
	</td>
	<td><%=nMaker.FItemList(i).FRegdate%></td>
	<td><%=nMaker.FItemList(i).FRegid%></td>
	<td><%=nMaker.FItemList(i).FLastupdate%></td>
	<td><%=nMaker.FItemList(i).FUpdateid%></td>
	<td>
		<select name="isusing" class="select">
			<option value="Y" <%=Chkiif(nMaker.FItemList(i).FIsusing = "Y","selected","")%> >Y</option>
			<option value="N" <%=Chkiif(nMaker.FItemList(i).FIsusing = "N","selected","")%> >N</option>
		</select>
	</td>
</tr>
<% Next %>
<tr height="30">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
	<% If nMaker.HasPreScroll Then %>
		<a href="javascript:goPage('<%= nMaker.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + nMaker.StartScrollPage To nMaker.FScrollCount + nMaker.StartScrollPage - 1 %>
		<% If i>nMaker.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If nMaker.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
		등록된 브랜드가 없습니다
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nMaker = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->