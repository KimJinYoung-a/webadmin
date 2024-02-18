<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/potal/potalCls.asp"-->
<%
Dim mallid, mode, NOTitemid, sqlStr, itemid, SavearrCnt, mallName
Dim nItem, page, itemidarr, isusingarr, makerid, bigo
mallid		= requestCheckvar(request("mallid"),32)
page		= request("page")
mode		= request("mode")
NOTitemid	= Trim(request("NOTitemid"))
makerid		= requestCheckvar(request("makerid"), 32)

Select Case mallid
	Case "ggshop"		mallName = "구글쇼핑"
	Case "naverEP"		mallName = "네이버EP"
	Case "daumEP"		mallName = "다음EP"
End Select

If Right(NOTitemid,1) = "," Then NOTitemid = Left(NOTitemid, Len(NOTitemid) - 1)

itemidarr	= request("itemidarr")
isusingarr	= request("isusingarr")
itemid		= request("itemid")
bigo 		= NullFillWith(Trim(requestCheckVar(request("bigo"),300)),"")
SavearrCnt 	= Ubound(Split(NOTitemid,",")) + 1

If page = "" Then page = 1

Dim iA2, tmpItemID, arrTemp2, arrItemid2, j
If mode = "I" Then
	If NOTitemid<>"" then
		tmpItemID = NOTitemid
		tmpItemID = replace(tmpItemID,",",chr(10))
		tmpItemID = replace(tmpItemID,chr(13),"")
		arrTemp2 = Split(tmpItemID,chr(10))
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
		arrItemid2 = left(arrItemid2,len(arrItemid2)-1)
	End If

	arrItemid2 = Split(arrItemid2, ",")
	for j = 0 to UBound(arrItemid2)
		if Trim(arrItemid2(j)) <> "" then
			NOTitemid = Trim(arrItemid2(j))
			strSql = "	DECLARE @Temp CHAR(1) " & _
						"	If NOT EXISTS(SELECT * FROM db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun = '" & mallid & "' AND itemid = '" & NOTitemid & "') " & _
						"		BEGIN " & _
						"			INSERT INTO db_temp.dbo.tbl_EpShop_not_in_itemid(itemid, mallgubun, isusing, regdate, regid, bigo) VALUES('" & NOTitemid & "','" & mallid & "', 'Y', getdate(),  '"&session("ssBctID")&"', '"& bigo &"') " & _
						"		END	"
			dbget.execute strSql


			strSql = "	DECLARE @Temp CHAR(1) " & _
						"	If NOT EXISTS(SELECT * FROM db_outmall.dbo.tbl_EpShop_not_in_itemid Where mallgubun = '" & mallid & "' AND itemid = '" & NOTitemid & "') " & _
						"		BEGIN " & _
						"			INSERT INTO db_outmall.dbo.tbl_EpShop_not_in_itemid(itemid, mallgubun, isusing, regdate, regid, bigo) VALUES('" & NOTitemid & "','" & mallid & "', 'Y', getdate(),  '"&session("ssBctID")&"', '"& bigo &"') " & _
						"		END	"
			dbCTget.execute strSql
		end if
	Next
	NOTitemid = Request("NOTitemid")
 	response.write "<script language='javascript'>alert('저장하였습니다.');location.href='/admin/etc/potal/notinitemid.asp?mallid="&mallid&"&menupos="&menupos&"';</script>"
ElseIf mode = "U" Then
	Dim cnt, tmpIsusing
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)

	isusingarr	=  split(isusingarr,",")
	For i = 0 to cnt
		tmpIsusing = isusingarr(i)
		sqlStr = "UPDATE db_temp.dbo.tbl_EpShop_not_in_itemid SET "
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,updateid = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)
		dbget.execute sqlStr

		sqlStr = "UPDATE db_outmall.dbo.tbl_EpShop_not_in_itemid SET "
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,updateid = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)
		dbCTget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('수정하였습니다.');location.href='/admin/etc/potal/notinitemid.asp?mallid="&mallid&"&menupos="&menupos&"';</script>"
End If

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

SET nItem = new CPotal
	nItem.FCurrPage					= page
	nItem.FPageSize					= 20
	nItem.FRectItemid				= itemid
	nItem.FMakerId					= makerid
	nItem.FRectMallGubun			= mallid
    nItem.getPotalNotInItemidList
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
		alert("상품코드를 입력하세요.");
		document.frm.NOTitemid.focus();
		return;
	}
	document.frm.mode.value = "I";
	document.frm.submit();
}
function goPage(pg){
    var frm = document.frmsearch;
    frm.page.value=pg;
	frm.submit();
}
</script>
<% If mallid = "ggshop" Then %>
<!-- #include virtual="/admin/etc/potal/inc_googleHead.asp" -->
<% ElseIf mallid = "naverEP" Then %>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
<% End If %>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmsearch" method="get" action="notinitemid.asp" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall 구분 : <%= mallName %></td>
		    <td rowspan="4" width="10%"><input type="button" value="검 색" onClick="goPage(1)" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
			&nbsp;
			브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<form name="frmIsusing" method="post" action="notinitemid.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mallid" value="<%= mallid %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<form name="frm" action="notinitemid.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr>
	<td>
		제외 상품코드 : <textarea class="textarea" name="NOTitemid" rows="2" cols="16"></textarea>
		&nbsp;&nbsp;
		코멘트 : <input type="text" class="text" name="bigo" size="40">
		<input type="button" class="button" value="저장" onClick="insert_makerid()">
	</td>
	<td align="right">
		<% If nItem.fresultcount >0 then %>
			<input class="button" type="button" id="btnEditSel" value="제외상품Y/N 수정" onClick="jsIsusing();">
	    <% End If %>
	<% If mallid = "naverEP" Then %>
	    (Y:네이버EP제외)
	<% End If %>
	</td>
</tr>
</form>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="30" align="LEFT" height="25">
	<td colspan="10">
		검색결과 : <b><%= FormatNumber(nItem.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(nItem.FTotalPage,0) %></b>
	</td>
</tr>
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
    <td>몰구분</td>
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>등록일</td>
	<td>등록자</td>
	<td>최종수정일</td>
	<td>최종수정자</td>
	<td>제외상품Y/N</td>
	<td>코멘트</td>
</tr>
<% If nItem.FResultCount > 0 Then %>
<% For i = 0 To nItem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nItem.FItemlist(i).FIdx %>"></td>
	<td><%=nItem.FItemList(i).FMallgubun%></td>
	<td><%=nItem.FItemList(i).FItemid%></td>
	<td><%=nItem.FItemList(i).FMakerid%></td>
	<td><%=nItem.FItemList(i).FRegdate%></td>
	<td><%=nItem.FItemList(i).FRegid%></td>
	<td><%=nItem.FItemList(i).FLastupdate%></td>
	<td><%=nItem.FItemList(i).FUpdateid%></td>
	<td>
		<select name="isusing" class="select">
			<option value="Y" <%=Chkiif(nItem.FItemList(i).FIsusing = "Y","selected","")%> >Y</option>
			<option value="N" <%=Chkiif(nItem.FItemList(i).FIsusing = "N","selected","")%> >N</option>
		</select>

		<%=CHKIIF(nItem.FItemList(i).FIsusing="N","판매함","판매안함")%>
	</td>
	<td><%=nItem.FItemList(i).Fbigo%></td>
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