<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/potal/potalCls.asp"-->
<%
Dim mallid, mode, couponmakerid, sqlStr, SavearrCnt, mallName
Dim nItem, page, makeridarr, makerid, bigo
mallid		= requestCheckvar(request("mallid"),32)
page		= request("page")
mode		= request("mode")
couponmakerid	= Trim(request("couponmakerid"))
makerid		= requestCheckvar(request("makerid"), 32)

Select Case mallid
	Case "ggshop"		mallName = "구글쇼핑"
	Case "naverEP"		mallName = "네이버EP"
	Case "daumEP"		mallName = "다음EP"
End Select

If Right(couponmakerid,1) = "," Then couponmakerid = Left(couponmakerid, Len(couponmakerid) - 1)

makeridarr	= request("makeridarr")
bigo 		= NullFillWith(Trim(requestCheckVar(request("bigo"),300)),"")
SavearrCnt 	= Ubound(Split(couponmakerid,",")) + 1

If page = "" Then page = 1

Dim iA2, tmpMakerID, arrTemp2, arrMakerid2, j
If mode = "I" Then
rw couponmakerid & "!!"
	If couponmakerid<>"" then
		tmpMakerID = couponmakerid
		tmpMakerID = replace(tmpMakerID,",",chr(10))
		tmpMakerID = replace(tmpMakerID,chr(13),"")
		arrTemp2 = Split(tmpMakerID,chr(10))
		iA2 = 0
		Do While iA2 <= ubound(arrTemp2)
			If Trim(arrTemp2(iA2))<>"" then
				arrMakerid2 = arrMakerid2 & trim(arrTemp2(iA2)) & ","
			End If
			iA2 = iA2 + 1
		Loop
		arrMakerid2 = left(arrMakerid2,len(arrMakerid2)-1)
	End If

	arrMakerid2 = Split(arrMakerid2, ",")

	for j = 0 to UBound(arrMakerid2)
		if Trim(arrMakerid2(j)) <> "" then
			couponmakerid = Trim(arrMakerid2(j))
			strSql = 	"	If NOT EXISTS(SELECT * FROM db_item.[dbo].[tbl_nvs_item_force_coupon_by_makerid] Where makerid = '" & couponmakerid & "') " & _
						"		BEGIN " & _
						"			INSERT INTO db_item.[dbo].[tbl_nvs_item_force_coupon_by_makerid] (makerid, regdate, adminid, comment) " & _
						"			SELECT '" & couponmakerid & "', GETDATE(), '"&session("ssBctID")&"', '"& bigo &"' " & _
						"			FROM db_partner.dbo.tbl_partner " & _
						"			WHERE id = '" & couponmakerid & "' " & _
						"		END	"
			dbget.execute strSql
		end if
	Next

	couponmakerid = Request("couponmakerid")
 	response.write "<script language='javascript'>alert('저장하였습니다.');location.href='/admin/etc/naverEp/couponMakerid.asp?mallid="&mallid&"&menupos="&menupos&"';</script>"
ElseIf mode = "U" Then
	Dim cnt
	makeridarr = split(makeridarr,",")
	cnt = ubound(makeridarr)
	For i = 0 to cnt
		sqlStr = "DELETE db_item.[dbo].[tbl_nvs_item_force_coupon_by_makerid] WHERE makerid = '"& makeridarr(i) &"' "
		dbget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('삭제 하였습니다.');location.href='/admin/etc/naverEp/couponMakerid.asp?mallid="&mallid&"&menupos="&menupos&"';</script>"
End If

If makerid<>"" then
	Dim iA, arrTemp, arrmakerid
	makerid = replace(makerid,",",chr(10))
	makerid = replace(makerid,chr(13),"")
	arrTemp = Split(makerid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			arrmakerid = arrmakerid & trim(arrTemp(iA)) & ","
		End If
		iA = iA + 1
	Loop
	makerid = left(arrmakerid,len(arrmakerid)-1)
End If

SET nItem = new CPotal
	nItem.FCurrPage					= page
	nItem.FPageSize					= 100
	nItem.FMakerId					= makerid
    nItem.getPotalCouponMakeridList
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
	var sValue;
	frm = document.fitem;
	sValue = "";
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
			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
		}
	}
	if(chkSel<=0) {
		alert("선택한 상품이 없습니다.");
		return;
	}

	if(confirm("삭제 하시겠습니까?")){
		document.frmIsusing.makeridarr.value = sValue;
		document.frmIsusing.mode.value = "U";
		document.frmIsusing.submit();
	}
}

function insert_makerid()
{
	if(document.frm.couponmakerid.value == "")
	{
		alert("브랜드ID를 입력하세요.");
		document.frm.couponmakerid.focus();
		return;
	}
	if(confirm("저장 하시겠습니까?")){
		document.frm.mode.value = "I";
		document.frm.submit();
	}
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
<form name="frmsearch" method="get" action="couponMakerid.asp" style="margin:0px;">
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
			브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<form name="frmIsusing" method="post" action="couponMakerid.asp" style="margin:0px;">
	<input type="hidden" name="makeridarr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mallid" value="<%= mallid %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<form name="frm" action="couponMakerid.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr>
	<td>
		쿠폰 적용 브랜드ID : <textarea class="textarea" name="couponmakerid" rows="2" cols="16"></textarea>
		&nbsp;&nbsp;
		코멘트 : <input type="text" class="text" name="bigo" size="40">
		<input type="button" class="button" value="저장" onClick="insert_makerid()">
	</td>
	<td align="right">
		<% If nItem.fresultcount >0 then %>
			<input class="button" type="button" id="btnEditSel" value="쿠폰적용삭제" onClick="jsIsusing();">
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
	<td>브랜드ID</td>
	<td>등록일</td>
	<td>등록자</td>
	<td>코멘트</td>
</tr>
<% If nItem.FResultCount > 0 Then %>
<% For i = 0 To nItem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nItem.FItemlist(i).FMakerid %>"></td>
	<td><%=nItem.FItemList(i).FMakerid%></td>
	<td><%=nItem.FItemList(i).FRegdate%></td>
	<td><%=nItem.FItemList(i).FRegid%></td>
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
		등록된 브랜드가 없습니다
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->