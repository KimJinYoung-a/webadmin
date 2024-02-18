<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim mallid, oOutmall, page, i
Dim itemid, apiAction, resultCode, lastUserid
mallid		= request("mallid")
itemid		= request("itemid")
apiAction	= request("apiAction")
resultCode	= request("resultCode")
page 		= request("page")
lastUserid	= request("lastUserid")

If page = "" Then page = 1
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

Set oOutmall = new COutmall
	If (session("ssBctID")="kjy8517") Then
		oOutmall.FPageSize 			= 50
	Else
		oOutmall.FPageSize 			= 20
	End If
	oOutmall.FCurrPage			= page
	oOutmall.FRectMallid 		= mallid
	oOutmall.FRectItemid 		= itemid
	oOutmall.FRectApiAction 	= apiAction
	oOutmall.FRectResultCode 	= resultCode
	oOutmall.FRectLastUserid 	= lastUserid
	oOutmall.getQueOptionLogList
%>
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
// 선택된 상품 판매여부 변경
function etcmallSellYnProcess(chkYn, imallid) {
	var chkSel=0, strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
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

	switch(chkYn) {
		case "Y": strSell="판매중";break;
		case "N": strSell="품절";break;
	}

	if (imallid == 'gsshop'){
	    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        if (chkYn=="X"){
	            if (!confirm(strSell + '로 변경하면 GSShop에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
	        }
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
	        document.frmSvArr.submit();
	    }
	}else if (imallid == 'lotteimall'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 롯데iMall에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
		}	
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td>몰구분 : <%= mallid %></td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		API액션 : 
		<select name="apiAction" class="select">
			<option value="">전체</option>
			<option value="REG"  <%= Chkiif(apiAction = "REG", "selected", "")%> >상품등록</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >상품수정</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >품절처리</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >가격수정</option>
			<% If (mallid <> "cjmall" and mallid <> "11stmy") Then %>
			<option value="ITEMNAME" <%= Chkiif(apiAction = "ITEMNAME", "selected", "")%> >상품명수정</option>
			<% End If %>
			<% If mallid = "gsshop" Then %>
			<option value="IMAGE"	 <%= Chkiif(apiAction = "IMAGE", "selected", "")%> >이미지수정</option>
			<option value="CONTENT"  <%= Chkiif(apiAction = "CONTENT", "selected", "")%> >상품설명수정</option>
			<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >정부고시수정</option>
			<% ElseIf mallid = "11stmy" Then %>
			<option value="VIEWOPT"  <%= Chkiif(apiAction = "VIEWOPT", "selected", "")%> >옵션조회</option>
			<% ElseIf mallid = "cjmall" Then %>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >신규상품조회</option>
			<% ElseIf (mallid = "lotteimall") OR (mallid = "lotteCom") Then %>
			<option value="CHKSTOCK"  <%= Chkiif(apiAction = "CHKSTOCK", "selected", "")%> >재고조회</option>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >신규상품조회</option>
				<% If mallid = "lotteimall" Then %>
					<option value="DISPVIEW"  <%= Chkiif(apiAction = "DISPVIEW", "selected", "")%> >전시상품조회</option>
				<% Else %>
					<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >정부고시수정</option>
				<% End If %>
			<% End If %>
		</select>
		&nbsp;
		성공여부 : 
		<select name="resultCode" class="select">
			<option value="">전체</option>
			<option value="OK"  	<%= Chkiif(resultCode = "OK", "selected", "")%> >성공</option>
			<option value="ERR"		<%= Chkiif(resultCode = "ERR", "selected", "")%> >에러</option>
			<option value="QNull"	<%= Chkiif(resultCode = "QNull", "selected", "")%> >예정</option>
		</select>
		&nbsp;
		수행ID : 
		<select name="lastUserid" class="select">
			<option value="">전체</option>
			<option value="system"	<%= Chkiif(lastUserid = "system", "selected", "")%> >스케줄</option>
			<option value="etc"		<%= Chkiif(lastUserid = "etc", "selected", "")%> >관리자</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="subcmd" value="">
<input type="hidden" name="chgSellYn" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14">
		검색결과 : <b><%= FormatNumber(oOutmall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOutmall.FTotalPage,0) %></b>
	</td>
	<td align="right" valign="top">
		선택상품을 품절로
		<input class="button" type="button" id="btnSellYn" value="변경" onClick="etcmallSellYnProcess('N', '<%= mallid %>');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td>몰구분</td>
	<td>API액션</td>
	<td>아웃몰코드</td>
	<td>상품코드</td>
	<td>옵션코드</td>
	<td>우선순위</td>
	<td>등록시간</td>
	<td>큐읽은시간</td>
	<td>API완료시간</td>
	<td>제휴판매</td>
	<td>실패수</td>
	<td>성공여부</td>
	<td>수행ID</td>
	<td width="300">Message</td>
</tr>
<% For i = 0 To oOutmall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oOutmall.FItemList(i).FMidx %>"></td>
	<td><%= oOutmall.FItemlist(i).FMallid %></td>
	<td><%= oOutmall.FItemlist(i).FApiAction %></td>
	<td><%= oOutmall.FItemlist(i).FOutmallGoodno %></td>
	<td><%= oOutmall.FItemlist(i).FItemid %></td>
	<td><%= oOutmall.FItemlist(i).FItemoption %></td>
	<td><%= oOutmall.FItemlist(i).FPriority %></td>
	<td><%= oOutmall.FItemlist(i).FRegdate %></td>
	<td><%= oOutmall.FItemlist(i).FReaddate %></td>
	<td><%= oOutmall.FItemlist(i).FFindate %></td>
	<td>
		<%
			If oOutmall.FItemlist(i).FGSShopSellyn = "Y" Then
				response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			Else
				response.write "<font color='RED'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			End If
		%>
	</td>
	<td><%= oOutmall.FItemlist(i).FAccFailCnt %></td>
	<td>
	<%
		Select Case oOutmall.FItemlist(i).FResultCode
			Case "OK"		response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case "ERR"		response.write "<font color='RED'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case Else		response.write "<font color='GRAY'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
		End Select
	%>
	</td>
	<td><%= oOutmall.FItemlist(i).FLastUserid %></td>
	<td width="300"><%= oOutmall.FItemlist(i).FLastErrMsg %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	<% If oOutmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oOutmall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oOutmall.StartScrollPage To oOutmall.FScrollCount + oOutmall.StartScrollPage - 1 %>
		<% If i>oOutmall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oOutmall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<% Set oOutmall = Nothing %>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="400"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
