<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Dim makerid, oOutmall, gubun, strSQL
Dim whyhope, adminText, adminRegdate, page, i, allHopeInsert, cisextusing, currstat, idx
Dim isBrandListOpenYN
isBrandListOpenYN	= "N"
'############### 브랜드리스트에서 팝업을 열었는 지 확인 ###############
If request("isBrandPage") = "Y" Then
	isBrandListOpenYN = "Y"
End If
'######################################################################
makerid = request("makerid")
gubun	= request("gubun")

If gubun = "D" Then
	Dim mallgubun, hopeidx
	mallgubun = request("mallgubun")
	hopeidx	= request("hopeidx")
	strSQL = ""
	strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET " & vbcrlf
	strSQL = strSQL & " isComplete = 'X' " & vbcrlf
	strSQL = strSQL & " WHERE idx = '"&hopeidx&"' " & vbcrlf
	dbget.Execute strSQL

	strSQL = ""
	strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
	strSQL = strSQL & " VALUES ('"&mallgubun&"', '"&makerid&"', '변경 요청 취소', 'X', '"&session("ssBctID")&"', getdate()) " & vbcrlf
	dbget.Execute strSQL
	response.redirect("/admin/etc/outmall/popJaehyu_Not_In_Makerid.asp?makerid="&makerid&"")
End If

SET oOutMall = new cOutmall
	Call oOutMall.fnGetIsExtusing(makerid, cisextusing, allHopeInsert, currstat, whyhope, adminText, adminRegdate, idx)
	oOutMall.FCurrPage			= 1
	oOutMall.FPageSize			= 100
	oOutMall.FRectMakerid		= makerid
	oOutMall.getOutmallList
%>
<script language='javascript'>
function popLogView(imallid, imakerid){
    var pwin = window.open("/designer/itemmaster/popHopeLog.asp?mallid="+imallid+"&makerid="+imakerid,"popHopeLog","width=850,height=700,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function SugiSell(imallid, isellyn, imakerid){
    var pwin2 = window.open("/admin/etc/outmall/popAdminComment.asp?mallid="+imallid+"&sellyn="+isellyn+"&makerid="+imakerid,"popSugiSell","width=850,height=300,scrollbars=yes,resizable=yes");
	pwin2.focus();
}

function SugiMultiSell(isellyn, imakerid) {
	var imallid, chk, currstate, i;

	imallid = "";
	for (i = 0; ; i++) {
		chk = document.getElementById("chk" + i);
		currstate = document.getElementById("currstate" + i);
		if (chk == undefined) { break; }
		if (currstate == undefined) { break; }

		if (chk.checked == true) {
			imallid = imallid + "," + chk.value;
			if ((currstate.value === "Y") && (isellyn === "Y")) {
				chk.checked = false;
				AnCheckClick(chk);
				continue;
			}
			if ((currstate.value === "N") && (isellyn === "N")) {
				chk.checked = false;
				AnCheckClick(chk);
				continue;
			}
		}
	}

	if (imallid === "") {
		alert('선택된 사이트가 없습니다.');
		return;
	}

    var pwin2 = window.open("/admin/etc/outmall/popAdminComment.asp?mallid="+imallid+"&sellyn="+isellyn+"&makerid="+imakerid,"popSugiMultiSell","width=850,height=300,scrollbars=yes,resizable=yes");
	pwin2.focus();
}

function cancelHope(imallid, ihopeidx){
	var frm = document.frmDel;
	frm.mallgubun.value = imallid;
	frm.hopeidx.value = ihopeidx;
	frm.submit();
}
</script>
<form name="frmDel" method="POST" action="<%=CurrURL()%>" style="margin:0px;">
	<input type="hidden" name="gubun" value="D">
	<input type="hidden" name="mallgubun">
	<input type="hidden" name="hopeidx">
	<input type="hidden" name="makerid" value="<%=makerid%>">
</form>
<% If isBrandListOpenYN <> "Y" Then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<% Else %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <%= makerid %>
	</td>
</tr>
</table>
<% End If %>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 판매 설정 변경 요청 버튼 클릭 시, <font color="red">팝업창 내에 기재하는 사유(필수 입력)</font>는 업체 및 제휴업무 담당자에게 노출됩니다.<br>
		* 변경 요청 후 1주일 이내로, <font color="red">마케팅팀 제휴 담당자의 승인 절차를 거쳐 판매설정이 변경 완료됩니다.</font>
	</td>
</tr>
</table>
<br />
<%
	If makerid <> "" Then
		If oOutMall.isUsingMakerid(makerid) = 0  Then
			response.write "<script>alert('등록된 브랜드가 아닙니다.');location.replace('/admin/etc/outmall/popJaehyu_Not_In_Makerid.asp');</script>"
		End If
%>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a">
	<tr>
		<td align="left" valign="bottom"></td>
		<td align="right">
			<input type='button' value='선택 판매안함' class='button' onclick="SugiMultiSell('N', '<%=makerid%>');">
			&nbsp;
			<input type='button' value='선택 판매' class='button' onclick="SugiMultiSell('Y', '<%=makerid%>');">
		</td>
	</tr>
</table>
<form name="frm1" method="post" onSubmit="return false;" action="" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frm1.cksel);"></td>
	<td width="60">구분</td>
	<td width="100">사이트</td>
	<td width="100"><font color="BLUE"><strong>현재판매설정</strong></font></td>
	<td width="100">등록자</td>
	<td width="300">최종설정일</td>
	<td width="130"><font color="RED">변경요청</font></td>
	<td>상태</td>
	<td width="100">내역확인</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td></td>
	<td>제휴사</td>
	<td>제휴사전체</td>
	<td><strong><%= Chkiif(cisextusing="N", "판매안함", "판매") %></strong></td>
	<td width="400" colspan="2">제휴사 전체 [판매안함] 인경우 아래 몰별 설정과 관계없이 판매안함 </td>
	<td>
	<%
		Dim disableChk : disableChk = false
		If allHopeInsert = "Y" and currstat = 0 Then
			disableChk = true
		End If

		If cisextusing = "Y" Then
			response.write "<input type='button' value='판매안함' class='button' onclick=""SugiSell('all', 'N', '"&makerid&"');"">"
		Else
			response.write "<input type='button' value='판매' class='button' onclick=""SugiSell('all', 'Y', '"&makerid&"');"">"
		End If
	%>
	</td>
	<td>
	<%
		If allHopeInsert = "Y" Then
			Select Case currstat
				Case "1"	response.write "<font title='"& whyhope &"'>승인대기</font>"
							response.write "<br><span style='cursor:pointer;' onclick=""cancelHope('all', "&idx&");""><font color='gray'>[변경 요청 취소]</font></span>"
				Case "2"	response.write "<font title='"& whyhope &"'>반려</font>"
							response.write "<br><font color='RED'>("&adminText&")</font>"
							response.write "<br>"&adminRegdate
			End Select
		End If
	%>
	</td>
	<td><input type='button' value='보기' class='button' onclick="popLogView('all', '<%=makerid%>');"></td>
</tr>
<tr height="2" bgcolor="#FFFFFF" >
	<td colspan="9"></td>
</tr>
<% For i = 0 To oOutMall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<input type="checkbox" name="cksel" id="chk<%= i %>" value="<%= oOutMall.FItemList(i).FMallid %>" onClick="AnCheckClick(this);">
		<input type="hidden" name="currstate<%= i %>" id="currstate<%= i %>" value="<%= CHKIIF(oOutMall.FItemList(i).FIdx = "0", "Y", "N")%>">
	</td>
	<td>제휴사</td>
	<td><%= oOutMall.FItemList(i).FMallid %></td>
	<td><strong><%= Chkiif(oOutMall.FItemList(i).FIdx = "0", "판매", "판매안함") %></strong></td>
	<td><%= oOutMall.FItemList(i).FReguserid %></td>
	<td><%= oOutMall.FItemList(i).FRegdate %></td>
	<td>
	<%
	If oOutMall.FItemList(i).FIdx = "0" Then
		response.write "<input type='button' value='판매안함' class='button' onclick=""SugiSell('"&oOutMall.FItemList(i).FMallid&"', 'N', '"&makerid&"');"">"
	Else
		response.write "<input type='button' value='판매' class='button' onclick=""SugiSell('"&oOutMall.FItemList(i).FMallid&"', 'Y', '"&makerid&"');"">"
	End If
	%>
	<td>
	<%
		Select Case oOutMall.FItemList(i).FCurrstat
			Case "1"	response.write "<font title='"& oOutMall.FItemList(i).FWhyhope &"'>승인대기</font>"
						response.write "<br><span style='cursor:pointer;' onclick=cancelHope('"&oOutMall.FItemList(i).FMallid&"','"&oOutMall.FItemList(i).FHopeidx&"');><font color='gray'>[변경 요청 취소]</font></span>"
			Case "2"	response.write "<font color='RED' title='"& oOutMall.FItemList(i).FWhyhope &"'>반려</font>"
						response.write "<br><font color='RED'>("&oOutMall.FItemList(i).FadminText&")</font>"
						response.write "<br>"&oOutMall.FItemList(i).FadminRegdate
		End Select
	%>
	</td>
	</td>
	<td><input type='button' value='보기' class='button' onclick="popLogView('<%= oOutMall.FItemList(i).FMallid %>', '<%= makerid %>');"></td>
</tr>
<% Next %>
<tr height="2" bgcolor="#FFFFFF" >
	<td colspan="9"></td>
</tr>
<%
	Dim ospcialOutmall
	SET ospcialOutmall = new cOutmall
		ospcialOutmall.FCurrPage		= 1
		ospcialOutmall.FPageSize		= 100
		ospcialOutmall.FRectMakerid		= makerid
		ospcialOutmall.getSpecialOutmallList
%>
<% For i = 0 To ospcialOutmall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<input type="checkbox" name="cksel" id="chk<%= oOutMall.FResultCount + 99999 + i %>" value="<%= ospcialOutmall.FItemList(i).FMallid %>" onClick="AnCheckClick(this);">
		<input type="hidden" name="currstate<%= oOutMall.FResultCount + 99999 + i %>" id="currstate<%= oOutMall.FResultCount + 99999 + i %>" value="<%= CHKIIF(ospcialOutmall.FItemList(i).FIdx = "0", "Y", "N")%>">
	</td>
	<td>제휴사 + </td>
	<td><%= ospcialOutmall.FItemList(i).FMallid %></td>
	<td><strong><%= Chkiif(ospcialOutmall.FItemList(i).FIdx = "0", "판매", "판매안함") %></strong></td>
	<td><%= ospcialOutmall.FItemList(i).FReguserid %></td>
	<td><%= ospcialOutmall.FItemList(i).FRegdate %></td>
	<td>
	<%
	If ospcialOutmall.FItemList(i).FIdx = "0" Then
		response.write "<input type='button' value='판매안함' class='button' onclick=""SugiSell('"&ospcialOutmall.FItemList(i).FMallid&"', 'N', '"&makerid&"');"">"
	Else
		response.write "<input type='button' value='판매' class='button' onclick=""SugiSell('"&ospcialOutmall.FItemList(i).FMallid&"', 'Y', '"&makerid&"');"">"
	End If
	%>
	<td>
	<%
		Select Case ospcialOutmall.FItemList(i).FCurrstat
			Case "1"	response.write "<font title='"& ospcialOutmall.FItemList(i).FWhyhope &"'>승인대기</font>"
						response.write "<br><span style='cursor:pointer;' onclick=cancelHope('"&ospcialOutmall.FItemList(i).FMallid&"','"&ospcialOutmall.FItemList(i).FHopeidx&"');><font color='gray'>[변경 요청 취소]</font></span>"
			Case "2"	response.write "<font color='RED' title='"& ospcialOutmall.FItemList(i).FWhyhope &"'>반려</font>"
						response.write "<br><font color='RED'>("&ospcialOutmall.FItemList(i).FadminText&")</font>"
						response.write "<br>"&ospcialOutmall.FItemList(i).FadminRegdate
		End Select
	%>
	</td>
	</td>
	<td><input type='button' value='보기' class='button' onclick="popLogView('<%= ospcialOutmall.FItemList(i).FMallid %>', '<%= makerid %>');"></td>
</tr>
<% Next %>
<%	SET ospcialOutmall = nothing %>
<% SET oPotalsite = nothing %>
<tr height="2" bgcolor="#FFFFFF" >
	<td colspan="9"></td>
</tr>
<%
	Dim oPotalsite
	SET oPotalsite = new cOutmall
		oPotalsite.FCurrPage		= 1
		oPotalsite.FPageSize		= 100
		oPotalsite.FRectMakerid		= makerid
		oPotalsite.getPotalSiteList
%>
<% For i = 0 To oPotalsite.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<input type="checkbox" name="cksel" id="chk<%= oOutMall.FResultCount + i %>" value="<%= oPotalsite.FItemList(i).FMallid %>" onClick="AnCheckClick(this);">
		<input type="hidden" name="currstate<%= oOutMall.FResultCount + i %>" id="currstate<%= oOutMall.FResultCount + i %>" value="<%= CHKIIF(oPotalsite.FItemList(i).FIsusing="Y", "Y", "N")%>">
	</td>
	<td>EP</td>
	<td>
	<%
		Select Case oPotalsite.FItemList(i).FMallID
			Case "naverep" response.write "네이버"
			Case "daumep" response.write "다음"
			Case "shodocep" response.write "쇼닥"
			Case "wemakepriceep" response.write "위메프"
			Case "ggshop" response.write "구글쇼핑"
		End Select
	%>
	</td>
	<td><strong><%= Chkiif(oPotalsite.FItemList(i).FIsusing = "Y", "판매", "판매안함") %></strong></td>
	<td>
		<%= Chkiif(isnull(oPotalsite.FItemList(i).FUpdateid), oPotalsite.FItemList(i).FReguserid, oPotalsite.FItemList(i).FUpdateid) %>
	</td>
	<td>
		<%= Chkiif(isnull(oPotalsite.FItemList(i).FLastUpdate), oPotalsite.FItemList(i).FRegdate, oPotalsite.FItemList(i).FLastUpdate) %>
	</td>
	<td>
	<%
		If oPotalsite.FItemList(i).FIsusing = "Y" Then
			response.write "<input type='button' value='판매안함' class='button' onclick=""SugiSell('"&oPotalsite.FItemList(i).FMallid&"', 'N', '"&makerid&"');""> "
		Else
			response.write "<input type='button' value='판매' class='button' onclick=""SugiSell('"&oPotalsite.FItemList(i).FMallid&"', 'Y', '"&makerid&"');""> "
		End If
	%>
	</td>
	<td>
	<%
		Select Case oPotalsite.FItemList(i).FCurrstat
			Case "1"	response.write "<font title='"& oPotalsite.FItemList(i).FWhyhope &"'>승인대기</font>"
						response.write "<br><span style='cursor:pointer;' onclick=cancelHope('"&oPotalsite.FItemList(i).FMallid&"','"&oPotalsite.FItemList(i).FHopeidx&"');><font color='gray'>[변경 요청 취소]</font></span>"
			Case "2"	response.write "<font color='RED' title='"& oPotalsite.FItemList(i).FWhyhope &"'>반려</font>"
						response.write "<br><font color='RED'>("&oPotalsite.FItemList(i).FadminText&")</font>"
						response.write "<br>"&oPotalsite.FItemList(i).FadminRegdate
		End Select
	%>
	</td>
	<td><input type='button' value='보기' class='button' onclick="popLogView('<%= oPotalsite.FItemList(i).FMallid %>', '<%= makerid %>');"></td>
</tr>
<% Next %>
<% SET oPotalsite = nothing %>
<tr height="2" bgcolor="#FFFFFF" >
	<td colspan="9"></td>
</tr>
<%
	Dim oOverseassite
	SET oOverseassite = new cOutmall
		oOverseassite.FCurrPage		= 1
		oOverseassite.FPageSize		= 100
		oOverseassite.FRectMakerid		= makerid
		oOverseassite.getOverseasOutmallList
%>
<% For i = 0 To oOverseassite.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<input type="checkbox" name="cksel" id="chk<%= oOutMall.FResultCount + i %>" value="<%= oOverseassite.FItemList(i).FMallid %>" onClick="AnCheckClick(this);">
		<input type="hidden" name="currstate<%= oOutMall.FResultCount + i %>" id="currstate<%= oOutMall.FResultCount + i %>" value="<%= CHKIIF(oOverseassite.FItemList(i).FIsusing="Y", "Y", "N")%>">
	</td>
	<td>해외몰</td>
	<td><%= oOverseassite.FItemList(i).FMallid %></td>
	<td><strong><%= Chkiif(oOverseassite.FItemList(i).FIdx = "0", "판매", "판매안함") %></strong></td>
	<td>
		<%= Chkiif(isnull(oOverseassite.FItemList(i).FUpdateid), oOverseassite.FItemList(i).FReguserid, oOverseassite.FItemList(i).FUpdateid) %>
	</td>
	<td>
		<%= Chkiif(isnull(oOverseassite.FItemList(i).FLastUpdate), oOverseassite.FItemList(i).FRegdate, oOverseassite.FItemList(i).FLastUpdate) %>
	</td>
	<td>
	<%
		If oOverseassite.FItemList(i).FIdx = "0" Then
			response.write "<input type='button' value='판매안함' class='button' onclick=""SugiSell('"&oOverseassite.FItemList(i).FMallid&"', 'N', '"&makerid&"');""> "
		Else
			response.write "<input type='button' value='판매' class='button' onclick=""SugiSell('"&oOverseassite.FItemList(i).FMallid&"', 'Y', '"&makerid&"');""> "
		End If
	%>
	</td>
	<td>
	<%
		Select Case oOverseassite.FItemList(i).FCurrstat
			Case "1"	response.write "<font title='"& oOverseassite.FItemList(i).FWhyhope &"'>승인대기</font>"
						response.write "<br><span style='cursor:pointer;' onclick=cancelHope('"&oOverseassite.FItemList(i).FMallid&"','"&oOverseassite.FItemList(i).FHopeidx&"');><font color='gray'>[변경 요청 취소]</font></span>"
			Case "2"	response.write "<font color='RED' title='"& oOverseassite.FItemList(i).FWhyhope &"'>반려</font>"
						response.write "<br><font color='RED'>("&oOverseassite.FItemList(i).FadminText&")</font>"
						response.write "<br>"&oOverseassite.FItemList(i).FadminRegdate
		End Select
	%>
	</td>
	<td><input type='button' value='보기' class='button' onclick="popLogView('<%= oOverseassite.FItemList(i).FMallid %>', '<%= makerid %>');"></td>
</tr>
<% Next %>
<% SET oOverseassite = nothing %>
</table>
</form>
<% End If %>
<% SET oOutMall = nothing %>
<form name="frmSvArr" method="post">
<input type="hidden" name="cmdparam">
<input type="hidden" name="sugimallid">
<input type="hidden" name="sugisellyn">
<input type="hidden" name="sugiadminid" value="<%=session("ssBctID")%>">
<input type="hidden" name="sugimakerid" value="<%=makerid%>">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100"></iframe>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
