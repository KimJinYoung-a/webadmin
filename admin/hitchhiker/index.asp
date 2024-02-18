<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	PageName 	: /admin/hitchhiker/index.asp
'	Description : 히치하이커 신청회원리스트 다운및 발송확인
'	History		: 2006.11.30 정윤정 생성
'                 2012.02.13 허진원 - 미니달력 교체
'				  2016.07.19 한용민 수정 SSL 적용
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhikerCls.asp"-->
<%
Dim clsHList, arrHVol, intHV, iHVol, arrAVol, intA , iAVol, blnSend, searchTxt, search
Dim arrHList, intH, iPageSize, iCurrpage, iTotCnt, iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim chkList,chkView, startDate, endDate
	iHVol = Request("iHV")
	iAVol =	 Request("iAV")
	startDate	= Request("startDate")
	endDate	= Request("endDate")
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	searchTxt = Request("searchTxt")
	search= Request("search")
	chkList = Request("chkList")
	IF chkList ="view" THEN
		chkView  = ""
	ELSE
		chkView  = "none"
	END IF

set clsHList =  new Chitchhiker
	arrHVol	= clsHList.fnGetHVol	'1.발행회차 가져오기
	IF iHVol = "" THEN
		IF isArray(arrHVol) THEN
		iHVol	= arrHVol(0,0)
		END IF
	END IF

	clsHList.FHVol = iHVol			'Set 발행회차
	arrAVol = clsHList.fnGetApplyVol	'2.신청회차 가져오기

	IF iAVol = "" and blnSend ="" THEN
		IF isArray(arrAVol) THEN
			iAVol = arrAVol(0,0)
		END IF
	END IF

	clsHList.FAVol = iAVol			'Set 신청회차
	clsHList.FisSend = blnSend		'Set 발송여부

	IF chkList = "view" THEN
	clsHList.FPSize = iPageSize		'Set 페이지 사이즈
	clsHList.FCPage = iCurrpage		'Set 현재 페이지 번호
	clsHList.FSearch = search
	clsHList.FSearchTxt = searchTxt	'Set 검색어
	clsHList.FSDate = startDate		'검색시작일
	clsHList.FEDate = endDate		'검색종료일
	arrHList = clsHList.fnGetList	'3.신청 리스트 가져오기
	iTotCnt = clsHList.FTotCnt 		'신청리스트 총 갯수 가져오기
	ELSE
		arrHList = NULL
	END IF
set clsHList = nothing

'전체 페이지 수
iTotalPage 	=  Int(iTotCnt/iPageSize)
IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1

%>
<script type="text/javascript">
<!--
	function jsSetPage(frm, sVal){
		eval("frm."+sVal).value = "";
		frm.submit();
	}

	function jsGetList(frm){
		frm.chkList.value = "view";
		frm.submit();
	}


	function jsDownList(frm){
	    var iHV = frm.iHV.value;
	    var iAV = frm.iAV.value;
	    var iType = frm.blnS.value;
	    var sSdt = frm.startdate.value;
	    var sEdt = frm.enddate.value;
//alert(iType);
//return;
	    var popwin = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/downHitchhiker.asp?iHV=' + iHV + '&iAV='+iAV+'&iType='+iType+'&startDate='+sSdt+'&endDate='+sEdt,'downHitchhiker','width=400, height=300');
		popwin.focus();

	}

	function jsSend(frm){
		var smsyn='';
		smsyn = 'N';
		if (frm.smsyn.value=='Y'){
			if(confirm("고객님께 발송처리 문자발송을 선택 하셨습니다. 발송 하시겠습니까?") == true) {
				smsyn = 'Y';
			}else{
				return false;
			}
		}

		if(frm.blnS.value == 2){
			alert("이미 발송된 모든 리스트에 대해서는 발송확인 불가능합니다.");
			return;
		}
		if(confirm("발송확인 처리하시겠습니까?")){
			frm.pMode.value = "C";
			frm.action ="<%= getSCMSSLURL %>/admin/hitchhiker/processHitchhiker.asp";
			frm.submit();
		}
	}

	function jsApply(iH){
		var winApply;
		winApply = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/registHitchhiker.asp?pMode=A&iHV='+iH,'popWin','width=400, height=300');
		winApply.focus();
	}

	//파라메타에 개인정보 넘기지 말것. 키 생성후 idx 으로 넘김.		2017.07.06 한용민
	function jsReApply(iH, temp, iA, idx){
		var winApply;
		winApply = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/registHitchhiker.asp?pMode=R&iHV='+iH+'&idx='+idx+'&iAV='+iA,'popWin','width=400, height=300');
		winApply.focus();
	}

	function jsGoPage(iP){
		document.frmList.iC.value = iP;
		document.frmList.submit();
	}
	function jsLogView(iHVol,iAVol){
		var winApply;
		winApply = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/LogListHitchhiker.asp?iHV='+iHVol+'&iAV='+iAVol,'popWin','width=800, height=800');
		winApply.focus();
	}

	//파라메타에 개인정보 넘기지 말것. 키 생성후 idx 으로 넘김.		2017.07.06 한용민
	function jsAddrUPdate(iH,temp,idx){
		var winAddr;
		winAddr = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/updateHitchhikerAddr.asp?iHV='+iH+'&idx='+idx,'popWin','width=400, height=300');
		winAddr.focus();
	}

//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- 표 상단바 시작-->
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1">
<tr>
	<td width="95"><font color="gray">+ 리스트보기 :</td>
	<td> <font color="gray">히치하이커 발행회차 선택 - 신청발송회차 또는 발송여부를 선택 - 리스트보기버튼 - 해당조건에 해당하는 신청자 리스트 확인</td>
</tr>

<tr>
	<td  width="95" valign="top"><font color="gray">+ 발송확인 처리 :</td>
	<td> <font color="gray">히치하이커 발행회차 선택 - 발송여부를 미발송으로 선택 - 발송확인처리버튼 - 미발송처리건에 대해 발송확인처리<br>
		 히치하이커 발행회차 선택 - 신청발송회차 선택 - 발송확인처리버튼 - 발송여부에 상관없이 발송확인 재처리
	</td>
</tr>
<tr>
	<td width="95"><font color="gray">+ 발송신청 :</td>
	<td><font color="gray">히치하이커 발행회차 선택 - 새발송신청버튼 - 해당 발행회차의 미발송 회차로 신청처리</td>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#F3F3FF" >
	<td>
		<input type = "button" class="button" onclick="javascript:location.href='<%= getSCMSSLURL %>/admin/eventmanage/hitchhiker/index.asp';" value="히치하이커 VIP 주소입력 관리">
	</td>
</tr>
</table>

<form name="frmH" method="post" action="index.asp" style="margin:0px;">
<input type="hidden" name="chkList" value="<%=chkList%>">
<input type="hidden" name="pMode" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#F3F3FF" >
		<td rowspan="2" width="200" align="center">
			히치하이커 발행회차 Vol.
        	<select name="iHV" onChange="jsSetPage(document.frmH,'blnS');">
			<% IF isArray(arrHVol) THEN %>
				<%FOR intHV = 0 TO UBound(arrHVol,2) %>
			<option value="<%=arrHVol(0,intHV)%>" <%IF Cint(iHVol) = Cint(arrHVol(0,intHV)) THEN %> selected<%END IF%>><%=arrHVol(0,intHV)%></option>
				<%NEXT%>
			<% END IF %>
			</select>
		</td>
		<td width="180" align="right">
			신청발송 회차
		<select name="iAV" onChange="jsSetPage(document.frmH,'blnS');" style="width:75px;">
		<option value="">--선택--</option>
		<% IF isArray(arrAVol) THEN %>
			<%FOR intA = 0 TO UBound(arrAVol,2) %>
		<option value="<%=arrAVol(0,intA)%>" <%IF CInt(iAVol) = CInt(arrAVol(0,intA)) THEN %> selected<%END IF%>><%=arrAVol(0,intA)%></option>
			<%NEXT%>
		<% END IF %>
		</select>
		&nbsp;&nbsp;
		</td>
		<td rowspan="2" align="center">
	        시작일 <input id="startdate" name="startdate" value="<%=startdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /><br>
	        종료일 <input id="enddate" name="enddate" value="<%=enddate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "startdate", trigger    : "startdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "enddate", trigger    : "enddate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
		<td rowspan="2">&nbsp;&nbsp;
			<input type="button" value="리스트보기" class="a" onClick="jsGetList(document.frmH);">
		<% If C_ADMIN_AUTH or C_SYSTEM_Part or session("ssAdminPsn") = "14" or session("ssAdminPsn") = "23" or C_logics_Part or C_CriticInfoUserLV3 Then %>
			<input type="button" value="엑셀다운" class="a" onClick="jsDownList(document.frmH);">
			<br />
			문자발송 :
			<select name="smsyn" class="select">
				<option value="">-CHOICE-</option>
				<option value="Y">Y</option>
			</select>
			<input type="button" value="발송확인처리" class="a" onClick="jsSend(document.frmH);">
			<br />
		<% End If %>
			<input type="button" value="새발송신청" class="a" onClick="jsApply(<%=iHVol%>);">
			<% If chkList = "view" Then %>
			<input type="button" value="재발송Log보기" class="a" onClick="jsLogView(<%=iHVol%>,<%=iAVol%>);">
			<% End If %>
		</td>

	</tr>
	<tr bgcolor="#F3F3FF" >
		<td align="right">
		발송여부
		<!--
		<select name="blnS" onChange="jsSetPage(document.frmH,'iAV');" class="a" style="width:75px;">
		-->
		<select name="blnS" class="a" style="width:75px;">
			<option value="">--선택--</option>
			<option value="1" <%IF blnSend ="1" THEN%>selected<%END IF%>>미발송</option>
			<option value="2" <%IF blnSend ="2" THEN%>selected<%END IF%>>발송</option>
		</select>
		&nbsp;&nbsp;
		</td>
	</tr>
</table>
</form>
<!-- 표 상단바 끝-->
<br>
<div id="hlist" style="display:<%=chkView%>;">
<form name="frmList" method="post" action="<%= getSCMSSLURL %>/admin/hitchhiker/index.asp">
<input type="hidden" name="iHV" value="<%=iHVol%>">
<input type="hidden" name="iAV" value="<%=iAVol%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="chkList" value="<%=chkList%>">
<input type="hidden" name="startdate" value="<%= startdate %>">
<input type="hidden" name="enddate" value="<%= enddate %>">
<!---- 검색------------->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="30" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<select name="search">
        		<option value="">-선택-</option>
        		<option value="userid" <% If Search = "userid" Then response.write "selected" End If %>>아이디</option>
        		<option value="username" <% If Search = "username" Then response.write "selected" End If %>>이름</option>
        		<option value="receviename" <% If Search = "receviename" Then response.write "selected" End If %>>수령인</option>
        	</select>
			<input type="text" name="searchTxt" maxlength="32" value="<%=searchTxt%>">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!---- /검색------------->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>아이디</td>
	<td>이름</td>

	<% if (FALSE) then %>
	<td>수령인</td>
	<td>우편번호</td>
	<td>주소</td>
	<td>전화번호</td>
	<td>핸드폰</td>
	<% end if %>

	<td>신청일</td>
	<td>발송일</td>
	<td>재발송</td>
</tr>
<%IF isArray(arrHList) THEN %>
	<%FOR intH =0 TO UBound(arrHList,2)%>
	<tr align="center" bgcolor="ffffff">
	<td><%=iTotCnt-intH-(iPageSize*(iCurrpage-1))%></td>
	<td><%= printUserId(arrHList(3,intH), 2, "*") %></td>

	<% IF isNull(arrHList(10,intH)) THEN %>
		<td onclick="jsAddrUPdate('<%=arrHList(0,intH)%>','','<%=arrHList(13,intH)%>');" style="cursor:pointer"><%=arrHList(4,intH)%></td>
	<% else %>
		<td ><%=arrHList(4,intH)%></td>
	<% end if %>

	<% if (FALSE) then %>
	<td><%=arrHList(12,intH)%></td>
	<td><%=arrHList(5,intH)%></td>

	<% IF isNull(arrHList(10,intH)) THEN %>
		<td align="left" onclick="jsAddrUPdate('<%=arrHList(0,intH)%>','','<%=arrHList(13,intH)%>');" style="cursor:pointer"><%=arrHList(6,intH)%>&nbsp;<%=db2html(arrHList(7,intH))%></td>
	<% Else %>
		<td align="left"><%=arrHList(6,intH)%>&nbsp;<%=db2html(arrHList(7,intH))%></td>
	<% End If %>

	<td><%=arrHList(8,intH)%></td>
	<td><%=arrHList(9,intH)%></td>
	<% end if %>

	<td><%=arrHList(2,intH)%></td>
	<td><%IF isNull(arrHList(10,intH)) THEN%><font color="red">미발송</font><%ELSE%><%=arrHList(10,intH)%><%END IF%></td>
	<td>
		<% IF (isNull(arrHList(10,intH)) OR DateDiff("d", NOW, arrHList(10,intH)) > -7) and Not(C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or session("ssAdminPsn")="14" or session("ssAdminPsn")="23") THEN %>
			&nbsp;
		<% ElseIf C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or session("ssAdminPsn")="14" or session("ssAdminPsn")="23" Then %>
			<input type="button" value="신청" onClick="jsReApply(<%=iHVol%>,'','<%=iAVol%>','<%=arrHList(13,intH)%>');" class="a">
		<% Else %>
			<input type="button" value="신청" onClick="jsReApply(<%=iHVol%>,'','<%=iAVol%>','<%=arrHList(13,intH)%>');" class="a">
		<% End If %>
	</td>
	</tr>
	<%NEXT%>
	<%ELSE%>
	<tr align="center" bgcolor="ffffff">
	<td colspan="10" align="center">등록된 내용이 없습니다.</td>
	</tr>
<%END IF%>
</table>
<!-- 페이징처리 -->
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
    <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
	<% else %>[pre]<% end if %>
    <%
		for ix = iStartPage  to iEndPage
			if (ix > iTotalPage) then Exit for
			if Cint(ix) = Cint(iCurrpage) then
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
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td  background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
</form>
<!-- /페이징처리 -->
</div>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
