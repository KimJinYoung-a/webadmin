<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : PLAYing
' Hieditor : 이종화 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<%
	Dim i, cPl, vMIdx, vType, vVolNum, vTitle, vOpenDate, vState, vMoBGColor, vWorkText, vPartWDID, vPartMKID, vPartPBID
	Dim vArrDetail
	vMIdx = requestCheckVar(Request("midx"),10)
	
	If vMIdx <> "" Then
		SET cPl = New CPlay
		cPl.FRectMIdx = vMIdx
		cPl.FRectImgGubun = "1" '### 1 : 모바일리스트이미지
		cPl.sbPlayMasterDetail
		
		vVolNum = cPl.FOneItem.Fvolnum
		vTitle = cPl.FOneItem.Ftitle
		vOpenDate = cPl.FOneItem.Fstartdate
		vState = cPl.FOneItem.Fstate
		vMoBGColor = cPl.FOneItem.Fmobgcolor
		vWorkText = cPl.FOneItem.Fworktext
		vPartWDID = cPl.FOneItem.FpartWDID
		vPartMKID = cPl.FOneItem.FpartMKID
		vPartPBID = cPl.FOneItem.FpartPBID
		
		vArrDetail = cPl.FDetailList
		SET cPl = Nothing

	End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script type='text/javascript'>
function goSavePlay(){
	if(frm1.volnum.value == ""){
		alert("Vol 을 입력하세요.");
		frm1.volnum.focus();
		return;
	}
	if(isNaN(frm1.volnum.value)){
		alert("Vol 을 숫자로만 입력하세요.");
		frm1.volnum.value = "";
		frm1.volnum.focus();
		return;
	}
	if(frm1.opendate.value == ""){
		alert("오픈일을 입력하세요.");
		return;
	}
	if(frm1.state.value == ""){
		frm1.state.focus();
		alert("상태를 선택하세요.");
		return;
	}
	
	frm1.submit();
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function goPlaylist(){
	location.href = "index.asp";
}

function goNewReg(didx,cate){
	var popCorner;
	var wsize;
	
	wsize = "1200";
	popCorner = window.open('cornerwrite.asp?volnum=<%=vVolNum%>&midx=<%=vMIdx%>&didx='+didx+'&cate='+cate+'','popCorner','width='+wsize+',height=1000,scrollbars=yes,resizable=yes');
	popCorner.focus();
}

function jsPlayView(device,didx,state,sdate){
	var playVieww;
	var playsite;
	
	if(device == "w"){
		playsite = "http://<%=CHKIIF(application("Svr_Info")="Dev","2015","")%>www.10x10.co.kr";
	}else{
		playsite = "http://<%=CHKIIF(application("Svr_Info")="Dev","test","")%>m.10x10.co.kr";
	}

	playVieww = window.open(''+playsite+'/playing/view.asp?isadmin=o&didx='+didx+'&state='+state+'&sdate='+sdate+'','playVieww','width=1024, height=768, toolbar=yes, location=yes, directories=yes, status=yes, menubar=yes, scrollbars=yes, copyhistory=yes, resizable=yes');
	playVieww.focus();
}

function goThingThingUser(didx,title){
	var popThingThingUser;

	popThingThingUser = window.open('thingthing_entry_list.asp?didx='+didx+'&title='+title+'','popThingThingUser','width=750,height=850,scrollbars=yes,resizable=yes');
	popThingThingUser.focus();
}

function goPlaylistUser(didx){
	var popPlaylistUser;

	popPlaylistUser = window.open('playlist_comment_list.asp?didx='+didx+'','popPlaylistUser','width=750,height=900,scrollbars=yes,resizable=yes');
	popPlaylistUser.focus();
}
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>[ON] PLAY &gt; <strong>PLAYing</strong> vol <%=CHKIIF(vMIdx<>"","수정","등록")%></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="1836">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l 
			<!-- 마스터이상 메뉴권한 설정 //-->
			<a href="Javascript:PopMenuEdit('1836');">권한변경</a> l 
			<!-- Help 설정 //-->
			<a href="Javascript:PopMenuHelp('1836');">HELP</a>
		</div>
	</div>
	
	<div class="searchWrap">
	<form name="frm1" action="volproc.asp" method="post" style="margin:0px;">
	<input type="hidden" name="action" value="<%=CHKIIF(vMIdx="","insert","update")%>">
	<input type="hidden" name="midx" value="<%=vMIdx%>">
	<table class="tbType1 writeTb" bgcolor="#FFFFFF">
		<tbody>
			<tr>
				<th width="15%">Vol.</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="volnum" value="<%=vVolNum%>" size="10" maxlength="3"> * 1 부터 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">타이틀</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="title" value="<%= ReplaceBracket(vTitle) %>" size="30" maxlength="96"> * ex. 2016.10.10 - 10.20
				</td>
			</tr>
			<tr>
				<th width="15%">오픈일</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="opendate" value="<%=vOpenDate%>" onClick="jsPopCal('opendate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
				</td>
			</tr>
			<tr>
				<th width="15%">상 태</th>
				<td height="30" style="padding-left:5px;">
					<select name="state" class="formSlt">
						<%=fnStateSelectBox("select",vState)%>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">BG 컬러</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #을 제외한 6자 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">담당자</th>
				<td height="30" style="padding-left:5px;">
					<select name="partmkid" >
						<option value="">선택</option>
						<option value="shaeiou" <%=CHKIIF(vPartMKID="shaeiou","selected","")%>>김시화</option>
						<option value="ascreem" <%=CHKIIF(vPartMKID="ascreem","selected","")%>>남찬</option>
						<option value="sss162000" <%=CHKIIF(vPartMKID="sss162000","selected","")%>>손아름</option>
						<option value="madebyash" <%=CHKIIF(vPartMKID="madebyash","selected","")%>>안서연</option>
						<option value="heejong1013" <%=CHKIIF(vPartMKID="heejong1013","selected","")%>>최희종</option>
						<option value="ppono2" <%=CHKIIF(vPartMKID="ppono2","selected","")%>>한유민</option>
						<option value="torymilk" <%=CHKIIF(vPartMKID="torymilk","selected","")%>>이서영</option>
						<option value="spinel93" <%=CHKIIF(vPartMKID="spinel93","selected","")%>>이수진</option>
						<option value="dhalsdud57" <%=CHKIIF(vPartMKID="dhalsdud57","selected","")%>>오민영</option>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">작업자</th>
				<td height="30" style="padding-left:5px;">
					WD:<% sbGetpartid "partwdid",vPartWDID,"","12" %>
					&nbsp;&nbsp;&nbsp;
					퍼블리셔:
					<select name="partpbid">
						<option value="">선택</option>
						<option value="happyngirl" <%=CHKIIF(vPartPBID="happyngirl","selected","")%>>최선미</option>
						<option value="kyungae13" <%=CHKIIF(vPartPBID="kyungae13","selected","")%>>조경애</option>
						<option value="jinyeonmi" <%=CHKIIF(vPartPBID="jinyeonmi","selected","")%>>진연미</option>
						<option value="jj999a" <%=CHKIIF(vPartPBID="jj999a","selected","")%>>김송이</option>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">작업 전달 사항</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<textarea name="worktext" rows="10" cols="70"><%= ReplaceBracket(vWorkText) %></textarea>
				</td>
			</tr>
		</tbody>
	</table>
	<table width="100%">
	<tr>
		<td style="padding-top:5px;float:left;"><input type="button" style="width:100px;height:30px;" value="리스트로" onClick="goPlaylist();" /></td>
		<td style="padding-top:5px;float:right;"><input type="button" style="width:100px;height:30px;" value="저 장" onClick="goSavePlay();" /></td>
	</tr>
	</table>
	</form>
	</div>
	
	<% If vMIdx <> "" Then %>
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt"><strong>* idx 높은 숫자로 정렬되어 있습니다. 등록시 순서를 고려해주세요. 오픈일 기준으로 하니 너무 느려서 idx 값 기준으로 했습니다.</strong></div>
				<div class="ftRt">
					<p class="btn2 cBk1 ftLt"><a href="javascript:goNewReg('','');"><span class="eIcon"><em class="fIcon">신규등록</em></span></a></p>
				</div>
			</div>
			<div class="tPad15">
				<table class="tbType1 listTb">
					<thead>
					<tr>
						<th><div>idx</div></th>
						<th><div>코너</div></th>
						<th><div>M 리스트이미지</div></th>
						<th><div>타이틀</div></th>
						<th><div>상태</div></th>
						<th><div>작업자</div></th>
						<th><div>View Count</div></th>
						<th><div></div></th>
						<th><div>미리보기<br>(<strong>미리 로그인必</strong>)</div></th>
					</tr>
					</thead>
					<tbody>
					<% IF isArray(vArrDetail) THEN
							'd.didx, d.cate, d.title, d.startdate, d.state, imgurl, linkurl, catename, partWDname, partPBname
							For i = 0 To UBound(vArrDetail,2)
					%>
							<tr>
								<td><%=vArrDetail(0,i)%></td>
								<td><%=fnPlayCateName(vArrDetail(1,i))%></td>
								<td>
									<%
										If vArrDetail(1,i) = "3" OR vArrDetail(1,i) = "41" OR vArrDetail(1,i) = "42" OR vArrDetail(1,i) = "43" Then
											If vArrDetail(5,i) <> "" Then
												Response.Write "<img src='" & vArrDetail(5,i) & "' width='50'>"
											End If
										Else
											If vArrDetail(1,i) <> "5" Then
												Response.Write "<img src='" & fnPlayImage(vArrDetail(0,i),vArrDetail(1,i),"11","","","i") & "' width='50'>"
											End If
										End If
									%>
								</td>
								<td><%=vArrDetail(2,i)%></td>
								<td><%=fnStateSelectBox("one",vArrDetail(4,i))%><br />실제오픈일 : <strong><%=vArrDetail(3,i)%></strong></td>
								<td>
									WD:<%=vArrDetail(8,i)%><br />
									PB:<%=vArrDetail(9,i)%>
								</td>
								<td>W:<%=vArrDetail(10,i)%>, M:<%=vArrDetail(11,i)%>, A:<%=vArrDetail(12,i)%></td>
								<td>
									<input type="button" onClick="goNewReg('<%=vArrDetail(0,i)%>','<%=vArrDetail(1,i)%>');" value="수 정">
									<% If vArrDetail(1,i) = "1" Then %>
									&nbsp;<input type="button" onClick="goPlaylistUser('<%=vArrDetail(0,i)%>');" value="List">
									<% End If %>
									<% If vArrDetail(1,i) = "42" Then %>
									&nbsp;<input type="button" onClick="goThingThingUser('<%=vArrDetail(0,i)%>','<%=Server.URLencode(vArrDetail(2,i))%>');" value="List">
									<% End If %>
								</td>
								<td>
									<input type="button" onClick="jsPlayView('w','<%=vArrDetail(0,i)%>','<%=vArrDetail(4,i)%>','<%=vArrDetail(3,i)%>');" value="W">&nbsp;
									<input type="button" onClick="jsPlayView('m','<%=vArrDetail(0,i)%>','<%=vArrDetail(4,i)%>','<%=vArrDetail(3,i)%>');" value="M">
								</td>
							</tr>
					<% 
						Next
					Else
						Response.Write "<tr><td colspan='9' align='center'>등록된게 없네요~</td></tr>"
					End If %>
					</tbody>
				</table>
			</div>
		</div>
	<% End If %>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->