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
' Hieditor : ����ȭ ����
'			 2022.07.07 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
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
		cPl.FRectImgGubun = "1" '### 1 : ����ϸ���Ʈ�̹���
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
		alert("Vol �� �Է��ϼ���.");
		frm1.volnum.focus();
		return;
	}
	if(isNaN(frm1.volnum.value)){
		alert("Vol �� ���ڷθ� �Է��ϼ���.");
		frm1.volnum.value = "";
		frm1.volnum.focus();
		return;
	}
	if(frm1.opendate.value == ""){
		alert("�������� �Է��ϼ���.");
		return;
	}
	if(frm1.state.value == ""){
		frm1.state.focus();
		alert("���¸� �����ϼ���.");
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
		<div class="locate"><h2>[ON] PLAY &gt; <strong>PLAYing</strong> vol <%=CHKIIF(vMIdx<>"","����","���")%></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="1836">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">���ã��</a> l 
			<!-- �������̻� �޴����� ���� //-->
			<a href="Javascript:PopMenuEdit('1836');">���Ѻ���</a> l 
			<!-- Help ���� //-->
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
					<input type="text" name="volnum" value="<%=vVolNum%>" size="10" maxlength="3"> * 1 ���� ���ڷθ� �Է��ϼ���.
				</td>
			</tr>
			<tr>
				<th width="15%">Ÿ��Ʋ</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="title" value="<%= ReplaceBracket(vTitle) %>" size="30" maxlength="96"> * ex. 2016.10.10 - 10.20
				</td>
			</tr>
			<tr>
				<th width="15%">������</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="opendate" value="<%=vOpenDate%>" onClick="jsPopCal('opendate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
				</td>
			</tr>
			<tr>
				<th width="15%">�� ��</th>
				<td height="30" style="padding-left:5px;">
					<select name="state" class="formSlt">
						<%=fnStateSelectBox("select",vState)%>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">BG �÷�</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #�� ������ 6�� ���ڷθ� �Է��ϼ���.
				</td>
			</tr>
			<tr>
				<th width="15%">�����</th>
				<td height="30" style="padding-left:5px;">
					<select name="partmkid" >
						<option value="">����</option>
						<option value="shaeiou" <%=CHKIIF(vPartMKID="shaeiou","selected","")%>>���ȭ</option>
						<option value="ascreem" <%=CHKIIF(vPartMKID="ascreem","selected","")%>>����</option>
						<option value="sss162000" <%=CHKIIF(vPartMKID="sss162000","selected","")%>>�վƸ�</option>
						<option value="madebyash" <%=CHKIIF(vPartMKID="madebyash","selected","")%>>�ȼ���</option>
						<option value="heejong1013" <%=CHKIIF(vPartMKID="heejong1013","selected","")%>>������</option>
						<option value="ppono2" <%=CHKIIF(vPartMKID="ppono2","selected","")%>>������</option>
						<option value="torymilk" <%=CHKIIF(vPartMKID="torymilk","selected","")%>>�̼���</option>
						<option value="spinel93" <%=CHKIIF(vPartMKID="spinel93","selected","")%>>�̼���</option>
						<option value="dhalsdud57" <%=CHKIIF(vPartMKID="dhalsdud57","selected","")%>>���ο�</option>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">�۾���</th>
				<td height="30" style="padding-left:5px;">
					WD:<% sbGetpartid "partwdid",vPartWDID,"","12" %>
					&nbsp;&nbsp;&nbsp;
					�ۺ���:
					<select name="partpbid">
						<option value="">����</option>
						<option value="happyngirl" <%=CHKIIF(vPartPBID="happyngirl","selected","")%>>�ּ���</option>
						<option value="kyungae13" <%=CHKIIF(vPartPBID="kyungae13","selected","")%>>�����</option>
						<option value="jinyeonmi" <%=CHKIIF(vPartPBID="jinyeonmi","selected","")%>>������</option>
						<option value="jj999a" <%=CHKIIF(vPartPBID="jj999a","selected","")%>>�����</option>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">�۾� ���� ����</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<textarea name="worktext" rows="10" cols="70"><%= ReplaceBracket(vWorkText) %></textarea>
				</td>
			</tr>
		</tbody>
	</table>
	<table width="100%">
	<tr>
		<td style="padding-top:5px;float:left;"><input type="button" style="width:100px;height:30px;" value="����Ʈ��" onClick="goPlaylist();" /></td>
		<td style="padding-top:5px;float:right;"><input type="button" style="width:100px;height:30px;" value="�� ��" onClick="goSavePlay();" /></td>
	</tr>
	</table>
	</form>
	</div>
	
	<% If vMIdx <> "" Then %>
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt"><strong>* idx ���� ���ڷ� ���ĵǾ� �ֽ��ϴ�. ��Ͻ� ������ ������ּ���. ������ �������� �ϴ� �ʹ� ������ idx �� �������� �߽��ϴ�.</strong></div>
				<div class="ftRt">
					<p class="btn2 cBk1 ftLt"><a href="javascript:goNewReg('','');"><span class="eIcon"><em class="fIcon">�űԵ��</em></span></a></p>
				</div>
			</div>
			<div class="tPad15">
				<table class="tbType1 listTb">
					<thead>
					<tr>
						<th><div>idx</div></th>
						<th><div>�ڳ�</div></th>
						<th><div>M ����Ʈ�̹���</div></th>
						<th><div>Ÿ��Ʋ</div></th>
						<th><div>����</div></th>
						<th><div>�۾���</div></th>
						<th><div>View Count</div></th>
						<th><div></div></th>
						<th><div>�̸�����<br>(<strong>�̸� �α�����</strong>)</div></th>
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
								<td><%=fnStateSelectBox("one",vArrDetail(4,i))%><br />���������� : <strong><%=vArrDetail(3,i)%></strong></td>
								<td>
									WD:<%=vArrDetail(8,i)%><br />
									PB:<%=vArrDetail(9,i)%>
								</td>
								<td>W:<%=vArrDetail(10,i)%>, M:<%=vArrDetail(11,i)%>, A:<%=vArrDetail(12,i)%></td>
								<td>
									<input type="button" onClick="goNewReg('<%=vArrDetail(0,i)%>','<%=vArrDetail(1,i)%>');" value="�� ��">
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
						Response.Write "<tr><td colspan='9' align='center'>��ϵȰ� ���׿�~</td></tr>"
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