<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �÷��̸����
' Hieditor : ����ȭ ����
'			 2022.07.07 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->
<%
	Dim cPlay, vIdx, vType, vIsUsing, vColorCD, vState, vViewNo, vTitle, vStartDate, vListImg, vContents, vRegdate, vLastUpdate, vLastAdminId
	Dim vPartWDID, vPartMDID, vPartPBID, vIsComment, vStyle, vSubCopy, vWorkComm, vViewNoTxt, vContentsIdx, vURLl, vSortNo
	vIdx = requestCheckVar(Request("idx"),10)
	vType = requestCheckVar(Request("type"),10)
	
	If vIdx <> "" Then
		SET cPlay = New CPlayMoContents
		cPlay.FRectIdx = vIdx
		cPlay.sbPlayMoDetail
		
		vViewNo = cPlay.FOneItem.Fviewno
		vViewNoTxt = cPlay.FOneItem.Fviewnotxt
		vType = cPlay.FOneItem.Ftype
		vState = cPlay.FOneItem.Fstate
		vTitle = cPlay.FOneItem.Ftitle
		vSubCopy = cPlay.FOneItem.Fsubcopy
		vStartDate = cPlay.FOneItem.Fstartdate
		vIsUsing = cPlay.FOneItem.Fisusing
		vIsComment = cPlay.FOneItem.Fiscomment
		vListImg = cPlay.FOneItem.Flistimg
		vContents = cPlay.FOneItem.Fcontents
		vColorCD = cPlay.FOneItem.Fcolorcd
		vRegdate = cPlay.FOneItem.Fregdate
		vLastUpdate = cPlay.FOneItem.Flastupdate
		vLastAdminId = cPlay.FOneItem.Flastadminid
		vPartWDID = cPlay.FOneItem.FpartWDID
		vPartMDID = cPlay.FOneItem.FpartMDID
		vPartPBID = cPlay.FOneItem.FpartPBID
		vStyle = cPlay.FOneItem.Fstyle
		vWorkComm = cPlay.FOneItem.Fworkcomm
		vContentsIdx = cPlay.FOneItem.Fcontents_idx
		vSortNo = cPlay.FOneItem.Fsortno
		SET cPlay = Nothing
		
		If CStr(vType) <> CStr(Request("type")) AND Request("type") <> "" Then
			vType = requestCheckVar(Request("type"),10)
		End If
	End If
	
	vIsUsing = NullFillWith(vIsUsing,"Y")
	vIsComment = NullFillWith(vIsComment,"N")
	
	If vType = "" Then
		vType = ""
	End If
	
	IF vPartMDID = "" Then
		vPartMDID = session("ssBctId")
	End If
	
	Select Case vType
		Case "1" : vURLl = "http://www.10x10.co.kr/play/playGround.asp?"
		Case "2" : vURLl = "http://www.10x10.co.kr/play/playStylePlusView.asp?idx="
		Case "3" : vURLl = ""
		Case "4" : vURLl = "http://www.10x10.co.kr/play/playdesignfingers.asp?fingerid="
		Case "5" : vURLl = "http://www.10x10.co.kr/play/playPicDiary.asp?"
		Case "6" : vURLl = "http://www.10x10.co.kr/play/playVideoClip.asp?idx="
		Case "7" : vURLl = "http://www.10x10.co.kr/play/playtEpisodePhotopick.asp?"
	End Select
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
body {font:9pt/135% "dotum";color:#000000}
.tbType1 {width:100%;}
.tbType1 th, .tbType1 td {color:#444;}
.tbType1 th {background-color:#eaeaea;}
.tbType1 th a, .tbType1 td a {color:#444;}
.tbType1 th a:hover, .tbType1 td a:hover {text-decoration:underline;}


.writeTb {border-top:2px solid #b9b9b9; border-bottom:2px solid #b9b9b9;}
.writeTb th, .writeTb td {border-bottom:1px solid #c9c9c9; vertical-align:middle;}
.writeTb th {font-weight:bold; text-align:center;}
.writeTb th div {padding:9px 10px 7px 10px; vertical-align:middle;}


.btn2 {display:inline-block; font-size:11px !important; letter-spacing:-0.025em; line-height:110%; border-left:1px solid #f0f0f0; border-top:1px solid #f0f0f0; border-right:1px solid #cdcdcd; border-bottom:1px solid #cdcdcd; background-color:#f2f2f2; background-image:-webkit-linear-gradient(#fff, #e1e1e1); background-image:-moz-linear-gradient(#fff, #e1e1e1); background-image:-ms-linear-gradient(#fff, #e1e1e1); background-image:linear-gradient(#fff, #e1e1e1); text-align:center; cursor:pointer;}
.cBk1, .cBk1 a {color:#000 !important;}
.ftLt {float:left;}
</style>
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<% if session("sslgnMethod")<>"S" then %>
<!-- USBŰ ó�� ���� (2008.06.23;������) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USBŰ ó�� �� -->
<% end if %>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

function goSavePlay(){
	if(frm1.playtype.value == ""){
		alert("�з��� �����ϼ���.");
		return;
	}
	if(frm1.viewno.value == ""){
		alert("# No.�� �Է��ϼ���.");
		return;
	}
	if(frm1.contentsidx.value == ""){
		alert("<%=fnTypeSelectBox("one",vType,"Y")%> ����Ʈ�� idx �Ǵ� ��ȣ ���� �Է��ϼ���.");
		return;
	}
	if(frm1.title.value == ""){
		alert("������ �Է��ϼ���.");
		return;
	}
	if(frm1.sortno.value == ""){
		alert("�켱������ȣ�� �Է��ϼ���.");
		return;
	}

	frm1.submit();
}
function goTypeOnChange(a){
	location.href = "write.asp?idx=<%=vIdx%>&type="+a+"";
}

//�����ڵ� ����
function selColorChip(cd) {
	var i;
	document.frm1.colorcd.value= cd;
	for(i=0;i<=30;i++) {
		document.all("cline"+i).bgColor='#DDDDDD';
	}
	if(!cd) document.all("cline0").bgColor='#DD3300';
	else document.all("cline"+cd).bgColor='#DD3300';
}

function jsTagview(gidx , idx){	
	var poptagm;
	poptagm = window.open('pop_tagReg.asp?idx='+gidx+'&playcate=<%=vType%>','poptagm','width=500,height=400,scrollbars=yes,resizable=yes');
	poptagm.focus();
}

function jsUploadImg(a,b){
	document.domain ="10x10.co.kr";
	var popupl;
	popupl = window.open('/admin/sitemaster/play_mo/pop_uploadimg.asp?type=<%=vType%>&folder='+a+'&span='+b,'popupl','width=370,height=150');
	popupl.focus();
}

//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

function jsManagePlayImage(){
    var playManageDir = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","test","")%>upload.10x10.co.kr/linkweb/play/playManageDir.asp?idx=<%=vIdx%>','playManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
    playManageDir.focus();
}

function confirmidx(){
	var wconfirmidx;
	var idx = frm1.contentsidx.value;
	wconfirmidx = window.open('<%=vURLl%>'+idx,'wconfirmidx','width=1000,height=1000,toolbar=yes, location=yes, directories=yes, status=yes, menubar=yes, scrollbars=yes, copyhistory=yes, resizable=yes');
	wconfirmidx.focus();
}
</script>
</head>
<body TOPMARGIN="0" <% if session("sslgnMethod")<>"S" then %>onload="checkUSBKey()"<% end if %>>
<br /><font size="5" color="red"><strong>* ����� PLAY ����Ʈ ���� ��</strong></font> (play ����Ʈ���� ������ ����)<br /><br />
<form name="frm1" action="proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="action" value="<%=CHKIIF(vIdx="","insert","update")%>">
<input type="hidden" name="idx" value="<%=vIdx%>">
<table class="tbType1 writeTb">
	<tbody>
		<% If vIdx <> "" Then %>
		<tr>
			<th width="15%">Idx</th>
			<td height="25"><%=vIdx%></td>
		</tr>
		<% End If %>
		<tr>
			<th>�� ��</th>
			<td height="25">
				<select name="playtype" class="formSlt" onChange="goTypeOnChange(this.value);">
					<%=fnTypeSelectBox("select",vType,"Y")%>
				</select>
			</td>
		</tr>
		<% If vType <> "" Then %>
			<tr>
				<th width="15%">No | �ؽ�Ʈ</th>
				<td height="25">
					<input type="text" name="viewno" value="<%= ReplaceBracket(vViewNo) %>" style="width:7%;" maxlength="20" />&nbsp;&nbsp;|&nbsp;
					<input type="text" name="viewnotxt" value="<%= ReplaceBracket(vViewNoTxt) %>" style="width:20%;" maxlength="20" />
				</td>
			</tr>
			<tr>
				<th width="15%"><%=fnTypeSelectBox("one",vType,"Y")%>�� �۹�ȣ</th>
				<td height="25"><input type="text" name="contentsidx" value="<%=vContentsIdx%>" style="width:10%;" maxlength="10" />
				<% If vType <> "1" Then %>
					[<a href="javascript:confirmidx();">�۹�ȣȮ��</a>]
					<font color="blue"><strong>* [ON] Play �޴��� <%=fnTypeSelectBox("one",vType,"Y")%> ����Ʈ�� idx �Ǵ� ��ȣ ��. �ݵ�� �Է��ؾ���.</strong></font>
				<% End If %>
				</td>
			</tr>
			<tr>
				<th>�� ��</th>
				<td height="25"><input type="text" name="title" value="<%= ReplaceBracket(vTitle) %>" style="width:100%;" maxlength="50" /></td>
			</tr>
			<tr>
				<th>����ī��</th>
				<td height="25"><input type="text" name="subcopy" value="<%= ReplaceBracket(vSubCopy) %>" style="width:100%;" maxlength="100" /></td>
			</tr>
			<tr>
				<th>������</th>
				<td height="25">
					<input type="text" id="startdate" name="startdate" value="<%=vStartDate%>" style="width:100px" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="startdate_trigger" alt="�޷����� �˻�" style="cursor:pointer;" />
					<script>
						var CAL_Start = new Calendar({
							inputField : "startdate", trigger    : "startdate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</td>
			</tr>
			<tr>
				<th>�����</th>
				<td height="25">
					<% sbGetpartid "partmdid",vPartMDID,"","11,14,21,22,23" %>
				</td>
			</tr>
			<tr>
				<th>�۾���</th>
				<td height="25">
					WD:<% sbGetpartid "partwdid",vPartWDID,"","12" %>
					&nbsp;&nbsp;&nbsp;
					�ۺ���:
					<select name="partpbid">
						<option value="">����</option>
						<option value="happyngirl" <%=CHKIIF(vPartPBID="happyngirl","selected","")%>>�ּ���</option>
						<option value="kyungae13" <%=CHKIIF(vPartPBID="kyungae13","selected","")%>>�����</option>
						<option value="jinyeonmi" <%=CHKIIF(vPartPBID="jinyeonmi","selected","")%>>������</option>
					</select>
				</td>
			</tr>
			<tr>
				<th>�� ��</th>
				<td height="25">
					<select name="state" class="formSlt">
						<%=fnStateSelectBox("select",vState)%>
					</select>
					�� ������ �ؼ� �����Ͽ��� ������ =< ���� �̾�߸� ������ �˴ϴ�. 
				</td>
			</tr>
			<tr>
				<th>����Ʈ �̹���</th>
				<td>
					<input type="button" value="�̹������" onClick="jsUploadImg('listimg','listimgspan');" /><br /><br />
					<span id="listimgspan" style="padding:5px 5px 5px 0;"><%
						If vListImg <> "" Then
							Response.Write "<img src='" & vListImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vListImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('listimg','listimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br /><br />
					�̹����ּ� : <%=vListImg%>
					<input type="hidden" name="listimg" value="<%=vListImg%>">
				</td>
			</tr>
			<tr>
				<th width="15%">�켱������ȣ</th>
				<td height="25">
					<input type="text" name="sortno" value="<%=vSortNo%>" style="width:7%;" maxlength="10" />
					�� ���� �����ϼ��� ��ܿ� �ö󰩴ϴ�.
				</td>
			</tr>
			<tr>
				<th>��뿩��</th>
				<td height="25">
					<input type="radio" name="isusing" value="Y" <%=CHKIIF(vIsUsing="Y","checked","")%> /> Y&nbsp;&nbsp;&nbsp;
					<input type="radio" name="isusing" value="N" <%=CHKIIF(vIsUsing="N","checked","")%> /> N
				</td>
			</tr>
			<!--
			<tr>
				<th>�ڸ�Ʈ �Խ���<br />��뿩��</th>
				<td height="25">
					<input type="radio" name="iscomment" value="Y" <%=CHKIIF(vIsComment="Y","checked","")%> /> Y&nbsp;&nbsp;&nbsp;
					<input type="radio" name="iscomment" value="N" <%=CHKIIF(vIsComment="N","checked","")%> /> N
				</td>
			</tr>
			<% If vType = "2" Then %>
			<tr>
				<th>�÷�����</th>
				<td>
					<input type="hidden" name="colorcd" value="<%= vColorCD %>">
					<%=FnSelectColorBar(vColorCD,32)%>
				</td>
			</tr>
			<tr>
				<th>��Ÿ�ϼ���</th>
				<td>
					<select name="playstyle" class="formSlt">
						<%=fnStyleSelectBox("select",vStyle,"Y")%>
					</select>
				</td>
			</tr>
			<% End If %>
			<tr>
				<th>Tag</th>
				<td height="25">
					<% If vIdx <> "" Then %>
					<input type="button" name="btnviewImg" value="�±� ����" onClick="jsTagview('<%=vIdx%>', '')" class="button"/>
					�� �±װ����� �˾����� ���� �մϴ� ���� ��� ���ּ���.��
					<% Else %>
					�� �±״� ������ ��ϵǾ�� �Է� �����մϴ�.(DB�� ����� �۹�ȣ�� �־�� ������ ����)
					<% End If %>
				</td>
			</tr>
			//-->
			<tr>
				<th>�۾����޻���</th>
				<td><textarea name="workcomment" style="width:100%; height:120px;"><%= ReplaceBracket(vWorkComm) %></textarea></td>
			</tr>
			<!--
			<tr>
				<th>�� ��(html)</th>
				<td><textarea name="contents" style="width:100%; height:400px;"><%= ReplaceBracket(vContents) %></textarea></td>
			</tr>
			//-->
		<% End If %>
	</tbody>
</table>
<table width="100%">
<tr>
	<td style="padding-top:5px;float:right;"><input type="button" style="width:100px;height:30px;" value="�� ��" onClick="goSavePlay();" /></td>
</tr>
</table>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->