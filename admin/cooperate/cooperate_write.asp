<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<%
'###########################################################
' Description : 업무협조
' Hieditor : 강준구 생성
'			 2023.05.22 한용민 수정(권한 체크 추가. 남의글 못보게)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateUtf8Cls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->

<%
	Dim iTotCnt, arrList, intLoop, arrFileList, i
	Dim iPageSize, iCurrentpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, sDoc_ViewList, sDoc_ViewListRef, sDoc_ReferView
	Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Start, sDoc_End, sDoc_Type, sDoc_Import, sDoc_Diffi, sDoc_Subj, sDoc_Content
	Dim sDoc_WorkerName, sDoc_Worker, sDoc_UseYN, sDoc_Regdate, sDoc_WorkerView, sDoc_Refer, sDoc_ReferName, vIsPop
	Dim sDoc_reportidx ,sDoc_reportstate,sSys_reportidx  ,sSys_reportstate

	vIsPop			= Request("ispop")
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)

	If iDoc_Idx = "" Then
		sDoc_Id 		= session("ssBctId")
		sDoc_Name		= session("ssBctCname")
		sDoc_Regdate	= Left(now(),10)
		sDoc_WorkerName	= ""
		sDoc_Worker		= ""
		sDoc_ReferView	= ""
		sDoc_Refer		= ""
		sDoc_Status		= "1"
	Else
		'####### 맨 처음 확인 날짜 저장 #######
		Call WorkerView(iDoc_Idx)
		Call ReferView(iDoc_Idx)
		'####### 맨 처음 확인 날짜 저장 #######

		Dim cooperateView, cooperateFile
		Set cooperateView = New CCooperate
		cooperateView.FDoc_Idx = iDoc_Idx
		cooperateView.fnGetCooperateView

		sDoc_Id 		= cooperateView.FDoc_Id
		sDoc_Name		= cooperateView.FDoc_Name
		sDoc_Status		= cooperateView.FDoc_Status
		sDoc_Start		= cooperateView.FDoc_Start
		sDoc_End		= cooperateView.FDoc_End
		sDoc_Type		= cooperateView.FDoc_Type
		sDoc_Import		= cooperateView.FDoc_Import
		sDoc_Diffi		= cooperateView.FDoc_Diffi
		sDoc_Subj		= cooperateView.FDoc_Subj
		sDoc_Content	= cooperateView.FDoc_Content
		sDoc_UseYN		= cooperateView.FDoc_UseYN
		sDoc_Regdate	= cooperateView.FDoc_Regdate
		sDoc_WorkerName	= cooperateView.FDoc_WorkerName
		sDoc_Worker		= cooperateView.FDoc_Worker
		sDoc_WorkerView	= cooperateView.FDoc_WorkerViewdate
		sDoc_Refer		= cooperateView.FDoc_Refer
		sDoc_ReferName	= cooperateView.FDoc_ReferName
		sDoc_ReferView	= cooperateView.FDoc_ReferViewdate
		sDoc_reportidx  = cooperateView.FDoc_reportidx
		sDoc_reportstate= cooperateView.FDoc_reportstate
		sSys_reportidx  = cooperateView.FSys_reportidx
		sSys_reportstate= cooperateView.FSys_reportstate

		set cooperateFile = new CCooperate
	 	cooperateFile.FDoc_Idx = iDoc_Idx
		arrFileList = cooperateFile.fnGetFileList

		For i=0 To UBOUND(Split(sDoc_WorkerName,","))
			if Not(sDoc_WorkerView="" or isNull(sDoc_WorkerView)) then
				'확인일이 없는경우 Pass (2009.06.03;허진원)
				'sDoc_ViewList = sDoc_ViewList & "&nbsp;" & Split(sDoc_WorkerName,",")(i) & " : " & Split(sDoc_WorkerView,",")(i) & "<br>"
				sDoc_ViewList = Split(sDoc_WorkerView,",")(i)
			end if
		Next

		For i=0 To UBOUND(Split(sDoc_ReferName,","))
			if Not(sDoc_ReferView="" or isNull(sDoc_ReferView)) then
				'확인일이 없는경우 Pass (2009.06.03;허진원)
				sDoc_ViewListRef = sDoc_ViewListRef & "&nbsp;" & Split(sDoc_ReferName,",")(i) & " : " & Split(sDoc_ReferView,",")(i) & "<br>"
			end if
		Next

		'//현재 페이지를 요청한 유저가 등록자인지 확인
		If Trim(sDoc_Id) <> Trim(session("ssBctId")) Then
			Response.Write "<script>alert('협조문 등록자만 수정할 수 있습니다.');history.back();</script>"
			Response.End
		End If

	End If
%>

<html>
<head>
		 <!-- daumeditor head -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=10" />
<link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="utf-8"/>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="utf-8"></script>
<script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="utf-8"></script>
<!-- daumeditor  -->
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript">
    var config = {
        initializedId: "",
        wrapper: "tx_trex_container",
        form: 'frm',
        txIconPath: "/lib/util/daumeditor/images/icon/editor/",
        txDecoPath: "/lib/util/daumeditor/images/deco/contents/",
        events: {
            preventUnload: false
        },
        sidebar: {
            attachbox: {
                show: true
            },
            attacher: {
                 image: {
                    popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
                }
            }
        }
    };
</script>
<!-- //daumeditor head -->

<script type="text/javascript" >
var openWorker = null;

function workerlist()
{
	var worker = frm.doc_worker.value;
	var refer = frm.doc_refer.value;
	openWorker = window.open('PopWorkerList.asp?worker='+worker+'&refer='+refer+'&didx=<%=iDoc_Idx%>','openWorker','width=570,height=570,scrollbars=yes');
	openWorker.focus();
}

function fileupload()
{
	window.open('popUpload.asp','worker','width=420,height=200,scrollbars=yes');
}

function clearRow(tdObj) {
	if(confirm("선택하신 파일을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}

var blockChar=["&lt;script","<scrip"];
 function chkContent(p) {
 for (var i=0; i<blockChar.length; i++) {
  if (p.indexOf(blockChar[i])>=0) {
   return blockChar[i];
  }
 }
 return null;
}

function checkform()
{
	var frm= document.frm;
	if (frm.doc_start.value == "")
	{
		alert("처리기간의 시작일을 입력하세요!");
		return ;
	}
	if (frm.doc_end.value == "")
	{
		alert("처리기간의 종료일을 입력하세요!");
		return ;
	}
	if (frm.doc_type.value == "")
	{
		alert("업무 구분을 선택해 주세요!");
		return ;
	}

	count = 0;
	num = frm.doc_important.length;

	for(i=0; i<num; i++)
	{
		if(frm.doc_important[i].checked == true)
		{
			count +=1;
		}
	}
	if(count==0)
	{
		alert("업무 중요도를 선택해 주세요!");
		return ;
	}
<%
'	count = 0;
'	num = frm.doc_difficult.length;
'
'	for(i=0; i<num; i++)
'	{
'		if(frm.doc_difficult[i].checked == true)
'		{
'			count +=1;
'		}
'	}
'	if(count==0)
'	{
'		alert("업무 난이도를 선택해 주세요!");
'		return false;
'	}
%>
	if (frm.doc_worker.value == "")
	{
		alert("작업자를 선택해 주세요!");
		return ;
	}

	if (frm.doc_subject.value == "")
	{
		alert("제목을 입력해 주세요!");
		frm.doc_subject.focus();
		return ;
	}

	if (frm.doc_important[0].checked&&!(frm.sms_send.checked)) {
	    alert("업무중요도가 [긴급]일때는 SMS 발송을 체크해야만 됩니다.");
	    return ;
	}

	if(frm.doc_important[0].checked == true) {
		if(!confirm("업무중요도를 [긴급]을 선택하셨습니다.\n정말 긴급한 상황이 맞습니까?\n\n※긴급한 상황이 아니라면 [빠른시일내]를 선택해주세요.")) {
			return ;
		}
	}

	//daum editor start---------
		var content = Editor.getContent();
		var str = chkContent(content);
		  if (str) {
		   alert("script태그는 사용할 수 없는 문자열 입니다.");
		   return ;
		  }

     document.getElementById("doc_content").value = content;
 //daum editor end -----------

	if(!(openWorker == null))
	{
		if(!(openWorker.closed))
		{
			openWorker.close();
		}
	}

	document.frm.submit();
}

function filedownload(idx)
{
	filefrm.file_idx.value = idx;
	filefrm.submit();
}

function cooperate_del()
{
	if(confirm("현재 협조문을 삭제하시겠습니까?") == true) {
		frm.doc_useyn.value = "N";
		frm.submit();
	} else {
		return false;
	}
}

function cooperate_status_change(v,m)
{
	if(confirm(""+m+"을 재요청하시겠습니까?") == true) {
		frm.doc_status.value = v;
		frm.submit();
	} else {
		return false;
	}
}

function issystem(value)
{
	if(value == "3")
	{
		document.getElementById("onlysystem").innerHTML = "<br><br><font color='red'><b>※ PC 신청, 업그레이드, 수리, POS 등 모든 장비 관련 문의는 시스템장애신청에서 하시기 바랍니다.</b></font>]";
		<%if iDoc_Idx <> "" THEN%>
		document.all.divEappN.style.display="";
	  document.all.divEappC.style.display="";
	  <%end if%>
	}
	else
	{
		document.getElementById("onlysystem").innerHTML = "";
		<%if iDoc_Idx <> "" THEN%>
		document.all.divEappN.style.display="none";
		document.all.divEappC.style.display="none";
		<%end if%>
	}
}

//전자결재 품의서 등록 - 업무협조고유번호(scmidx)
function jsRegEapp(scmidx){
	var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes");
	document.frmEapp.iSL.value = scmidx;
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

//전자결재 품의서 내용보기
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}
</script>

</head>
<body LEFTMARGIN="5" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<table width="100%" cellpadding="0" cellspacing="1" class="a" border="0">
<tr>
	<td height="25"><font color="#4E9FC6"><b>보낸업무협조 > <%=NaviTitle(sDoc_Status)%> > <%=CHKIIF(sDoc_Subj="","신규등록",sDoc_Subj)%></font></b></font></td>
</tr>
</table>
<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
<input type="hidden" name="tC" value="">
<input type="hidden" name="ieidx" value="37">
<input type="hidden" name="iSL" value="">
</form>
<form name="frm" action="cooperate_proc.asp" method="post"  style="margin:0px;">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<input type="hidden" name="gubun" value="write">
<input type="hidden" name="doc_useyn" value="<%=sDoc_UseYN%>">
<input type="hidden" name="ispop" value="<%=vIsPop%>">
<input type="hidden" name="doc_status" value="<%=sDoc_Status%>">
<table border="0" cellpadding="5" cellspacing="0" class="a" width="100%">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" width="850">
		<% If iDoc_Idx <> "" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">협조문 번호</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=iDoc_Idx%><%=CHKIIF(sDoc_UseYN="N","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<u><b><font color='red'>삭제된 문서</font></b></u>]","")%></td>
		</tr>
		<% End If %>
		<% '<input type="hidden" name="doc_useyn" value="Y"> %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Name%>(<%=sDoc_Id%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%=sDoc_Regdate%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">현재 상태</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=NaviTitle(sDoc_Status)%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">처리 기간</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" name="doc_start" size="10" maxlength=10 readonly value="<%= sDoc_Start %>">
				<a href="javascript:calendarOpen(frm.doc_start);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
				&nbsp;~&nbsp;
				<input type="text" name="doc_end" size="10" maxlength=10 readonly value="<%= sDoc_End %>">
				<a href="javascript:calendarOpen(frm.doc_end);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">업무 구분</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td><%=CommonCode("w","doc_type",sDoc_Type)%><span id="onlysystem"></span> </td>
					<td align="right" height="30" width="130" style="cursor:pointer"><div class='mainMenu33' flg='A'>[업무구분 설명보기]</div></td>
					<div id='subAID' style='display:none; position:absolute; border:solid 1px #000000; width:200px; padding:3px; background-color:#ffffff;'><%=MyTeamDocTypeExpl()%></div>
				 </tr>
				</table>
			</td>
		</tr>
		<%if iDoc_Idx <> "" THEN%>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">
				<div id="divEappN" style="display:<%if sDoc_Type <>"3" then%>none<% end if%> ;padding:5px;">결재</div>
			</td>
			<td  bgcolor="#FFFFFF" >
				<div id="divEappC" style="display:<%if sDoc_Type <>"3" then%>none<% end if%> ;padding:5px;"">
					<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
						<Tr>
							<Td>
								<% if isNull(sDoc_reportidx) or sDoc_reportidx="" then %>
								<input type="button" class="button"  value="품의서 작성" onClick="jsRegEapp('<%=iDoc_Idx%>');" >
								<% else %>
								<%=fnGetReportState(sDoc_reportstate)%>&nbsp;
								<input type="button" class="button"   value="품의서 보기" onClick="jsViewEapp('<%=sDoc_reportidx%>','<%=sDoc_reportstate%>');">
								<% end if%>
						</td>
					</tr>
				</table>
				</div>
				</td>
		</tr>
		<%END IF%>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">업무 중요도</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("w","doc_important",sDoc_Import)%></td>
		</tr>
	<%
	'				<tr height="30">
	'					<td width="100" align="center"  bgcolor="<= adminColor("tabletop") >">업무 난이도</td>
	'					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><=CommonCode("w","doc_difficult",sDoc_Diffi)></td>
	'				</tr>
	%>				<input type="hidden" name="doc_difficult" value="2">
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">작업자 선택</td>
			<td bgcolor="#FFFFFF" style="padding: 3 0 3 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td>
						<input type="text" class="text" name="doc_workername" value="<%=sDoc_WorkerName%>" size="60" readonly>
						<input type="hidden" name="doc_worker" value="<%=sDoc_Worker%>">
						<input type="button" class="button" value="작업자리스트" onClick="workerlist()">
						<% If iDoc_Idx <> "" Then %>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;확인일 : <%=sDoc_ViewList%>
						<% End If %>
						<br><label id="sms_send_label" style="cursor:pointer"><input type="checkbox" id="sms_send_label" name="sms_send" value="o" <% If iDoc_Idx = "" Then %>checked<% End If %>>선택된 작업자에게 SMS 전송</label>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">참조자 선택<br>※ <font color="blue">필수입력아님</font></td>
			<td bgcolor="#FFFFFF" style="padding: 3 0 3 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td>
						<input type="text" class="text" name="doc_refername" value="<%=sDoc_ReferName%>" size="60" readonly>
						<input type="hidden" name="doc_refer" value="<%=sDoc_Refer%>">
						<input type="button" class="button" value="참조자리스트" onClick="workerlist()">
						<br><label id="sms_r_send_label" style="cursor:pointer"><input type="checkbox" id="sms_r_send_label" name="sms_r_send" value="o">선택된 참조자에게 SMS 전송</label>
					</td>
					<% If iDoc_Idx <> "" Then %>
					<td align="right" height="30" width="130" style="cursor:pointer"><div class='mainMenu33' flg='C'>[참조자별 확인일 보기]</div></td>
      				<div id='subCID' style='display:none; position:absolute; border:solid 1px #000000; width:200px; padding:3px; background-color:#ffffff;'><%=sDoc_ViewListRef%></div>
					<% End If %>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="doc_subject" value="<%=sDoc_Subj%>" size="95" maxlength="148">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내 용</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 5 4 2 5">
			<textarea name="doc_content" id="doc_content" style="width: 100%; height: 490px;" ><%=sDoc_Content%></textarea>
					 <script type="text/javascript">
					    EditorCreator.convert(document.getElementById("doc_content"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                            EditorJSLoader.ready(function (Editor) {
                                new Editor(config);
                                Editor.modify({
                                    content: document.getElementById("doc_content")
                                });
                            });
                        });
					    </script>
                	<!-- daumeditor   -->
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td width="100" valign="top" style="padding:5 0 0 0"><input type="button" value="파일업로드" onClick="fileupload()" class="button"></td>
					<td width="100%" style="padding:3 0 3 10">
						<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
						<%
						IF isArray(arrFileList) THEN
							For intLoop =0 To UBound(arrFileList,2)
						%>
							<tr>
								<td>
									<input type='hidden' name='doc_file' value='<%=arrFileList(1,intLoop)%>'>
									<input type='hidden' name='doc_realfile' value='<%=arrFileList(2,intLoop)%>'>
									<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
									<!--<a href='<%=arrFileList(0,intLoop)%>' target='_blank'><%=Split(Replace(arrFileList(1,intLoop),"http://",""),"/")(3)%></a>//-->
									<span id="<%=intLoop%>" class="a" onClick="filedownload(<%=arrFileList(0,intLoop)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,intLoop),"http://",""),"/")(3)%></span>
								</td>
							</tr>
						<%
							Next
							Response.Write "<input type='hidden' name='isfile' value='o'>"
						Else
						%>
							<tr>
								<td>
								</td>
							</tr>
						<% End If %>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="2">
						<br><b>※ 파일 삭제시 하단의 확인버튼을 클릭하셔야 적용이 됩니다. 새로고침을 하면 다시 나타납니다.<br>
						※ 파일명을 될 수 있는한 특수문자를 - _ ( ) 이 정도로 제한해 주시고, 가급적 한글명보다는<br>&nbsp;&nbsp;&nbsp;&nbsp;영문명으로 해주시기 바랍니다.<br>
						※ 파일명을 될 수 있는한 짧게 해주시기 바랍니다.<br>
						※ 똑같은 파일명이 존재할 경우 파일명 앞에 현재 시간이 자동으로 붙게 됩니다.</b>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left">
	<% If sDoc_Status <> "3" Then %>
		<input type="image" src="/images/<%=CHKIIF(iDoc_Idx="","icon_save","coop_modify")%>.gif" border="0" onClick="checkform();">
		&nbsp;&nbsp;&nbsp;
		<% If iDoc_Idx <> "" AND sDoc_Id = session("ssBctId") Then %><img src="/images/icon_delete.gif" border="0" style="cursor:pointer" onClick="cooperate_del()"><% End If %>
	<% Else %>
		※ 작업완료된 것은 수정이 불가합니다.
	<% End If %>
	</td>
	<td width="50%" align="right">
	<% If sDoc_Status = "4" Then %>
		<img src="/images/coop_re.gif" border="0" style="cursor:pointer" onClick="cooperate_status_change('1','반려된 협조문')">
	<% End If %>
	</td>
</tr>
</table>
</form>


<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/cooperate_admin/cooperate_download.asp" target="fileiframe">
<input type="hidden" name="doc_idx" value="<%=iDoc_Idx%>">
<input type="hidden" name="file_idx" value="">
</form>
<iframe src=""   name="fileiframe" width="0" height="0" frameborder="0"></iframe>
<br>
<% If iDoc_Idx <> "" Then %>
<!-- ####### 답변쓰기 ####### //-->
<br>
<div style="text-align:center;width:850px">
<iframe src="iframe_cooperate_ans.asp?ispop=<%=vIsPop%>&didx=<%=iDoc_Idx%>" name="iframeDB1" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="5" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
</div>
<!-- ####### 답변쓰기 ####### //-->
<% End If %>
<br><br>

<%
	set cooperateView = nothing
	set cooperateFile = nothing
%>
<script>
$(document).ready(function()
{
     $('.mainMenu33').mouseover(function(){
            setClientPos($(this));
     });
     $('.mainMenu33').mouseout(function(){
            $('#sub'+$(this).attr('flg')+'ID').hide();
     });
});
function setClientPos(main)
{
	window.status = $(document.body).position().top;
    var sub = $('#sub'+main.attr('flg')+'ID');
    sub.css( "zIndex", 1000 ).show();
    sub.css('left',(main.position().left+main.width()-199)).css('top',main.position().top+12);
}
</script>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
