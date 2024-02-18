<%@ language=vbscript %>
<% option explicit %>
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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
	Dim iTotCnt, arrList, intLoop, arrFileList, i
	Dim iPageSize, iCurrentpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, sDoc_ViewList, sDoc_ViewListRef, sDoc_ReferView
	Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Start, sDoc_End, sDoc_Type, sDoc_Import, sDoc_Diffi, sDoc_Subj, sDoc_Content
	Dim sDoc_WorkerName, sDoc_Worker, sDoc_UseYN, sDoc_Regdate, sDoc_WorkerView, sDoc_WorkerView_temp, sDoc_Refer, sDoc_ReferName
	Dim vParam, s_search_team, s_status, s_type, s_ans_ox, s_onlymine
	Dim sDoc_reportidx ,sDoc_reportstate,sSys_reportidx  ,sSys_reportstate, existsdoc_workerYN,existsdoc_referYN
	
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	
	s_search_team	= NullFillWith(requestCheckVar(Request("s_search_team"),20),"")
	s_status		= NullFillWith(requestCheckVar(Request("s_status"),10),"")
	s_type			= NullFillWith(requestCheckVar(Request("s_type"),10),"")
	s_ans_ox		= NullFillWith(requestCheckVar(Request("s_ans_ox"),1),"")
	s_onlymine		= NullFillWith(requestCheckVar(Request("s_onlymine"),1),"")
	vParam = "&iC="&iCurrentpage&"&search_team="&s_search_team&"&doc_status="&s_status&"&doc_type="&s_type&"&ans_ox="&s_ans_ox&"&onlymine="&s_onlymine&""
	'<!-- 읽는 페이지에 내용도 저장이 되어서 혹시나 모를 일에 대비하여 파라메터명을 바꿔서 주고 받았슴. //-->

existsdoc_workerYN="N"
existsdoc_referYN="N"

	If iDoc_Idx = "" Then
		sDoc_Id 		= session("ssBctId")
		sDoc_Name		= session("ssBctCname")
		sDoc_Regdate	= Left(now(),10)
		sDoc_WorkerName	= ""
		sDoc_Worker		= ""
		sDoc_ReferView	= ""
		sDoc_Refer		= ""
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
		sDoc_Content	= ReplaceScript(cooperateView.FDoc_Content)
		sDoc_UseYN		= cooperateView.FDoc_UseYN
		sDoc_Regdate	= cooperateView.FDoc_Regdate
		sDoc_WorkerName	= cooperateView.FDoc_WorkerName
		sDoc_Worker		= cooperateView.FDoc_Worker
		sDoc_Refer		= cooperateView.FDoc_Refer
		sDoc_ReferName	= cooperateView.FDoc_ReferName
		sDoc_WorkerView	= cooperateView.FDoc_WorkerViewdate
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
	End If

if instr(sdoc_worker,Trim(session("ssBctId")))>0 then existsdoc_workerYN = "Y"
if instr(sdoc_refer,Trim(session("ssBctId")))>0 then existsdoc_referYN = "Y"

' 작업자 , 참조자 둘다 아닐경우에 팅겨냄
If existsdoc_workerYN="N" and existsdoc_referYN="N" Then
	Response.Write "<script>alert('작업자나,참조자만 보기 권한이 있습니다.');history.back();</script>"
	Response.End
End If
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="Javascript">
var openWorker = null;

function workerlist()
{
	var worker = frm.doc_worker.value;
	openWorker = window.open('PopWorkerList.asp?worker='+worker+'&didx=<%=iDoc_Idx%>','openWorker','width=570,height=527,scrollbars=yes');
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

function checkform(frm)
{


	if (frm.doc_worker.value == "")
	{
		alert("작업자를 선택해 주세요!");
		return false;
	}
	
	if(!(openWorker == null))
	{
		if(!(openWorker.closed))
		{
			openWorker.close();
		}
	}
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

function filedownload(idx)
{
	filefrm.file_idx.value = idx;
	filefrm.submit();
}

function goProgram(){
	var popprogram = window.open('/admin/cooperate/program/write.asp?didx=<%=iDoc_Idx%>','popprogram','width=850,height=190,scrollbars=yes');
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
<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
<input type="hidden" name="tC" value="">
<input type="hidden" name="ieidx" value="37">  
<input type="hidden" name="iSL" value="">
</form>
<form name="frm" action="cooperate_proc.asp" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="read" value="o">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>협조문 내용</b></td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td style="padding-bottom:10"> 
			<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<% If iDoc_Idx <> "" Then %>
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">협조문 번호</td>
					<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=iDoc_Idx%></td>
				</tr>
				<% End If %>
				<input type="hidden" name="doc_useyn" value="Y">
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
					<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Name%>(<%=sDoc_Id%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%=sDoc_Regdate%></td>
				</tr>
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">현재 상태</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("w","doc_status",sDoc_Status)%></td>
				</tr>
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">처리 기간</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Start%> ~ <%=sDoc_End%>
						<input type="hidden" name="doc_start" value="<%=sDoc_Start%>">
						<input type="hidden" name="doc_end" value="<%=sDoc_End%>">
					</td>
				</tr>
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">업무 구분</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("v","doc_type",sDoc_Type)%>
						<input type="hidden" name="doc_type" value="<%=sDoc_Type%>">
					</td>
				</tr>
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
							    <font color="Gray">품의서 미작성</font>
								<% else %>
								<%=fnGetReportState(sDoc_reportstate)%>&nbsp;
								<input type="button" class="button"   value="품의서 보기" onClick="jsViewEapp('<%=sDoc_reportidx%>','<%=sDoc_reportstate%>');">
								<% end if%>  
								<%IF sDoc_reportstate = 7 THEN%> 
								<% if isNull(sSys_reportidx) or sSys_reportidx="" then %>
								<input type="button" class="button"  value="▷ 개발계획서 품의" onClick="jsRegEapp('<%=iDoc_Idx%>');" >
								<% else %>
								/ <%=fnGetReportState(sSys_reportstate)%>&nbsp; 
								<input type="button" class="button"  value="개발계획서 보기" onClick="jsViewEapp('<%=sSys_reportidx%>','<%=sSys_reportstate%>');">
								<% end if%>  
							 <%END IF%>
						</td>
					</tr>
				</table>
				</div>
				</td>
		</tr>
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">업무 중요도</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("v","doc_important",sDoc_Import)%>
						<input type="hidden" name="doc_important" value="<%=sDoc_Import%>">
					</td>
				</tr>
				<input type="hidden" name="doc_difficult" value="2">
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">작업자 선택</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
						<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
						<tr>
							<td>
								<input type="text" name="doc_workername" value="<%=sDoc_WorkerName%>" size="60" readonly>
								<input type="hidden" name="doc_worker" value="<%=sDoc_Worker%>">
								<input type="button" class="button" value="작업자리스트" onClick="workerlist()">
								&nbsp;&nbsp;&nbsp;확인일 : <%=sDoc_ViewList%>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">참조자 선택</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
						<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
						<tr>
							<td>
								<input type="text" name="doc_refername" value="<%=sDoc_ReferName%>" size="60" readonly>
								<input type="hidden" name="doc_refer" value="<%=sDoc_Refer%>">
								<input type="button" class="button" value="참조자리스트" onClick="workerlist()">
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
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Subj%>
					</td>
				</tr>
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내 용</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Content%>
					</td>
				</tr>
				<tr height="30">
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
						<table cellpadding="0" cellspacing="0" border="0" class="a">
						<tr>
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
											<!--· <a href='<%=arrFileList(0,intLoop)%>' target='_blank'><%'Split(Replace(arrFileList(0,intLoop),"http://",""),"/")(3)%></a>//-->
											· <span id="<%=intLoop%>" class="a" onClick="filedownload(<%=arrFileList(0,intLoop)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,intLoop),"http://",""),"/")(3)%></span>
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
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<table width="813" border="0" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td width="50%" align="left"><a href="index.asp?menupos=<%=g_MenuPos%><%=vParam%>"><img src="/images/icon_list.gif" border="0"></a></td>
		<td width="50%" align="right">
		<% If iDoc_Idx <> "" AND sDoc_Id = session("ssBctId") Then %><!--<img src="/images/icon_delete.gif" border="0" style="cursor:pointer" onClick="cooperate_del()">&nbsp;&nbsp;&nbsp;//--><% End If %>
			<table cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td style="padding-right:15">
				<% If sDoc_Type = "3" AND (CInt(g_MyPart) = CInt("7") OR CInt(g_MyPart) = CInt("30")) Then %>
					<% If fnProgramWriteCount(iDoc_Idx) = 0 Then %>
						<input type="button" value="프로그램변경내역작성" onClick="goProgram();">
					<% End If %>
				<% End If %>
				</td>
				<td><input type="image" src="/images/icon_confirm.gif" border="0"></td>
			</tr>
			</table>
		</td>
	</tr>
</table>
</form>


<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/cooperate_admin/cooperate_download.asp" target="fileiframe">
<input type="hidden" name="doc_idx" value="<%=iDoc_Idx%>">
<input type="hidden" name="file_idx" value="">
</form>
<iframe src="" width="0" height="0" name="fileiframe" width="0" height="0"></iframe>


<% If iDoc_Idx <> "" Then %>
<!-- ####### 답변쓰기 ####### //-->
<br>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td width="14"><img src="/images/icon_arrow_link.gif"></td>
	<td width="800" style="padding-top:3">&nbsp;<b>협조문 답변</b></td>
</tr>
<tr>
	<td colspan="2"><iframe src="iframe_cooperate_ans.asp?didx=<%=iDoc_Idx%>&registid=<%=sDoc_Id%>" name="iframeDB1" width="814" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" style="width:814px;" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
</table>
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
     sub.show(); 
     sub.css('left',(main.position().left+main.width()-199)).css('top',main.position().top+12); 
} 
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
