<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 설문관리
' Hieditor : 허진원 생성
'			 2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim page, lp, qstNo, using, strType, strDel, srv_sn
	dim btcid
	btcid= session("ssBctID")

	srv_sn = Request("sn")

	'기본값 지정
	page=1
	using="Y"

	'// 설문내용 접수
	dim oSurveyMaster
	Set oSurveyMaster = new CSurvey

	oSurveyMaster.FRectSn = srv_sn
	
	oSurveyMaster.GetSurveyCont

	'// 설문문항 목록
	dim oSurveyQuestion
	Set oSurveyQuestion = new CSurvey

	oSurveyQuestion.FRectSn = srv_sn
	oSurveyQuestion.FPagesize = 100
	oSurveyQuestion.FCurrPage = page
	oSurveyQuestion.FRectUsing = using
	oSurveyQuestion.FRectOrder = "asc"

	oSurveyQuestion.GetSurveyQstList
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<title><%=oSurveyMaster.FitemList(1).Fsrv_subject%></title>
<style type="text/css">
<!--
body,table,tr,td {font-family: 돋음,AppleGothic,sans-serif; color:#888888; font-size:10px; word-spacing: -2px;scrollbar-face-color: F2F2F2; scrollbar-shadow-color:#bbbbbb; scrollbar-highlight-color: #bbbbbb; scrollbar-3dlight-color: #FFFFFF; scrollbar-darkshadow-color: #FFFFFF; scrollbar-track-color: #F2F2F2; scrollbar-arrow-color: #bbbbbb; scrollbar-base-color:#E9E8E8;}
td {word-break:break-all;}
img,table {border:0px;}
b {letter-spacing:-1px;}
input {padding-top:3px; height:21px;}
textarea {line-height:18px; padding:3px;}

.redtitle{font-family: 돋움; FONT-SIZE: 12px; COLOR: #c3080a; font-weight:bold;}
.graytext{font-family: 돋움; FONT-SIZE: 12px; COLOR: #888888; font-weight:bold;}
.grayNomal{font-family: 돋움; FONT-SIZE: 12px; COLOR: #888888;}
.input_text {border:1px #cccccc solid; FONT-FAMILY: "돋움"; font-size: 12px; color="#888888"; padding:1px;}
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>
<script language="javascript">
<!--
	// 추가답변 여부 검사
	function chkPollAdd(pollSn,dsp)	{
		if(document.all["addPoll"+pollSn])
		{
			if(dsp=='Y')
				document.all["addPoll"+pollSn].style.display="";
			else
				document.all["addPoll"+pollSn].style.display="none";
		}
	}

	// 답변 확인 및 실행
	function chkSubmit()	{
		var frm = document.frmSurvey;

	<%
		'문항 도돌이(답변여부 확인용)
		qstNo = 1
		for lp=0 to oSurveyQuestion.FResultCount - 1
			Select Case oSurveyQuestion.FitemList(lp).Fqst_type
				Case "1"
					if oSurveyQuestion.FitemList(lp).Fqst_isNull="N" then
						Response.Write "	if(!chkRadio(" & oSurveyQuestion.FitemList(lp).Fqst_sn & "," & qstNo & ")) return false;" & vbCrLf
					end if
				Case "2", "3"
					if oSurveyQuestion.FitemList(lp).Fqst_isNull="N" then
						Response.Write "	if(!chkText(" & oSurveyQuestion.FitemList(lp).Fqst_sn & "," & qstNo & ")) return false;" & vbCrLf
					end if
				Case "9"
					qstNo = qstNo - 1
			end Select
			qstNo = qstNo + 1
		next
	%>
		return true;
	}

	// 체크박스 선택 여부 확인
	function chkRadio(rid,rno)	{
		var chk=0;
		var robj = MM_findObj("qst" + rid);
   	   	if(!robj.length){
   	   		if(robj.checked){
   	   			chk++;
   	   		}
   	    }else{
   	    	for(i=0;i<robj.length;i++){
   	    		if(robj[i].checked) {	   	    			
					chk++;
   	    		}	
   	    	}
   	    	
   	    	if (chk==0){
   	    		alert(rno+"번 문항의 답변이 없습니다. 답변을 선택해주세요.");
   	   			return false;
   	    	} else {
   	    		return true;
   	    	}
   	    }
	}

	// 주관식 답변여부 확인
	function chkText(rid,rno)	{
		var robj = MM_findObj("qst" + rid);
		if(!robj.value) {
    		alert(rno+"번 문항의 답변이 없습니다. 답변을 작성해주세요.");
   			return false;
    	} else {
    		return true;
    	}
	}

	function MM_findObj(n, d) { //v4.01
	  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
	    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
	  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
	  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
	  if(!x && d.getElementById) x=d.getElementById(n); return x;
	}
//-->
</script>
</head>
<body>
<!-- // 바디 시작 // -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<!-- // 머리말 // -->
	<td class="grayNomal"><%= nl2br(ReplaceBracket(db2html(oSurveyMaster.FitemList(1).Fsrv_head))) %></td>
</tr>
<%
	if oSurveyQuestion.FResultCount>0 then
%>
<form name="frmSurvey" method="POST" action="popup_survey_process.asp" onSubmit="return chkSubmit()">
<input type="hidden" name="sn" value="<%=srv_sn%>">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td style="padding:0 20 20 20;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			'## 문항 도돌이
			qstNo = 1
			for lp=0 to oSurveyQuestion.FResultCount - 1
			
				'//문항 구분별 출력
				Select Case oSurveyQuestion.FitemList(lp).Fqst_type
					Case "1"	'객관식
		%>
				<tr>
					<td style="padding:0 20px 0 20px;">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td style="padding:20px 0 20px 0">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td class="graytext" style="padding-bottom:8px;"><%=qstNo & ". " & ReplaceBracket(oSurveyQuestion.FitemList(lp).Fqst_content) %></td>
								</tr>
								<tr>
									<td><%=oSurveyQuestion.PrintSurveyPollList(oSurveyQuestion.FitemList(lp).Fqst_sn)%></td>
								</tr>
								</table>
							</td>
						</tr>
						<tr height="1"><td height="1" bgcolor="#dddddd"></td></tr>
						</table>
					</td>
				</tr>
		<%
					Case "2"	'주관식
		%>
				<tr>
					<td style="padding:0 20px 0 20px;">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td style="padding:20px 0 20px 0">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td class="graytext" style="padding-bottom:8px;"><%=qstNo & ". " & ReplaceBracket(oSurveyQuestion.FitemList(lp).Fqst_content) %></td>
								</tr>
								<tr>
									<td class="graytext"><textarea name="qst<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>" class="input_text" style="width:100%;height:120px;"></textarea></td>
								</tr>
								</table>
							</td>
						</tr>
						<tr height="1"><td height="1" bgcolor="#dddddd"></td></tr>
						</table>
					</td>
				</tr>
		<%
					Case "3"	'단답형
		%>
				<tr>
					<td style="padding:0 20px 0 20px;">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td style="padding:20px 0 20px 0">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td class="graytext" style="padding-bottom:8px;"><%=qstNo & ". " & ReplaceBracket(oSurveyQuestion.FitemList(lp).Fqst_content) %></td>
								</tr>
								<tr>
									<td class="graytext" style="padding-left:15px;"><input type="text" name="qst<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>" class="input_text" style='width:450px;height:16px;'></td>
								</tr>
								</table>
							</td>
						</tr>
						<tr height="1"><td height="1" bgcolor="#dddddd"></td></tr>
						</table>
					</td>
				</tr>
		<%
					Case "9"	'문항구분
						'문항번호에서 제외
						qstNo = qstNo - 1
		%>
				<tr>
					<td class="redtitle" style="padding:20 0 0 0">[<%= ReplaceBracket(oSurveyQuestion.FitemList(lp).Fqst_content) %>]</td>
				</tr>
		<%
				End Select
				qstNo = qstNo + 1
			Next
		%>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" style="padding-bottom:20px;"><input type="image" src="http://fiximage.10x10.co.kr/web2008/etc/2008poll/poll_btn2.jpg" style="width:189;height:47" border="0"></td>
</tr>
</form>
<% end if %>
<tr>
	<!-- // 꼬리말 // -->
	<td class="grayNomal"><%=nl2br(ReplaceBracket(db2html(oSurveyMaster.FitemList(1).Fsrv_tail)))%></td>
</tr>
</table>
<!-- // 바디 끝 // -->
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->