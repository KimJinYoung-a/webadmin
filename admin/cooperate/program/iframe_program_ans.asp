<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 프로그램변경내역
' Hieditor : 강준구 생성
'			 2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/programchangeCls.asp"-->
<%
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrentpage ,iDelCnt, sRegistId
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, iDoc_Idx, iAns_Idx, sAns_Content, vIsPop
	Dim cPrCh, vIdx
	
	vIdx			= NullFillWith(requestCheckVar(Request("pidx"),10),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sRegistId	 	= NullFillWith(requestCheckVar(Request("registid"),50),"")
	iPageSize 		= 100
	iPerCnt 		= 10
	
	Dim cooperateAns, i
	
		set cPrCh = new CProgramChange
	 	cPrCh.FCPage = iCurrentpage
	 	cPrCh.FPSize = iPageSize
	 	cPrCh.FPIdx = vIdx
		arrList = cPrCh.fnGetPrChAnsList
		iTotCnt = cPrCh.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type='text/javascript'>
function jsGoPage(iP){
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}
function ans_edit(aidx)
{
	location.href = "iframe_cooperate_ans.asp?ispop=<%=vIsPop%>&didx=<%=iDoc_Idx%>&iC=<%=iCurrentpage%>&aidx="+aidx+"&registid=<%=sRegistId%>";
}
function ans_del(aidx)
{
	if(confirm("선택하신 글을 삭제하시겠습니까?") == true) {
		location.href = "program_ans_proc.asp?aidx="+aidx+"&del=o&registid=<%=sRegistId%>";
	} else {
		return false;
	}
}
function checkform(frm)
{
	if (frm.ans_content.value == "")
	{
		alert("답변을 입력하세요!");
		frm.ans_content.focus();
		return false;
	}
}
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<table align="center" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
<tr>
	<td bgcolor="#FFFFFF">
		<form name="frm" action="program_ans_proc.asp" method="post" onSubmit="return checkform(this);" style="margin:0px;">
		<input type="hidden" name="pidx" value="<%=vIdx%>">
		<input type="hidden" name="aidx" value="<%=iAns_Idx%>">
		<input type="hidden" name="registid" value="<%=sRegistId%>">
		<table width="812"  cellpadding="1" cellspacing="1" class="a" border="0" >
		<tr>
			<td style="padding:5 5 10 5;"> * COMMENT </td> 
		</tr>
		<tr>
			<td align="center" style="padding:0 0 15 0;">
				<textarea name="ans_content" rows="3" cols="<%=CHKIIF(InStr(UCase(cstr(request.ServerVariables("HTTP_USER_AGENT"))),"MSIE"),"97","85")%>"><%= ReplaceBracket(sAns_Content) %></textarea>
				&nbsp;<input type="submit" value="등록" class="button" style="height:50px;width:80px;vertical-align:top;">
			</td>
		</tr>
		<tr>
			<td align="right">
			</td>
		</tr>
		<tr>
			<td>	
				<div id="Cmtlist" style="padding-left:18px;padding-right:18px;"> 
				<%
					'### A.idx, A.userid, A.comment, A.regdate, B.username
					IF isArray(arrList) THEN
						For intLoop =0 To UBound(arrList,2)
							Response.Write "<span style=""font-size:11px;color:#696969"">" & arrList(4,intLoop) & "(" & arrList(1,intLoop) & ")&nbsp;" & arrList(3,intLoop) & "</span>&nbsp;"
							If arrList(1,intLoop) = session("ssBctId") Then
								'Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_modify.gif' style='cursor:pointer' onClick='ans_edit(" & arrList(0,intLoop) & ")'>"
								Response.Write "&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='ans_del(" & arrList(0,intLoop) & ")'>"
							End If
							Response.Write "<br><div style=""padding:5px;border-bottom:1px solid #BABABA;width:100%"">" & replace(arrList(2,intLoop),vbCrLf,"<br>") & "</div><Br>"
						Next
					Else
						Response.Write "<center><div style=""padding:5px;border-bottom:1px solid #BABABA;width:100%;"">[답변이 없습니다.]</div></center><Br>"
					End If
				%>
				</div>
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>


<%
	set cooperateAns = nothing
%>

</body>
</html>