<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->

<%
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrentpage ,iDelCnt, sRegistId
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, iDoc_Idx, iAns_Idx, sAns_Content, vIsPop
	
	vIsPop			= Request("ispop")
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iAns_Idx		= NullFillWith(requestCheckVar(Request("aidx"),10),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sRegistId	 	= NullFillWith(requestCheckVar(Request("registid"),50),"")
	iPageSize 		= 100
	iPerCnt 		= 10
	
	Dim cooperateAns, i
	
		set cooperateAns = new CCooperate
	 	cooperateAns.FCPage = iCurrentpage
	 	cooperateAns.FPSize = iPageSize
	 	cooperateAns.FDoc_Idx = iDoc_Idx
		arrList = cooperateAns.fnGetCooperateAnsList
		iTotCnt = cooperateAns.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
	
	If iAns_Idx <> "" Then
		cooperateAns.FAns_Idx = iAns_Idx
		cooperateAns.fnGetCooperateAnsView
		sAns_Content = cooperateAns.FAns_Content
		
		If sAns_Content = "" Then
			Response.Write "<script>alert('잘못된 접근입니다.');location.href='iframe_cooperate_ans.asp?didx="&iDoc_Idx&"&iC="&iCurrentpage&"';</script>"
			dbget.close()
			Response.End
		End IF
	End If
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=10" /> 
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
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
		location.href = "cooperate_ans_proc.asp?ispop=<%=vIsPop%>&didx=<%=iDoc_Idx%>&iC=<%=iCurrentpage%>&aidx="+aidx+"&del=o&registid=<%=sRegistId%>";
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
<table align="center" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA width="100%">
<tr>
	<td bgcolor="#FFFFFF">
		<form name="frm" action="cooperate_ans_proc.asp" method="post" onSubmit="return checkform(this);" style="margin:0px;">
		<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
		<input type="hidden" name="aidx" value="<%=iAns_Idx%>">
		<input type="hidden" name="registid" value="<%=sRegistId%>">
		<input type="hidden" name="ispop" value="<%=vIsPop%>">
		<table    cellpadding="3" cellspacing="1" class="a" border="0" align="center">
		<tr>
			<td style="padding:5 5 10 5;"> * COMMENT </td> 
		</tr>
		<tr>
			<td align="right">
				<textarea name="ans_content" rows="3" cols="80"><%=sAns_Content%></textarea> 
			</td>
			<td align="left">	
				<input type="submit" value="등록" class="button" style="height:50px;width:80px;vertical-align:top;">
			</td>
		</tr>
		<tr>
			<td align="right" colspan="2">
				<% If sRegistId <> "" Then %>
				<label id="sms_send_label" style="cursor:pointer;"><input type="checkbox" id="sms_send_label" name="sms_send" value="o" checked>등록자에게 SMS 전송</label>&nbsp;&nbsp;&nbsp;
				<% End IF %>
			</td>
		</tr>
		<tr>
			<td  colspan="2">	
				<div id="Cmtlist" style="padding-left:18px;padding-right:18px;"> 
				<%
					IF isArray(arrList) THEN
						For intLoop =0 To UBound(arrList,2)
							Response.Write "<span style=""font-size:11px;color:#696969"">" & arrList(4,intLoop) & "(" & arrList(5,intLoop) & ")&nbsp;" & arrList(3,intLoop) & "</span>&nbsp;"
							If arrList(5,intLoop) = session("ssBctId") Then
								Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_modify.gif' style='cursor:pointer' onClick='ans_edit(" & arrList(0,intLoop) & ")'>"
								Response.Write "&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='ans_del(" & arrList(0,intLoop) & ")'>"
							End If
							Response.Write "<br><div style=""padding:5px;border-bottom:1px solid #BABABA;width:100%"">" & ReplaceScript(replace(arrList(2,intLoop),vbCrLf,"<br>")) & "</div><Br>"
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
<%
	session.codePage = 949
%>
