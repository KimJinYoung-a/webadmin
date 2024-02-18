<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/common/comment/commentCls.asp"-->

<%
	Dim scmcommentCls, arrList, i, iTotalPage, iTotCnt, iCurrentpage, iPerCnt, iPageSize, intLoop
	Dim vCols, vRows, vBtnWidth, vBtnHeight, vCommentIdx, vParentIdx, vRegistId, vComment, vBoardType, vBoardGubun, vEtc1, vEtc2
	vCols			= NullFillWith(requestCheckVar(Request("cols"),3),97)
	vRows			= NullFillWith(requestCheckVar(Request("rows"),3),3)
	vBtnWidth		= NullFillWith(requestCheckVar(Request("btnwidth"),3),80)
	vBtnHeight		= NullFillWith(requestCheckVar(Request("btnheight"),3),50)
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	vCommentIdx		= requestCheckVar(Request("cidx"),10)
	vRegistId	 	= requestCheckVar(Request("registid"),50)
	vBoardType		= requestCheckVar(Request("boardtype"),2)
	vBoardGubun		= requestCheckVar(Request("boardgubun"),50)
	vParentIdx		= requestCheckVar(Request("pidx"),10)
	vEtc1			= requestCheckVar(Request("etc1"),100)
	vEtc2			= requestCheckVar(Request("etc2"),100)
	iPageSize 		= 300
	iPerCnt 		= 10


	If vBoardGubun = "" Then
		Response.End
		dbget.close()
	End IF
	
	If vParentIdx = "" Then
		Response.End
		dbget.close()
	Else
		If IsNumeric(vParentIdx) = False Then
			Response.End
			dbget.close()
		End If
	End IF

	set scmcommentCls = new CSCMComment
 	scmcommentCls.FCPage = iCurrentpage
 	scmcommentCls.FPSize = iPageSize
 	scmcommentCls.FBoardGubun = vBoardGubun
 	scmcommentCls.FParentIdx = vParentIdx
 	scmcommentCls.FDeleteyn = "n"
	arrList = scmcommentCls.fnGetSCMCommentList
	iTotCnt = scmcommentCls.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
	
	
	If vCommentIdx <> "" Then
		scmcommentCls.FCIdx = vCommentIdx
		scmcommentCls.fnGetSCMCommentView
		vComment = scmcommentCls.FComment
	End IF
	set scmcommentCls = nothing
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
function jsGoPage(iP){
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}
function ans_edit(cidx)
{
	location.href = "comment.asp?pidx=<%=vParentIdx%>&iC=<%=iCurrentpage%>&cidx="+cidx+"&registid=<%=vRegistId%>&boardtype=<%=vBoardType%>&boardgubun=<%=vBoardGubun%>&cols=<%=vCols%>&rows=<%=vRows%>&btnwidth=<%=vBtnWidth%>&btnheight=<%=vBtnHeight%>";
}
function ans_del(cidx)
{
	if(confirm("선택하신 글을 삭제하시겠습니까?") == true) {
		location.href = "comment_proc.asp?pidx=<%=vParentIdx%>&iC=<%=iCurrentpage%>&del=o&cidx="+cidx+"&registid=<%=vRegistId%>&boardtype=<%=vBoardType%>&boardgubun=<%=vBoardGubun%>&cols=<%=vCols%>&rows=<%=vRows%>&btnwidth=<%=vBtnWidth%>&btnheight=<%=vBtnHeight%>";
	} else {
		return false;
	}
}
function checkform(frm)
{
	if (frm.comment.value == "")
	{
		alert("답변을 입력하세요!");
		frm.comment.focus();
		return false;
	}
}
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<form name="frm" action="comment_proc.asp" method="post" onSubmit="return checkform(this);" style="margin:0px;">
<input type="hidden" name="cols" value="<%=vCols%>">
<input type="hidden" name="rows" value="<%=vRows%>">
<input type="hidden" name="btnwidth" value="<%=vBtnWidth%>">
<input type="hidden" name="btnheight" value="<%=vBtnHeight%>">
<input type="hidden" name="pidx" value="<%=vParentIdx%>">
<input type="hidden" name="cidx" value="<%=vCommentIdx%>">
<input type="hidden" name="registid" value="<%=vRegistId%>">
<input type="hidden" name="boardtype" value="<%=vBoardType%>">
<input type="hidden" name="boardgubun" value="<%=vBoardGubun%>">

<!-- ################################ 예비 입력란 ################################ //-->
<input type="hidden" name="etc1" value="<%=vEtc1%>">
<input type="hidden" name="etc2" value="<%=vEtc2%>">
<!-- ################################ 예비 입력란 ################################ //-->

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
<tr>
	<td bgcolor="#FFFFFF" style="padding:5px 0 0 0;" align="center">
		<table cellpadding="1" cellspacing="1" class="a" border="0" >
		<tr>
			<td style="padding:5 5 10 5;"> * COMMENT </td> 
		</tr>
		<tr>
			<td align="center" style="padding:0 0 15 0;">
				<textarea name="comment" rows="<%=vRows%>" cols="<%=vCols%>"><%=vComment%></textarea>&nbsp;<input type="submit" value="등 록" class="button" style="height:<%=vBtnHeight%>px;width:<%=vBtnWidth%>px;vertical-align:top;">
			</td>
		</tr>
		<tr>
			<td align="right">
				<% If vRegistId <> "" Then %>
				<label id="sms_send_label" style="cursor:pointer;"><input type="checkbox" id="sms_send_label" name="sms_send" value="o" checked>등록자에게 SMS 전송</label>&nbsp;&nbsp;&nbsp;
				<% End IF %>
			</td>
		</tr>
		<tr>
			<td>	
				<div id="Cmtlist" style="padding-left:18px;padding-right:18px;"> 
				<%	'배열번호	0			1					2						3						4			5			6			7
					'####### A.cIdx, A.ans_content, isNull(A.etc1,'') AS etc1, isNull(A.etc2,'') AS etc2, A.regUserid, A.deleteyn, A.regdate, B.username
					
					IF isArray(arrList) THEN
						For intLoop =0 To UBound(arrList,2)
							Response.Write "<span style=""font-size:11px;color:#696969"">" & arrList(7,intLoop) & "(" & arrList(4,intLoop) & ")&nbsp;" & arrList(6,intLoop) & "</span>&nbsp;"
							If arrList(4,intLoop) = session("ssBctId") Then
								Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_modify.gif' style='cursor:pointer' onClick='ans_edit(" & arrList(0,intLoop) & ")'>"
								Response.Write "&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='ans_del(" & arrList(0,intLoop) & ")'>"
							End If
							Response.Write "<br><div style=""padding:5px;border-bottom:1px solid #BABABA;width:100%"">" & replace(arrList(1,intLoop),vbCrLf,"<br>") & "</div><Br>"
						Next
					Else
						Response.Write "<center><div style=""padding:5px;border-bottom:1px solid #BABABA;width:100%;"">[답변이 없습니다.]</div></center><Br>"
					End If
				%>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->