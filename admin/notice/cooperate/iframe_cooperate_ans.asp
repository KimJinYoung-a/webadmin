<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��������
' History : ���ر� ����
'			2022.07.11 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<%
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrentpage ,iDelCnt, sRegistId
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, iDoc_Idx, iAns_Idx, sAns_Content
	
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iAns_Idx		= NullFillWith(requestCheckVar(Request("aidx"),10),"")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sRegistId	 	= NullFillWith(requestCheckVar(Request("registid"),50),"")
	iPageSize 		= 20
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
			Response.Write "<script type='text/javascript'>alert('�߸��� �����Դϴ�.');location.href='iframe_cooperate_ans.asp?didx="&iDoc_Idx&"&iC="&iCurrentpage&"';</script>"
			dbget.close()
			Response.End
		End IF
	End If
%>

<script type='text/javascript'>

function jsGoPage(iP){
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}
function ans_edit(aidx)
{
	location.href = "iframe_cooperate_ans.asp?didx=<%=iDoc_Idx%>&iC=<%=iCurrentpage%>&aidx="+aidx+"&registid=<%=sRegistId%>";
}
function ans_del(aidx)
{
	if(confirm("�����Ͻ� ���� �����Ͻðڽ��ϱ�?") == true) {
		location.href = "cooperate_ans_proc.asp?didx=<%=iDoc_Idx%>&iC=<%=iCurrentpage%>&aidx="+aidx+"&del=o&registid=<%=sRegistId%>";
	} else {
		return false;
	}
}
function checkform(frm)
{
	if (frm.ans_content.value == "")
	{
		alert("�亯�� �Է��ϼ���!");
		frm.ans_content.focus();
		return false;
	}
}
</script>

<form name="frm" action="cooperate_ans_proc.asp" method="post" onSubmit="return checkform(this);" style="margin:0px;">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<input type="hidden" name="aidx" value="<%=iAns_Idx%>">
<input type="hidden" name="registid" value="<%=sRegistId%>">
<table width="800" align="center" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�亯����</td>
	<td align="left"><textarea class="textarea" name="ans_content" cols="112" rows="5"><%=sAns_Content%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right">
		<% If sRegistId <> "" Then %>
		<label id="sms_send_label" style="cursor:pointer"><input type="checkbox" id="sms_send_label" name="sms_send" value="o" checked>����ڿ��� SMS ����</label>&nbsp;&nbsp;&nbsp;
		<% End IF %>
		<input type="submit" value="�亯����" class="button">
	</td>
</tr>
</table>
</form>

<br>

<table width="800" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" width="140">�ۼ���</td>
	<td align="center">��&nbsp;&nbsp;&nbsp;��</td>
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>
	    	<tr align="center" bgcolor="#FFFFFF" height="30">
				<td align="center" valign="top" style="padding:3 0 0 3">
					<%
						Response.Write arrList(4,intLoop) & "(" & arrList(5,intLoop) & ")"
						Response.Write "<br>" & arrList(3,intLoop)
						If arrList(5,intLoop) = session("ssBctId") Then
							Response.Write "<br><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_modify.gif' style='cursor:pointer' onClick='ans_edit(" & arrList(0,intLoop) & ")'>"
							Response.Write "&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='ans_del(" & arrList(0,intLoop) & ")'>"
						End If
					%>
				</td>
				<td align="left" style="padding:3 3 3 3"><%= ReplaceBracket(replace(arrList(2,intLoop),vbCrLf,"<br>")) %></td>
	    	</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="2" align="center" class="page_link">[�亯�� �����ϴ�.]</td>
		</tr>
<%
	End If
%>
<tr>
	<td colspan="2">
	<!-- ����¡ó�� -->
	<%
	iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
	
	If (iCurrentpage mod iPerCnt) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	<form name="frmpage" method="post" style="margin:0px;">
	<input type="hidden" name="iC" value="<%=iCurrentpage%>">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	    <tr height="25">        
	        <td align="center">
	         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrentpage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
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
