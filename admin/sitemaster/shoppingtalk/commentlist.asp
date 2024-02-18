<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/shoppingtalk/classes/shoppingtalkCls.asp" -->

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>

<%
	Dim cTalkComm, i, vTalkIdx, vCurrPage, vUserID, vContents
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	vTalkIdx = requestCheckVar(request("talkidx"),10)
	vCurrPage = requestCheckVar(Request("cpg"),5)
	If vCurrPage = "" Then vCurrPage = 1
	iPerCnt 		= 10
		
	SET cTalkComm = New CShoppingTalk
	cTalkComm.FPageSize = 10
	cTalkComm.FCurrpage = vCurrPage
	cTalkComm.FRectTalkIdx = vTalkIdx
	'cTalkComm.FRectUserId = vUserID
	'cTalkComm.FRectUseYN = "y"
	cTalkComm.fnShoppingTalkCommList
%>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language="javascript">

function jsGoPage(iP){
	document.frmpage.cpg.value = iP;
	document.frmpage.submit();
}

function talkuseyn(tidx, idx, v)
{
	var m = "";
	var a = "";
	if(v == "n"){
		m = "삭제";
		a = "y";
	}else{
		m = "OPEN";
		a = "n";
	}
	if(confirm("선택하신 쇼핑톡 댓글을 "+m+"하시겠습니까?") == true) {
		$('input[name="action"]').val("update");
		$('input[name="talkidx"]').val(tidx);
		$('input[name="idx"]').val(idx);
		$('input[name="useyn"]').val(v);
		frm1.submit();
     }else{
		$('input[name="action"]').val("");
		$('input[name="talkidx"]').val("");
		$('input[name="idx"]').val("");
		$('input[name="useyn"]').val("");
		$("select[name=useyn"+idx+"]").val(a);
		return;
     }
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="20">
		검색결과 : <b><%= cTalkComm.FTotalCount %></b>
	</td>
</tr>
<% If (cTalkComm.FResultCount < 1) Then %>
<%
	Else
		For i = 0 To cTalkComm.FResultCount-1
%>
		<tr>
			<td bgcolor="#FFFFFF">
				<%=cTalkComm.FItemList(i).FUserID%> / <%=cTalkComm.FItemList(i).FRegdate%>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<select name="useyn<%=cTalkComm.FItemList(i).FIdx%>" class="select" onChange="talkuseyn('<%=vTalkIdx%>','<%=cTalkComm.FItemList(i).FIdx%>',this.value);">
				<option value="y" <%=CHKIIF(cTalkComm.FItemList(i).FUseYN="y","selected","")%>>OPEN</option>
				<option value="n" <%=CHKIIF(cTalkComm.FItemList(i).FUseYN="n","selected","")%>>삭제처리</option>
				</select>
				<br>
				<%=cTalkComm.FItemList(i).FContents%>
			</td>
		</tr>
		<% Next %>
	<%
	iTotalPage 	=  int((cTalkComm.FTotalCount-1)/10) +1
	iStartPage = (Int((vCurrPage-1)/iPerCnt)*iPerCnt) + 1
	
	If (vCurrPage mod iPerCnt) = 0 Then
		iEndPage = vCurrPage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	
	<form name="frmpage" method="get" action="<%=CurrURL()%>">
	<input type="hidden" name="cpg" value="">
	<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(vCurrPage) then
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
    </form>
<% End If %>
</table>
<form name="frm1" action="proc.asp" target="prociframe">
<input type="hidden" name="gubun" value="comment">
<input type="hidden" name="action" value="">
<input type="hidden" name="talkidx" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="useyn" value="">
</form>
<iframe name="prociframe" id="prociframe" src="" width="0" height="0"></iframe>
<% SET cTalkComm = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->