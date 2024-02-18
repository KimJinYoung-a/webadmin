<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 메일진 통계
' History : 2008.05.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim page, i, mailergubun
	page = request("page")
	mailergubun = request("mailergubun")
	
if page="" then page=1

dim omd
set omd = New CMailzine
	omd.FCurrPage = page
	omd.FPageSize=20
	omd.frectmailergubun = mailergubun
	omd.GetMailingList
%>

<link href="/report.css" rel="stylesheet" type="text/css">

<script language="javascript">

//신규등록 팝업시작
function popup()
{
	var popup = window.open('/admin/mailopen/mail_reg.asp?mode=add','popup','width=1024,height=768,scrollbars=yes,resizable=yes');
	popup.focus();
}

function popupedit(idx)
{
	var popupedit = window.open('/admin/mailopen/mail_edit.asp?idx='+idx+'&mode=edit','popupedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	popupedit.focus();
	
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 발송메일러 : <% drawmailergubun "mailergubun" , mailergubun , "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>
<!-- 표 중간바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		※ 발송후 3일이 지난후에 발송현황이 등록되어 짐니다.
		<Br>&nbsp;&nbsp;&nbsp;- EMS메일러 매일 새벽3시15분에 자동 등록
		<Br>&nbsp;&nbsp;&nbsp;- TMS메일러 매일 새벽3시에 자동 등록
    </td>
    <td align="right">
    	<!--<input type="button" value="THUNDERMAIL등록" class="button" onclick="javascript:popup();">-->
    </td>        
</tr>
</table>
<!-- 표 중간바 끝-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= omd.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= omd.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>발송 이름</td>
	<td>총대상자수</td>
	<td>발송시간</td>
	<td>완료시간</td>
	<td>발송<br>메일러</td>
	<td>비고</td> 
</tr>
<% if omd.FresultCount>0 then %>
<% for i=0 to omd.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->	 
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td><% = omd.FItemList(i).Ftitle %></td>
	<td align="right"><% = FormatNumber(omd.FItemList(i).Ftotalcnt,0) %></td>
	<td><% = omd.FItemList(i).Fstartdate %></td>
	<td><% = omd.FItemList(i).Fenddate %></td>
	<td><% = omd.FItemList(i).fmailergubun %></td>
	<td width=60>
		<input type="button" value="수정" onclick="javascript:popupedit(<% = omd.FItemList(i).Fidx %>);" class="button">
	</td>					
</tr>   
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if omd.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit(<%= omd.StartScrollPage-1 %>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + omd.StartScrollPage to omd.StartScrollPage + omd.FScrollCount - 1 %>
			<% if (i > omd.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(omd.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit(<%= i %>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if omd.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit(<%= i %>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>	
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	</td>
</tr>
</table>

<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->