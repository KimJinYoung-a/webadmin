<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/csdailyreportcls.asp" -->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2,i

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

dim temp

if (yyyy1="") then
	
	temp=dateadd("y",now(),-7)
	temp=split(temp,"-")
	
	yyyy1=temp(0)
	mm1=temp(1)
	dd1=temp(2)

end if

if (yyyy2="") then
	
	temp=dateadd("y",now(),0)
	temp=split(temp,"-")
	
	yyyy2=temp(0)
	mm2=temp(1)
	dd2=temp(2)
	
end if





dim qna
set qna = new CsTotal
qna.yyyy1=yyyy1
qna.mm1=mm1
qna.dd1=dd1
qna.yyyy2=yyyy2
qna.mm2=mm2
qna.dd2=dd2
qna.GetCsTotal

%>
<table width="100%" border="0" cellpadding="0" cellspacing="0"class="a">
	<tr>
		<td align=left>접수건수 :	<img src="/images/dot1.gif" height="4" width="20"> 지연건수 : <img src="/images/dot2.gif" height="4" width="20">  평균 답변시간 : <img src="/images/dot4.gif" height="4" width="20"></td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
		<input type="hidden" name="showtype" value="showtype">
		<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>
<table width="100%" height=50 border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td width="10%" align=center>날짜</td>
	<!--<td width="20%" align=center>총접수건수</td>-->
	<!--<td width="20%" align=center>총처리건수</td>-->
	<td width="90%" align=center>내용</td>
	<!--<td width="20%" align=center>평균답변시간</td>-->
</tr>
<% if qna.FTotalCount < 1  then %>
<% else %>
<% For i=0 to qna.FTotalCount-1 %>

<tr bgcolor="#FFFFFF">
	<td width="7%" align=center><%= qna.Items(i).Fday %></td>
	<td>
		<table width="90%" height=50 border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align=left>
					<img src="/images/dot1.gif" height="4" width="<%= qna.Items(i).FRegcnt/qna.maxregcnt*90 %>%"><%= qna.Items(i).FRegcnt %><br>
					<img src="/images/dot2.gif" height="4" width="<%= qna.Items(i).FDelaycnt/qna.maxregcnt*90 %>%"><%= qna.Items(i).FDelaycnt%><br>
				<div align="left">
					<img src="/images/dot4.gif" height="4" width="<%= left(qna.Items(i).FAvgtime,5)/40*90 %>%"><%= left(qna.Items(i).FAvgtime,5)%></div>
				</td>
			</tr>
		</table>
	</td>
	
<% next %>
<% end if %>
</tr>
<tr bgcolor=#FFFFFF>
	<td width="7%" align=center>합&nbsp;&nbsp;계</td>
	<td align=center width="100%">
	<table  width="100%" border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
			<td width="25%" align=center>총 접수 : <%= qna.regtotal %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="25%" align=center>처리 건수 합계 : <%= qna.fintotal %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="25%" align=center>지연 건수 합계 : <%= qna.delaytotal %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="25%" align=center>평균 답변 시간: <%= left(qna.avgtotal,5) %></td>
		</tr>
	</table>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->