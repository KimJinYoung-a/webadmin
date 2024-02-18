<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->


<%
dim ck_joinmall,ck_ipjummall,ck_pointmall
dim research
dim opt_rect

ck_joinmall = request("ck_joinmall")
ck_ipjummall = request("ck_ipjummall")
ck_pointmall = request("ck_pointmall")
opt_rect = request("opt_rect")
research = request("research")

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
	if ck_pointmall="" then ck_pointmall="on"
	if opt_rect="" then opt_rect="24"
end if

dim oreport
set oreport = new CJumunMaster
oreport.FRectJoinMallNotInclude = ck_joinmall
oreport.FRectExtMallNotInclude = ck_ipjummall
oreport.FRectPointNotInclude = ck_pointmall
oreport.FRectSearchType = opt_rect

oreport.SearchMallSellrePort_Week


dim i,p1,p2,p3,p4
dim maybe_monthcount
dim maybe_monthsum
dim dayno, currno

dim nowdate,nowyyyymm

if opt_rect="all" then
	nowdate = CStr(date)
	nowyyyymm = left(nowdate,7)
	currno = CInt(right(nowdate,2))

	nowdate = dateserial(Left(nowdate,4),Mid(nowdate,6,2)+1,0)
	dayno = CInt(right(nowdate,2))
end if

%>


<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		�˻��Ⱓ :
		<input type="radio" name="opt_rect" value="24" <% if opt_rect="24" then response.write "checked" %> >24��
		<input type="radio" name="opt_rect" value="48" <% if opt_rect="48" then response.write "checked" %> >48��
		<input type="radio" name="opt_rect" value="all" <% if opt_rect="all" then response.write "checked" %> >��ü

		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="checkbox" name="ck_joinmall" <% if ck_joinmall="on" then response.write "checked" %> >���޸� ����
		<input type="checkbox" name="ck_ipjummall" <% if ck_ipjummall="on" then response.write "checked" %> >������ ����
		<input type="checkbox" name="ck_pointmall" <% if ck_pointmall="on" then response.write "checked" %> >����Ʈ�� ����
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">�Ⱓ</font></td>
          <td class="a"><font color="#FFFFFF"></font></td>
          <td class="a" width="100"><font color="#FFFFFF">�ݾ�(��)</font></td>
          <td class="a" width="50"><font color="#FFFFFF">�Ǽ�</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then

				maybe_monthsum	 = CLng(oreport.FMasterItemList(i).Fselltotal * dayno / currno)
				maybe_monthcount = CLng(oreport.FMasterItemList(i).Fsellcnt * dayno / currno)


				if maybe_monthcount>oreport.maxc then
					oreport.maxc = maybe_monthcount
				end if

				if maybe_monthsum>oreport.maxt then
					oreport.maxt = maybe_monthsum
				end if

				p3 = Clng(maybe_monthsum/oreport.maxt*100)
				p4 = Clng(maybe_monthcount/oreport.maxc*100)
			end if

			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if
		%>
        <tr bgcolor="#FFFFFF" height="35" class="a">
		<td>
          	<%= oreport.FMasterItemList(i).Fsitename %>��
          	</td>
          	<td>
          	<% if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then %>
          		<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p3 %>%"></div><br>
          		<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p4 %>%"></div><br>
          	<% end if %>
			<div align="left"> <img src="/images/dot1.gif" height="3" width="<%= p1 %>%"></div><br>
          		<div align="left"> <img src="/images/dot2.gif" height="3" width="<%= p2 %>%"></div>
          	</td>
		<td class="a" width="100" align="right">
		  	<% if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then %>
		  	<font color="#AAAAAA"><%= FormatNumber(maybe_monthsum,0) %></font><br>
		  	<% end if %>
		    	<%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %> <br>
		</td>
		<td class="a" width="50" align="right">
		  	<% if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then %>
		  	<font color="#AAAAAA"><%= FormatNumber(maybe_monthcount,0) %></font><br>
		  	<% end if %>
          		<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>
		</td>
        </tr>
        <% next %>
</table>
<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->