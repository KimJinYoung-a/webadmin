<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� �������
' History : 2007.12.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<%
dim dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2 
dim i ,defaultdate,defaultdate1 , olddata
	ipkumdatesucc = request("ipkumdatesucc")
	olddata = request("olddata")
	sitename = request("sitenamebox")
	accountdiv = request("accountdiv")
	bancancle = request("bancancle")
		if bancancle = "" then bancancle = "1"
	datecancle = request("datecancle") 	
	dateview1 = request("dateview1")
	if dateview1 = "" then dateview1 = "yes"
	defaultdate1 = dateadd("d",-60,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 60�������� �˻�
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = mid(defaultdate1,6,2)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = right(defaultdate1,2)	
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	session("yyyy2") = yyyy2
	session("datecancle") = datecancle
	session("bancancle") = bancancle
	session("accountdiv") = accountdiv			
	session("sitename") = sitename
	session("dateview1") = dateview1
	
	mm2 = request("mm2")
	if mm2 = "" then 
		mm2 = month(now)
	else
		if Len(mm2) = 2 then
			mm2 = request("mm2")
		else
			mm2 = "0"&request("mm2")
		end if		
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)
		
dim Omaechul_list
set Omaechul_list = new Cmaechul_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
	Omaechul_list.FRectEndDate = yyyy2 & "-" & mm2 & "-" & dd2
	Omaechul_list.frectdatecancle = datecancle
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc		
	Omaechul_list.fmaechul_list()


if olddata = "no" then 
	dim Omaechul_list_old
	set Omaechul_list_old = new Cmaechul_list
		Omaechul_list_old.FRectStartdate = (yyyy1-1) & "-" & mm1 & "-" & dd1
		Omaechul_list_old.FRectEndDate = (yyyy2-1) & "-" & mm2 & "-" & dd2
		Omaechul_list_old.frectdatecancle = datecancle
		Omaechul_list_old.frectbancancle = bancancle
		Omaechul_list_old.frectaccountdiv = accountdiv
		Omaechul_list_old.frectsitename = sitename	
		Omaechul_list_old.fmaechul_list()
end if 

''����Ʈ����
Sub Drawsitename(selectboxname, sitename)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�._selectboxname�� sub���������� ����
	dim userquery, tem_str 

	response.write "<select name='" & selectboxname & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if sitename ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">��ü</option>"								'�����̶� �ܾ ��������.

	'����� �˻� �ɼ� ���� DB���� ��������
	userquery = " select id from [db_partner].[dbo].tbl_partner" 
	userquery = userquery + " where 1=1" 
	userquery = userquery + " and id <> ''" 
	userquery = userquery + " and id is not null"
	userquery = userquery + " and userdiv= '999'" 
	userquery = userquery + " group by id"

	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("id")) then 	'�˻��� �̸��� db�� ����� �̸��� ���ؼ� �´ٸ�, //
				tem_str = " selected"								'// �˻���� ����
			end if

			response.write "<option value='" & rsget("id") & "' " & tem_str & ">" & rsget("id") & "</option>"
			tem_str = ""				'rsget�� id�� �����ϰ� �˻��� ������ ����
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
%>

<script language="javascript" src="/admin/maechul/daumchart/FusionCharts.js"></script>		<!-- �׷����� ���� �ڹٽ�ũ��Ʈ����-->
<script language="javascript">

function submit()
{
	frm.submit();
}

<!--���� ���� �󼼺��� ����-->
function monthsum(yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var monthsum = window.open('/admin/maechul/maechul_month_sum.asp?yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'monthsum','width=1024,height=768,scrollbars=yes,resizable=yes');
	monthsum.focus();
}
<!--���� ���� �󼼺��� ��-->

<!--���� ���� �󼼺��� ����-->
function weeksum(yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var weeksum = window.open('/admin/maechul/maechul_week_sum.asp?yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'weeksum','width=1024,height=768,scrollbars=yes,resizable=yes');
	weeksum.focus();
}
<!--���� ���� �󼼺��� ��-->

<!--������� ����-->
function excelprint(olddata,yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var excelprint = window.open('/admin/maechul/maechul_sum_excel.asp?olddata='+olddata+'&yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'excelprint','width=1024,height=768,scrollbars=yes,resizable=yes');
	excelprint.focus();
}
<!--���� ���  ��-->

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">

		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">			
			<input type="radio" name="dateview1" name="dateview1" value="yes"  <% if dateview1="yes" then  response.write "checked" %>>�ֹ���
			<input type="radio" name="dateview1" name="dateview1" value="no"  <% if dateview1="no" then  response.write "checked" %>>�Ա���
        	/ ��¥ <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        	<br>
        	<!--<input type=checkbox name="datecancle" value="on" <% if datecancle="on" then  response.write "checked" %>>��ҰǸ�-->
        	<input type=radio name="bancancle" value="1" <% if bancancle="1" then  response.write "checked" %>>��ǰ����
        	<input type=radio name="bancancle" value="2" <% if bancancle="2" then  response.write "checked" %>>��ǰ�Ǹ�
        	<input type=radio name="bancancle" value="3" <% if bancancle="3" then  response.write "checked" %>>��ǰ����        		
        	<input type=checkbox name="ipkumdatesucc" value="on" <% if ipkumdatesucc="on" then  response.write "checked" %>>�̰���������	
        	/ �������� <select name="accountdiv">
        		<option value="" <% if accountdiv = "" then response.write "selected" %>>��ü</option>
        		<option value="7" <% if accountdiv = "7" then response.write "selected" %>>������</option>
        		<option value="20" <% if accountdiv = "20" then response.write "selected" %>>�ǽð�</option>
        		<option value="50" <% if accountdiv = "50" then response.write "selected" %>>�ܺθ�</option>
        		<option value="80" <% if accountdiv = "80" then response.write "selected" %>>�ÿ�</option>
        		<option value="100" <% if accountdiv = "100" then response.write "selected" %>>�ſ�ī��</option>        		        		        		        		        		
        	</select>
        	����Ʈ���� <% Drawsitename "sitenamebox",sitename %>        	 
		</td>
	</tr>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�������" onclick="javascript:excelprint('<%=olddata%>','<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">			
		</td>
		<td align="right">	

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<%
dim totalsum_totalsum, totalcount_totalsum, subtotalprice_totalsum, totalbuysum_totalsum, spendScoupon_totalsum, spendMileage_totalsum
dim discountEtc_totalsum, tendeliversum_totalsum, tendeliverCount_totalsum, sunsuik_totalsum, magin_totalsum
%>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if Omaechul_list.ftotalcount > 0 then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<% if dateview1 = "yes" then %>
			<td align="center">�ֹ���</td>
		<% elseif dateview1 = "no" then %>
			<td align="center">�Ա���</td>				
		<% end if %>
		<% if datecancle <> "" then %>
			<td align="center">�����</td>			
		<% end if %>	
		<td align="center">�ѱݾ�</td>
		<td align="center">�ѰǼ�</td>
		<td align="center">�Ǳݾ�</td>
		<td align="center">���԰�</td>
		<td align="center">��������</td>
		<td align="center">���ϸ���</td>		
		<td align="center">��Ÿ����</td>
			
		<td align="center">�ٹ�ۺ�</td>
		<td align="center">�ٹ�ۼ�</td>
		<td align="center">�������</td>
			
		<td align="center">����</td>
    </tr>
	<% for i = 0 to Omaechul_list.ftotalcount -1 %>
    <tr align="center" bgcolor="#FFFFFF">
		<td align="center">
		<% if right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "�����" then %>
			<font color="blue"><%= Omaechul_list.flist(i).forderdate %></font>
		<% elseif right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "�Ͽ���" then %>
			<font color="red"><%= Omaechul_list.flist(i).forderdate %></font>				
		<% else %>
			<%= Omaechul_list.flist(i).forderdate %>			
		<% end if %>
		</td>
		<% if datecancle <> "" then %>
			<td align="center"><%= Omaechul_list.flist(i).fcanceldate %></td>			
		<% end if %>				
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalsum) %></td>
		<td align="center"><%= Omaechul_list.flist(i).ftotalcount %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsubtotalprice) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalbuysum) %></td>			
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendScoupon) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendMileage) %></td>		
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdiscountEtc) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftendeliversum) %></td>
		<td align="center"><%= Omaechul_list.flist(i).ftendeliverCount %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsunsuik) %></td>
		<td align="center"><%= Omaechul_list.flist(i).fmagin %>%</td>
    </tr>   
	<% totalsum_totalsum = totalsum_totalsum + Omaechul_list.flist(i).ftotalsum %>
	<% totalcount_totalsum = totalcount_totalsum + Omaechul_list.flist(i).ftotalcount %>
	<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list.flist(i).fsubtotalprice %>
	<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list.flist(i).ftotalbuysum %>
	<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list.flist(i).fspendScoupon %>
	<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list.flist(i).fspendMileage %>			
	<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list.flist(i).fdiscountEtc %>
	<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list.flist(i).ftendeliversum %>
	<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list.flist(i).ftendeliverCount %>
	<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list.flist(i).fsunsuik %>
	<% next %>
	
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
		�� �հ�		
		</td>
		<td align="right">
			<%= CurrFormat(totalsum_totalsum) %>		
		</td>
		<td align="center">
			<%= totalcount_totalsum %>		
		</td>      		    
		<td align="right">
			<%= CurrFormat(subtotalprice_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(totalbuysum_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(spendScoupon_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(spendMileage_totalsum) %>		
		</td>   
		<td align="right">
			<%= CurrFormat(discountEtc_totalsum) %>		
		</td>	
		<td align="right">
			<%= CurrFormat(tendeliversum_totalsum) %>		
		</td>	
		<td align="center">
			<%= CurrFormat(tendeliverCount_totalsum) %>		
		</td>	
		<td align="right">
			<%= CurrFormat(sunsuik_totalsum) %>		
		</td>
		<td align="center">
			<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>
			<%= round(magin_totalsum,2) %>%		
		</td>
		<%
		totalsum_totalsum = 0
		totalcount_totalsum = 0
		subtotalprice_totalsum = 0
		totalbuysum_totalsum = 0
		spendScoupon_totalsum = 0
		spendMileage_totalsum = 0
		discountEtc_totalsum = 0
		tendeliversum_totalsum = 0
		tendeliverCount_totalsum = 0
		sunsuik_totalsum = 0
		magin_totalsum = 0
		%>
	</tr>
	<tr bgcolor="#DDDDFF">
		<td colspan="15">
			&nbsp;&nbsp;&nbsp;���⵵ �� ���� ǥ��
			<input type=checkbox name="olddata" value="no" onclick=
			"submit();"<% if olddata="no" then  response.write "checked" %>>
		</td>       
	</tr>
		
	<% if olddata = "no" then %>
		<% if Omaechul_list_old.ftotalcount > 0 then %>					
			<% for i = 0 to Omaechul_list_old.ftotalcount -1 %>
			<tr bgcolor="#FFFFFF">
				<td align="right">
				<% if right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "�����" then %>
					<font color="blue"><%= Omaechul_list_old.flist(i).forderdate %></font>
				<% elseif right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "�Ͽ���" then %>
					<font color="red"><%= Omaechul_list_old.flist(i).forderdate %></font>				
				<% else %>
					<%= Omaechul_list_old.flist(i).forderdate %>			
				<% end if %>
				</td>
				<% if datecancle <> "" then %>
					<td align="center"><%= Omaechul_list_old.flist(i).fcanceldate %></td>			
				<% end if %>				
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalsum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftotalcount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsubtotalprice) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalbuysum) %></td>			
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendScoupon) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendMileage) %></td>		
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fdiscountEtc) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftendeliversum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftendeliverCount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsunsuik) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).fmagin %>%</td>
			</tr>
			<% totalsum_totalsum = totalsum_totalsum + Omaechul_list_old.flist(i).ftotalsum %>
			<% totalcount_totalsum = totalcount_totalsum + Omaechul_list_old.flist(i).ftotalcount %>
			<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list_old.flist(i).fsubtotalprice %>
			<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list_old.flist(i).ftotalbuysum %>
			<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list_old.flist(i).fspendScoupon %>
			<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list_old.flist(i).fspendMileage %>			
			<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list_old.flist(i).fdiscountEtc %>
			<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list_old.flist(i).ftendeliversum %>
			<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list_old.flist(i).ftendeliverCount %>
			<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list_old.flist(i).fsunsuik %>
			<% next %>
			<tr bgcolor="#F4F4F4">
				<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
				�� �հ�		
				</td>
				<td align="right">
					<%= CurrFormat(totalsum_totalsum) %>		
				</td>
				<td align="center">
					<%= totalcount_totalsum %>		
				</td>      		    
				<td align="right">
					<%= CurrFormat(subtotalprice_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(totalbuysum_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(spendScoupon_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(spendMileage_totalsum) %>		
				</td>   
				<td align="right">
					<%= CurrFormat(discountEtc_totalsum) %>		
				</td>	
				<td align="right">
					<%= CurrFormat(tendeliversum_totalsum) %>		
				</td>	
				<td align="center">
					<%= CurrFormat(tendeliverCount_totalsum) %>		
				</td>	
				<td align="right">
					<%= CurrFormat(sunsuik_totalsum) %>		
				</td>
				<td align="center">
					<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>				
					<%= round(magin_totalsum,2) %>%		
				</td>
			</tr>			

		<% else %>
			<tr align="center" bgcolor="#DDDDFF">
		    	<td align=center bgcolor="#FFFFFF" colspan="15">���⵵ �˻� ����� �����ϴ�.</td>
		    </tr>		
		<% end if %>
	<% end if %>
		

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>

</form>	
</table>

<br>
<!-- �׷��� ����-->	
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor=FFFFFF>
		<td align="left">
			&nbsp;&nbsp;<a href="javascript:monthsum('<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
			���� ������� �󼼳��� ���� [Ŭ��]</a>
		</td>
		<td align="right">
			<a href="javascript:weeksum('<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
			�ֺ� ������� �󼼳��� ���� [Ŭ��]</a>&nbsp;&nbsp;		
		</td>		
	</tr>	
	<tr bgcolor=FFFFFF>
		<td align="center" colspan="2">	
			<div align="center>"><br><font size="3"><%= yyyy2 %>�� ���� ���</font></div>
			<br><div id="chartdiv3" align="center"></div>
			<script type="text/javascript">	
			var chart = new FusionCharts("/admin/maechul/daumchart/MSCombiDY2D.swf", "chartdiv3", "640", "480", "0", "0");
			chart.setDataURL("/admin/maechul/daumchart/MSCombiDY2D.asp");
			chart.render("chartdiv3");
			</script>
		</td>
	</tr>
</table>
<!-- �׷��� ��-->
	

<%
	set Omaechul_list = nothing
%>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
