<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� CS���� �� Ŭ����(~��)
' History : 2007.08.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->


<%
response.write "�������"
response.end

dim yyyy,graphyyyy
	yyyy = request("yyyy")
		if (yyyy="") then yyyy = Cstr(Year(now()))
graphyyyy = request("graphyyyy")


session("yyyy") = yyyy

dim omonthcsclaim , i
	set omonthcsclaim = new Cchulgoitemlist
	omonthcsclaim.frectyyyy = yyyy
	omonthcsclaim.fmonthcsclaim()

dim omonthcssangdam 
	set omonthcssangdam = new Cchulgoitemlist
	omonthcssangdam.frectyyyy = yyyy
	omonthcssangdam.fmonthcssangdam()	
%>

<script language="javascript">

<!--cs������Ŭ������踦���� �˾���ũ ����-->
function jsyyyy(yyyy)
{
var popup = window.open('/admin/chulgo/monthcsclaim_detail.asp?yyyy='+yyyy,'jsyyyy','width=1024,height=768,scrollbars=yes,resizable=yes');
popup.focus();
}
<!--cs������Ŭ������踦���� �˾���ũ ��-->

function submit()
{
document.frm.submit();
}
function chk(){
	document.frm.graphyyyy.value = document.graph.graphyyyy.value;
	document.frm.submit();
}
</script>
<script language="javascript" src="/admin/chulgo/daumchart/FusionCharts.js"></script>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>����CS���� �� Ŭ����</strong></font>
			</td>
			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>

</table>
<!--ǥ ��峡-->

<!-- ǥ �˻��κ� ����-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	
	<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	
		<tr bgcolor="#FFFFFF" valign="top">
	        <td background="/images/tbl_blue_round_04.gif" width="1%" bgcolor="F4F4F4"></td>
	        <td width="54%" bgcolor="F4F4F4"> 
	       		��: &nbsp;<% DrawYBox yyyy %>
	        	<input type="submit" value="�˻�">
	        </td>
	        <td valign="top" align="right" width="40%" bgcolor="F4F4F4">
	      	</td>
	        <td background="/images/tbl_blue_round_05.gif" bgcolor="F4F4F4" width="1%"></td>
	    </tr>
    </form>
    
	<!--<form name="graph" method="get" action="/admin/chulgo/SmartChart_line2/SmartChart_line2(Beta).asp">
	<tr><td><input type="text" name="graphyyyy" value="<%=yyyy%>"></td></tr>
	</form>-->
</table>
<!-- ǥ �˻��κ� ��-->		

<!-- ���� �����Ǽ� ����-->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	
<% if omonthcsclaim.ftotalcount > 0 then %>
	<tr>
		<td bgcolor="ffffff" colspan=9>
		���� �����Ǽ�
		</td>
	</tr>
	<tr bgcolor=#DDDDFF>
		<td align="center">�� | ����</td>
		<td align="center">�±�ȯ���</td>
		<td align="center">������߼�</td>
		<td align="center">���񽺹߼�</td>
		<td align="center">��ǰ</td>
		<td align="center">ȸ��</td>
		<td align="center">�±�ȯȸ��</td>
		<td align="center">�ֹ����</td>
		<td align="center">�հ�</td>
	</tr>
	<% dim fitemtotal %>
	<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
		<tr bgcolor=#FFFFFF>
			<td align="center"><a href="javascript:jsyyyy('<%= omonthcsclaim.flist(i).fyyyy %>')"><%= omonthcsclaim.flist(i).fyyyy %></a></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd0 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd1 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd2 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd3 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd4 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd5 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd6 %></td>
			<td align="center"><% fitemtotal = omonthcsclaim.flist(i).fitemd0+omonthcsclaim.flist(i).fitemd1+omonthcsclaim.flist(i).fitemd2+omonthcsclaim.flist(i).fitemd3+omonthcsclaim.flist(i).fitemd4+omonthcsclaim.flist(i).fitemd5+omonthcsclaim.flist(i).fitemd6 %>
			<%= fitemtotal %></td>
		</tr>
	<% next %>
<!-- ���� �����Ǽ� ��-->	
	</table>
	
	<!--�׷��� ��� ����-->
	<br>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#FFFFFF>
		<!--<td align="center">
			<embed
			src="/admin/chulgo/SmartChart_line2/SmartChart_line2(Beta).swf?page_name=/admin/chulgo/SmartChart_line2/SmartChart_line2(Beta).asp&data_1=&data_2="
			quality="high" scale="noscale"
			bgcolor="#ffffff" width="800" height="600" name="barchart" align="middle" 
			allowScriptAccess="sameDomain" type="application/x-shockwave-flash"
			pluginspage="http://www.macromedia.com/go/getflashplayer">
			</embed> 
		</td>-->
		<td align="center">
		<div align="right"><input type="button" value="�׷�������Ʈ" onclick="javascript:window.print();"></div><br>
			<div id="chartdiv3" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/admin/chulgo/daumchart/MSCombiDY2D.swf", "chartdiv3", "800", "600", "0", "0");
				chart.setDataURL("/admin/chulgo/daumchart/MSCombiDY2D.asp");
				chart.render("chartdiv3");
			</script>
		</td>
	</tr>
	</table><br>
	<!--�׷��� ��� ��-->
	
<% else %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr align="center" bgcolor="#DDDDFF">
	<td align=center bgcolor="#FFFFFF"><%= yyyy %>�� �������� �˻� ����� �����ϴ�.</td>
	</tr>
	</table>
<% end if %>
<% if omonthcssangdam.ftotalcount > 0 then %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td bgcolor="ffffff" colspan=11>
			1:1 ���
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
		<td align="center">����</td>
		<td align="center">���</td>
		<td align="center">�ֹ�</td>
		<td align="center">��ǰ</td>
		<td align="center">���</td>
		<td align="center">���</td>
		<td align="center">ȯ��</td>
		<td align="center">��ȯ</td>
		<td align="center">AS</td>
		<td align="center">�̺�Ʈ</td>
		<td align="center">��������</td>
		</tr>
	<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
		<tr bgcolor=#FFFFFF>
			<td align="center"><%= omonthcssangdam.flist(i).fyyyy %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd0 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd1 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd2 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd3 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd4 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd5 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd6 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd7 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd8 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd9 %></td>
		</tr>
	<% next %>
		</tr>
		<tr bgcolor=#DDDDFF>
		<td align="center">����</td>
		<td align="center">�ý���</td>
		<td align="center">ȸ������</td>
		<td align="center">ȸ������</td>
		<td align="center">��÷</td>
		<td align="center">��ǰ</td>
		<td align="center">�Ա�</td>
		<td align="center">��������</td>
		<td align="center">����/���ϸ���</td>
		<td align="center">�������</td>
		<td align="center">��Ÿ</td>
		</tr>
	<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
		<tr bgcolor=#FFFFFF>
			<td align="center"><%= omonthcssangdam.flist(i).fyyyy %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd10 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd11 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd12 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd13 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd14 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd15 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd16 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd17 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd18 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd20 %></td>
		</tr>
	<% next %>	
	</table>
	<!--������ <table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td bgcolor="ffffff" colspan=11>
			1:1 ����հ� �� �ֹ������
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
		<td align="center">��</td>
		<td align="center">�����հ�</td>
		<td align="center">�ֹ����</td>
		</tr>
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
			<tr bgcolor=#FFFFFF>
			<td align="center"><%= omonthcssangdam.flist(i).fyyyy %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemdtot %></td>
			<td align="center">�ֹ����</td>
			</tr>
		<% next %>	
	</table>-->
	
	<!--1:1 ��� �׷��� ���1 ����-->
	<br>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#FFFFFF>
		<td align="center">
		<div align="right"><input type="button" value="�׷�������Ʈ" onclick="javascript:window.print();"></div><br>
			<div id="chartdiv4" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/admin/chulgo/daumchart/MSCombiDY2D.swf", "chartdiv3", "800", "600", "0", "0");
				chart.setDataURL("/admin/chulgo/daumchart/MSCombiDY2D1.asp");
				chart.render("chartdiv4");
			</script>
		</td>
	</tr>
	</table>
	<!--1:1 ��� �׷��� ���1 ��-->
	<!--1:1 ��� �׷��� ���2 ����-->
	<br>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#FFFFFF>
		<td align="center">
		<div align="right"><input type="button" value="�׷�������Ʈ" onclick="javascript:window.print();"></div><br>
			<div id="chartdiv5" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/admin/chulgo/daumchart/MSCombiDY2D.swf", "chartdiv3", "800", "600", "0", "0");
				chart.setDataURL("/admin/chulgo/daumchart/MSCombiDY2D2.asp");
				chart.render("chartdiv5");
			</script>
		</td>
	</tr>
	</table>
	<!--1:1 ��� �׷��� ���2 ��-->	
	<% else %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr align="center" bgcolor="#DDDDFF">
	<td align=center bgcolor="#FFFFFF"><%= yyyy %>�� 1:1 ��� �˻� ����� �����ϴ�.</td>
	</tr>
	</table>		
<% end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->