<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim yyyy1, mm1
Dim makerID, onlyChkUpcheOnly
dim i

yyyy1	= req("yyyy1", Left(DateAdd("m", -1, Now()),4))
mm1		= req("mm1", Mid(DateAdd("m", -1, Now()),6,2))
makerID = req("makerID", "")
onlyChkUpcheOnly = req("onlyChkUpcheOnly", "")


Dim strSql
strSql = " db_datamart.dbo.usp_Ten_DeliveryDelayList_List ('" & yyyy1 & "-" & mm1 & "', '" & makerID & "', '" & onlyChkUpcheOnly & "')"

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

Dim rs
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If
db3_rsget.close

class CBrandServiceItem
	public Fyyyymm
	public Fmakerid
	public FbaljuCnt
	public FstockoutCnt
	public FdelayCnt
	public FbaditemCnt
	public FerrdeliveryCnt
	public FchulgoCnt
	public FchulgoNDaySum
	public FrealOverNDaySum
	public FfalsehoodSongjangCnt

	public function GetSUM
		GetSUM = (FstockoutCnt + FdelayCnt + FbaditemCnt + FerrdeliveryCnt)
	end function

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

function toClass(rs, i)
	dim result
	'// yyyymm, makerid, baljuCnt, stockoutCnt, delayCnt, baditemCnt, errdeliveryCnt, regdate, lastupdate
	'// , chulgoCnt, chulgoNDaySum, realOverNDaySum, falsehoodSongjangCnt
	set result = new CBrandServiceItem
	result.Fyyyymm 			= rs(0,i)
	result.Fmakerid 		= rs(1,i)
	result.FbaljuCnt 		= rs(2,i)
	result.FstockoutCnt 	= rs(3,i)
	result.FdelayCnt 		= rs(4,i)
	result.FbaditemCnt 		= rs(5,i)
	result.FerrdeliveryCnt 	= rs(6,i)
	result.FchulgoCnt 		= rs(7,i)
	result.FchulgoNDaySum 	= rs(8,i)
	result.FrealOverNDaySum 		= rs(9,i)
	result.FfalsehoodSongjangCnt 	= rs(10,i)

	set toClass = result
end function

dim rowCnt, item, val
dim totbaljuCnt, totstockoutCnt, totdelayCnt, totbaditemCnt, toterrdeliveryCnt

%>

<script language='javascript'>
function jsPopDashBoard(makerid) {
    var popwin = window.open("/admin/brandStatic/brandServicePointDashBoard.asp?menupos=4024&makerID=" + makerid,"jsPopDashBoard","width=1400 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
		<td align="left">
	       	��� :
			<% DrawYMBox yyyy1,mm1 %>
			&nbsp;
			�귣��ID :
			<input type="text" class="text" name="makerID" value="<%=makerID%>">
		</td>

		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="onlyChkUpcheOnly" value="Y" <%= CHKIIF(onlyChkUpcheOnly="Y", "checked", "") %>> Ȯ�δ�� �귣��(���Ҹ� ���(��ǰ) ���� 5% �̻� �Ǵ� �������� ���� 5% �̻� �ִ� �귣��)��
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

<pre>
* ���ְǼ��� �귣�庰 ��ǰ������ �Դϴ�.(�� �ֹ��� 3���� ��ǰ�� �ֹ��ϸ� 3���� ī��Ʈ�մϴ�.)
* 1���� ���ֵ� �ֹ��� ����, 2���� ��Ұ� �̷������ ����� ���� �и��˴ϴ�.
* ��չ�ۼҿ����� ��ü�뺸 ���� �ù�����ϰ� �Ǵ� �������� ����ϼ� �Դϴ�.
* ���������� ������ ���� �ù�����ϰ� �̷�����ų�, ����ϰ� �ù���������� ���̰� 3�� �̻� ���ų�, 5�ϰ� �����ȸ�� �ȵǴ� ���Դϴ�.
</pre>

<p />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="60">
			���
		</td>
		<td rowspan="2" width="250">�귣��</td>
		<td width="80" rowspan="2">�ѹ��ְǼ�<br>(��ü���)</td>
        <td colspan="6">���Ҹ� ���(��ǰ)�Ǽ�</td>
        <td colspan="4" width="80">��չ�ۼҿ���</td>
		<td rowspan="2" width="80"><b>��������</b></td>
		<td rowspan="2">���</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80">ǰ��</td>
		<td width="80">�������</td>
		<td width="80">��ǰ�ҷ�</td>
		<td width="80">�����</td>
		<td width="80">�հ�</td>
		<td width="80"><b>����<br />(���ִ��)</b></td>
		<td width="80">���Ǽ�</td>
		<td width="80">����ϱ���</td>
		<td width="80">������ȸ����</td>
		<td width="80">��������Ǽ�</td>
	</tr>
	<%
	If IsArray(rs) Then
		rowCnt = UBound(rs,2) + 1
		For i = 0 To UBound(rs,2)
			set item = toClass(rs, i)

			totbaljuCnt = totbaljuCnt + item.FbaljuCnt
			totstockoutCnt = totstockoutCnt + item.FstockoutCnt
			totdelayCnt = totdelayCnt + item.FdelayCnt
			totbaditemCnt = totbaditemCnt + item.FbaditemCnt
			toterrdeliveryCnt = toterrdeliveryCnt + item.FerrdeliveryCnt
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= item.Fyyyymm %></td>
		<td><a href="javascript:jsPopDashBoard('<%= item.Fmakerid %>')"><%= item.Fmakerid %></a></td>
		<td><%= item.FbaljuCnt %></td>
		<td><%= item.FstockoutCnt %></td>
		<td><%= item.FdelayCnt %></td>
		<td><%= item.FbaditemCnt %></td>
		<td><%= item.FerrdeliveryCnt %></td>
		<td><%= item.GetSUM %></td>
		<td>
			<%
			if item.FbaljuCnt > 0 then
				val = Round((1.0 * item.GetSUM / item.FbaljuCnt * 100), 1)
				if (val >= 5) then
					response.write "<font color='red'><b>" & val & "%</b></font>"
				else
					response.write val & "%"
				end if
			else
				response.write "-"
			end if
			%>
		</td>
		<td><%= item.FchulgoCnt %></td>
		<% if item.FchulgoCnt > 0 then %>
		<td><%= Round(1.0*item.FchulgoNDaySum/item.FchulgoCnt,1) %></td>
		<td><%= Round(1.0*(item.FchulgoNDaySum+item.FrealOverNDaySum)/item.FchulgoCnt,1) %></td>
		<td>
			<% if (item.FfalsehoodSongjangCnt > 0) then %>
			<font color="red"><b><%= item.FfalsehoodSongjangCnt %></b></font>
			<% else %>
			<%= item.FfalsehoodSongjangCnt %>
			<% end if %>
		</td>
		<% else %>
		<td>-</td>
		<td>-</td>
		<td>-</td>
		<% end if %>
		<td></td>
		<td></td>
	</tr>
	<%
		next
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td><%= totbaljuCnt %></td>
		<td><%= totstockoutCnt %></td>
		<td><%= totdelayCnt %></td>
		<td><%= totbaditemCnt %></td>
		<td><%= toterrdeliveryCnt %></td>
		<td>
			<%= (totstockoutCnt + totdelayCnt + totbaditemCnt + toterrdeliveryCnt) %>
		</td>
		<td>
			<%
			if totbaljuCnt > 0 then
				val = Round((1.0 * (totstockoutCnt + totdelayCnt + totbaditemCnt + toterrdeliveryCnt) / totbaljuCnt * 100), 1)
				if (val >= 5) then
					response.write "<font color='red'><b>" & val & "%</b></font>"
				else
					response.write val & "%"
				end if
			else
				response.write "-"
			end if
			%>
		</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<%
	end if
	%>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
