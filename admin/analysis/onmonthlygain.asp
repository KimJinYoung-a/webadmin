<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysiscls.asp"-->

<%
response.write "������ : <a href='/admin/maechul/maechul_sum.asp?menupos=1013'><font color=blue>�������v2&gt;&gt;������� ����</font></a>"
dbget.close()	:	response.End

dim yyyy1,mm1
dim yyyy2,mm2
yyyy1 = request("yyyy1")
mm1 = request("mm1")
'yyyy2 = request("yyyy2")
'mm2 = request("mm2")


dim dt
if yyyy1="" then
	'dt = dateserial(year(Now),month(now)-3,1)
	'yyyy1 = Left(CStr(dt),4)
	'mm1 = Mid(CStr(dt),6,2)

	dt = dateserial(year(Now),month(now),1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)

	dt = dateserial(year(Now),month(now)+1,1)
	yyyy2 = Left(CStr(dt),4)
	mm2 = Mid(CStr(dt),6,2)
else
	dt = CStr(dateserial(yyyy1,mm1+1,1))
	yyyy2 = Left(CStr(dt),4)
	mm2 = Mid(CStr(dt),6,2)
end if


dim nextyyyymm
'nextyyyymm = CStr(dateserial(yyyy2,mm2+1,1))
'response.write yyyy1 + "-" + mm1 + "-01"
'response.write yyyy2 + "-" + mm2 + "-01"
dim oanal
set oanal = new CAnalysis
oanal.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
oanal.FRectYYYYMMDD2 = yyyy2 + "-" + mm2 + "-01"

oanal.FBeasongPay = 2700
oanal.GetMeachulWithCopons
oanal.GetMinusMeachulSum
oanal.GetTenBeasongcount
oanal.GetMeaipSum

'oanal.getOnlineMonthlyGain

dim i

'response.write "FRectYYYYMMDD : " + oanal.FRectYYYYMMDD + "<br>"
'response.write "FRectYYYYMMDD2 : " + oanal.FRectYYYYMMDD2 + "<br>"
%>
<table width="800" border="0" cellpadding="2" cellspacing="1" class="a">
<tr>
	<td> (��ۺ�ܰ� 2,700), ī�������, ���޸������� �� ����</td>
</tr>
</table>
<table width="800" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		�˻��Ⱓ:<% DrawYMBox yyyy1,mm1 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF" align=center>
	<td width=150>�˻���</td>
	<td width=200><%= oanal.FOneItem.Fyyyymm %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>�����Ǽ�</td>
	<td><%= FormatNumber(oanal.FOneItem.FMCnt,0) %> ��</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>�ֹ��ݾ�</td>
	<td><%= FormatNumber(oanal.FOneItem.FTotalSum,0) %> ��</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>�����ݾ�</td>
	<td><%= FormatNumber(oanal.FOneItem.FSubTotalPrice,0) %> ��</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.FSubTotalPrice/oanal.FOneItem.FTotalSum*100*100)/100 %> %
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>���ܰ�</td>
	<td>
		<% if oanal.FOneItem.FMCnt<>0 then %>
		<%= FormatNumber(oanal.FOneItem.FSubTotalPrice/oanal.FOneItem.FMCnt,0) %> ��
		<% end if %>
	</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>��ǰ�Ǽ�</td>
	<td><%= FormatNumber(oanal.FOneItem.Fminuscnt,0) %> ��</td>
	<td>
		<% if oanal.FOneItem.FMCnt<>0 then %>
		<%= clng(oanal.FOneItem.Fminuscnt/oanal.FOneItem.FMCnt*100*100)/100 %> % (�����Ǽ� ���)
		<% end if %>
	</td>
</tr>
<!--
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>��ǰ�ֹ��ݾ�</td>
	<td><%= FormatNumber(oanal.FOneItem.FminusTotalSum,0) %> ��</td>
	<td>

	</td>
</tr>
-->
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>��ǰ�����ݾ�</td>
	<td><%= FormatNumber(oanal.FOneItem.FminusSubTotalPrice,0) %> ��</td>
	<td>
		<% if oanal.FOneItem.FSubTotalPrice<>0 then %>
		<%= clng(Abs(oanal.FOneItem.FminusSubTotalPrice)/Abs(oanal.FOneItem.FSubTotalPrice)*100*100)/100 %> % (������ ���)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>�������</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Ftencardspend,0) %> ��
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Ftencardspend/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (�ֹ��״��)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>���ϸ���</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Fmiletotalprice,0) %> ��
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Fmiletotalprice/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (�ֹ��״��)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>SKT�����</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Fspendmembership,0) %> ��
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Fspendmembership/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (�ֹ��״��)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>�ÿ�����</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Fallatdiscountprice,0) %> ��
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Fallatdiscountprice/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (�ֹ��״��)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>���μҰ�</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.getTotalDiscountsum,0) %> ��
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.getTotalDiscountsum/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (�ֹ��״��)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
</table>


<br>
<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>��۰Ǽ�<br>(�ٹ� ���� �Ǽ�)</td>
	<td width=200><%= Formatnumber(oanal.FOneItem.FBeasongCnt,0) %></td>
	<td>
		<% if oanal.FOneItem.FMCnt<>0 then %>
		<%= clng(oanal.FOneItem.FBeasongCnt/oanal.FOneItem.FMCnt*100*100)/100 %> %
		<% end if %>
		(�����Ǽ� ���)
	</td>
</tr>
</table>

<br>
<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>�ǰ�����</td>
	<td width=200><%= Formatnumber(oanal.FOneItem.getRealSubTotalPrice,0) %></td>
	<td>(�����ݾ� - ��ǰ�����ݾ�)</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>��ǰ���԰�</td>
	<td width=200><%= Formatnumber(oanal.FOneItem.FMeaipTotal,0) %></td>
	<td>
		<% if oanal.FOneItem.getRealSubTotalPrice<>0 then %>
		<%= clng(oanal.FOneItem.FMeaipTotal/(oanal.FOneItem.getRealSubTotalPrice)*100*100)/100 %> %  (�ǰ����� ���)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>��ۺ�</td>
	<td width=200><%= formatnumber(oanal.FOneItem.GetBeasongTotal,0) %></td>
	<td>(�ٹ����ٹ�۰Ǽ� * ��۴ܰ�(<%= formatnumber(oanal.FBeasongPay,0) %>)</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>ī�������</td>
	<td width=200></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>���޼�����</td>
	<td width=200></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>����</td>
	<td width=200><%= formatnumber(oanal.FOneItem.GetSuic,0) %></td>
	<td>
		<% if oanal.FOneItem.getRealSubTotalPrice<>0 then %>
		<%= clng(oanal.FOneItem.GetSuic/(oanal.FOneItem.getRealSubTotalPrice)*100*100)/100 %> %  (�ǰ����� ���)
		<% end if %>
		(�ǰ����״��)
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
</table>

<br><br>
<!--
<br><br>

<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF" align=center>
	<td width=60>�Ѱ����Ǽ�</td>
	<td width=70>���ֹ��ݾ�</td>
	<td width=70>�Ѱ����ݾ�<br>(A)</td>
	<td width=70>���ܰ�</td>
	<td width=70>�Ѹ��԰�<br>(B)</td>
	<td width=70>��۰Ǽ�</td>
	<td width=70>��۱ݾ�<br>(C)</td>
	<td width=70>���ϸ������</td>
	<td width=70>�������</td>
	<td width=70>����<br>(A-B-C)</td>
	<td width=70>������</td>
</tr>
<% if false then %>
<% for i=0 to oanal.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td align=center><%= FormatNumber(oanal.FItemList(i).FMCnt,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).FTotalSum,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).FSubTotalPrice,0) %></td>
	<td align=right>
	<% if oanal.FItemList(i).FMCnt<>0 then %>
	<%= FormatNumber(oanal.FItemList(i).FSubTotalPrice/oanal.FItemList(i).FMCnt,0) %>
	<% end if %>
	</td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).FMeaipTotal,0) %></td>
	<td align=center><%= FormatNumber(oanal.FItemList(i).FBeasongCnt,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).GetBeasongTotal,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).Fmiletotalprice,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).Ftencardspend,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).GetSuic,0) %></td>
	<td align=center>
	<% if oanal.FItemList(i).FSubTotalPrice<>0 then %>
		<%= FormatNumber(oanal.FItemList(i).GetSuic/oanal.FItemList(i).FSubTotalPrice*100,0) %> %
	<% end if %>
	</td>
</tr>
<% next %>
<% end if %>
</table>
-->
<%
set oanal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->