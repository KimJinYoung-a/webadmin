<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id, yyyymm,designerid
id = RequestCheckvar(request("id"),10)

dim ojungsan, ojungsanmaster
set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = id
ojungsanmaster.FrectDesigner = session("ssBctID")
ojungsanmaster.JungsanMasterList

if ojungsanmaster.FresultCount <1 then
	dbget.close()	:	response.End
end if

yyyymm = ojungsanmaster.FItemList(0).FYYYYmm
designerid = ojungsanmaster.FItemList(0).FDesignerid

set ojungsan = new CUpcheJungsan
ojungsan.FRectid = id
ojungsan.FrectDesigner = session("ssBctID")
'ojungsan.FRectgubun = "upche"
'if (id>=179504) then ''2014/02 ����
'  ojungsan.FRectgubun = "lecture"
'end if
ojungsan.JungsanDetailListSum

dim i, suplysum, suplytotalsum, duplicated

suplytotalsum = 0
%>
<table width="760" cellspacing="0" class="a">
<tr>
  <td align="right"><a href="jungsanmaster.asp?menupos=<%= menupos %>&id=<%= id %>">����Ȯ��&gt;&gt;</a></td>
</tr>
</table>
<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ ��ü��� �����ۺ� �հ� ]</td>
	<td align="right">�հ� <%= FormatNumber(ojungsanmaster.FitemList(0).Fub_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="40">��ǰID</td>
      <td width="200">��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="70">�ǸŰ�</td>
      <td width="70">���ް�</td>
      <td width="70">���ް��հ�</td>
    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
    <tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="6"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<% end if %>

<% if ojungsan.FResultCount>0 then %>
<%
ojungsan.FRectOrder = "orderserial"
ojungsan.JungsanDetailList
%>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ ��ü��� ���� ] - (<font color="#FF0000">����� ����</font>�Դϴ�. ������� �������� ��� ������ ���꿡 ���Ե˴ϴ�.)</td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="80">�ֹ���ȣ</td>
      <td width="50">������</td>
      <td width="50">������</td>
      <td width="120">�����۸�</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="70">�ǸŰ�</td>
      <td width="70">���ް�</td>
      <td width="100">�����</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td ><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><%= ojungsan.FItemList(i).FExecDate %></td>
    </tr>
    <% next %>
</table>
<% end if %>

<%
ojungsan.FRectgubun = "maeip"
ojungsan.JungsanDetailListSum
%>
<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ �����԰� ��ǰ�� �հ� ]</td>
	<td align="right">�հ� <%= FormatNumber(ojungsanmaster.FitemList(0).Fme_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1" class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="40">��ǰID</td>
      <td>��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="70">�ǸŰ�</td>
      <td width="70">���ް�</td>
      <td width="80">���ް��հ�</td>

    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td align="center"><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="6"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>

<%
ojungsan.FRectOrder = "orderserial"
ojungsan.JungsanDetailList
%>
<table border="0" width="760" class="a">
<tr>
	<td>[ �����԰� �󼼳��� ]</td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1" class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">�԰��ڵ�</td>
      <td width="80">�԰���</td>
      <td>��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="70">�ǸŰ�</td>
      <td width="70">���ް�</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td align="center"><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td align="center"><%= ojungsan.FItemList(i).FExecDate %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
    </tr>
    <% next %>
</table>
<% end if %>


<% if id>1357 then %>
<!-- Modify 20040305 -->
<%
'ojungsan.FrectDesigner = designerid
'ojungsan.FRectStartDay = yyyymm + "-" + "01"
'ojungsan.FRectEndDay   = CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))+1,1))
'ojungsan.FRectYYYYMM   = yyyymm
'ojungsan.FRectPreYYYYMM   = Left(CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))-1,1)),7)

ojungsan.GetWitakJungSanSummary
%>

<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ Ư�� �����ۺ� �հ� ] - (<font color="#FF0000">��ۿϷ��� ����</font>�Դϴ�. ��ۿϷ����� �������� ��� ������ ���꿡 ���Ե˴ϴ�.)</td>
	<td align="right">�հ� <%= FormatNumber(ojungsanmaster.FitemList(0).Fwi_totalsuplycash + ojungsanmaster.FitemList(0).Fsh_totalsuplycash + ojungsanmaster.FitemList(0).Fet_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="40">��ǰID</td>
      <td>��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="60">�¶����Ǹ�</td>
      <td width="60">��Ÿ�Ǹ�</td>
      <td width="60">�ǸŰ�</td>
      <td width="60">���ް�</td>
      <td width="80">���ް��հ�</td>
    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).FSuplycash * (ojungsan.FItemList(i).Fsellno + ojungsan.FItemList(i).Foffsellno + ojungsan.FItemList(i).FChulgoNo)
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
    <tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td align="center" ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).Fsellno %></td>
      <td align="center"><%= ojungsan.FItemList(i).FChulgoNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).FSellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).FSuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="7"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<%
ojungsan.FRectgubun = "witaksell"
ojungsan.FRectOrder = "orderserial"
ojungsan.JungsanDetailList

dim sumttl1, sumttl2
sumttl1 = 0
sumttl2 = 0
%>
<table border="0" width="760" class="a">
<tr>
	<td>[ Ư�� �¶����Ǹ� �󼼳��� ] - (<font color="#FF0000">����� ����</font>�Դϴ�. ������� �������� ��� ������ ���꿡 ���Ե˴ϴ�.)</td>
</tr>
</table>
<% if ojungsan.FResultCount>0 then %>

<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="80">�ֹ���ȣ</td>
      <td width="50">������</td>
      <td width="50">������</td>
      <td>��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="60">�ǸŰ�</td>
      <td width="60">���ް�</td>
      <td width="80">���ް���</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
	sumttl1 = sumttl1 + ojungsan.FItemList(i).FItemNo*ojungsan.FItemList(i).Fsellcash
	sumttl2 = sumttl2 + ojungsan.FItemList(i).FItemNo*ojungsan.FItemList(i).Fsuplycash
	%>
    <tr bgcolor="#FFFFFF">
      <td align="center" ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td align="center" ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td align="center" ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center" ><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsan.FItemList(i).FItemNo*ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td colspan=8></td>
    	<td align=right><%= formatNumber(sumttl2,0) %></td>
    </tr>
</table>
<% else %>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
	<tr bgcolor="#FFFFFF"><td align=center>������ �������� �ʽ��ϴ�.</td></tr>
</table>
<% end if %>

<% end if %>

<% else %>
<%
ojungsan.FrectDesigner = designerid
ojungsan.FRectStartDay = yyyymm + "-" + "01"
ojungsan.FRectEndDay   = CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))+1,1))
ojungsan.FRectYYYYMM   = yyyymm
ojungsan.FRectPreYYYYMM   = Left(CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))-1,1)),7)

ojungsan.GetWitakJungSanByItemView
%>
<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ Ư�� �����ۺ� �հ� ]</td>
	<td align="right">�հ� <%= FormatNumber(ojungsanmaster.FitemList(0).Fwi_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="40">��ǰID</td>
      <td width="200">��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">���/�Ǹŷ�</td>
      <td width="80">�ǸŰ�</td>
      <td width="80">���ް�</td>
      <td width="80">���ް��հ�</td>
    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).FSuplycash_sell * ojungsan.FItemList(i).FjungsanNo
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
    <tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FjungsanNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSellcash_sell) %>"><%= FormatNumber(ojungsan.FItemList(i).FSellcash_sell,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSuplycash_sell) %>"><%= FormatNumber(ojungsan.FItemList(i).FSuplycash_sell,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="6"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<% end if %>
<% end if %>



<%
ojungsan.FRectOrder = "orderserial"
ojungsan.FRectgubun = "witakchulgo"
ojungsan.JungsanDetailList
%>
<% if ojungsan.FResultCount>0 then %>
<table border="0" width="760" class="a">
<tr>
	<td>[ Ư�� ��Ÿ�Ǹ� �󼼳��� ] - ����, ���� �� ��Ÿ �Ǹ�</td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">�԰��ڵ�</td>
      <td width="80">�����</td>
      <td>��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="60">�ǸŰ�</td>
      <td width="60">���ް�</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td align="center"><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td align="center"><%= ojungsan.FItemList(i).FExecDate %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>

    </tr>
    <% next %>
</table>
<% end if %>
<%
set ojungsan = Nothing
set ojungsanmaster = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->