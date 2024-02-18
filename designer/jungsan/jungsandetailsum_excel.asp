<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id,gubun
id = request("id")
gubun = request("gubun")

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.FRectgubun = gubun
ojungsan.FRectdesigner = session("ssBctID")
ojungsan.JungsanMasterList

if ojungsan.FresultCount <1 then
	dbget.close()	:	response.End
end if

dim gubunstr
if (gubun = "upche") then
	gubunstr = "��ü���"
elseif (gubun = "maeip") then
	gubunstr = "����"
elseif (gubun = "witaksell") then
	gubunstr = "Ư��"
elseif (gubun = "witakchulgo") then
	gubunstr = "��Ÿ���"
end if


%>
<!-- �������Ϸ� ���� ��� �κ� -->
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=" & "�¶��� " & ojungsan.FItemList(0).Ftitle & " " & gubunstr & ".xls"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
/* ���� �ٿ�ε�� ����� ���ڷ� ǥ�õ� ��� ���� */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">����</td>
		<td width="100">�ѰǼ�</td>
		<td width="100">�Һ��ڰ��Ѿ�</td>
		<td width="100">���ް��Ѿ�</td>
		<td width="70">��ո���</td>
		<% if gubun="maeip" then %>
		<td colspan=4>���</td>
		<% else %>
		<td colspan=6>���</td>
		<% end if %>
	</tr>
	<% if gubun="upche" then %>
	<tr bgcolor="#CCCCFF">
		<td>��ü���</td>
		<td align=right><%= ojungsan.FItemList(0).Fub_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fub_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fub_totalsuplycash/ojungsan.FItemList(0).Fub_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td colspan=6><%= nl2br(ojungsan.FItemList(0).Fub_comment) %></td>
	</tr>
	<% end if %>
	<% if gubun="maeip" then %>
	<tr bgcolor="#CCCCFF">
		<td>����</td>
		<td align=right><%= ojungsan.FItemList(0).Fme_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fme_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fme_totalsuplycash/ojungsan.FItemList(0).Fme_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td colspan=4><%= nl2br(ojungsan.FItemList(0).Fme_comment) %></td>
	</tr>
	<% end if %>
	<% if gubun="witaksell" then %>
	<tr bgcolor="#CCCCFF">
		<td>Ư��</td>
		<td align=right><%= ojungsan.FItemList(0).Fwi_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fwi_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fwi_totalsuplycash/ojungsan.FItemList(0).Fwi_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td colspan=6><%= nl2br(ojungsan.FItemList(0).Fwi_comment) %></td>
	</tr>
	<% end if %>
	<!--
	<tr bgcolor="#FFFFFF">
		<td>Ư�� ��������</td>
		<td><%= ojungsan.FItemList(0).Fsh_cnt %></td>
		<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsellcash,0) %></td>
		<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fsh_totalsellcash<>0 then %>
		<td><%= CLng((1-ojungsan.FItemList(0).Fsh_totalsuplycash/ojungsan.FItemList(0).Fsh_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td>?</td>
		<% end if %>
		<td><%= nl2br(ojungsan.FItemList(0).Fsh_comment) %></td>
		<td align="center"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
	</tr>
	-->
	<% if gubun="witakchulgo" then %>
	<tr bgcolor="#CCCCFF">
		<td>��Ÿ���</td>
		<td align=right><%= ojungsan.FItemList(0).Fet_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fet_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fet_totalsuplycash/ojungsan.FItemList(0).Fet_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=right></td>
		<% end if %>
		<td colspan=6><%= nl2br(ojungsan.FItemList(0).Fet_comment) %></td>
	</tr>
	<% end if %>
</table>


<p>

<%
set ojungsan = Nothing


dim ojungsansummary
set ojungsansummary = new CUpcheJungsan
ojungsansummary.FRectId = id
ojungsansummary.FRectgubun = gubun
ojungsansummary.FRectdesigner = session("ssBctID")

'' 1357 ���������� �������� �ٸ�(����������)
if (id>1357) and (gubun<>"") then
    ojungsansummary.JungsanDetailListSum
end if
%>

<!-- �����ۺ� �հ� ����Ʈ ����-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
		<% if gubun="maeip" then %>
		<td colspan="9" align="left">
		<% else %>
		<td colspan="11" align="left">
		<% end if %>

			<b>��ǰ(������)�� �հ踮��Ʈ</b>
			&nbsp;&nbsp;
			<% if ojungsansummary.FRectgubun="maeip" then %>
			â���԰�Ȯ���� �������� ��ϵ˴ϴ�.
			<% else %>
			����� �������� ��ϵ˴ϴ�.
			<% end if %>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">��ǰ�ڵ�</td>
		<td colspan=3>��ǰ��</td>
		<% if gubun="maeip" then %>
		<td>�ɼǸ�</td>
		<% else %>
		<td colspan=3>�ɼǸ�</td>
		<% end if %>
		<td width="40">����</td>
		<td width="70">�ǸŰ�</td>
		<td width="70">���ް�</td>
		<td width="80">���ް��հ�</td>
    </tr>
<% if ojungsansummary.FResultCount>0 and ojungsansummary.FRectgubun<>"" then %>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsansummary.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsansummary.FItemList(i).Fsuplycash * ojungsansummary.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

    %>
    <tr bgcolor="#FFFFFF" align="center">
      <td class="txt"><%= ojungsansummary.FItemList(i).FItemID %></td>
      <td align="left" class="txt" colspan=3><%= ojungsansummary.FItemList(i).FItemName %></td>
		<% if gubun="maeip" then %>
		<td class="txt"><%= ojungsansummary.FItemList(i).FItemOptionName %></td>
		<% else %>
		<td class="txt" colspan=3><%= ojungsansummary.FItemList(i).FItemOptionName %></td>
		<% end if %>
      <td><%= ojungsansummary.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsansummary.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsansummary.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsansummary.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsansummary.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td align="center">�հ�</td>
      <td colspan="7"></td>
		<% if gubun="maeip" then %>
		<% else %>
		<td colspan="2"></td>
		<% end if %>
      <td align="right"><font color="<%= MinusFont(suplytotalsum) %>"><%= FormatNumber(suplytotalsum,0) %></font></td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<% if gubun="maeip" then %>
    	<td colspan="9" align="center">&nbsp;�˻������� �����ϴ�.</td>
    	<% else %>
    	<td colspan="11" align="center">&nbsp;�˻������� �����ϴ�.</td>
    	<% end if %>
    </tr>
<% end if %>
</table>
<!-- �����ۺ� �հ� ����Ʈ ��-->
<p>


<%
set ojungsansummary = Nothing


dim i, suplysum, suplytotalsum, duplicated
dim sumttl1, sumttl2
sumttl1 = 0
sumttl2 = 0

dim ojungsandetail
set ojungsandetail = new CUpcheJungsan
ojungsandetail.FRectId = id
ojungsandetail.FRectgubun = gubun
ojungsandetail.FRectdesigner = session("ssBctID")
ojungsandetail.FRectOrder = "orderserial"


'' 1357 ���������� �������� �ٸ�(����������)
if (id>1357) and (gubun<>"")   then
    ojungsandetail.JungsanDetailList
end if
%>
<!-- �ֹ��Ǻ� ����Ʈ ����-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
		<% if gubun="maeip" then %>
		<td colspan="9" align="left">
		<% else %>
		<td colspan="11" align="left">
		<% end if %>

			<b>�ֹ�/���/�԰�Ǻ� �󼼸���Ʈ</b>
			&nbsp;&nbsp;
			<% if ojungsandetail.FRectgubun="maeip" then %>
			â���԰�Ȯ���� �������� ��ϵ˴ϴ�.
			<% else %>
			����� �������� ��ϵ˴ϴ�.
			<% end if %>

			<% if ojungsandetail.FResultCount>=5000 then %>
			(�ִ� <%= ojungsandetail.FResultCount %> �� ǥ��)
			<% end if %>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <% if ojungsandetail.FRectgubun="maeip" then %>
      <td width="70">�԰��ڵ�</td>
      <% elseif ojungsandetail.FRectgubun="witakchulgo" then %>
      <td width="70">����ڵ�</td>
      <% else %>
      <td width="70">�ֹ���ȣ</td>
      <% end if %>

      <% if (ojungsandetail.FRectgubun<>"maeip") and (ojungsandetail.FRectgubun<>"witakchulgo") then %>
      <td width="45">������</td>
      <td width="45">������</td>
      <% elseif (ojungsandetail.FRectgubun="witakchulgo") then %>
      <td width="45"></td>
      <td width="45"></td>
      <% end if %>
      <td colspan=2>��ǰ��</td>
      <td>�ɼǸ�</td>
      <td width="35">����</td>
      <td width="50">�ǸŰ�</td>
      <td width="50">���ް�</td>
      <td width="65">���ް���</td>

      <% if ojungsandetail.FRectgubun="maeip" then %>
      <td width="65">�԰���</td>
      <% else %>
      <td width="65">�����</td>
      <% end if %>
    </tr>
<% if ojungsandetail.FResultCount>0 and ojungsandetail.FRectgubun<>"" then %>
    <% for i=0 to ojungsandetail.FResultCount-1 %>

    <%
	sumttl1 = sumttl1 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsellcash
	sumttl2 = sumttl2 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td class="txt"><%= ojungsandetail.FItemList(i).Fmastercode %></td>
      <% if ojungsandetail.FRectgubun<>"maeip" and ojungsandetail.FRectgubun<>"witakchulgo" then %>
      <td><%= ojungsandetail.FItemList(i).FBuyname %></td>
      <td><%= ojungsandetail.FItemList(i).FReqname %></td>
      <% elseif (ojungsandetail.FRectgubun="witakchulgo") then %>
      <td><%= ojungsandetail.FItemList(i).FBuyname %></td>
      <td><%= ojungsandetail.FItemList(i).FReqname %></td>
      <% end if %>
      <td align="left" class="txt" colspan=2><%= ojungsandetail.FItemList(i).FItemName %></td>
      <td class="txt"><%= ojungsandetail.FItemList(i).FItemOptionName %></td>
      <td><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo) %>"><%= ojungsandetail.FItemList(i).FItemNo %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>

      <td><%= ojungsandetail.FItemList(i).FExecDate %></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF" align="center">
    	<td>�հ�</td>
    	<% if gubun="maeip" then %>
    	<td colspan="6"></td>
    	<% else %>
    	<td colspan="8"></td>
    	<% end if %>
    	<td align="right"><font color="<%= MinusFont(sumttl2) %>"><%= formatNumber(sumttl2,0) %></font></td>
    	<td></td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<% if gubun="maeip" then %>
    	<td colspan="9" align="center">&nbsp;�˻������� �����ϴ�.</td>
    	<% else %>
    	<td colspan="11" align="center">&nbsp;�˻������� �����ϴ�.</td>
    	<% end if %>
    </tr>
<% end if %>
</table>
<!-- �ֹ��Ǻ� ����Ʈ ��-->

<%
set ojungsandetail = Nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
