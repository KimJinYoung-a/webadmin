<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id
id = RequestCheckvar(request("id"),10)
dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.FRectdesigner = session("ssBctID")
ojungsan.JungsanMasterList

if ojungsan.FresultCount <1 then
	dbget.close()	:	response.End
end if

dim rd_state
rd_state = ojungsan.FItemList(0).Ffinishflag
%>
<script language='javascript'>
function savestate(frm){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}
}
</script>
<br>
<!--
<table width="760" cellspacing="0" class="a">
<tr>
  <td align="right"><a href="popshowdetail.asp?menupos=<%= menupos %>&id=<%= id %>">�󼼳���&gt;&gt;</a></td>
</tr>
</table>
-->
<!--
<div class="a">[���� ������ ������ <b>��üȮ�οϷ�</b>�� �������� �����Ͻñ� �ٶ��ϴ�.]</div>
-->
<br>
<div class="a">1.��������</div>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<form name="statefrm" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="mode" value="statechange">
<input type="hidden" name="idx" value="<%= ojungsan.FItemList(0).FId %>">
    <tr >
      <td width="100" bgcolor="#DDDDFF">�귣��ID</td>
      <td bgcolor="#FFFFFF"><%= ojungsan.FItemList(0).Fdesignerid %></td>
    </tr>
    <tr >
      <td width="100" bgcolor="#DDDDFF">��������</td>
      <td bgcolor="#FFFFFF"><%= ojungsan.FItemList(0).FYYYYMM %></td>
    </tr>
    <tr >
      <td width="100" bgcolor="#DDDDFF">�������</td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="rd_state" value="1" <% if rd_state="1" then response.write "checked" %> >��üȮ�δ��
		<input type="radio" name="rd_state" value="2" <% if rd_state<>"1" then response.write "checked" %> >��üȮ�οϷ�
		<input type="button" value="����" onclick="savestate(statefrm);" <% if rd_state<>"1" then response.write "disabled" %> >
      </td>
    </tr>
    <tr>
      <td width="100" bgcolor="#DDDDFF">���ݰ�꼭������</td>
      <td bgcolor="#FFFFFF">
      	<%= ojungsan.FItemList(0).Ftaxregdate %>
      </td>
    </tr>
    <tr>
      <td width="100" bgcolor="#DDDDFF">�Ա���</td>
      <td bgcolor="#FFFFFF">
      	<%= ojungsan.FItemList(0).Fipkumdate %>
      </td>
    </tr>
</form>
</table>

<br>
<div class="a"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=upche">2.���곻��</a></div>
<table width="760" cellspacing="1" cellpadding=2 class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width=100 align=left>����</td>
	<td width=100>���ֹ��Ǽ�</td>
	<td width=100>�Һ��ڰ��Ѿ�</td>
	<td width=100>���ް��Ѿ�</td>
	<td width=70>����</td>
	<td>��Ÿ</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=upche">��ü���</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fub_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fub_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fub_totalsuplycash/ojungsan.FItemList(0).Fub_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fub_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=maeip">���Գ���</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fme_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fme_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fme_totalsuplycash/ojungsan.FItemList(0).Fme_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fme_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=witaksell">Ư���¶��γ���</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fwi_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fwi_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fwi_totalsuplycash/ojungsan.FItemList(0).Fwi_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fwi_comment) %></td>
</tr>
<!--
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">Ư�� ��������</td>
	<td><%= ojungsan.FItemList(0).Fsh_cnt %></td>
	<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsellcash,0) %></td>
	<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fsh_totalsellcash<>0 then %>
	<td><%= CLng((1-ojungsan.FItemList(0).Fsh_totalsuplycash/ojungsan.FItemList(0).Fsh_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td>?</td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fsh_comment) %></td>
</tr>
-->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">Ư�� ��Ÿ</td>
	<td align=right><%= ojungsan.FItemList(0).Fet_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fet_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fet_totalsuplycash/ojungsan.FItemList(0).Fet_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=right></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fet_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">�Ѱ�</td>
	<td></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSuplycash,0) %></td>
	<% if ojungsan.FItemList(0).GetTotalSellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).GetTotalSuplycash/ojungsan.FItemList(0).GetTotalSellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=right></td>
	<% end if %>
	<td></td>
</tr>
</table>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->