<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �Աݳ���
' History : ������ ����
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%
dim jungsanidx
	jungsanidx 		= requestCheckVar(Request("jungsanidx"),10)

dim oipkum
set oipkum = new IpkumChecklist
	oipkum.FCurrpage=1
	oipkum.FPagesize=100
	oipkum.FScrollCount = 10

	oipkum.FOrderby = "desc"

	oipkum.FRectJungsanIDX = jungsanidx

	oipkum.GetMatchedIpkumlistAccounts

dim i
dim totmatchprice

totmatchprice = 0

%>

<script language='javascript'>

function SubmitSearch(frm) {

	document.frm.submit();
}

function SubmitDelete(frm) {

	if (confirm("������ �����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="jungsanidx" value="<%= jungsanidx %>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" height="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">

		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="SubmitSearch(frm)">
		</td>
	</tr>

	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%= oipkum.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td>IDX</td>
	<td width="70">�����</td>
	<td width="100">���¹�ȣ</td>
	<td width="70">�������</td>
	<td>����</td>
  	<td width="80">�Աݱݾ�</td>
  	<td width="80">��ݱݾ�</td>
  	<td width="80">��Ī�ݾ�</td>
  	<td>���</td>
</tr>
<% if oipkum.FResultCount > 0 then %>
<% for i=0 to oipkum.FResultCount-1 %>
	<% totmatchprice = totmatchprice + oipkum.Fipkumitem(i).Fmatchprice %>
<form name="frmmatch<%= i %>" method="post" action="pop_ipkum_search_process.asp">
<input type="hidden" name="mode" value="delmatch">
<input type="hidden" name="jungsanidx" value="<%= jungsanidx %>">
<input type="hidden" name="inoutidx" value="<%= oipkum.Fipkumitem(i).Finoutidx %>">
<input type="hidden" name="matchdetailidx" value="<%= oipkum.Fipkumitem(i).Fmatchdetailidx %>">
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td><%= oipkum.Fipkumitem(i).Finoutidx %></td>
	<td>
		<%= oipkum.Fipkumitem(i).Fbkname %>
	</td>
	<td>
		<%= oipkum.Fipkumitem(i).Fbkacctno %>
	</td>
	<td>
		<%= mid(oipkum.Fipkumitem(i).Fbkdate,1,4) %>-<%= mid(oipkum.Fipkumitem(i).Fbkdate,5,2) %>-<%= mid(oipkum.Fipkumitem(i).Fbkdate,7,2) %>
	</td>
	<td>
		<%= oipkum.Fipkumitem(i).Fbkjukyo %>
	</td>
  	<td>
		<% if oipkum.Fipkumitem(i).finout_gubun = "2" then %>
			<%= FormatNumber(oipkum.Fipkumitem(i).Fbkinput,0) %>
		<% end if %>
  	</td>
  	<td>
		<% if oipkum.Fipkumitem(i).finout_gubun = "1" then %>
			<%= FormatNumber(oipkum.Fipkumitem(i).Fbkinput,0) %>
		<% end if %>
  	</td>
  	<td>
  		<%= FormatNumber(oipkum.Fipkumitem(i).Fmatchprice,0) %>
  	</td>
	<td>
		<input type="button" class="button_s" value="�����ϱ�" onClick="SubmitDelete(frmmatch<%= i %>)">
	</td>
</tr>
</form>
<% next %>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td>�Ѿ�</td>
	<td colspan="6"></td>
	<td>
		<%= FormatNumber(totmatchprice, 0) %>
	</td>
  	<td></td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>




<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->