<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ �Է�
' History : 2007.09.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->

<%
dim idx, itemid , i
	idx = request("idx")

dim oip
	set oip = new Cauctionlist        			'Ŭ���� ����
	oip.fauctionedit()		

%>
				
	<script language="javascript">
	function sendit(){
	if(document.frm.auction_cate_code.value==""){
	alert("����ī�װ����� �Է��ϼ���.")
	document.frm.auction_cate_code.focus();
	}
	else if(document.frm.auction_realsel.value==""){
	alert("���ǿ� ��� �Ͻ� ������ �Է��ϼ���.")
	document.frm.auction_realsel.focus();
	}
	else if(document.frm.auction_realsel.value==0){
	alert("���ǿ� ��� �Ͻ� ���� 1�� �̻� �Է��ϼ���.")
	document.frm.auction_realsel.focus();
	}
	else if(document.frm.auction_isusing.value==""){
	alert("���ǿ� ��� ���θ� �Է� �ϼ���.")
	document.frm.auction_isusing.focus();
	}
	else
	document.frm.submit();
	}
	</script>
	
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">
			<input type="button" value="���" onclick="sendit();" class="button">
		</td>
	</tr>
</form>	
</table>
<!-- �׼� �� -->
	
<!--��ǰ���̺����-->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form method="get" name="frm" action="auctionedit_submit.asp">  	
	<tr bgcolor="#FFFFFF">
		<td rowspan=5><input type="hidden" name="mode"><img src="<%= oip.flist(0).FImageList %>" width="100" height="100"></td>
		<td><font size=2>��������ȣ :</font></td>
		<td><font size=2><%= idx %><input type="hidden" name="idx" value="<%= idx %>"></font></td>
		<td><font size=2>������ �ɼ� : </font></td>
		<td><font size=2><%= oip.flist(0).ten_option %></font>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>��ǰ��ȣ :</font></td> 
		<td><font size=2><%= oip.flist(0).ten_itemid %><input type="hidden" name="ten_itemid" value="<%= oip.flist(0).ten_itemid %>"></font></td>
		<td><font size=2>����ī�װ��� :</font></td> 
		<td><font size=2><input type="text" name="auction_cate_code" value="<%= oip.flist(0).auction_cate_code %>" size="13"></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>��ǰ�� : </font></td>
		<td><font size=2><%= oip.flist(0).ten_itemname %></font></td>
		<td><font size=2>�귣�� :</font></td>
		<td><font size=2><%= oip.flist(0).ten_makerid %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>������� : </font></td>
		<td>
			<font size=2>
			<% if oip.flist(0).ten_jaego >= 10 then
			response.write "Y"
			else
			response.write "N"
			end if %></font>
		</td>
		<td><font size=2>�ٹ�������� : </font></td>
		<td><font size=2><%= oip.flist(0).ten_jaego %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><font size=2>���ǵ�Ͽ��� : </font></td>
		<td><font size=2><input type="text" name="auction_isusing" value="<%= oip.flist(0).auction_isusing %>" size="2"> ex)y,n</font></td>
		<td><!--<font size=2>���ǵ�ϼ��� : </font>--></td> 
		<td><font size=2><input type="hidden" name="auction_realsel" value="1" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=1>�̹������ : </td>
		<td colspan=5><%= oip.flist(0).FImageList %></td>	
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=7>��ǰ �� ���� : </td>	
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=7><%= nl2br(oip.flist(0).ten_itemcontent) %></td>
	</tr>
	</form>
</table>
<!--��ǰ���̺�-->
	
<!-- #include virtual="/lib/db/dbclose.asp" -->